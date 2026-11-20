import os
import unittest
from asyncio import gather, sleep
from pathlib import Path
from unittest.mock import patch
from uuid import uuid4

import httpx2

from api import workbook_mcp

FIXTURE = Path(__file__).resolve().parents[1] / "data" / "requests.xlsx"


async def fixture_download(path):
    path.write_bytes(FIXTURE.read_bytes())


class WorkbookMcpTests(unittest.IsolatedAsyncioTestCase):
    async def test_transport_serves_the_real_upstream_read_tools(self):
        headers = {"Authorization": "Bearer local-test-only", "Accept": "application/json, text/event-stream"}
        route = "/api/workbook_mcp"

        def rpc(method, params=None):
            payload = {"jsonrpc": "2.0", "id": 1, "method": method}
            if params is not None:
                payload["params"] = params
            return payload

        with patch.dict(os.environ, {"MCP_HTTP_TOKEN": "local-test-only"}), \
                patch.object(workbook_mcp, "download_workbook", fixture_download):
            async with httpx2.AsyncClient(
                transport=httpx2.ASGITransport(app=workbook_mcp.app), base_url="http://test",
            ) as client:
                denied = await client.post(route, json=rpc("tools/list"))
                self.assertEqual(denied.status_code, 401)

                init = await client.post(route, headers=headers, json=rpc("initialize", {
                    "protocolVersion": "2025-11-25", "capabilities": {},
                    "clientInfo": {"name": "test", "version": "1"},
                }))
                self.assertEqual(init.json()["result"]["serverInfo"]["name"], "excel-mcp-server")

                catalog = await client.post(route, headers=headers, json=rpc("tools/list"))
                names = {tool["name"] for tool in catalog.json()["result"]["tools"]}
                self.assertEqual(names, {"describe_workbook", "describe_sheet", "read_range"})

                args = {"path": "requests.xlsx", "sheet": "Заявки", "range": "A1:D4"}
                result = await client.post(route, headers=headers, json=rpc("tools/call", {
                    "name": "read_range", "arguments": args,
                }))
                self.assertEqual([row[0] for row in result.json()["result"]["structuredContent"]["values"][1:]], [101, 102, 103])

                wrong_file = await client.post(route, headers=headers, json=rpc("tools/call", {
                    "name": "read_range", "arguments": {**args, "path": "other.xlsx"},
                }))
                self.assertEqual(wrong_file.status_code, 400)

                denied_write = await client.post(route, headers=headers, json=rpc("tools/call", {
                    "name": "write_range", "arguments": {"path": "requests.xlsx", "sheet": "Заявки", "at": "D2", "rows": [["Готово"]]},
                }))
                self.assertTrue(denied_write.json()["result"]["isError"])

    async def test_concurrent_edits_preserve_the_winning_version_and_retry_is_idempotent(self):
        head = None
        versions = {None: FIXTURE.read_bytes()}
        receipts = {}
        delayed = set()

        async def download(path, revision=None):
            path.write_bytes(versions[revision])

        async def upload(path, revision):
            versions[revision] = path.read_bytes()
            await sleep(0)

        async def state(operation=None, request_hash=None, expected=None, candidate=None, write_at=None, ticket_id=None):
            nonlocal head
            if operation is None:
                return {"revision": head}
            if operation in receipts:
                saved_hash, saved_receipt = receipts[operation]
                return ({**saved_receipt, "replayed": True} if request_hash == saved_hash
                        else {"conflict": True})
            if candidate is None:
                return {"pending": True}
            if head != expected:
                return {"conflict": True}
            if candidate not in delayed:
                delayed.add(candidate)
                request = httpx2.Request("POST", "https://example.test/rpc")
                response = httpx2.Response(400, request=request, json={
                    "code": "22023", "message": "Workbook version was not uploaded",
                })
                response.raise_for_status()
            self.assertIn(candidate, versions)
            head = candidate
            receipt = {"revision": candidate, "replayed": False, "at": write_at, "ticketId": ticket_id}
            receipts[operation] = (request_hash, receipt)
            return receipt

        headers = {"Authorization": "Bearer local-test-only", "Accept": "application/json, text/event-stream"}
        route = "/api/workbook_mcp"

        def write(at, status):
            return {"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {
                "_meta": {"workbookOperation": str(uuid4())},
                "name": "write_range", "arguments": {
                    "path": "requests.xlsx", "sheet": "Заявки", "at": at, "rows": [[status]],
                },
            }}

        def read():
            return {"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {
                "name": "read_range", "arguments": {"path": "requests.xlsx", "sheet": "Заявки", "range": "D2:D3"},
            }}

        with patch.dict(os.environ, {"MCP_HTTP_TOKEN": "local-test-only", "WORKBOOK_WRITE_TOKEN": "test-writes"}), \
                patch.object(workbook_mcp, "download_workbook", download), \
                patch.object(workbook_mcp, "upload_version", upload), \
                patch.object(workbook_mcp, "workbook_state", state):
            async with httpx2.AsyncClient(
                transport=httpx2.ASGITransport(app=workbook_mcp.app), base_url="http://test", headers=headers,
            ) as client:
                catalog = await client.post(route, json={"jsonrpc": "2.0", "id": 1, "method": "tools/list"})
                names = {tool["name"] for tool in catalog.json()["result"]["tools"]}
                self.assertEqual(names, {"describe_workbook", "describe_sheet", "read_range", "write_range"})
                first_read = await client.post(route, json=read())
                self.assertEqual(first_read.headers["x-workbook-revision"], "seed")

                writes = [write("D2", "Готово"), write("D3", "Готово")]
                revision_header = {"x-workbook-revision": "seed"}
                replies = await gather(*(client.post(route, json=payload, headers=revision_header) for payload in writes))
                successful = [i for i, reply in enumerate(replies) if not reply.json()["result"].get("isError")]
                self.assertEqual(len(successful), 1)
                winner = successful[0]
                latest_read = await client.post(route, json=read())
                self.assertEqual(latest_read.headers["x-workbook-revision"], head)
                statuses = [["Новый"], ["В работе"]]
                statuses[winner] = ["Готово"]
                self.assertEqual(latest_read.json()["result"]["structuredContent"]["values"], statuses)
                retry = await client.post(route, json=writes[winner], headers=revision_header)
                self.assertEqual(retry.headers["x-workbook-revision"], head)
                self.assertIn('"replayed": true', retry.json()["result"]["content"][0]["text"])
                self.assertEqual(len(receipts), 1)


if __name__ == "__main__":
    unittest.main()
