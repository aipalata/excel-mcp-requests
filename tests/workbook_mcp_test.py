import os
import unittest
from pathlib import Path
from unittest.mock import patch

import httpx2

from api import workbook_mcp

FIXTURE = Path(__file__).resolve().parents[1] / "data" / "requests.xlsx"


async def fixture_download(path):
    path.write_bytes(FIXTURE.read_bytes())


class ReadOnlyMcpTests(unittest.IsolatedAsyncioTestCase):
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


if __name__ == "__main__":
    unittest.main()
