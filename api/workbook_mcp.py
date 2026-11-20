"""Stateless HTTP adapter for the ready-made Excel MCP server."""

import hashlib
import json
import os
import re
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4

import httpx2
from starlette.applications import Starlette
from starlette.requests import ClientDisconnect, Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from excel_mcp.config import Limits, Settings
from excel_mcp.server import create_server
from excel_mcp.server.areas import AREAS
from excel_mcp.server.http import BearerTokenMiddleware

WORKBOOK = "requests.xlsx"
MAX_WORKBOOK_BYTES = 1024 * 1024
MAX_REQUEST_BYTES = 32 * 1024
UUID_PATTERN = re.compile(r"[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}")
XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
STATUSES = {"Новый", "В работе", "Готово"}


def storage_config() -> tuple[str, dict[str, str]]:
    project = os.environ["WORKBOOK_SUPABASE_PROJECT"]
    if not re.fullmatch(r"[a-z0-9]{20}", project):
        raise ValueError("Invalid Supabase project")
    return f"https://{project}.supabase.co", {
        "apikey": os.environ["WORKBOOK_SUPABASE_PUBLISHABLE_KEY"],
        "x-workbook-access": os.environ["WORKBOOK_STORAGE_TOKEN"],
        "x-workbook-write": os.environ.get("WORKBOOK_WRITE_TOKEN", ""),
    }


def object_name(revision: str | None) -> str:
    if revision is None:
        return WORKBOOK
    if not UUID_PATTERN.fullmatch(revision):
        raise ValueError("Invalid workbook revision")
    return f"versions/{revision}.xlsx"


async def download_workbook(path: Path, revision: str | None = None) -> None:
    base, headers = storage_config()
    url = f"{base}/storage/v1/object/authenticated/workbooks/{object_name(revision)}"
    async with httpx2.AsyncClient(timeout=4, follow_redirects=False) as client:
        async with client.stream("GET", url, headers=headers) as response:
            response.raise_for_status()
            size = 0
            with path.open("wb") as destination:
                async for chunk in response.aiter_bytes():
                    size += len(chunk)
                    if size > MAX_WORKBOOK_BYTES:
                        raise ValueError("Workbook exceeds the size limit")
                    destination.write(chunk)


async def workbook_state(**parameters) -> dict:
    base, headers = storage_config()
    async with httpx2.AsyncClient(timeout=4, follow_redirects=False) as client:
        response = await client.post(f"{base}/rest/v1/rpc/requests_workbook_state", headers=headers, json=parameters)
        response.raise_for_status()
        return response.json()


async def upload_version(path: Path, revision: str) -> None:
    if path.stat().st_size > MAX_WORKBOOK_BYTES:
        raise ValueError("Workbook exceeds the size limit")
    base, headers = storage_config()
    headers.update({"Content-Type": XLSX_CONTENT_TYPE, "Cache-Control": "no-store"})
    async with httpx2.AsyncClient(timeout=4, follow_redirects=False) as client:
        url = f"{base}/storage/v1/object/workbooks/{object_name(revision)}"
        response = await client.post(url, headers=headers, content=path.read_bytes())
        response.raise_for_status()


def valid_write(args: dict) -> bool:
    if args.keys() != {"path", "sheet", "at", "rows"} or args["sheet"] != "Заявки":
        return False
    at, rows = args["at"], args["rows"]
    if not isinstance(at, str) or not re.fullmatch(r"[AD](?:[2-9]|[1-9]\d{1,2}|100[01])", at):
        return False
    if not isinstance(rows, list) or len(rows) != 1 or not isinstance(rows[0], list):
        return False
    row = rows[0]
    if at.startswith("D"):
        return len(row) == 1 and isinstance(row[0], str) and row[0] in STATUSES
    if len(row) != 4:
        return False
    ticket_id, client, problem, status = row
    return (
        type(ticket_id) is int and 1 <= ticket_id <= 999_999_999 and status == "Новый"
        and isinstance(client, str) and 0 < len(client) <= 100 and not client.startswith("=")
        and isinstance(problem, str) and 0 < len(problem) <= 450 and not problem.startswith("=")
    )


def tool_result(message: dict, revision: str | None, *, error: str | None = None, receipt: dict | None = None):
    content = error or json.dumps({"workbookReceipt": {key: receipt[key] for key in ("at", "ticketId", "replayed")}})
    return JSONResponse({
        "jsonrpc": "2.0", "id": message.get("id"),
        "result": {"content": [{"type": "text", "text": content}], "isError": error is not None},
    }, headers={"x-workbook-revision": revision or "seed"})


async def serve_mcp(scope, receive, send):
    if scope["method"] != "POST":
        await JSONResponse({"error": "POST required"}, status_code=405, headers={"Allow": "POST"})(scope, receive, send)
        return

    request = Request(scope, receive)
    body = bytearray()
    try:
        async for chunk in request.stream():
            if len(body) + len(chunk) > MAX_REQUEST_BYTES:
                await JSONResponse({"error": "Request too large"}, status_code=413)(scope, receive, send)
                return
            body.extend(chunk)
    except ClientDisconnect:
        return

    try:
        message = json.loads(body)
    except (ValueError, UnicodeError):
        message = None
    is_tool_call = isinstance(message, dict) and message.get("method") == "tools/call"
    writable = bool(os.environ.get("WORKBOOK_WRITE_TOKEN", "").strip())
    mutation = False
    operation = None
    requested_revision = request.headers.get("x-workbook-revision")
    if is_tool_call:
        params = message.get("params")
        args = params.get("arguments") if isinstance(params, dict) else None
        if not isinstance(args, dict) or args.get("path") != WORKBOOK:
            await JSONResponse({"error": "Unknown workbook"}, status_code=400)(scope, receive, send)
            return
        mutation = writable and params.get("name") == "write_range"
        if requested_revision not in {None, "seed"} and not UUID_PATTERN.fullmatch(requested_revision):
            await JSONResponse({"error": "Invalid workbook revision"}, status_code=400)(scope, receive, send)
            return
        if mutation:
            metadata = params.get("_meta")
            operation = metadata.get("workbookOperation") if isinstance(metadata, dict) else None
            if not requested_revision or not isinstance(operation, str) or not UUID_PATTERN.fullmatch(operation) or not valid_write(args):
                await JSONResponse({"error": "Invalid workbook update"}, status_code=400)(scope, receive, send)
                return

    first_chunk = True

    async def replay_body():
        nonlocal first_chunk
        if first_chunk:
            first_chunk = False
            return {"type": "http.request", "body": bytes(body), "more_body": False}
        return await receive()

    with TemporaryDirectory(prefix="requests-mcp-") as directory:
        workspace = Path(directory)
        path = workspace / WORKBOOK
        revision = None
        if is_tool_call:
            try:
                if writable:
                    revision = (
                        (await workbook_state())["revision"] if requested_revision is None
                        else requested_revision if requested_revision != "seed" else None
                    )
                    if mutation:
                        if args["at"].startswith("A"):
                            intent = ["add", *args["rows"][0][1:]]
                        else:
                            intent = ["status", args["at"], args["rows"]]
                        request_hash = hashlib.sha256(json.dumps(intent, ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()
                        receipt = await workbook_state(operation=operation, request_hash=request_hash)
                        if receipt.get("conflict") or receipt.get("replayed"):
                            await tool_result(message, receipt.get("revision"),
                                              error="Operation already used" if receipt.get("conflict") else None,
                                              receipt=receipt)(scope, receive, send)
                            return
                if revision is None:
                    await download_workbook(path)
                else:
                    await download_workbook(path, revision)
            except (KeyError, ValueError, OSError, httpx2.HTTPError):
                await JSONResponse({"error": "Workbook storage unavailable"}, status_code=503)(scope, receive, send)
                return

        server = create_server(Settings(
            allowed_dirs=[workspace],
            read_only=not writable,
            tool_areas=("data",) if writable else (),
            limits=Limits(max_file_bytes=MAX_WORKBOOK_BYTES, max_cells=4004),
        ))
        if writable:
            for name in AREAS["data"]:
                if name != "write_range":
                    server.remove_tool(name)
        upstream = server.streamable_http_app(
            streamable_http_path=scope["path"],
            host="0.0.0.0",
            stateless_http=True,
            json_response=True,
            max_request_body_size=MAX_REQUEST_BYTES,
        )
        outgoing = []

        async def send_response(event):
            if mutation:
                outgoing.append(event)
                return
            if writable and is_tool_call and event["type"] == "http.response.start":
                event["headers"] = [*event.get("headers", []), (b"x-workbook-revision", (revision or "seed").encode())]
            await send(event)

        async with server.session_manager.run():
            await upstream(scope, replay_body, send_response)

        if not mutation:
            return
        try:
            body = b"".join(event["body"] for event in outgoing if event["type"] == "http.response.body")
            result = json.loads(body).get("result", {})
            if result and not result.get("isError"):
                candidate = str(uuid4())
                await upload_version(path, candidate)
                ticket_id = args["rows"][0][0] if args["at"].startswith("A") else None
                receipt = await workbook_state(operation=operation, request_hash=request_hash,
                                               expected=revision, candidate=candidate,
                                               write_at=args["at"], ticket_id=ticket_id)
                if receipt.get("conflict"):
                    await tool_result(message, revision, error="Workbook changed. Refresh and retry.")(scope, receive, send)
                    return
                await tool_result(message, receipt["revision"], receipt=receipt)(scope, receive, send)
                return
        except (KeyError, ValueError, OSError, httpx2.HTTPError):
            await tool_result(message, revision, error="Workbook update could not be verified. Refresh and retry.")(scope, receive, send)
            return

        for event in outgoing:
            if event["type"] == "http.response.start":
                event["headers"] = [*event.get("headers", []), (b"x-workbook-revision", (revision or "seed").encode())]
            await send(event)


class ProtectedMcp:
    async def __call__(self, scope, receive, send):
        bearer = os.environ.get("MCP_HTTP_TOKEN", "").strip()
        if not bearer:
            await JSONResponse({"error": "MCP not configured"}, status_code=503)(scope, receive, send)
            return
        await BearerTokenMiddleware(serve_mcp, bearer)(scope, receive, send)


app = Starlette(routes=[Route("/api/workbook_mcp", endpoint=ProtectedMcp())])
