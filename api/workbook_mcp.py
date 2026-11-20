"""Stateless HTTP adapter for the ready-made Excel MCP server."""

import json
import os
import re
from pathlib import Path
from tempfile import TemporaryDirectory

import httpx2
from starlette.applications import Starlette
from starlette.requests import ClientDisconnect, Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from excel_mcp.config import Limits, Settings
from excel_mcp.server import create_server
from excel_mcp.server.http import BearerTokenMiddleware

WORKBOOK = "requests.xlsx"
MAX_WORKBOOK_BYTES = 1024 * 1024
MAX_REQUEST_BYTES = 32 * 1024


async def download_workbook(path: Path) -> None:
    project = os.environ["WORKBOOK_SUPABASE_PROJECT"]
    if not re.fullmatch(r"[a-z0-9]{20}", project):
        raise ValueError("Invalid Supabase project")
    headers = {
        "apikey": os.environ["WORKBOOK_SUPABASE_PUBLISHABLE_KEY"],
        "x-workbook-access": os.environ["WORKBOOK_STORAGE_TOKEN"],
    }
    url = f"https://{project}.supabase.co/storage/v1/object/authenticated/workbooks/{WORKBOOK}"
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
    if is_tool_call:
        params = message.get("params")
        args = params.get("arguments") if isinstance(params, dict) else None
        if not isinstance(args, dict) or args.get("path") != WORKBOOK:
            await JSONResponse({"error": "Unknown workbook"}, status_code=400)(scope, receive, send)
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
        if is_tool_call:
            try:
                await download_workbook(workspace / WORKBOOK)
            except (KeyError, ValueError, OSError, httpx2.HTTPError):
                await JSONResponse({"error": "Workbook storage unavailable"}, status_code=503)(scope, receive, send)
                return

        server = create_server(Settings(
            allowed_dirs=[workspace],
            read_only=True,
            tool_areas=(),
            limits=Limits(max_file_bytes=MAX_WORKBOOK_BYTES, max_cells=4004),
        ))
        upstream = server.streamable_http_app(
            streamable_http_path=scope["path"],
            host="0.0.0.0",
            stateless_http=True,
            json_response=True,
            max_request_body_size=MAX_REQUEST_BYTES,
        )
        async with server.session_manager.run():
            await upstream(scope, replay_body, send)


class ProtectedMcp:
    async def __call__(self, scope, receive, send):
        bearer = os.environ.get("MCP_HTTP_TOKEN", "").strip()
        if not bearer:
            await JSONResponse({"error": "MCP not configured"}, status_code=503)(scope, receive, send)
            return
        await BearerTokenMiddleware(serve_mcp, bearer)(scope, receive, send)


app = Starlette(routes=[Route("/api/workbook_mcp", endpoint=ProtectedMcp())])
