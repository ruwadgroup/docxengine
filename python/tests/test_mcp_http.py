"""MCP Streamable HTTP transport tests (algorithms.md §25).

A real :class:`ThreadingHTTPServer` is bound to an ephemeral port (``port 0``)
and driven over a live socket with ``urllib.request``: initialize mints a
session id, that id unlocks tools/call + resources, and an expired session
yields HTTP 410.
"""

from __future__ import annotations

import json
import threading
import urllib.error
import urllib.request
from collections.abc import Iterator
from pathlib import Path

import pytest

from docxengine.mcp import _SESSION_HEADER, _SessionStore, serve_http


@pytest.fixture
def server() -> Iterator[tuple[str, _SessionStore]]:
    """A running HTTP server on an ephemeral port; yields its base URL + store."""
    store = _SessionStore()
    httpd = serve_http("127.0.0.1", 0, store=store)
    host, port = httpd.server_address[0], httpd.server_address[1]
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://{host}:{port}", store
    finally:
        httpd.shutdown()
        httpd.server_close()
        thread.join(timeout=5)


@pytest.fixture
def docx_path(docx_bytes: bytes, tmp_path: Path) -> Path:
    path = tmp_path / "fixture.docx"
    path.write_bytes(docx_bytes)
    return path


def post(
    base_url: str, body: dict[str, object], *, session_id: str | None = None
) -> tuple[int, dict[str, object], str | None]:
    """POST a JSON-RPC body; returns ``(status, parsed_body, session_id_header)``."""
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        base_url + "/", data=data, headers={"Content-Type": "application/json"}
    )
    if session_id is not None:
        req.add_header(_SESSION_HEADER, session_id)
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw = resp.read().decode("utf-8")
            parsed = json.loads(raw) if raw else {}
            return resp.status, parsed, resp.headers.get(_SESSION_HEADER)
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode("utf-8")
        parsed = json.loads(raw) if raw else {}
        return exc.code, parsed, exc.headers.get(_SESSION_HEADER)


def rpc(req_id: int | None, method: str, **params: object) -> dict[str, object]:
    message: dict[str, object] = {"jsonrpc": "2.0", "method": method}
    if req_id is not None:
        message["id"] = req_id
    if params:
        message["params"] = params
    return message


INITIALIZE = rpc(
    1,
    "initialize",
    protocolVersion="2025-03-26",
    capabilities={},
    clientInfo={"name": "pytest", "version": "0"},
)


def test_health(server: tuple[str, _SessionStore]) -> None:
    base_url, _ = server
    with urllib.request.urlopen(base_url + "/health", timeout=10) as resp:
        assert resp.status == 200
        assert json.loads(resp.read().decode("utf-8")) == {"status": "ok"}


def test_initialize_mints_session_id_in_header(server: tuple[str, _SessionStore]) -> None:
    base_url, _ = server
    status, body, session_id = post(base_url, INITIALIZE)
    assert status == 200
    assert body["result"]["protocolVersion"] == "2025-03-26"  # type: ignore[index]
    assert body["result"]["serverInfo"]["name"] == "docxengine"  # type: ignore[index]
    assert session_id and len(session_id) >= 8


def test_full_session_flow(server: tuple[str, _SessionStore], docx_path: Path) -> None:
    base_url, _ = server
    _, _, session_id = post(base_url, INITIALIZE)
    assert session_id is not None

    # tools/call docx_open within the session's own doc store.
    status, body, _ = post(
        base_url,
        rpc(2, "tools/call", name="docx_open", arguments={"path": str(docx_path)}),
        session_id=session_id,
    )
    assert status == 200
    result = body["result"]  # type: ignore[index]
    assert result["isError"] is False
    opened = json.loads(result["content"][0]["text"])
    assert opened["doc_id"] == "d1"

    # resources/list reflects the now-open doc.
    status, body, _ = post(base_url, rpc(3, "resources/list"), session_id=session_id)
    assert status == 200
    uris = {r["uri"] for r in body["result"]["resources"]}  # type: ignore[index]
    assert uris == {"docx://d1/outline", "docx://d1/projection"}

    # resources/read returns text/markdown from the projector.
    status, body, _ = post(
        base_url,
        rpc(4, "resources/read", uri="docx://d1/projection"),
        session_id=session_id,
    )
    assert status == 200
    contents = body["result"]["contents"][0]  # type: ignore[index]
    assert contents["mimeType"] == "text/markdown"
    assert "Master Services Agreement" in contents["text"]


def test_missing_session_id_is_jsonrpc_error(server: tuple[str, _SessionStore]) -> None:
    base_url, _ = server
    status, body, _ = post(base_url, rpc(2, "tools/list"))
    assert status == 200
    assert body["error"]["code"] == -32600  # type: ignore[index]
    assert _SESSION_HEADER in body["error"]["message"]  # type: ignore[index]


def test_unknown_session_id_is_jsonrpc_error(server: tuple[str, _SessionStore]) -> None:
    base_url, _ = server
    status, body, _ = post(base_url, rpc(2, "tools/list"), session_id="deadbeef")
    assert status == 200
    assert body["error"]["code"] == -32600  # type: ignore[index]


def test_expired_session_is_410(server: tuple[str, _SessionStore]) -> None:
    base_url, store = server
    _, _, session_id = post(base_url, INITIALIZE)
    assert session_id is not None
    assert store.expire(session_id) is True
    status, _, _ = post(base_url, rpc(2, "tools/list"), session_id=session_id)
    assert status == 410


def test_sessions_have_isolated_doc_stores(
    server: tuple[str, _SessionStore], docx_path: Path
) -> None:
    base_url, _ = server
    _, _, sid_a = post(base_url, INITIALIZE)
    _, _, sid_b = post(base_url, INITIALIZE)
    assert sid_a != sid_b

    post(
        base_url,
        rpc(2, "tools/call", name="docx_open", arguments={"path": str(docx_path)}),
        session_id=sid_a,
    )
    # Session B never opened anything: its resources list is empty.
    _, body, _ = post(base_url, rpc(3, "resources/list"), session_id=sid_b)
    assert body["result"]["resources"] == []  # type: ignore[index]
