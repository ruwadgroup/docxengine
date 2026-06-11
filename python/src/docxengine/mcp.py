"""MCP server (docs/mcp/server.md): ``docxengine-mcp``.

Two transports share one JSON-RPC dispatch (:func:`_handle`):

* **stdio** (default, the conformance transport): dependency-free JSON-RPC 2.0
  with newline-delimited framing per the MCP stdio transport (2025-03-26 spec):
  one JSON-RPC message per line, no Content-Length headers, no embedded
  newlines. EOF on stdin exits 0.
* **Streamable HTTP** (``--http --port {p}``, algorithms.md §25): a stdlib
  threading ``http.server``. ``POST /`` carries a JSON-RPC body; ``initialize``
  allocates an ``Mcp-Session-Id`` (response header) and every later POST must
  echo it (missing/unknown → JSON-RPC error; an *expired* session → HTTP 410).
  Each session owns its own doc store. ``GET /health`` → ``{"status":"ok"}``.

``tools/list`` is generated from the packaged spec schemas; ``tools/call``
dispatches to the same :func:`~docxengine._dispatch.call` the CLI uses. Tool
failures are MCP tool results with ``isError: true`` carrying the structured
error JSON, never JSON-RPC errors. ``resources/list`` / ``resources/read``
expose, per open doc, ``docx://{doc_id}/outline`` and ``/projection`` rendered
by the projector. doc_ids persist for the session lifetime.
"""

from __future__ import annotations

import argparse
import json
import socketserver
import sys
import threading
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import IO, Any

from . import __version__, _projector
from ._dispatch import call
from ._errors import ToolError
from ._session import Session
from ._spec import tool_schemas

PROTOCOL_VERSION = "2025-03-26"

_PARSE_ERROR = -32700
_INVALID_REQUEST = -32600
_METHOD_NOT_FOUND = -32601
_INVALID_PARAMS = -32602
_INTERNAL_ERROR = -32603

#: docx://{doc_id}/{view} resource URIs (one per open doc, per view).
_RESOURCE_VIEWS = ("outline", "projection")


def _error(req_id: object, code: int, message: str) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}}


def _result(req_id: object, result: dict[str, Any]) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def _mcp_tools() -> list[dict[str, Any]]:
    """The spec tool schemas in the MCP shape (``input_schema`` -> ``inputSchema``)."""
    return [
        {
            "name": schema["name"],
            "description": schema["description"],
            "inputSchema": schema.get("input_schema", {"type": "object"}),
        }
        for schema in tool_schemas()
    ]


def _initialize(req_id: object) -> dict[str, Any]:
    return _result(
        req_id,
        {
            "protocolVersion": PROTOCOL_VERSION,
            "capabilities": {
                "tools": {"listChanged": False},
                "resources": {"listChanged": False, "subscribe": False},
            },
            "serverInfo": {"name": "docxengine", "version": __version__},
        },
    )


def _tools_call(req_id: object, params: Any, session: Session) -> dict[str, Any]:
    if not isinstance(params, dict) or not isinstance(params.get("name"), str):
        return _error(req_id, _INVALID_PARAMS, "tools/call requires params with a string 'name'.")
    arguments = params.get("arguments")
    if arguments is not None and not isinstance(arguments, dict):
        return _error(req_id, _INVALID_PARAMS, "tools/call 'arguments' must be an object.")
    try:
        payload: dict[str, Any] = dict(call(params["name"], arguments, session=session))
        is_error = False
    except ToolError as exc:
        payload = exc.to_payload()
        is_error = True
    return _result(
        req_id,
        {
            "content": [{"type": "text", "text": json.dumps(payload, ensure_ascii=False)}],
            "isError": is_error,
        },
    )


# ---------------------------------------------------------------------------
# Resources (algorithms.md §25): docx://{doc_id}/{outline,projection}
# ---------------------------------------------------------------------------


def _resources_list(session: Session) -> list[dict[str, Any]]:
    """One resource per (open doc, view): ``docx://{doc_id}/{view}``."""
    resources: list[dict[str, Any]] = []
    for doc_id in session.doc_ids():
        for view in _RESOURCE_VIEWS:
            resources.append(
                {
                    "uri": f"docx://{doc_id}/{view}",
                    "name": f"{doc_id} {view}",
                    "description": f"The §2{'a' if view == 'outline' else ''} {view} of {doc_id}.",
                    "mimeType": "text/markdown",
                }
            )
    return resources


def _parse_resource_uri(uri: object) -> tuple[str, str]:
    """``docx://{doc_id}/{view}`` → ``(doc_id, view)``; ``ToolError`` otherwise."""
    if not isinstance(uri, str) or not uri.startswith("docx://"):
        raise ToolError(
            "invalid_args",
            f"Unknown resource URI: {uri!r}.",
            ["Use docx://{doc_id}/outline or docx://{doc_id}/projection."],
        )
    rest = uri[len("docx://") :]
    doc_id, _, view = rest.partition("/")
    if not doc_id or view not in _RESOURCE_VIEWS:
        raise ToolError(
            "invalid_args",
            f"Unknown resource URI: {uri!r}.",
            ["Use docx://{doc_id}/outline or docx://{doc_id}/projection."],
        )
    return doc_id, view


def _resources_read(uri: object, session: Session) -> dict[str, Any]:
    """The ``text/markdown`` body of one resource; raises ``ToolError`` on failure."""
    doc_id, view = _parse_resource_uri(uri)
    package = session.get(doc_id).package
    if view == "outline":
        text = _projector.render_outline_markdown(package)
    else:
        text = _projector.render_projection_markdown(package)
    return {
        "contents": [
            {"uri": f"docx://{doc_id}/{view}", "mimeType": "text/markdown", "text": text}
        ]
    }


def _handle(message: Any, session: Session) -> dict[str, Any] | None:
    """One JSON-RPC message in, one response object out (or None for notifications)."""
    if not isinstance(message, dict) or not isinstance(message.get("method"), str):
        return _error(None, _INVALID_REQUEST, "Message must be a JSON-RPC request object.")
    method: str = message["method"]
    req_id = message.get("id")
    if "id" not in message:  # notification: never answered, even if unknown
        return None
    if method == "initialize":
        return _initialize(req_id)
    if method == "ping":
        return _result(req_id, {})
    if method == "tools/list":
        return _result(req_id, {"tools": _mcp_tools()})
    if method == "tools/call":
        try:
            return _tools_call(req_id, message.get("params"), session)
        except Exception as exc:  # never kill the transport on a handler bug
            return _error(req_id, _INTERNAL_ERROR, f"{type(exc).__name__}: {exc}")
    if method == "resources/list":
        return _result(req_id, {"resources": _resources_list(session)})
    if method == "resources/read":
        params = message.get("params")
        uri = params.get("uri") if isinstance(params, dict) else None
        try:
            return _result(req_id, _resources_read(uri, session))
        except ToolError as exc:
            return _error(req_id, _INVALID_PARAMS, f"{exc.code}: {exc.message}")
        except Exception as exc:  # never kill the transport on a handler bug
            return _error(req_id, _INTERNAL_ERROR, f"{type(exc).__name__}: {exc}")
    return _error(req_id, _METHOD_NOT_FOUND, f"Method not found: {method}")


# ---------------------------------------------------------------------------
# stdio transport (the conformance transport — unchanged framing)
# ---------------------------------------------------------------------------


def serve(stdin: IO[str], stdout: IO[str], *, session: Session | None = None) -> int:
    """Run the JSON-RPC loop until EOF; one session per call."""
    session = session if session is not None else Session()
    for raw in stdin:
        line = raw.strip()
        if not line:
            continue
        try:
            message = json.loads(line)
        except json.JSONDecodeError as exc:
            response: dict[str, Any] | None = _error(None, _PARSE_ERROR, f"Parse error: {exc}.")
        else:
            response = _handle(message, session)
        if response is not None:
            stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            stdout.flush()
    return 0


# ---------------------------------------------------------------------------
# Streamable HTTP transport (algorithms.md §25)
# ---------------------------------------------------------------------------

_SESSION_HEADER = "Mcp-Session-Id"


class _SessionStore:
    """Thread-safe ``Mcp-Session-Id`` → :class:`Session` registry.

    ``initialize`` mints a fresh id with its own doc store; ``expire`` drops a
    session so a later POST with that id reports it as *expired* (HTTP 410), as
    opposed to *unknown* (a JSON-RPC error).
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._sessions: dict[str, Session] = {}
        self._expired: set[str] = set()

    def create(self) -> str:
        session_id = uuid.uuid4().hex
        with self._lock:
            self._sessions[session_id] = Session()
        return session_id

    def get(self, session_id: str) -> Session | None:
        with self._lock:
            return self._sessions.get(session_id)

    def is_expired(self, session_id: str) -> bool:
        with self._lock:
            return session_id in self._expired

    def expire(self, session_id: str) -> bool:
        with self._lock:
            if session_id in self._sessions:
                del self._sessions[session_id]
                self._expired.add(session_id)
                return True
            return False


def _is_initialize(message: Any) -> bool:
    return isinstance(message, dict) and message.get("method") == "initialize"


class _Handler(BaseHTTPRequestHandler):
    """JSON-RPC-over-HTTP handler sharing :func:`_handle` with the stdio loop."""

    protocol_version = "HTTP/1.1"
    store: _SessionStore  # set on the subclass by :func:`_make_handler`

    def log_message(self, *_args: Any) -> None:  # silence default stderr logging
        return

    def _send_json(
        self, status: int, body: dict[str, Any], *, session_id: str | None = None
    ) -> None:
        payload = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        if session_id is not None:
            self.send_header(_SESSION_HEADER, session_id)
        self.end_headers()
        self.wfile.write(payload)

    def _send_empty(self, status: int) -> None:
        self.send_response(status)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self) -> None:  # noqa: N802 - http.server dispatch name
        if self.path == "/health":
            self._send_json(200, {"status": "ok"})
        else:
            self._send_empty(404)

    def do_POST(self) -> None:  # noqa: N802 - http.server dispatch name
        if self.path not in ("/", ""):
            self._send_empty(404)
            return
        length = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(length) if length else b""
        try:
            message = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            self._send_json(200, _error(None, _PARSE_ERROR, f"Parse error: {exc}."))
            return

        if _is_initialize(message):
            session_id = self.store.create()
            response = _handle(message, self.store.get(session_id))  # type: ignore[arg-type]
            self._send_json(200, response or {}, session_id=session_id)
            return

        header_id = self.headers.get(_SESSION_HEADER)
        if not header_id:
            self._send_json(
                200,
                _error(
                    message.get("id") if isinstance(message, dict) else None,
                    _INVALID_REQUEST,
                    f"Missing {_SESSION_HEADER} header.",
                ),
            )
            return
        if self.store.is_expired(header_id):
            self._send_empty(410)
            return
        session = self.store.get(header_id)
        if session is None:
            self._send_json(
                200,
                _error(
                    message.get("id") if isinstance(message, dict) else None,
                    _INVALID_REQUEST,
                    f"Unknown {_SESSION_HEADER}: {header_id}.",
                ),
            )
            return
        response = _handle(message, session)
        if response is None:  # notification: 202 Accepted, no JSON-RPC body
            self._send_empty(202)
        else:
            self._send_json(200, response, session_id=header_id)


def _make_handler(store: _SessionStore) -> type[_Handler]:
    return type("_BoundHandler", (_Handler,), {"store": store})


class _Server(ThreadingHTTPServer):
    """Threading HTTP server that skips the reverse-DNS ``getfqdn`` on bind.

    The stdlib base resolves ``server_name`` via :func:`socket.getfqdn`, which
    can stall for tens of seconds on a host with no reverse record; we never use
    ``server_name``, so binding plainly keeps startup instant.
    """

    daemon_threads = True
    allow_reuse_address = True

    def server_bind(self) -> None:
        # Bypass HTTPServer.server_bind's getfqdn(host) lookup.
        socketserver.TCPServer.server_bind(self)
        host, port = self.server_address[:2]
        self.server_name = str(host)
        self.server_port = int(port)


def serve_http(host: str, port: int, *, store: _SessionStore | None = None) -> _Server:
    """Build (do not start) a threading HTTP server bound to ``host:port``."""
    store = store if store is not None else _SessionStore()
    return _Server((host, port), _make_handler(store))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="docxengine-mcp", description="DocxEngine MCP server.")
    parser.add_argument(
        "--http", action="store_true", help="Serve over Streamable HTTP instead of stdio."
    )
    parser.add_argument("--port", type=int, default=8765, help="HTTP port (with --http).")
    parser.add_argument("--host", default="127.0.0.1", help="HTTP bind host (with --http).")
    args = parser.parse_args(argv)
    if args.http:
        server = serve_http(args.host, args.port)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
        finally:
            server.server_close()
        return 0
    return serve(sys.stdin, sys.stdout)


if __name__ == "__main__":
    raise SystemExit(main())
