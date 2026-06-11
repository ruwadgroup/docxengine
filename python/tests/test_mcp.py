"""MCP stdio server tests — real subprocess speaking newline-delimited JSON-RPC 2.0."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest


def run_mcp(lines: list[str]) -> tuple[list[dict[str, object]], int]:
    """Pipe JSON-RPC lines into ``python -m docxengine.mcp``; parsed responses + rc."""
    proc = subprocess.run(
        [sys.executable, "-m", "docxengine.mcp"],
        input="".join(line + "\n" for line in lines),
        capture_output=True,
        text=True,
        timeout=60,
    )
    responses = [json.loads(line) for line in proc.stdout.splitlines() if line.strip()]
    return responses, proc.returncode


def rpc(req_id: int | None, method: str, **params: object) -> str:
    message: dict[str, object] = {"jsonrpc": "2.0", "method": method}
    if req_id is not None:
        message["id"] = req_id
    if params:
        message["params"] = params
    return json.dumps(message)


INITIALIZE = rpc(
    1,
    "initialize",
    protocolVersion="2025-03-26",
    capabilities={},
    clientInfo={"name": "pytest", "version": "0"},
)
INITIALIZED = rpc(None, "notifications/initialized")


@pytest.fixture
def docx_path(docx_bytes: bytes, tmp_path: Path) -> Path:
    path = tmp_path / "fixture.docx"
    path.write_bytes(docx_bytes)
    return path


class TestHandshake:
    def test_initialize_tools_list_call_and_clean_eof(self, docx_path: Path) -> None:
        responses, rc = run_mcp(
            [
                INITIALIZE,
                INITIALIZED,
                rpc(2, "tools/list"),
                rpc(3, "tools/call", name="docx_open", arguments={"path": str(docx_path)}),
            ]
        )
        assert rc == 0  # clean shutdown on stdin EOF
        assert [r["id"] for r in responses] == [1, 2, 3]  # notification gets no response

        init = responses[0]["result"]
        assert init["protocolVersion"] == "2025-03-26"
        assert init["serverInfo"]["name"] == "docxengine"
        assert "tools" in init["capabilities"]

        tools = responses[1]["result"]["tools"]
        assert len(tools) == 24
        assert all(set(t) == {"name", "description", "inputSchema"} for t in tools)
        by_name = {t["name"]: t for t in tools}
        assert by_name["docx_open"]["inputSchema"]["properties"]["path"]["type"] == "string"

        result = responses[2]["result"]
        assert result["isError"] is False
        assert [c["type"] for c in result["content"]] == ["text"]
        opened = json.loads(result["content"][0]["text"])
        assert opened["doc_id"] == "d1"
        assert opened["n_paragraphs"] == 3

    def test_resources_list_and_read_over_stdio(self, docx_path: Path) -> None:
        responses, rc = run_mcp(
            [
                INITIALIZE,
                INITIALIZED,
                rpc(2, "tools/call", name="docx_open", arguments={"path": str(docx_path)}),
                rpc(3, "resources/list"),
                rpc(4, "resources/read", uri="docx://d1/outline"),
                rpc(5, "resources/read", uri="docx://d1/projection"),
            ]
        )
        assert rc == 0
        by_id = {r["id"]: r for r in responses}
        resources = by_id[3]["result"]["resources"]
        assert {r["uri"] for r in resources} == {
            "docx://d1/outline",
            "docx://d1/projection",
        }
        assert all(r["mimeType"] == "text/markdown" for r in resources)

        outline = by_id[4]["result"]["contents"][0]
        assert outline["uri"] == "docx://d1/outline"
        assert outline["mimeType"] == "text/markdown"
        assert "Master Services Agreement" in outline["text"]

        projection = by_id[5]["result"]["contents"][0]["text"]
        assert "H1] Master Services Agreement" in projection
        assert "The term is five (5) years" in projection

    def test_resources_read_unknown_doc_is_invalid_params(self) -> None:
        responses, _ = run_mcp(
            [INITIALIZE, INITIALIZED, rpc(2, "resources/read", uri="docx://d404/outline")]
        )
        assert responses[1]["error"]["code"] == -32602
        assert "doc_not_found" in responses[1]["error"]["message"]

    def test_ping(self) -> None:
        responses, rc = run_mcp([INITIALIZE, INITIALIZED, rpc(2, "ping")])
        assert rc == 0
        assert responses[1] == {"jsonrpc": "2.0", "id": 2, "result": {}}

    def test_doc_state_persists_across_calls(self, docx_path: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        responses, rc = run_mcp(
            [
                INITIALIZE,
                INITIALIZED,
                rpc(2, "tools/call", name="docx_open", arguments={"path": str(docx_path)}),
                rpc(
                    3,
                    "tools/call",
                    name="docx_replace",
                    arguments={"doc_id": "d1", "old": "five (5) years", "new": "two (2) years"},
                ),
                rpc(
                    4, "tools/call", name="docx_save", arguments={"doc_id": "d1", "path": str(out)}
                ),
            ]
        )
        assert rc == 0
        saved = json.loads(responses[3]["result"]["content"][0]["text"])
        assert saved == {"ok": True, "validated": True, "bytes": out.stat().st_size}


class TestErrors:
    def test_tool_error_is_iserror_result_with_structured_json(self) -> None:
        read = rpc(2, "tools/call", name="docx_read", arguments={"doc_id": "d404"})
        responses, _ = run_mcp([INITIALIZE, INITIALIZED, read])
        result = responses[1]["result"]
        assert result["isError"] is True
        assert json.loads(result["content"][0]["text"]) == {
            "error": "doc_not_found",
            "message": "Unknown or expired doc_id: d404.",
            "suggestions": ["Call docx_open again."],
        }

    def test_every_spec_tool_is_implemented(self) -> None:
        # Phase 2 complete: a spec tool with missing args reports invalid_args, not
        # not_implemented (docx_template_fill needs 'template').
        table = rpc(2, "tools/call", name="docx_template_fill", arguments={"data": {}})
        responses, _ = run_mcp([INITIALIZE, INITIALIZED, table])
        result = responses[1]["result"]
        assert result["isError"] is True
        payload = json.loads(result["content"][0]["text"])
        assert payload["error"] == "invalid_args"
        assert "template" in payload["message"]

    def test_unknown_method_is_method_not_found(self) -> None:
        responses, _ = run_mcp([INITIALIZE, INITIALIZED, rpc(2, "prompts/list")])
        assert responses[1]["error"]["code"] == -32601

    def test_unknown_notification_is_ignored(self) -> None:
        responses, rc = run_mcp([INITIALIZE, rpc(None, "notifications/cancelled"), rpc(2, "ping")])
        assert rc == 0
        assert [r["id"] for r in responses] == [1, 2]

    def test_parse_error_does_not_kill_the_process(self) -> None:
        responses, rc = run_mcp(["{not json", INITIALIZE])
        assert rc == 0
        assert responses[0]["error"]["code"] == -32700
        assert responses[0]["id"] is None
        assert responses[1]["id"] == 1

    def test_non_request_object_is_invalid_request(self) -> None:
        responses, _ = run_mcp(['["jsonrpc"]', '{"jsonrpc": "2.0", "id": 7}'])
        assert [r["error"]["code"] for r in responses] == [-32600, -32600]

    def test_call_without_tool_name_is_invalid_params(self) -> None:
        responses, _ = run_mcp([INITIALIZE, INITIALIZED, rpc(2, "tools/call", arguments={})])
        assert responses[1]["error"]["code"] == -32602
