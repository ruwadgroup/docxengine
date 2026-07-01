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
        assert len(tools) == 23  # facade surface: 24 spec tools minus docx_save
        assert all(set(t) == {"name", "description", "inputSchema"} for t in tools)
        by_name = {t["name"]: t for t in tools}
        assert "docx_save" not in by_name  # saving is automatic on every edit
        assert by_name["docx_open"]["inputSchema"]["properties"]["path"]["type"] == "string"
        # The path surface never exposes doc_id on any tool.
        assert not any(
            "doc_id" in t["inputSchema"].get("properties", {}) for t in tools
        )

        result = responses[2]["result"]
        assert result["isError"] is False
        assert [c["type"] for c in result["content"]] == ["text"]
        opened = json.loads(result["content"][0]["text"])
        assert "doc_id" not in opened  # the file IS the handle
        assert opened["path"] == str(docx_path)
        assert opened["n_paragraphs"] == 3

    def test_resources_templates_and_read_over_stdio(self, docx_path: Path) -> None:
        outline_uri = f"docx://{docx_path}/outline"
        projection_uri = f"docx://{docx_path}/projection"
        responses, rc = run_mcp(
            [
                INITIALIZE,
                INITIALIZED,
                rpc(2, "resources/list"),
                rpc(3, "resources/templates/list"),
                rpc(4, "resources/read", uri=outline_uri),
                rpc(5, "resources/read", uri=projection_uri),
            ]
        )
        assert rc == 0
        by_id = {r["id"]: r for r in responses}
        # The filesystem is not enumerated: resources/list is empty; reads are by path.
        assert by_id[2]["result"]["resources"] == []
        templates = by_id[3]["result"]["resourceTemplates"]
        assert {t["uriTemplate"] for t in templates} == {
            "docx://{path}/outline",
            "docx://{path}/projection",
        }

        outline = by_id[4]["result"]["contents"][0]
        assert outline["uri"] == outline_uri
        assert outline["mimeType"] == "text/markdown"
        assert "Master Services Agreement" in outline["text"]

        projection = by_id[5]["result"]["contents"][0]["text"]
        assert "H1] Master Services Agreement" in projection
        assert "The term is five (5) years" in projection

    def test_resources_read_missing_file_is_invalid_params(self, tmp_path: Path) -> None:
        missing = tmp_path / "nope.docx"
        responses, _ = run_mcp(
            [INITIALIZE, INITIALIZED, rpc(2, "resources/read", uri=f"docx://{missing}/outline")]
        )
        assert responses[1]["error"]["code"] == -32602
        assert "open_failed" in responses[1]["error"]["message"]

    def test_ping(self) -> None:
        responses, rc = run_mcp([INITIALIZE, INITIALIZED, rpc(2, "ping")])
        assert rc == 0
        assert responses[1] == {"jsonrpc": "2.0", "id": 2, "result": {}}

    def test_edits_persist_to_disk_across_calls(self, docx_path: Path) -> None:
        # Each call opens the file fresh: an edit in one call is on disk for the next,
        # with no doc_id threaded and no explicit save step.
        responses, rc = run_mcp(
            [
                INITIALIZE,
                INITIALIZED,
                rpc(
                    2,
                    "tools/call",
                    name="docx_replace",
                    arguments={
                        "path": str(docx_path),
                        "old": "five (5) years",
                        "new": "two (2) years",
                    },
                ),
                rpc(
                    3,
                    "tools/call",
                    name="docx_search",
                    arguments={"path": str(docx_path), "query": "two (2) years"},
                ),
            ]
        )
        assert rc == 0
        replaced = json.loads(responses[1]["result"]["content"][0]["text"])
        assert replaced["n_replaced"] == 1
        assert replaced["saved"] is True
        assert replaced["bytes"] == docx_path.stat().st_size
        # docx_search is text-first → markdown, not JSON (§26).
        found = responses[2]["result"]["content"][0]["text"]
        assert "two (2) years" in found  # a fresh open finds the persisted edit
        assert found.endswith("-->")  # the trailing metadata comment


class TestMarkdownProjection:
    """Text-first tools emit markdown over tools/call (§26); structured tools stay JSON."""

    def test_render_preview_and_convert_are_markdown(self, docx_path: Path) -> None:
        responses, rc = run_mcp(
            [
                INITIALIZE,
                INITIALIZED,
                rpc(
                    2,
                    "tools/call",
                    name="docx_render_preview",
                    arguments={"path": str(docx_path)},
                ),
                rpc(
                    3,
                    "tools/call",
                    name="docx_convert",
                    arguments={"path": str(docx_path), "to": "md"},
                ),
                rpc(4, "tools/call", name="docx_outline", arguments={"path": str(docx_path)}),
            ]
        )
        assert rc == 0
        by_id = {r["id"]: r["result"] for r in responses if "result" in r and "id" in r}

        preview = by_id[2]["content"][0]["text"]
        assert by_id[2]["isError"] is False
        assert preview.startswith("[P1#")  # the structural projection, not JSON
        assert "structural projection, no render adapter" in preview  # trailer
        assert preview.rstrip().endswith("-->")

        converted = by_id[3]["content"][0]["text"]
        assert "five (5) years" in converted  # body text, raw
        assert not converted.lstrip().startswith("{")  # raw markdown, not a JSON envelope

        outline = by_id[4]["content"][0]["text"]
        assert "# Master Services Agreement [P1#" in outline

    def test_structured_tool_stays_json(self, docx_path: Path) -> None:
        responses, _ = run_mcp(
            [
                INITIALIZE,
                INITIALIZED,
                rpc(2, "tools/call", name="docx_validate", arguments={"path": str(docx_path)}),
            ]
        )
        payload = json.loads(responses[1]["result"]["content"][0]["text"])  # parses as JSON
        assert "valid" in payload


class TestErrors:
    def test_tool_error_is_iserror_result_with_structured_json(self, tmp_path: Path) -> None:
        missing = tmp_path / "nope.docx"
        read = rpc(2, "tools/call", name="docx_read", arguments={"path": str(missing)})
        responses, _ = run_mcp([INITIALIZE, INITIALIZED, read])
        result = responses[1]["result"]
        assert result["isError"] is True
        payload = json.loads(result["content"][0]["text"])
        assert payload["error"] == "open_failed"
        assert str(missing) in payload["message"]
        assert payload["suggestions"]

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
