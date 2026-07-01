"""Dispatcher, spec-loading, and Document facade tests."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from conftest import build_docx

from docxengine import (
    Document,
    Session,
    ToolError,
    anthropic_tools,
    call,
    openai_tools,
    tool_schemas,
)
from docxengine._dispatch import MVP_TOOLS
from docxengine._spec import spec_tool_names
from docxengine.errors import ERROR_CODES


@pytest.fixture
def docx_path(docx_bytes: bytes, tmp_path: Path) -> Path:
    path = tmp_path / "fixture.docx"
    path.write_bytes(docx_bytes)
    return path


class TestDispatch:
    def test_routes_mvp_tool(self, docx_path: Path) -> None:
        session = Session()
        result = call("docx_open", {"path": str(docx_path)}, session=session)
        assert result["doc_id"] == "d1"
        outline = call("docx_outline", {"doc_id": "d1"}, session=session)
        assert isinstance(outline["outline"], list)

    def test_missing_required_arg_is_invalid_args(self) -> None:
        with pytest.raises(ToolError) as err:
            call("docx_read", {}, session=Session())
        assert err.value.code == "invalid_args"
        assert "doc_id" in err.value.message
        assert err.value.message.startswith("docx_read:")

    def test_multiple_missing_args_all_named(self) -> None:
        with pytest.raises(ToolError) as err:
            call("docx_replace", {"doc_id": "d1"}, session=Session())
        assert err.value.code == "invalid_args"
        assert "old" in err.value.message
        assert "new" in err.value.message

    def test_non_object_args_is_invalid_args(self) -> None:
        with pytest.raises(ToolError) as err:
            call("docx_read", ["d1"], session=Session())  # type: ignore[arg-type]
        assert err.value.code == "invalid_args"

    def test_every_spec_tool_is_implemented(self) -> None:
        # Phase 2 complete: nothing returns not_implemented for a spec tool.
        assert spec_tool_names() == MVP_TOOLS

    def test_unknown_tool_is_not_implemented(self) -> None:
        with pytest.raises(ToolError) as err:
            call("docx_frobnicate", {}, session=Session())
        assert err.value.code == "not_implemented"

    def test_every_spec_tool_dispatches(self, docx_path: Path) -> None:
        session = Session()
        call("docx_open", {"path": str(docx_path)}, session=session)
        # Every spec tool routes to a handler (missing-arg → invalid_args, never
        # not_implemented). docx_create/docx_template_fill take no doc_id.
        for tool in sorted(spec_tool_names()):
            try:
                call(tool, {}, session=session)
            except ToolError as err:
                assert err.code != "not_implemented", tool

    def test_extra_args_are_ignored(self, docx_path: Path) -> None:
        result = call(
            "docx_open", {"path": str(docx_path), "unexpected": True}, session=Session()
        )
        assert result["doc_id"] == "d1"

    def test_tool_errors_carry_spec_codes(self, docx_path: Path) -> None:
        session = Session()
        with pytest.raises(ToolError) as err:
            call("docx_read", {"doc_id": "d404"}, session=session)
        assert err.value.code == "doc_not_found"
        assert err.value.code in ERROR_CODES

    def test_default_session_persists_across_calls(self, docx_path: Path) -> None:
        opened = call("docx_open", {"path": str(docx_path)})
        read = call("docx_read", {"doc_id": opened["doc_id"]})
        assert "Master Services Agreement" in str(read["content"])


class TestSpecSurface:
    def test_tool_schemas_cover_all_spec_tools(self) -> None:
        schemas = tool_schemas()
        names = [schema["name"] for schema in schemas]
        assert names == sorted(names)
        assert set(names) == set(spec_tool_names())
        assert set(names) >= MVP_TOOLS
        for schema in schemas:
            assert isinstance(schema["description"], str)
            assert schema["input_schema"]["type"] == "object"

    def test_openai_tools_shape(self) -> None:
        tools = openai_tools()
        assert all(tool["type"] == "function" for tool in tools)
        by_name = {tool["function"]["name"]: tool["function"] for tool in tools}
        assert by_name["docx_replace"]["parameters"]["required"] == ["doc_id", "old", "new"]

    def test_anthropic_tools_shape(self) -> None:
        tools = anthropic_tools()
        by_name = {tool["name"]: tool for tool in tools}
        assert "doc_id" in by_name["docx_save"]["input_schema"]["properties"]
        assert by_name["docx_save"]["input_schema"]["required"] == ["doc_id", "path"]

    def test_returned_schemas_are_copies(self) -> None:
        first = anthropic_tools()
        first[0]["input_schema"]["properties"]["injected"] = {}
        assert "injected" not in anthropic_tools()[0]["input_schema"]["properties"]

    def test_error_codes_include_spec_set(self) -> None:
        assert {
            "anchor_stale",
            "doc_not_found",
            "invalid_args",
            "not_implemented",
            "validation_failed",
        } <= ERROR_CODES


class TestDocument:
    def test_open_paragraphs_find_replace_save(self, docx_path: Path, tmp_path: Path) -> None:
        doc = Document.open(docx_path)
        paragraphs = doc.paragraphs()
        assert paragraphs[0].anchor == "P1#515a"
        assert paragraphs[0].text == "Master Services Agreement"

        # search() returns match dicts; find() returns a Paragraph view (or None).
        hits = doc.search("five (5)")
        assert hits["matches"] and hits["matches"][0]["anchor"] == paragraphs[1].anchor
        para = doc.find("five (5)")
        assert para is not None and para.anchor == paragraphs[1].anchor

        assert not doc.dirty
        result = doc.replace("five (5) years", "three (3) years")
        assert result["n_replaced"] == 1
        assert doc.dirty

        out = tmp_path / "out.docx"
        saved = doc.save(out)
        assert saved["ok"] is True
        assert not doc.dirty
        with zipfile.ZipFile(out) as zf:
            assert b"three (3) years" in zf.read("word/document.xml")

    def test_document_from_bytes(self, docx_bytes: bytes) -> None:
        doc = Document(docx_bytes)
        assert len(doc.paragraphs()) == 3

    def test_save_runs_the_gate(self, tmp_path: Path) -> None:
        from conftest import DOCUMENT_XML, FIXTURE_PARTS

        parts = dict(FIXTURE_PARTS)
        parts["word/document.xml"] = DOCUMENT_XML.replace('w:ins w:id="2"', 'w:ins w:id="1"')
        doc = Document(build_docx(parts))
        with pytest.raises(ToolError) as err:
            doc.save(tmp_path / "out.docx")
        assert err.value.code == "validation_failed"
