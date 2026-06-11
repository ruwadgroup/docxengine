"""Unit tests for the MCP path facade (``_mcp_facade``).

These exercise the facade directly (no transport): the schema transform that
projects the doc_id contract onto a path surface, and the per-call
open -> run -> validate -> save-back lifecycle, including the sandbox and the
refusal-leaves-file-untouched guarantee.
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from pathlib import Path

import pytest
from conftest import DOCUMENT_RELS_XML, FIXTURE_PARTS

from docxengine._errors import ToolError
from docxengine._mcp_facade import FacadeContext, call_path_tool, facade_tool_schemas


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_docx(path: Path, parts: dict[str, str]) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, text in parts.items():
            zf.writestr(name, text.encode("utf-8"))
    path.write_bytes(buf.getvalue())


@pytest.fixture
def docx_file(tmp_path: Path) -> Path:
    path = tmp_path / "fixture.docx"
    _write_docx(path, FIXTURE_PARTS)
    return path


@pytest.fixture
def corrupt_docx_file(tmp_path: Path) -> Path:
    """A package that opens cleanly but fails validation (orphan relationship)."""
    orphan = DOCUMENT_RELS_XML.replace(
        "</Relationships>",
        '<Relationship Id="rId9" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/missing.png"/></Relationships>',
    )
    parts = {**FIXTURE_PARTS, "word/_rels/document.xml.rels": orphan}
    path = tmp_path / "corrupt.docx"
    _write_docx(path, parts)
    return path


@pytest.fixture
def ctx() -> FacadeContext:
    return FacadeContext(None)


# ---------------------------------------------------------------------------
# Schema transform
# ---------------------------------------------------------------------------


def test_facade_surface_is_23_path_tools() -> None:
    schemas = facade_tool_schemas()
    names = {s["name"] for s in schemas}
    assert len(schemas) == 23
    assert "docx_save" not in names  # saving is folded into every edit
    # No tool exposes doc_id; the file path is the handle.
    assert not any("doc_id" in s["input_schema"].get("properties", {}) for s in schemas)


def test_per_tool_transforms() -> None:
    by_name = {s["name"]: s["input_schema"] for s in facade_tool_schemas()}

    assert by_name["docx_create"]["required"] == ["path"]
    assert "path" in by_name["docx_create"]["properties"]

    # docx_open is a stateless inspect: path required, base64 bytes dropped.
    assert by_name["docx_open"]["required"] == ["path"]
    assert "bytes" not in by_name["docx_open"]["properties"]

    # convert/media: source is `path`, the file-producing output is `output_path`.
    assert by_name["docx_convert"]["required"] == ["path", "to"]
    assert "output_path" in by_name["docx_convert"]["properties"]
    assert by_name["docx_media"]["required"] == ["path", "op"]
    assert "output_path" in by_name["docx_media"]["properties"]

    # template_fill: template in, filled document written to `path`.
    assert set(by_name["docx_template_fill"]["required"]) == {"template", "path", "data"}

    # a representative generic edit tool
    assert by_name["docx_replace"]["required"] == ["path", "old", "new"]


def test_descriptions_drop_stale_handle_references() -> None:
    for schema in facade_tool_schemas():
        desc = schema["description"]
        assert "doc_id" not in desc
        assert "docx_save" not in desc


# ---------------------------------------------------------------------------
# Dispatch guards
# ---------------------------------------------------------------------------


def test_docx_save_is_rejected(ctx: FacadeContext, docx_file: Path) -> None:
    with pytest.raises(ToolError) as exc:
        call_path_tool("docx_save", {"path": str(docx_file)}, ctx)
    assert exc.value.code == "invalid_args"
    assert "automatically" in exc.value.message


def test_unknown_tool_is_rejected(ctx: FacadeContext) -> None:
    with pytest.raises(ToolError) as exc:
        call_path_tool("docx_frobnicate", {"path": "x.docx"}, ctx)
    assert exc.value.code == "invalid_args"


def test_missing_path_is_invalid_args(ctx: FacadeContext) -> None:
    with pytest.raises(ToolError) as exc:
        call_path_tool("docx_outline", {}, ctx)
    assert exc.value.code == "invalid_args"


# ---------------------------------------------------------------------------
# Lifecycle: create / edit persist; reads don't
# ---------------------------------------------------------------------------


def test_create_writes_file_and_omits_doc_id(ctx: FacadeContext, tmp_path: Path) -> None:
    out = tmp_path / "new.docx"
    result = call_path_tool("docx_create", {"path": str(out), "content_md": "# Hi\n\nBody."}, ctx)
    assert out.exists()
    assert "doc_id" not in result
    assert result["saved"] is True
    assert result["bytes"] == out.stat().st_size
    assert result["path"] == str(out)


def test_edit_persists_and_reports_saved(ctx: FacadeContext, docx_file: Path) -> None:
    before = _sha(docx_file)
    result = call_path_tool(
        "docx_replace",
        {"path": str(docx_file), "old": "five (5) years", "new": "two (2) years"},
        ctx,
    )
    assert result["saved"] is True
    assert "doc_id" not in result
    assert _sha(docx_file) != before  # the edit reached disk


def test_read_only_tool_does_not_touch_file(ctx: FacadeContext, docx_file: Path) -> None:
    before = _sha(docx_file)
    out = call_path_tool("docx_outline", {"path": str(docx_file)}, ctx)
    assert "saved" not in out
    assert _sha(docx_file) == before  # outline never writes


def test_convert_md_returns_content(ctx: FacadeContext, docx_file: Path) -> None:
    result = call_path_tool("docx_convert", {"path": str(docx_file), "to": "md"}, ctx)
    assert "content" in result


def test_render_preview_uri_uses_path(ctx: FacadeContext, docx_file: Path) -> None:
    result = call_path_tool("docx_render_preview", {"path": str(docx_file)}, ctx)
    pages = result["pages"]
    assert pages  # page numbers are present even in the structural fallback
    # Any image link (only when a renderer ran) is rewritten to use the path, not a doc_id.
    assert all(
        "image" not in page or page["image"].startswith(f"docx://{docx_file}/preview/")
        for page in pages
    )


# ---------------------------------------------------------------------------
# Path resolution + sandbox
# ---------------------------------------------------------------------------


def test_relative_path_resolves_against_cwd(
    ctx: FacadeContext, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.chdir(tmp_path)
    result = call_path_tool("docx_create", {"path": "rel.docx", "content_md": "# Rel"}, ctx)
    assert (tmp_path / "rel.docx").exists()
    assert result["path"] == "rel.docx"  # the user-supplied form is echoed back


def test_sandbox_denies_escape_and_allows_inside(tmp_path: Path) -> None:
    root = tmp_path / "root"
    root.mkdir()
    sctx = FacadeContext(root.resolve())

    with pytest.raises(ToolError) as exc:
        call_path_tool("docx_create", {"path": "../escape.docx", "content_md": "# x"}, sctx)
    assert exc.value.code == "path_denied"

    call_path_tool("docx_create", {"path": "inside.docx", "content_md": "# ok"}, sctx)
    assert (root / "inside.docx").exists()


# ---------------------------------------------------------------------------
# Write-back refusal leaves the file byte-for-byte untouched
# ---------------------------------------------------------------------------


def test_write_back_refusal_leaves_file_untouched(
    ctx: FacadeContext, corrupt_docx_file: Path
) -> None:
    before = _sha(corrupt_docx_file)
    with pytest.raises(ToolError) as exc:
        call_path_tool(
            "docx_replace",
            {"path": str(corrupt_docx_file), "old": "five (5) years", "new": "two (2) years"},
            ctx,
        )
    assert exc.value.code == "validation_failed"
    assert any("not modified" in s for s in exc.value.suggestions)
    assert _sha(corrupt_docx_file) == before  # the refused write changed nothing
