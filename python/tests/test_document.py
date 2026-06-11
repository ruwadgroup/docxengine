"""Native ``Document`` API: full-coverage parity with the tool surface."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from conftest import (
    CONTENT_TYPES_XML,
    DOCUMENT_RELS_XML,
    DOCUMENT_XML,
    FIXTURE_PARTS,
    ROOT_RELS_XML,
    STYLES_XML,
    build_docx,
    document_xml,
)

from docxengine import Document, Paragraph, Session, ToolError

# Every Document method maps to a contract tool (or a native read/persist helper).
_TOOL_METHODS = [
    "outline",
    "read",
    "search",
    "replace",
    "edit_paragraph",
    "insert",
    "delete",
    "revision",
    "comment",
    "table",
    "style",
    "format",
    "list",
    "section",
    "media",
    "field",
    "validate",
    "repair",
    "convert",
    "render_preview",
    "save",
    "to_bytes",
]


@pytest.fixture
def docx_path(docx_bytes: bytes, tmp_path: Path) -> Path:
    path = tmp_path / "fixture.docx"
    path.write_bytes(docx_bytes)
    return path


def _template_bytes() -> bytes:
    body = document_xml('<w:p><w:r><w:t>Hello {{name}}</w:t></w:r></w:p>')
    return build_docx(
        {
            "[Content_Types].xml": CONTENT_TYPES_XML,
            "_rels/.rels": ROOT_RELS_XML,
            "word/document.xml": body,
            "word/_rels/document.xml.rels": DOCUMENT_RELS_XML,
            "word/styles.xml": STYLES_XML,
        }
    )


def test_every_method_exists() -> None:
    for name in _TOOL_METHODS:
        assert callable(getattr(Document, name)), name
    for ctor in ("open", "create", "fill_template", "attach"):
        assert callable(getattr(Document, ctor)), ctor


def test_open_from_path_and_bytes(docx_path: Path, docx_bytes: bytes) -> None:
    assert len(Document.open(docx_path).paragraphs()) == 3
    assert len(Document.open(docx_bytes).paragraphs()) == 3
    assert len(Document(docx_bytes).paragraphs()) == 3  # bare constructor still opens


def test_paragraphs_carry_style_id(docx_bytes: bytes) -> None:
    paras = Document.open(docx_bytes).paragraphs()
    assert paras[0].anchor == "P1#515a"
    assert paras[0].text == "Master Services Agreement"
    assert paras[0].style == "Heading1"  # the w:pStyle styleId, mirroring JS
    assert paras[1].style is None


def test_find_returns_paragraph_search_returns_matches(docx_bytes: bytes) -> None:
    doc = Document.open(docx_bytes)
    para = doc.find("five (5)")
    assert isinstance(para, Paragraph)
    assert doc.find("nonexistent text") is None
    hits = doc.search("five (5)")
    assert hits["matches"][0]["anchor"] == para.anchor


def test_create_then_edit_roundtrips() -> None:
    doc = Document.create(content_md="# Title\n\nThe term is five (5) years.")
    assert not doc.dirty
    doc.replace("five (5) years", "three (3) years")
    assert doc.dirty
    assert doc.find("three (3) years") is not None


def test_edit_surface_methods(docx_bytes: bytes) -> None:
    doc = Document.open(docx_bytes)
    anchor = doc.paragraphs()[0].anchor
    assert doc.insert("New intro.", after=anchor)["new_anchors"]
    assert doc.revision("list")["revisions"]  # P3 carries tracked changes
    # comment + table both round-trip against a paragraph anchor
    assert doc.comment("add", anchor=anchor, text="note", author="QA")
    assert doc.table("create", after=anchor, rows=2, cols=2, data=[["a", "b"], ["c", "d"]])
    assert doc.validate()["valid"] is True


def test_paragraph_primitives(docx_bytes: bytes) -> None:
    doc = Document.open(docx_bytes)
    para = doc.find("five (5)")
    assert para is not None
    para.replace("five (5) years", "two (2) years")
    assert doc.find("two (2) years") is not None
    # the held view is now stale; re-fetch and use the other primitives
    fresh = doc.paragraphs()[0]
    assert fresh.insert_after("After the heading.")["new_anchors"]


def test_convert_md(docx_bytes: bytes) -> None:
    assert "content" in Document.open(docx_bytes).convert("md")


def test_save_runs_validation_gate(tmp_path: Path) -> None:
    parts = dict(FIXTURE_PARTS)
    parts["word/document.xml"] = DOCUMENT_XML.replace('w:ins w:id="2"', 'w:ins w:id="1"')
    doc = Document(build_docx(parts))
    with pytest.raises(ToolError) as err:
        doc.save(tmp_path / "out.docx")
    assert err.value.code == "validation_failed"


def test_save_then_clean(docx_bytes: bytes, tmp_path: Path) -> None:
    doc = Document.open(docx_bytes)
    doc.replace("five (5) years", "three (3) years")
    out = tmp_path / "out.docx"
    saved = doc.save(out)
    assert saved["ok"] is True
    assert not doc.dirty
    with zipfile.ZipFile(out) as zf:
        assert b"three (3) years" in zf.read("word/document.xml")


def test_to_bytes_roundtrips_and_gates(docx_bytes: bytes) -> None:
    doc = Document.open(docx_bytes)
    doc.replace("five (5) years", "three (3) years")
    data = doc.to_bytes()
    assert data[:2] == b"PK"
    assert not doc.dirty  # export marks the document saved
    assert Document.open(data).find("three (3) years") is not None

    parts = dict(FIXTURE_PARTS)
    parts["word/document.xml"] = DOCUMENT_XML.replace('w:ins w:id="2"', 'w:ins w:id="1"')
    with pytest.raises(ToolError) as err:
        Document(build_docx(parts)).to_bytes()
    assert err.value.code == "validation_failed"


def test_fill_template_from_path_and_bytes(tmp_path: Path) -> None:
    from_bytes = Document.fill_template(_template_bytes(), {"name": "World"})
    assert from_bytes.find("Hello World") is not None

    template = tmp_path / "tpl.docx"
    template.write_bytes(_template_bytes())
    from_path = Document.fill_template(template, {"name": "Ada"})
    assert from_path.find("Hello Ada") is not None


def test_fill_template_strict_raises() -> None:
    with pytest.raises(ToolError) as err:
        Document.fill_template(_template_bytes(), {}, strict=True)
    assert err.value.code == "placeholder_unfilled"


def test_attach_shares_session(docx_bytes: bytes) -> None:
    session = Session()
    a = Document.open(docx_bytes, session=session)
    b = Document.attach(session, a.doc_id)
    assert b.doc_id == a.doc_id
    assert b.session is session
    a.replace("five (5) years", "three (3) years")
    assert b.dirty  # same underlying handle
