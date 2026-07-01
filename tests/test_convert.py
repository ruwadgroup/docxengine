"""docx_convert tests (algorithms.md §23/§23a).

md is produced from the §2 projection: headings, tight lists, GitHub tables, inline
bold/italic, comments as `<!-- comment:… -->`, revisions in accepted view with
`[ins]`/`[del]`. html emits semantic tags with inline alignment/color styles. The
create(md)→convert(md) round-trip preserves heading/list/table structure. pdf/png
without a render adapter raise render_unavailable.
"""

from __future__ import annotations

import base64

import pytest
from conftest import build_docx

from docxengine import (
    Session,
    ToolError,
    docx_comment,
    docx_convert,
    docx_create,
    docx_open,
)


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)
    # Ensure no real soffice is picked up by the pdf/png path tests.
    monkeypatch.delenv("DOCXENGINE_SOFFICE", raising=False)


def open_docx(parts: dict[str, str] | None = None) -> tuple[Session, str]:
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())
    return session, str(result["doc_id"])


# ---------------------------------------------------------------------------
# md round-trip (create → convert)
# ---------------------------------------------------------------------------


class TestMarkdownRoundTrip:
    def test_headings_preserved(self) -> None:
        session = Session()
        md = "# Title\n\n## Section\n\nBody paragraph."
        doc_id = str(docx_create(session, content_md=md)["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="md")["content"]
        assert "# Title" in out
        assert "## Section" in out
        assert "Body paragraph." in out

    def test_lists_are_tight(self) -> None:
        session = Session()
        md = "- one\n- two\n- three"
        doc_id = str(docx_create(session, content_md=md)["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="md")["content"]
        # Tight list: single newline between items.
        assert out == "- one\n- two\n- three"

    def test_ordered_list(self) -> None:
        session = Session()
        md = "1. alpha\n2. beta"
        doc_id = str(docx_create(session, content_md=md)["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="md")["content"]
        assert out == "1. alpha\n1. beta"

    def test_table_structure_preserved(self) -> None:
        session = Session()
        md = "| Term | Value |\n| --- | --- |\n| Fee | $100 |"
        doc_id = str(docx_create(session, content_md=md)["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="md")["content"]
        assert "| Term | Value |" in out
        assert "| --- | --- |" in out
        assert "| Fee | $100 |" in out

    def test_inline_bold_reconstructed(self) -> None:
        session = Session()
        md = "See **clause** here."
        doc_id = str(docx_create(session, content_md=md)["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="md")["content"]
        assert "**clause**" in out

    def test_block_join_blank_line(self) -> None:
        session = Session()
        md = "# H\n\nbody"
        doc_id = str(docx_create(session, content_md=md)["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="md")["content"]
        assert out == "# H\n\nbody"


# ---------------------------------------------------------------------------
# Revisions + comments annotation
# ---------------------------------------------------------------------------


class TestAnnotations:
    def test_tracked_changes_accepted_view_with_markers(self) -> None:
        session, doc_id = open_docx()
        # PARA_TRACKED: "Payment due in [del 30][ins 45] days"
        out = docx_convert(session, doc_id=doc_id, to="md")["content"]
        assert "[ins]45" in out
        assert "[del]" in out
        assert "30" not in out  # delText omitted in accepted view

    def test_comment_inline_note(self) -> None:
        session, doc_id = open_docx()
        anchor = "P1#515a"
        docx_comment(
            session, doc_id=doc_id, op="add", anchor=anchor, text="Confirm scope", author="Jane"
        )
        out = docx_convert(session, doc_id=doc_id, to="md")["content"]
        assert "<!-- comment:Jane: Confirm scope -->" in out


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------


class TestHtml:
    def test_html_semantic_tags(self) -> None:
        session = Session()
        md = "# Head\n\nbody **x**\n\n- a\n- b"
        doc_id = str(docx_create(session, content_md=md)["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="html")["content"]
        assert "<h1>Head</h1>" in out
        assert "<strong>x</strong>" in out
        assert "<ul>" in out and "<li>a</li>" in out and "</ul>" in out

    def test_html_table(self) -> None:
        session = Session()
        md = "| A | B |\n| --- | --- |\n| 1 | 2 |"
        doc_id = str(docx_create(session, content_md=md)["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="html")["content"]
        assert "<table><tr><td>A</td><td>B</td></tr>" in out

    def test_html_escapes_text(self) -> None:
        session = Session()
        doc_id = str(docx_create(session, content_md="a < b & c")["doc_id"])
        out = docx_convert(session, doc_id=doc_id, to="html")["content"]
        assert "a &lt; b &amp; c" in out


# ---------------------------------------------------------------------------
# Errors / render targets
# ---------------------------------------------------------------------------


class TestErrors:
    def test_unsupported_format(self) -> None:
        session = Session()
        doc_id = str(docx_create(session, content_md="x")["doc_id"])
        with pytest.raises(ToolError) as exc:
            docx_convert(session, doc_id=doc_id, to="rtf")
        assert exc.value.code == "unsupported_format"

    def test_pdf_without_adapter_is_render_unavailable(self, tmp_path, monkeypatch) -> None:  # noqa: ANN001
        # Force detection failure: empty PATH and no platform default present here.
        monkeypatch.setenv("PATH", str(tmp_path))
        session = Session()
        doc_id = str(docx_create(session, content_md="x")["doc_id"])
        from docxengine import _render

        monkeypatch.setattr(_render, "PLATFORM_DEFAULTS", ())
        with pytest.raises(ToolError) as exc:
            docx_convert(session, doc_id=doc_id, to="pdf", path=str(tmp_path / "o.pdf"))
        assert exc.value.code == "render_unavailable"

    def test_md_note_counts_annotations(self) -> None:
        session, doc_id = open_docx()
        note = docx_convert(session, doc_id=doc_id, to="md")["note"]
        # PARA_TRACKED carries one del + one ins.
        assert "tracked change" in note
