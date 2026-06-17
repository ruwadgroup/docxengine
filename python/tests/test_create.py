"""docx_create tests (algorithms.md §22/§23a).

A created document passes the §8 validator, reopens with the right outline/anchors,
ships the deterministic skeleton (Normal, Heading1-6, ListParagraph, TableGrid, Quote),
honors DOCXENGINE_FIXED_DATE in docProps/core.xml, and allocates numbering.xml only
when a list item exists. Inline markdown becomes the §22 run sequence.
"""

from __future__ import annotations

import zipfile

import pytest

from docxengine import (
    Session,
    ToolError,
    docx_create,
    docx_outline,
    docx_validate,
)
from docxengine._create import parse_inline


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def _part(session: Session, doc_id: str, name: str) -> str:
    pkg = session.get(doc_id).package
    return pkg.part(name).decode("utf-8")


def _create(session: Session, md: str) -> str:
    return str(docx_create(session, content_md=md)["doc_id"])


# ---------------------------------------------------------------------------
# Validation + outline
# ---------------------------------------------------------------------------


class TestCreateValidates:
    def test_created_doc_passes_validator(self) -> None:
        session = Session()
        doc_id = _create(session, "# Title\n\nBody text.")
        result = docx_validate(session, doc_id=doc_id)
        assert result["valid"] is True
        assert result["issues"] == []

    def test_n_paragraphs_counts_body_paragraphs(self) -> None:
        session = Session()
        result = docx_create(session, content_md="# A\n\nB\n\nC")
        assert result["n_paragraphs"] == 3

    def test_table_excluded_from_paragraph_count(self) -> None:
        session = Session()
        md = "# Head\n\n| A | B |\n| --- | --- |\n| 1 | 2 |"
        result = docx_create(session, content_md=md)
        # One heading paragraph; the table is not a paragraph.
        assert result["n_paragraphs"] == 1

    def test_reopens_with_outline(self) -> None:
        session = Session()
        md = "# Master Agreement\n\n## Scope\n\nClause one.\n\n## Term\n\nClause two."
        doc_id = _create(session, md)
        outline = docx_outline(session, doc_id=doc_id)["outline"]
        levels = [(o["level"], o["text"]) for o in outline]
        assert levels == [
            (1, "Master Agreement"),
            (2, "Scope"),
            (2, "Term"),
        ]

    def test_table_appears_in_outline(self) -> None:
        session = Session()
        md = "# Head\n\n| Term | Value |\n| --- | --- |\n| Fee | $100 |"
        doc_id = _create(session, md)
        tables = docx_outline(session, doc_id=doc_id)["tables"]
        assert len(tables) == 1
        assert tables[0]["dims"] == "2×2"


# ---------------------------------------------------------------------------
# Skeleton parts
# ---------------------------------------------------------------------------


class TestSkeleton:
    def test_ships_base_styles(self) -> None:
        session = Session()
        doc_id = _create(session, "Plain text.")
        styles = _part(session, doc_id, "word/styles.xml")
        for sid in ("Normal", "Heading1", "Heading6", "ListParagraph", "TableGrid", "Quote"):
            assert f'w:styleId="{sid}"' in styles

    def test_trailing_sectpr_with_a4_geometry(self) -> None:
        session = Session()
        doc_id = _create(session, "Body.")
        doc = _part(session, doc_id, "word/document.xml")
        assert '<w:pgSz w:w="11906" w:h="16838"/>' in doc
        assert '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"' in doc

    def test_bare_body_paragraph_has_no_spacing(self) -> None:
        session = Session()
        doc_id = _create(session, "Just text.")
        doc = _part(session, doc_id, "word/document.xml")
        assert "<w:spacing" not in doc

    def test_numbering_only_with_lists(self) -> None:
        session = Session()
        plain = _create(session, "No lists here.")
        assert not session.get(plain).package.has_part("word/numbering.xml")
        listed = _create(session, "- one\n- two")
        assert session.get(listed).package.has_part("word/numbering.xml")

    def test_list_numids_start_at_one(self) -> None:
        session = Session()
        doc_id = _create(session, "1. a\n2. b\n- x\n- y")
        numbering = _part(session, doc_id, "word/numbering.xml")
        # ol allocated first (numId 1), ul second (numId 2) in first-use order.
        assert '<w:num w:numId="1">' in numbering
        assert '<w:num w:numId="2">' in numbering


# ---------------------------------------------------------------------------
# Deterministic dates
# ---------------------------------------------------------------------------


class TestDeterministicDates:
    def test_fixed_date_in_core_props(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", "2026-06-11T00:00:00Z")
        session = Session()
        doc_id = _create(session, "Body.")
        core = _part(session, doc_id, "docProps/core.xml")
        fixed = "2026-06-11T00:00:00Z"
        assert f'<dcterms:created xsi:type="dcterms:W3CDTF">{fixed}</dcterms:created>' in core
        assert f"{fixed}</dcterms:modified>" in core

    def test_default_date_when_unset(self) -> None:
        session = Session()
        doc_id = _create(session, "Body.")
        core = _part(session, doc_id, "docProps/core.xml")
        assert "2026-01-01T00:00:00Z" in core


# ---------------------------------------------------------------------------
# Inline markdown (§22)
# ---------------------------------------------------------------------------


class TestInline:
    def test_bold_italic_code_runs(self) -> None:
        session = Session()
        doc_id = _create(session, "See **clause** `4a` and *more*.")
        doc = _part(session, doc_id, "word/document.xml")
        assert "<w:rPr><w:b/></w:rPr><w:t>clause</w:t>" in doc
        assert '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>4a</w:t>' in doc
        assert "<w:rPr><w:i/></w:rPr><w:t>more</w:t>" in doc

    def test_parse_inline_unmatched_marker_is_literal(self) -> None:
        runs = parse_inline("a *b")
        assert [r.text for r in runs] == ["a *b"]
        assert all(not r.italic for r in runs)

    def test_escaping_special_chars(self) -> None:
        session = Session()
        doc_id = _create(session, "Tom & Jerry < Bob")
        doc = _part(session, doc_id, "word/document.xml")
        assert "Tom &amp; Jerry &lt; Bob" in doc

    def test_horizontal_rule(self) -> None:
        session = Session()
        doc_id = _create(session, "---")
        doc = _part(session, doc_id, "word/document.xml")
        bdr = '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>'
        assert bdr in doc

    def test_task_list_glyph_prefixed_listparagraph(self) -> None:
        session = Session()
        doc_id = _create(session, "- [ ] todo\n- [x] done\n* [X] also done")
        doc = _part(session, doc_id, "word/document.xml")
        assert (
            '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>'
            "<w:r><w:t>☐ todo</w:t></w:r></w:p>" in doc
        )
        assert "<w:t>☒ done</w:t>" in doc
        assert "<w:t>☒ also done</w:t>" in doc
        # The checkbox replaces the bullet — task items allocate no numbering.
        assert not session.get(doc_id).package.has_part("word/numbering.xml")


# ---------------------------------------------------------------------------
# Spec / errors
# ---------------------------------------------------------------------------


class TestArgsAndSpec:
    def test_both_content_and_spec_is_invalid(self) -> None:
        session = Session()
        with pytest.raises(ToolError) as exc:
            docx_create(session, content_md="x", spec={"blocks": []})
        assert exc.value.code == "invalid_args"

    def test_spec_lowers_to_markdown(self) -> None:
        session = Session()
        spec = {
            "blocks": [
                {"type": "heading", "level": 2, "text": "Scope"},
                {"type": "paragraph", "text": "Body"},
            ]
        }
        result = docx_create(session, spec=spec)
        outline = docx_outline(session, doc_id=str(result["doc_id"]))["outline"]
        assert [(o["level"], o["text"]) for o in outline] == [(2, "Scope")]

    def test_empty_create(self) -> None:
        session = Session()
        result = docx_create(session)
        assert result["n_paragraphs"] == 0
        assert docx_validate(session, doc_id=str(result["doc_id"]))["valid"] is True


# ---------------------------------------------------------------------------
# Package integrity
# ---------------------------------------------------------------------------


def test_created_package_entry_order(tmp_path) -> None:  # noqa: ANN001
    session = Session()
    doc_id = _create(session, "- item")
    out = tmp_path / "created.docx"
    from docxengine import docx_save

    docx_save(session, doc_id=doc_id, path=str(out))
    with zipfile.ZipFile(out) as zf:
        names = zf.namelist()
    assert names == [
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
        "[Content_Types].xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
        "docProps/core.xml",
    ]
