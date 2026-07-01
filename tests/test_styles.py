"""Style & formatting tests: docx_style + docx_format (algorithms.md §16).

Covers list (effective-style in_use), define (closed prop set + id collision),
apply (pStyle splicing), the style_selector edit that touches styles.xml not the
paragraphs, direct anchor/range formatting, and the closed prop emission order.
"""

from __future__ import annotations

import base64

import pytest
from conftest import FIXTURE_PARTS, SECT_PR, build_docx, document_xml

from docxengine import (
    Session,
    ToolError,
    docx_format,
    docx_open,
    docx_read,
    docx_style,
    docx_validate,
    paragraph_anchor,
)

A1 = "P1#515a"
OLD_P2 = "The term is five (5) years from the Effective Date."
A2 = paragraph_anchor(2, OLD_P2)
H1_SEL = {"style": "Heading1"}

# A styles part with a basedOn chain so in_use can roll up through the cascade.
STYLES_CASCADE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1">'
    '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading2">'
    '<w:name w:val="heading 2"/><w:basedOn w:val="Heading1"/></w:style>'
    "</w:styles>"
)


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def open_docx(parts: dict[str, str] | None = None) -> tuple[Session, str]:
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())
    return session, str(result["doc_id"])


def main_xml(session: Session, doc_id: str) -> str:
    package = session.get(doc_id).package
    return package.part(package.main_document_part()).decode("utf-8")


def styles_xml(session: Session, doc_id: str) -> str:
    return session.get(doc_id).package.part("word/styles.xml").decode("utf-8")


# ---------------------------------------------------------------------------
# list (§16)
# ---------------------------------------------------------------------------


class TestList:
    def test_list_reports_id_name_type(self) -> None:
        session, doc_id = open_docx()
        styles = docx_style(session, doc_id=doc_id, op="list")["styles"]
        assert styles == [
            {"id": "Heading1", "name": "heading 1", "type": "paragraph", "in_use": 1}
        ]

    def test_in_use_rolls_up_the_based_on_cascade(self) -> None:
        # P1 uses Heading1 directly; its in_use counts toward Heading1 and Normal.
        parts = dict(FIXTURE_PARTS)
        parts["word/styles.xml"] = STYLES_CASCADE
        parts["word/document.xml"] = document_xml(
            '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>X</w:t></w:r></w:p>',
            "<w:p><w:r><w:t>plain</w:t></w:r></w:p>",
            SECT_PR,
        )
        session, doc_id = open_docx(parts)
        listed = docx_style(session, doc_id=doc_id, op="list")["styles"]
        styles = {s["id"]: s["in_use"] for s in listed}
        assert styles["Heading2"] == 1
        assert styles["Heading1"] == 1  # via Heading2's basedOn
        assert styles["Normal"] == 2  # Heading2 chain + the plain paragraph's default

    def test_based_on_key_present_only_when_set(self) -> None:
        parts = dict(FIXTURE_PARTS)
        parts["word/styles.xml"] = STYLES_CASCADE
        session, doc_id = open_docx(parts)
        styles = {s["id"]: s for s in docx_style(session, doc_id=doc_id, op="list")["styles"]}
        assert "based_on" not in styles["Normal"]
        assert styles["Heading2"]["based_on"] == "Heading1"


# ---------------------------------------------------------------------------
# define (§16)
# ---------------------------------------------------------------------------


class TestDefine:
    def test_define_matches_worked_example(self) -> None:
        session, doc_id = open_docx()
        result = docx_style(
            session,
            doc_id=doc_id,
            op="define",
            name="Clause",
            based_on="Heading1",
            props={"size": 11, "bold": True, "spacing_after": 6, "alignment": "justify"},
        )
        assert result["style_id"] == "Clause"
        xml = styles_xml(session, doc_id)
        assert (
            '<w:style w:type="paragraph" w:styleId="Clause">'
            '<w:name w:val="Clause"/><w:basedOn w:val="Heading1"/>'
            '<w:pPr><w:jc w:val="both"/><w:spacing w:after="120"/></w:pPr>'
            '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>'
        ) in xml

    def test_define_id_strips_whitespace(self) -> None:
        session, doc_id = open_docx()
        result = docx_style(session, doc_id=doc_id, op="define", name="My Clause")
        assert result["style_id"] == "MyClause"

    def test_define_id_collision_gets_suffix(self) -> None:
        session, doc_id = open_docx()
        # Heading1 already exists in the fixture → first reuse becomes Heading12,
        # the next collides with that and becomes Heading13 (§16 collision suffix).
        first = docx_style(session, doc_id=doc_id, op="define", name="Heading1")
        assert first["style_id"] == "Heading12"
        second = docx_style(session, doc_id=doc_id, op="define", name="Heading1")
        assert second["style_id"] == "Heading13"

    def test_defined_style_passes_validator(self) -> None:
        session, doc_id = open_docx()
        docx_style(session, doc_id=doc_id, op="define", name="Clause", props={"italic": True})
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# apply (§16)
# ---------------------------------------------------------------------------


class TestApply:
    def test_apply_splices_pstyle_first(self) -> None:
        session, doc_id = open_docx()
        docx_style(session, doc_id=doc_id, op="define", name="Clause")
        result = docx_style(session, doc_id=doc_id, op="apply", anchor=A2, style="Clause")
        assert result["new_anchor"] == A2  # text unchanged, hash stable
        xml = main_xml(session, doc_id)
        assert '<w:pPr><w:pStyle w:val="Clause"/></w:pPr>' in xml

    def test_apply_replaces_existing_pstyle(self) -> None:
        session, doc_id = open_docx()
        docx_style(session, doc_id=doc_id, op="define", name="Clause")
        docx_style(session, doc_id=doc_id, op="apply", anchor=A1, style="Clause")
        xml = main_xml(session, doc_id)
        assert '<w:pStyle w:val="Heading1"/>' not in xml
        assert '<w:pStyle w:val="Clause"/>' in xml

    def test_apply_resolves_style_by_name(self) -> None:
        session, doc_id = open_docx()
        # "heading 1" is the w:name of styleId Heading1 in the fixture.
        result = docx_style(session, doc_id=doc_id, op="apply", anchor=A2, style="heading 1")
        assert result["new_anchor"] == A2
        assert '<w:pStyle w:val="Heading1"/>' in main_xml(session, doc_id)

    def test_apply_unknown_style_is_style_unknown(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_style(session, doc_id=doc_id, op="apply", anchor=A2, style="Nonexistent")
        assert err.value.code == "style_unknown"


# ---------------------------------------------------------------------------
# docx_format — style_selector (§16)
# ---------------------------------------------------------------------------


class TestFormatStyleSelector:
    def test_style_selector_edits_styles_not_paragraphs(self) -> None:
        session, doc_id = open_docx()
        before_doc = main_xml(session, doc_id)
        result = docx_format(
            session,
            doc_id=doc_id,
            props={"color": "#1F4E79", "bold": True},
            style_selector={"style": "Heading1"},
        )
        assert result["affected"] == 0
        # The paragraph bytes are untouched; only styles.xml changed.
        assert main_xml(session, doc_id) == before_doc
        styles = styles_xml(session, doc_id)
        assert '<w:rPr><w:b/><w:color w:val="1F4E79"/></w:rPr>' in styles

    def test_style_selector_is_idempotent(self) -> None:
        session, doc_id = open_docx()
        docx_format(session, doc_id=doc_id, props={"bold": True}, style_selector=H1_SEL)
        once = styles_xml(session, doc_id)
        docx_format(session, doc_id=doc_id, props={"bold": True}, style_selector=H1_SEL)
        assert styles_xml(session, doc_id) == once

    def test_style_selector_replaces_same_named_prop(self) -> None:
        session, doc_id = open_docx()
        docx_format(session, doc_id=doc_id, props={"size": 10}, style_selector=H1_SEL)
        docx_format(session, doc_id=doc_id, props={"size": 14}, style_selector=H1_SEL)
        styles = styles_xml(session, doc_id)
        assert '<w:sz w:val="28"/>' in styles
        assert '<w:sz w:val="20"/>' not in styles


# ---------------------------------------------------------------------------
# docx_format — direct (§16)
# ---------------------------------------------------------------------------


class TestFormatDirect:
    def test_direct_anchor_merges_into_every_run(self) -> None:
        session, doc_id = open_docx()
        result = docx_format(session, doc_id=doc_id, props={"italic": True}, anchor=A1)
        assert result["affected"] == 1
        assert result["anchors"] == [A1]
        xml = main_xml(session, doc_id)
        # The bold run gains italic; the plain runs get a fresh rPr.
        assert "<w:rPr><w:b/><w:i/></w:rPr>" in xml
        assert xml.count("<w:i/>") == 3  # every run of P1

    def test_direct_range_formats_each_paragraph(self) -> None:
        session, doc_id = open_docx()
        result = docx_format(session, doc_id=doc_id, props={"bold": True}, range="P1..P2")
        assert result["affected"] == 2
        assert result["anchors"] == [A1, A2]

    def test_direct_alignment_splices_ppr(self) -> None:
        session, doc_id = open_docx()
        docx_format(session, doc_id=doc_id, props={"alignment": "center"}, anchor=A2)
        assert '<w:jc w:val="center"/>' in main_xml(session, doc_id)

    def test_bool_false_emits_toggle_off(self) -> None:
        session, doc_id = open_docx()
        docx_format(session, doc_id=doc_id, props={"bold": False}, anchor=A2)
        assert '<w:b w:val="0"/>' in main_xml(session, doc_id)

    def test_direct_format_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_format(session, doc_id=doc_id, props={"italic": True, "alignment": "right"}, anchor=A1)
        assert docx_validate(session, doc_id=doc_id)["valid"] is True

    def test_format_without_target_is_anchor_invalid(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_format(session, doc_id=doc_id, props={"bold": True})
        assert err.value.code == "anchor_invalid"


def test_format_projects_unchanged_text() -> None:
    # Formatting must not alter the projected text (only run/para props change).
    session, doc_id = open_docx()
    before = docx_read(session, doc_id=doc_id)["content"]
    docx_format(session, doc_id=doc_id, props={"bold": True}, anchor=A1)
    assert docx_read(session, doc_id=doc_id)["content"] == before
