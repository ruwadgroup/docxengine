"""List & numbering tests: docx_list (algorithms.md §17).

Covers create (ol/ul, multi-level) with numbering.xml creation + its content-type/
rels wiring, restart (new num, same abstractNum, lvlOverride), set_level, convert
(paragraphs ↔ list items), and the projection annotations the numbering drives.
"""

from __future__ import annotations

import base64

import pytest
from conftest import build_docx

from docxengine import (
    Session,
    ToolError,
    docx_list,
    docx_open,
    docx_read,
    docx_validate,
    paragraph_anchor,
)

A1 = "P1#515a"
OLD_P2 = "The term is five (5) years from the Effective Date."
A2 = paragraph_anchor(2, OLD_P2)


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def open_docx(parts: dict[str, str] | None = None) -> tuple[Session, str]:
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())
    return session, str(result["doc_id"])


def numbering_xml(session: Session, doc_id: str) -> str:
    return session.get(doc_id).package.part("word/numbering.xml").decode("utf-8")


def main_xml(session: Session, doc_id: str) -> str:
    package = session.get(doc_id).package
    return package.part(package.main_document_part()).decode("utf-8")


# ---------------------------------------------------------------------------
# create (§17)
# ---------------------------------------------------------------------------


class TestCreate:
    def test_create_makes_numbering_part_with_content_type_and_rel(self) -> None:
        session, doc_id = open_docx()
        docx_list(
            session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "First"}]
        )
        package = session.get(doc_id).package
        assert package.has_part("word/numbering.xml")
        ct = package.part("[Content_Types].xml").decode()
        assert "wordprocessingml.numbering+xml" in ct
        rels = package.part("word/_rels/document.xml.rels").decode()
        assert "relationships/numbering" in rels
        assert "Target=\"numbering.xml\"" in rels

    def test_ol_abstract_num_cascade(self) -> None:
        session, doc_id = open_docx()
        docx_list(session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "x"}])
        nb = numbering_xml(session, doc_id)
        assert (
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val="%1."/>'
            '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
        ) in nb
        assert (
            '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>'
            '<w:lvlText w:val="%2."/>'
            '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
        ) in nb

    def test_ul_uses_bullet_numfmt(self) -> None:
        session, doc_id = open_docx()
        docx_list(session, doc_id=doc_id, op="create", after=A2, kind="ul", items=[{"text": "x"}])
        nb = numbering_xml(session, doc_id)
        assert '<w:numFmt w:val="bullet"/>' in nb
        assert '<w:lvlText w:val="•"/>' in nb

    def test_num_points_at_abstract(self) -> None:
        session, doc_id = open_docx()
        docx_list(session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "x"}])
        nb = numbering_xml(session, doc_id)
        assert '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>' in nb

    def test_items_get_numpr_and_liststyle(self) -> None:
        session, doc_id = open_docx()
        docx_list(
            session,
            doc_id=doc_id,
            op="create",
            after=A2,
            kind="ol",
            items=[{"text": "First"}, {"text": "Sub", "level": 1}],
        )
        xml = main_xml(session, doc_id)
        assert '<w:pStyle w:val="ListParagraph"/>' in xml
        assert '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' in xml
        assert '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>' in xml

    def test_create_projects_list_annotations(self) -> None:
        session, doc_id = open_docx()
        result = docx_list(
            session,
            doc_id=doc_id,
            op="create",
            after=A2,
            kind="ol",
            items=[{"text": "First"}, {"text": "Sub", "level": 1}, {"text": "Second"}],
        )
        assert result["n_affected"] == 3
        content = docx_read(session, doc_id=doc_id)["content"]
        assert "List:ol L1] First" in content
        assert "List:ol L2] Sub" in content
        assert "List:ol L1] Second" in content

    def test_created_list_passes_validator(self) -> None:
        session, doc_id = open_docx()
        docx_list(session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "x"}])
        assert docx_validate(session, doc_id=doc_id)["valid"] is True

    def test_list_paragraph_style_ensured(self) -> None:
        session, doc_id = open_docx()
        docx_list(session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "x"}])
        styles = session.get(doc_id).package.part("word/styles.xml")
        assert b'w:styleId="ListParagraph"' in styles


# ---------------------------------------------------------------------------
# restart (§17)
# ---------------------------------------------------------------------------


class TestRestart:
    def test_restart_allocates_new_num_same_abstract(self) -> None:
        session, doc_id = open_docx()
        anchors = docx_list(
            session,
            doc_id=doc_id,
            op="create",
            after=A2,
            kind="ol",
            items=[{"text": "a"}, {"text": "b"}],
        )["new_anchors"]
        docx_list(session, doc_id=doc_id, op="restart", anchor=anchors[1], at=5)
        nb = numbering_xml(session, doc_id)
        assert (
            '<w:num w:numId="2"><w:abstractNumId w:val="1"/>'
            '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>'
        ) in nb
        # The target paragraph now points at the new numId.
        assert '<w:numId w:val="2"/>' in main_xml(session, doc_id)

    def test_restart_on_non_list_is_anchor_invalid(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_list(session, doc_id=doc_id, op="restart", anchor=A1, at=1)
        assert err.value.code == "anchor_invalid"


# ---------------------------------------------------------------------------
# set_level (§17)
# ---------------------------------------------------------------------------


class TestSetLevel:
    def test_set_level_rewrites_ilvl(self) -> None:
        session, doc_id = open_docx()
        anchors = docx_list(
            session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "a"}]
        )["new_anchors"]
        docx_list(session, doc_id=doc_id, op="set_level", anchor=anchors[0], level=2)
        assert '<w:ilvl w:val="2"/>' in main_xml(session, doc_id)
        content = docx_read(session, doc_id=doc_id)["content"]
        assert "List:ol L3]" in content  # level 2 → L3 (ilvl + 1)

    def test_set_level_over_range(self) -> None:
        session, doc_id = open_docx()
        docx_list(
            session,
            doc_id=doc_id,
            op="create",
            after=A2,
            kind="ol",
            items=[{"text": "a"}, {"text": "b"}],
        )
        # The two items are P3 and P4.
        result = docx_list(session, doc_id=doc_id, op="set_level", range="P3..P4", level=1)
        assert result["n_affected"] == 2
        assert main_xml(session, doc_id).count('<w:ilvl w:val="1"/>') == 2


# ---------------------------------------------------------------------------
# convert (§17)
# ---------------------------------------------------------------------------


class TestConvert:
    def test_convert_paragraph_to_ul(self) -> None:
        session, doc_id = open_docx()
        docx_list(session, doc_id=doc_id, op="convert", anchor=A2, to="ul")
        content = docx_read(session, doc_id=doc_id, anchor=A2, window=0)["content"]
        assert "List:ul L1]" in content
        assert session.get(doc_id).package.has_part("word/numbering.xml")

    def test_convert_list_item_back_to_paragraph(self) -> None:
        session, doc_id = open_docx()
        anchors = docx_list(
            session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "a"}]
        )["new_anchors"]
        docx_list(session, doc_id=doc_id, op="convert", anchor=anchors[0], to="paragraphs")
        content = docx_read(session, doc_id=doc_id, anchor=anchors[0], window=0)["content"]
        assert "List:" not in content
        assert "<w:numPr>" not in main_xml(session, doc_id)

    def test_convert_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_list(session, doc_id=doc_id, op="convert", range="P1..P2", to="ol")
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


class TestErrors:
    def test_create_without_after_is_anchor_invalid(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_list(session, doc_id=doc_id, op="create", kind="ol", items=[{"text": "x"}])
        assert err.value.code == "anchor_invalid"

    def test_convert_unknown_target(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_list(session, doc_id=doc_id, op="convert", anchor=A2, to="bogus")
        assert err.value.code == "anchor_invalid"


def test_list_create_existing_numbering_reuses_part() -> None:
    # A second create allocates the next abstractNum/num ids in the existing part.
    session, doc_id = open_docx()
    docx_list(session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "a"}])
    a2_after = paragraph_anchor(3, "a")
    docx_list(session, doc_id=doc_id, op="create", after=a2_after, kind="ul", items=[{"text": "b"}])
    nb = numbering_xml(session, doc_id)
    assert '<w:abstractNum w:abstractNumId="1">' in nb
    assert '<w:abstractNum w:abstractNumId="2">' in nb
    assert '<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>' in nb
