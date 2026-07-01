"""Section tests: docx_section (algorithms.md §15).

Covers list (S-ids, page size/orientation/columns/header-footer state), set_geometry
(page-size presets, orientation swap, cm→twip margins, columns), set_header/set_footer
(part creation + content-type Override + rel + sectPr reference), and insert_break
(clone the body sectPr into a paragraph's pPr). The validator stays green throughout.
"""

from __future__ import annotations

import base64

import pytest
from conftest import build_docx

from docxengine import (
    Session,
    docx_open,
    docx_read,
    docx_section,
    docx_validate,
    paragraph_anchor,
)

A1 = "P1#515a"
A2 = paragraph_anchor(2, "The term is five (5) years from the Effective Date.")


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


def part(session: Session, doc_id: str, name: str) -> str:
    return session.get(doc_id).package.part(name).decode("utf-8")


# ---------------------------------------------------------------------------
# list (§15)
# ---------------------------------------------------------------------------


class TestList:
    def test_list_reports_default_section(self) -> None:
        session, doc_id = open_docx()
        sections = docx_section(session, doc_id=doc_id, op="list")["sections"]
        assert sections == [
            {
                "id": "S1",
                "break_type": "nextPage",
                "page_size": "A4",
                "orientation": "portrait",
                "columns": 1,
                "has_header": False,
                "has_footer": False,
            }
        ]


# ---------------------------------------------------------------------------
# set_geometry (§15)
# ---------------------------------------------------------------------------


class TestSetGeometry:
    def test_page_size_preset(self) -> None:
        session, doc_id = open_docx()
        docx_section(session, doc_id=doc_id, op="set_geometry", section="S1", page_size="Letter")
        md = main_xml(session, doc_id)
        assert '<w:pgSz w:w="12240" w:h="15840"/>' in md
        listed = docx_section(session, doc_id=doc_id, op="list")["sections"][0]  # type: ignore[index]
        assert listed["page_size"] == "Letter"

    def test_landscape_swaps_dims_and_sets_orient(self) -> None:
        session, doc_id = open_docx()
        docx_section(
            session, doc_id=doc_id, op="set_geometry", section="S1",
            page_size="A4", orientation="landscape",
        )
        md = main_xml(session, doc_id)
        assert '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' in md
        listed = docx_section(session, doc_id=doc_id, op="list")["sections"][0]  # type: ignore[index]
        assert listed["orientation"] == "landscape"
        assert listed["page_size"] == "A4"

    def test_portrait_removes_orient(self) -> None:
        session, doc_id = open_docx()
        docx_section(
            session, doc_id=doc_id, op="set_geometry", section="S1",
            page_size="A4", orientation="landscape",
        )
        docx_section(
            session, doc_id=doc_id, op="set_geometry", section="S1", orientation="portrait"
        )
        md = main_xml(session, doc_id)
        assert 'w:orient="landscape"' not in md
        assert '<w:pgSz w:w="11906" w:h="16838"/>' in md

    def test_margins_cm_to_twips(self) -> None:
        session, doc_id = open_docx()
        docx_section(
            session, doc_id=doc_id, op="set_geometry", section="S1",
            margins={"top": 2.5, "left": 3.0},
        )
        md = main_xml(session, doc_id)
        # 2.5 cm × 567 = 1417.5 → round 1418; 3 cm × 567 = 1701.
        assert 'w:top="1418"' in md
        assert 'w:left="1701"' in md
        # Unspecified margins keep the defaults.
        assert 'w:right="1440"' in md
        assert 'w:header="708"' in md

    def test_columns(self) -> None:
        session, doc_id = open_docx()
        docx_section(session, doc_id=doc_id, op="set_geometry", section="S1", columns=2)
        md = main_xml(session, doc_id)
        assert '<w:cols w:num="2" w:space="708"/>' in md

    def test_geometry_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_section(
            session, doc_id=doc_id, op="set_geometry", section="S1",
            page_size="Legal", orientation="landscape", margins={"top": 1}, columns=3,
        )
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# set_header / set_footer (§15)
# ---------------------------------------------------------------------------


class TestHeaderFooter:
    def test_set_header_creates_part_and_wires_everything(self) -> None:
        session, doc_id = open_docx()
        docx_section(
            session, doc_id=doc_id, op="set_header", section="S1",
            content="Confidential", variant="default",
        )
        package = session.get(doc_id).package
        assert package.has_part("word/header1.xml")
        assert "word/header1.xml" in package.content_types().overrides
        hdr = part(session, doc_id, "word/header1.xml")
        assert "<w:hdr" in hdr and "Confidential" in hdr
        md = main_xml(session, doc_id)
        assert '<w:headerReference w:type="default"' in md
        listed = docx_section(session, doc_id=doc_id, op="list")["sections"][0]  # type: ignore[index]
        assert listed["has_header"] is True

    def test_set_footer_creates_footer_part(self) -> None:
        session, doc_id = open_docx()
        docx_section(
            session, doc_id=doc_id, op="set_footer", section="S1",
            content="Page footer", variant="default",
        )
        package = session.get(doc_id).package
        assert package.has_part("word/footer1.xml")
        md = main_xml(session, doc_id)
        assert '<w:footerReference w:type="default"' in md
        listed = docx_section(session, doc_id=doc_id, op="list")["sections"][0]  # type: ignore[index]
        assert listed["has_footer"] is True

    def test_header_reference_precedes_pgsz(self) -> None:
        session, doc_id = open_docx()
        docx_section(session, doc_id=doc_id, op="set_header", section="S1", content="H")
        md = main_xml(session, doc_id)
        assert md.index("<w:headerReference") < md.index("<w:pgSz")

    def test_header_content_readable(self) -> None:
        session, doc_id = open_docx()
        docx_section(session, doc_id=doc_id, op="set_header", section="S1", content="Confidential")
        result = docx_read(session, doc_id=doc_id, scope="headers")
        assert "Confidential" in str(result["content"])

    def test_header_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_section(session, doc_id=doc_id, op="set_header", section="S1", content="H")
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# insert_break (§15)
# ---------------------------------------------------------------------------


class TestInsertBreak:
    def test_insert_break_adds_sectpr_to_paragraph(self) -> None:
        session, doc_id = open_docx()
        res = docx_section(
            session, doc_id=doc_id, op="insert_break", after=A1, break_type="nextPage"
        )
        assert res["new_anchor"] == A1  # text unchanged → same anchor
        md = main_xml(session, doc_id)
        assert '<w:type w:val="nextPage"/>' in md
        # A new section now exists in the document.
        sections = docx_section(session, doc_id=doc_id, op="list")["sections"]
        assert len(sections) == 2  # type: ignore[arg-type]

    def test_insert_break_type_first_in_sectpr(self) -> None:
        session, doc_id = open_docx()
        docx_section(session, doc_id=doc_id, op="insert_break", after=A1, break_type="continuous")
        md = main_xml(session, doc_id)
        # The paragraph-level sectPr begins with w:type.
        ppr_sect = md.index("<w:sectPr>")
        assert md[ppr_sect:].startswith('<w:sectPr><w:type w:val="continuous"/>')

    def test_insert_break_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_section(session, doc_id=doc_id, op="insert_break", after=A2, break_type="evenPage")
        assert docx_validate(session, doc_id=doc_id)["valid"] is True
