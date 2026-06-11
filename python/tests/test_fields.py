"""Field tests: docx_field (algorithms.md §20).

Covers insert_toc (the run-triple TOC field paragraph after an anchor — never
fldSimple), insert_page_number (ensure the body section's footer, append a PAGE
field run-triple), and update (flip <w:updateFields w:val="true"/> in settings.xml,
creating the part on demand). The validator stays green on every produced document.
"""

from __future__ import annotations

import base64

import pytest
from conftest import build_docx

from docxengine import (
    Session,
    docx_field,
    docx_open,
    docx_validate,
    paragraph_anchor,
)

A1 = "P1#515a"
A2 = paragraph_anchor(2, "The term is five (5) years from the Effective Date.")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def open_docx() -> tuple[Session, str]:
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx()).decode())
    return session, str(result["doc_id"])


def main_xml(session: Session, doc_id: str) -> str:
    package = session.get(doc_id).package
    return package.part(package.main_document_part()).decode("utf-8")


def part(session: Session, doc_id: str, name: str) -> str:
    return session.get(doc_id).package.part(name).decode("utf-8")


# ---------------------------------------------------------------------------
# insert_toc (§20)
# ---------------------------------------------------------------------------


class TestInsertToc:
    def test_inserts_run_triple_field(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="insert_toc", after=A1, levels=3)
        md = main_xml(session, doc_id)
        assert '<w:fldChar w:fldCharType="begin"/>' in md
        assert '<w:fldChar w:fldCharType="separate"/>' in md
        assert '<w:fldChar w:fldCharType="end"/>' in md
        assert 'TOC \\o "1-3" \\h \\z \\u' in md
        assert "Right-click to update field." in md
        # Never fldSimple.
        assert "fldSimple" not in md

    def test_levels_drives_instr(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="insert_toc", after=A1, levels=2)
        md = main_xml(session, doc_id)
        assert 'TOC \\o "1-2" \\h \\z \\u' in md

    def test_new_anchor_is_the_inserted_paragraph(self) -> None:
        session, doc_id = open_docx()
        res = docx_field(session, doc_id=doc_id, op="insert_toc", after=A1, levels=3)
        # The TOC field text normalizes to "Right-click to update field." → fresh anchor.
        assert res["new_anchor"] != A1
        assert str(res["new_anchor"]).startswith("P2#")

    def test_toc_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="insert_toc", after=A2, levels=3)
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# insert_page_number (§20)
# ---------------------------------------------------------------------------


class TestInsertPageNumber:
    def test_creates_footer_with_page_field(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="insert_page_number", scope="footer")
        package = session.get(doc_id).package
        assert package.has_part("word/footer1.xml")
        ftr = part(session, doc_id, "word/footer1.xml")
        assert '<w:instrText xml:space="preserve"> PAGE </w:instrText>' in ftr
        assert '<w:fldChar w:fldCharType="begin"/>' in ftr
        # The body sectPr references the footer.
        md = main_xml(session, doc_id)
        assert "<w:footerReference" in md

    def test_no_placeholder_empty_paragraph(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="insert_page_number", scope="footer")
        ftr = part(session, doc_id, "word/footer1.xml")
        # The footer holds exactly the PAGE field paragraph, no stray empty <w:p/>.
        assert ftr.count("<w:p>") == 1
        assert "<w:p/>" not in ftr

    def test_header_scope(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="insert_page_number", scope="header")
        package = session.get(doc_id).package
        assert package.has_part("word/header1.xml")
        md = main_xml(session, doc_id)
        assert "<w:headerReference" in md

    def test_page_number_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="insert_page_number", scope="footer")
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# update (§20)
# ---------------------------------------------------------------------------


class TestUpdate:
    def test_update_creates_settings_with_updatefields(self) -> None:
        session, doc_id = open_docx()
        res = docx_field(session, doc_id=doc_id, op="update")
        assert res["updated"] == 1
        package = session.get(doc_id).package
        assert package.has_part("word/settings.xml")
        assert "word/settings.xml" in package.content_types().overrides
        settings = part(session, doc_id, "word/settings.xml")
        assert '<w:updateFields w:val="true"/>' in settings
        rels = package.rels(package.main_document_part())
        rel_types = {r.rel_type.rsplit("/", 1)[-1] for r in rels}
        assert "settings" in rel_types

    def test_updatefields_is_first_child_of_settings(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="update")
        settings = part(session, doc_id, "word/settings.xml")
        start = settings.index("<w:settings")
        first_child = settings.index(">", start) + 1
        assert settings[first_child:].startswith("<w:updateFields")

    def test_update_is_idempotent(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="update")
        docx_field(session, doc_id=doc_id, op="update")
        settings = part(session, doc_id, "word/settings.xml")
        assert settings.count("<w:updateFields") == 1

    def test_update_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_field(session, doc_id=doc_id, op="update")
        assert docx_validate(session, doc_id=doc_id)["valid"] is True
