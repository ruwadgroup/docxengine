"""Validator, repair, and save-gate tests (algorithms.md §8/§8a/§9)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from conftest import DOCUMENT_RELS_XML, DOCUMENT_XML, FIXTURE_PARTS, build_docx, document_xml

from docxengine import Package, Session, ToolError, repair_package, validate_package
from docxengine._tools_lifecycle import docx_repair, docx_save, docx_validate

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ORPHAN_REL = (
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/image" Target="media/image1.png"/>'
)


def with_orphan_rel(parts: dict[str, str]) -> dict[str, str]:
    parts = dict(parts)
    parts["word/_rels/document.xml.rels"] = DOCUMENT_RELS_XML.replace(
        "</Relationships>", ORPHAN_REL + "</Relationships>"
    )
    return parts


def with_duplicate_rev_ids(parts: dict[str, str]) -> dict[str, str]:
    # The standard fixture has w:del id=1 + w:ins id=2; collide them on id=1.
    parts = dict(parts)
    parts["word/document.xml"] = DOCUMENT_XML.replace('w:ins w:id="2"', 'w:ins w:id="1"')
    return parts


def corrupt_docx() -> bytes:
    """The task fixture: an orphaned relationship plus duplicate revision ids."""
    return build_docx(with_duplicate_rev_ids(with_orphan_rel(FIXTURE_PARTS)))


def errors_of(pkg: Package) -> list[str]:
    return [i.message for i in validate_package(pkg) if i.severity == "error"]


def warnings_of(pkg: Package) -> list[str]:
    return [i.message for i in validate_package(pkg) if i.severity == "warning"]


class TestValidate:
    def test_clean_fixture_is_valid_with_no_issues(self, docx_bytes: bytes) -> None:
        assert validate_package(Package.open(docx_bytes)) == []

    def test_orphaned_relationship_is_check_c_error(self) -> None:
        pkg = Package.open(build_docx(with_orphan_rel(FIXTURE_PARTS)))
        issues = validate_package(pkg)
        errors = [i for i in issues if i.severity == "error"]
        assert len(errors) == 1
        assert errors[0].part == "word/_rels/document.xml.rels"
        assert errors[0].message == (
            "Relationship rId9 targets missing part word/media/image1.png."
        )
        assert "docx_repair" in errors[0].fix_hint
        # The image rel is also never referenced: pinned §8a warning, never blocking.
        assert warnings_of(pkg) == ["Relationship rId9 (image) is never referenced."]

    def test_dangling_r_id_is_check_b_error(self) -> None:
        parts = dict(FIXTURE_PARTS)
        parts["word/document.xml"] = document_xml(
            f'<w:p><w:hyperlink xmlns:r="{R_NS}" r:id="rId8" w:history="1">'
            "<w:r><w:t>full report</w:t></w:r></w:hyperlink></w:p>"
        )
        issues = validate_package(Package.open(build_docx(parts)))
        assert [i.severity for i in issues] == ["error"]
        assert issues[0].part == "word/document.xml"
        assert issues[0].message == (
            "r:id rId8 is referenced in word/document.xml "
            "but not defined in word/_rels/document.xml.rels."
        )
        assert "not auto-repairable" in issues[0].fix_hint

    def test_duplicate_revision_ids_are_check_d_error(self) -> None:
        pkg = Package.open(build_docx(with_duplicate_rev_ids(FIXTURE_PARTS)))
        issues = validate_package(pkg)
        assert [i.severity for i in issues] == ["error"]
        assert issues[0].part == "word/document.xml"
        assert issues[0].message == "Duplicate revision id 1 on 2 w:ins/w:del elements."

    def test_uncovered_part_is_check_a_error(self) -> None:
        parts = dict(FIXTURE_PARTS)
        parts["word/media/image1.png"] = "not really a png"
        issues = validate_package(Package.open(build_docx(parts)))
        assert [i.severity for i in issues] == ["error"]
        assert issues[0].part == "word/media/image1.png"
        assert issues[0].message == (
            "Part word/media/image1.png is not covered by [Content_Types].xml "
            "(no Override, no Default for extension 'png')."
        )

    def test_comment_reference_without_definition_is_error(self) -> None:
        parts = dict(FIXTURE_PARTS)
        parts["word/document.xml"] = document_xml(
            '<w:p><w:r><w:t>Noted.</w:t></w:r><w:r><w:commentReference w:id="3"/></w:r></w:p>'
        )
        issues = validate_package(Package.open(build_docx(parts)))
        assert [i.severity for i in issues] == ["error"]
        assert issues[0].part == "word/comments.xml"
        assert issues[0].message == "Comment id=3 referenced in body but missing."

    def test_unreferenced_definitions_warn_and_separators_are_exempt(self) -> None:
        parts = dict(FIXTURE_PARTS)
        parts["word/document.xml"] = document_xml("<w:p><w:r><w:t>Plain.</w:t></w:r></w:p>")
        parts["word/comments.xml"] = (
            f'<w:comments xmlns:w="{W_NS}">'
            '<w:comment w:id="1" w:author="J.Doe"><w:p/></w:comment></w:comments>'
        )
        parts["word/footnotes.xml"] = (
            f'<w:footnotes xmlns:w="{W_NS}">'
            '<w:footnote w:type="separator" w:id="0"><w:p/></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="1"><w:p/></w:footnote>'
            "</w:footnotes>"
        )
        issues = validate_package(Package.open(build_docx(parts)))
        assert [i.severity for i in issues] == ["warning"]
        assert issues[0].message == "Comment id=1 defined but never referenced."

    def test_issue_order_is_pinned_a_then_c_then_d(self) -> None:
        parts = with_duplicate_rev_ids(with_orphan_rel(FIXTURE_PARTS))
        parts["docProps/thumbnail.jpeg"] = "binary"  # uncovered, unrelated to the orphan rel
        issues = list(validate_package(Package.open(build_docx(parts))))
        kinds = [i.message.split(" ", 1)[0] for i in issues if i.severity == "error"]
        assert kinds == ["Part", "Relationship", "Duplicate"]

    def test_docx_validate_tool_shape(self) -> None:
        session = Session()
        doc = session.open_doc(corrupt_docx())
        result = docx_validate(session, doc_id=doc.doc_id)
        assert result["valid"] is False
        issues = result["issues"]
        assert isinstance(issues, list)
        assert all(
            {"severity", "part", "message", "fix_hint"} == set(issue) for issue in issues
        )
        assert sum(1 for issue in issues if issue["severity"] == "error") == 2


class TestRepair:
    def test_repairs_orphaned_rel_and_duplicate_ids(self) -> None:
        pkg = Package.open(corrupt_docx())
        fixed, remaining = repair_package(pkg)
        assert fixed == [
            "removed orphaned relationship rId9 (word/_rels/document.xml.rels)",
            "renumbered duplicate revision id 1 -> 2",
        ]
        assert remaining == []
        assert validate_package(pkg) == []  # §8a: validate must then be clean

    def test_renumber_mirrors_corpus_semantics(self) -> None:
        # corrupt-dup-ids corpus: del id=5 + ins id=5 + ins id=6 -> second 5 becomes 7.
        parts = dict(FIXTURE_PARTS)
        parts["word/document.xml"] = document_xml(
            "<w:p>"
            '<w:del w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z">'
            "<w:r><w:delText>March 1</w:delText></w:r></w:del>"
            '<w:ins w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z">'
            "<w:r><w:t>April 1</w:t></w:r></w:ins>"
            '<w:ins w:id="6" w:author="A" w:date="2026-01-01T00:00:00Z">'
            "<w:r><w:t>promptly </w:t></w:r></w:ins>"
            "</w:p>"
        )
        pkg = Package.open(build_docx(parts))
        fixed, remaining = repair_package(pkg)
        assert fixed == ["renumbered duplicate revision id 5 -> 7"]
        assert remaining == []
        data = pkg.part("word/document.xml")
        assert b'<w:del w:id="5"' in data  # first in document order keeps its id
        assert b'<w:ins w:id="7"' in data

    def test_dangling_r_id_lands_in_remaining(self) -> None:
        parts = with_orphan_rel(FIXTURE_PARTS)
        parts["word/document.xml"] = document_xml(
            f'<w:p><w:hyperlink xmlns:r="{R_NS}" r:id="rId8" w:history="1">'
            "<w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>"
        )
        pkg = Package.open(build_docx(parts))
        fixed, remaining = repair_package(pkg)
        assert fixed == ["removed orphaned relationship rId9 (word/_rels/document.xml.rels)"]
        assert remaining == [
            "r:id rId8 is referenced in word/document.xml "
            "but not defined in word/_rels/document.xml.rels."
        ]

    def test_adds_missing_content_type_default(self) -> None:
        parts = dict(FIXTURE_PARTS)
        parts["word/media/image1.png"] = "binary"
        pkg = Package.open(build_docx(parts))
        fixed, remaining = repair_package(pkg)
        assert fixed == ["added content-type Default for extension 'png'"]
        assert remaining == []
        assert b'<Default Extension="png" ContentType="image/png"/>' in pkg.part(
            "[Content_Types].xml"
        )

    def test_removes_orphaned_comment_reference(self) -> None:
        parts = dict(FIXTURE_PARTS)
        parts["word/document.xml"] = document_xml(
            "<w:p>"
            '<w:commentRangeStart w:id="3"/>'
            "<w:r><w:t>Noted.</w:t></w:r>"
            '<w:commentRangeEnd w:id="3"/>'
            '<w:r><w:commentReference w:id="3"/></w:r>'
            "</w:p>"
        )
        pkg = Package.open(build_docx(parts))
        fixed, remaining = repair_package(pkg)
        assert fixed == ["removed orphaned comment reference id=3"]
        assert remaining == []
        data = pkg.part("word/document.xml")
        assert b"commentReference" not in data
        assert b"commentRangeStart" not in data
        assert validate_package(pkg) == []

    def test_docx_repair_tool_marks_dirty(self) -> None:
        session = Session()
        doc = session.open_doc(corrupt_docx())
        assert not doc.dirty
        result = docx_repair(session, doc_id=doc.doc_id)
        assert result["remaining"] == []
        assert doc.dirty


class TestSaveGate:
    def test_save_refuses_invalid_package(self, tmp_path: Path) -> None:
        session = Session()
        doc = session.open_doc(corrupt_docx())
        out = tmp_path / "out.docx"
        with pytest.raises(ToolError) as err:
            docx_save(session, doc_id=doc.doc_id, path=str(out))
        assert err.value.code == "validation_failed"
        assert err.value.suggestions[0] == "Run docx_repair, then re-validate."
        assert not out.exists()

    def test_save_succeeds_after_repair(self, tmp_path: Path) -> None:
        session = Session()
        doc = session.open_doc(corrupt_docx())
        docx_repair(session, doc_id=doc.doc_id)
        out = tmp_path / "out.docx"
        result = docx_save(session, doc_id=doc.doc_id, path=str(out))
        assert result["ok"] is True
        assert result["validated"] is True
        assert result["bytes"] == out.stat().st_size
        assert not doc.dirty
        assert validate_package(Package.open(out.read_bytes())) == []

    def test_warnings_never_block_save(self, tmp_path: Path) -> None:
        # An unreferenced image relationship to an existing part: warning only.
        parts = dict(FIXTURE_PARTS)
        parts["word/_rels/document.xml.rels"] = DOCUMENT_RELS_XML.replace(
            "</Relationships>",
            '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships/image" Target="styles.xml"/></Relationships>',
        )
        session = Session()
        doc = session.open_doc(build_docx(parts))
        assert warnings_of(doc.package) == ["Relationship rId9 (image) is never referenced."]
        out = tmp_path / "out.docx"
        result = docx_save(session, doc_id=doc.doc_id, path=str(out))
        assert result["ok"] is True

    def test_saved_output_round_trips(self, docx_bytes: bytes, tmp_path: Path) -> None:
        session = Session()
        doc = session.open_doc(docx_bytes)
        out = tmp_path / "out.docx"
        docx_save(session, doc_id=doc.doc_id, path=str(out))
        with zipfile.ZipFile(out) as zf:
            assert zf.namelist() == list(FIXTURE_PARTS)
