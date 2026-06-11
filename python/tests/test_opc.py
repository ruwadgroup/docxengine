"""OPC layer tests: lazy parts, rels/content-types, §9 save normalization."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from conftest import DOCUMENT_XML, FIXTURE_PARTS, STYLES_XML, build_docx

from docxengine import Package, ToolError, rels_part_for, resolve_rel_target


class TestOpen:
    def test_open_from_bytes(self, docx_bytes: bytes) -> None:
        pkg = Package.open(docx_bytes)
        assert pkg.part("word/document.xml") == DOCUMENT_XML.encode("utf-8")

    def test_open_from_path(self, docx_bytes: bytes, tmp_path: Path) -> None:
        path = tmp_path / "fixture.docx"
        path.write_bytes(docx_bytes)
        pkg = Package.open(path)
        assert pkg.source_path == str(path)
        assert pkg.part("word/styles.xml") == STYLES_XML.encode("utf-8")

    def test_part_accepts_leading_slash(self, docx_bytes: bytes) -> None:
        pkg = Package.open(docx_bytes)
        assert pkg.part("/word/document.xml") == pkg.part("word/document.xml")

    def test_missing_part_raises_keyerror(self, docx_bytes: bytes) -> None:
        pkg = Package.open(docx_bytes)
        with pytest.raises(KeyError):
            pkg.part("word/nonexistent.xml")

    def test_not_a_zip_raises_open_failed(self) -> None:
        with pytest.raises(ToolError) as err:
            Package.open(b"this is not a zip archive")
        assert err.value.code == "open_failed"
        payload = err.value.to_payload()
        assert payload["error"] == "open_failed"
        assert isinstance(payload["suggestions"], list)

    def test_zip_without_content_types_raises_open_failed(self) -> None:
        bad = build_docx({"word/document.xml": DOCUMENT_XML})
        with pytest.raises(ToolError) as err:
            Package.open(bad)
        assert err.value.code == "open_failed"

    def test_missing_path_raises_open_failed(self, tmp_path: Path) -> None:
        with pytest.raises(ToolError) as err:
            Package.open(tmp_path / "missing.docx")
        assert err.value.code == "open_failed"


class TestMetadata:
    def test_content_types_defaults_and_overrides(self, docx_bytes: bytes) -> None:
        ct = Package.open(docx_bytes).content_types()
        assert ct.defaults["rels"] == ("application/vnd.openxmlformats-package.relationships+xml")
        assert ct.content_type_of("word/document.xml") == (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        )
        # Falls back to the extension default when no Override exists.
        assert ct.content_type_of("word/fontTable.xml") == "application/xml"
        assert ct.content_type_of("word/_rels/document.xml.rels") == (
            "application/vnd.openxmlformats-package.relationships+xml"
        )
        assert ct.content_type_of("word/media/image1.png") is None

    def test_root_rels(self, docx_bytes: bytes) -> None:
        rels = Package.open(docx_bytes).rels()
        assert len(rels) == 1
        rel = rels[0]
        assert rel.rel_id == "rId1"
        assert rel.target == "word/document.xml"
        assert rel.target_mode == "Internal"
        assert not rel.is_external

    def test_part_rels_and_target_resolution(self, docx_bytes: bytes) -> None:
        pkg = Package.open(docx_bytes)
        rels = pkg.rels("word/document.xml")
        assert [r.rel_id for r in rels] == ["rId1"]
        assert resolve_rel_target("word/document.xml", rels[0].target) == "word/styles.xml"

    def test_rels_for_part_without_rels_is_empty(self, docx_bytes: bytes) -> None:
        assert Package.open(docx_bytes).rels("word/styles.xml") == []

    def test_rels_part_for(self) -> None:
        assert rels_part_for(None) == "_rels/.rels"
        assert rels_part_for("word/document.xml") == "word/_rels/document.xml.rels"
        assert rels_part_for("/word/document.xml") == "word/_rels/document.xml.rels"

    def test_resolve_rel_target_forms(self) -> None:
        assert resolve_rel_target(None, "word/document.xml") == "word/document.xml"
        assert resolve_rel_target("word/document.xml", "media/image1.png") == (
            "word/media/image1.png"
        )
        assert resolve_rel_target("word/document.xml", "/word/styles.xml") == "word/styles.xml"
        assert resolve_rel_target("word/document.xml", "../docProps/core.xml") == (
            "docProps/core.xml"
        )

    def test_main_document_part(self, docx_bytes: bytes) -> None:
        assert Package.open(docx_bytes).main_document_part() == "word/document.xml"


class TestSave:
    def test_roundtrip_untouched_parts_byte_stable(self, docx_bytes: bytes, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        Package.open(docx_bytes).save(out)

        reopened = Package.open(out)
        assert reopened.part_names == list(FIXTURE_PARTS)
        for name, text in FIXTURE_PARTS.items():
            assert reopened.part(name) == text.encode("utf-8"), name

    def test_save_is_deterministic(self, docx_bytes: bytes, tmp_path: Path) -> None:
        a, b = tmp_path / "a.docx", tmp_path / "b.docx"
        Package.open(docx_bytes).save(a)
        Package.open(docx_bytes).save(b)
        assert a.read_bytes() == b.read_bytes()
        # Saving the saved file again is a fixed point (archive bytes stable).
        c = tmp_path / "c.docx"
        Package.open(a).save(c)
        assert c.read_bytes() == a.read_bytes()

    def test_zip_metadata_normalized(self, docx_bytes: bytes, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        Package.open(docx_bytes).save(out)
        with zipfile.ZipFile(out) as zf:
            assert zf.comment == b""
            for info in zf.infolist():
                assert info.date_time == (1980, 1, 1, 0, 0, 0), info.filename
                assert info.compress_type == zipfile.ZIP_DEFLATED, info.filename
                assert info.extra == b"", info.filename
                assert info.comment == b"", info.filename

    def test_set_part_marks_dirty_and_saves(self, docx_bytes: bytes, tmp_path: Path) -> None:
        pkg = Package.open(docx_bytes)
        assert pkg.dirty_part_names == ()
        new_doc = DOCUMENT_XML.replace("Master", "Amended").encode("utf-8")
        pkg.set_part("word/document.xml", new_doc)
        assert pkg.is_dirty("word/document.xml")
        assert pkg.dirty_part_names == ("word/document.xml",)
        assert pkg.part("word/document.xml") == new_doc

        out = tmp_path / "out.docx"
        pkg.save(out)
        reopened = Package.open(out)
        assert reopened.part("word/document.xml") == new_doc
        # Untouched parts pass through with identical decompressed bytes.
        for name in ("word/styles.xml", "_rels/.rels", "[Content_Types].xml"):
            assert reopened.part(name) == FIXTURE_PARTS[name].encode("utf-8"), name

    def test_new_parts_append_after_originals_in_creation_order(
        self, docx_bytes: bytes, tmp_path: Path
    ) -> None:
        pkg = Package.open(docx_bytes)
        pkg.set_part("word/comments.xml", b"<w:comments/>")
        pkg.set_part("word/footnotes.xml", b"<w:footnotes/>")
        out = tmp_path / "out.docx"
        pkg.save(out)
        with zipfile.ZipFile(out) as zf:
            names = [i.filename for i in zf.infolist()]
        assert names == [*FIXTURE_PARTS, "word/comments.xml", "word/footnotes.xml"]

    def test_atomic_save_replaces_existing_file(self, docx_bytes: bytes, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        out.write_bytes(b"old contents")
        Package.open(docx_bytes).save(out)
        assert Package.open(out).part_names == list(FIXTURE_PARTS)
        leftovers = [p.name for p in tmp_path.iterdir() if p.name.startswith(".docxengine-")]
        assert leftovers == []

    def test_save_to_unwritable_path_raises_save_failed(self, docx_bytes: bytes) -> None:
        with pytest.raises(ToolError) as err:
            Package.open(docx_bytes).save("/nonexistent-dir-xyz/out.docx")
        assert err.value.code == "save_failed"

    def test_lazy_part_access_after_save(self, docx_bytes: bytes, tmp_path: Path) -> None:
        # Parts never read before save still round-trip byte-identically.
        pkg = Package.open(docx_bytes)
        out = tmp_path / "out.docx"
        pkg.save(out)
        assert Package.open(out).part("word/styles.xml") == STYLES_XML.encode("utf-8")
