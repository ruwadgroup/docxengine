"""Anchor tests: §1 normalization/hashing, index building, stability across save."""

from __future__ import annotations

import hashlib
from pathlib import Path

from conftest import (
    PARA_RSID_FRAGMENTED,
    PARA_SPLIT_RUN,
    PARA_TRACKED,
    SECT_PR,
    build_docx,
    document_xml,
)

from docxengine import (
    Package,
    _xml,
    anchor_hash,
    build_anchor_index,
    normalized_text,
    paragraph_anchor,
)


def _index_for(*body_children: str) -> list:
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/></Types>'
        ),
        "word/document.xml": document_xml(*body_children),
    }
    return build_anchor_index(Package.open(build_docx(parts)), "word/document.xml")


class TestNormalizedText:
    def test_whitespace_collapse_full_set(self) -> None:
        # Tab, NBSP, NEL, em-space, line separator, narrow NBSP, ideographic space:
        # every maximal §1 White_Space run collapses to one ASCII space.
        raw = "a\t\tb\xa0c\x85d\u2003e\u2028f\u202fg\u3000h"
        assert normalized_text(raw) == "a b c d e f g h"

    def test_strip_leading_and_trailing(self) -> None:
        assert normalized_text("  Master Services  Agreement ") == "Master Services Agreement"

    def test_nfc_applied_before_collapse(self) -> None:
        decomposed = "Cafe\u0301 menu"  # e + combining acute
        composed = "Caf\u00e9 menu"
        assert normalized_text(decomposed) == composed
        assert anchor_hash(normalized_text(decomposed)) == anchor_hash(composed)

    def test_empty_normalizes_to_empty_and_hashes_e3b0(self) -> None:
        assert normalized_text("") == ""
        assert normalized_text(" \t  ") == ""
        assert anchor_hash("") == "e3b0"

    def test_hash_is_sha256_prefix(self) -> None:
        text = "The quick brown fox"
        assert anchor_hash(text) == hashlib.sha256(text.encode("utf-8")).hexdigest()[:4]


class TestAnchorIndex:
    def test_spec_worked_example_split_run_paragraph(self) -> None:
        # algorithms.md §1: the split-run Heading paragraph anchors as P1#515a.
        entries = _index_for(PARA_SPLIT_RUN)
        assert entries[0].anchor == "P1#515a"
        assert entries[0].normalized == "Master Services Agreement"

    def test_rsid_fragmentation_does_not_change_hash(self) -> None:
        fragmented = _index_for(PARA_RSID_FRAGMENTED)[0]
        single_run = _index_for(
            "<w:p><w:r><w:t>The term is five (5) years from the Effective Date.</w:t></w:r></w:p>"
        )[0]
        assert fragmented.normalized == "The term is five (5) years from the Effective Date."
        assert fragmented.anchor == single_run.anchor

    def test_deltext_excluded_insertions_included(self) -> None:
        # The tracked paragraph reads as-if-accepted: del("30") out, ins("45") in.
        entry = _index_for(PARA_TRACKED)[0]
        assert entry.normalized == "Payment due in 45 days"
        assert entry.anchor == paragraph_anchor(1, "Payment due in 45 days")

    def test_empty_paragraph_anchor(self) -> None:
        entries = _index_for("<w:p/>", '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>')
        assert [e.anchor for e in entries] == ["P1#e3b0", "P2#e3b0"]

    def test_ordinals_sectpr_and_tables(self) -> None:
        table = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        entries = _index_for(PARA_SPLIT_RUN, table, PARA_TRACKED, SECT_PR)
        assert [(e.kind, e.ordinal, e.anchor) for e in entries] == [
            ("paragraph", 1, "P1#515a"),
            ("table", 1, "T1"),
            ("paragraph", 2, paragraph_anchor(2, "Payment due in 45 days")),
        ]
        # Paragraphs nested inside the table get no body ordinal; sectPr is skipped.

    def test_full_fixture_index(self, docx_bytes: bytes) -> None:
        entries = build_anchor_index(Package.open(docx_bytes))
        assert len(entries) == 3  # three paragraphs; sectPr excluded
        assert [e.ordinal for e in entries] == [1, 2, 3]
        assert entries[0].anchor == "P1#515a"

    def test_anchors_stable_across_save_and_reopen(self, docx_bytes: bytes, tmp_path: Path) -> None:
        pkg = Package.open(docx_bytes)
        before = [e.anchor for e in build_anchor_index(pkg)]
        out = tmp_path / "out.docx"
        pkg.save(out)
        after = [e.anchor for e in build_anchor_index(Package.open(out))]
        assert before == after

        # A second save/reopen cycle is also stable.
        out2 = tmp_path / "out2.docx"
        Package.open(out).save(out2)
        assert [e.anchor for e in build_anchor_index(Package.open(out2))] == before

    def test_anchors_unaffected_by_unrelated_part_edits(
        self, docx_bytes: bytes, tmp_path: Path
    ) -> None:
        pkg = Package.open(docx_bytes)
        before = [e.anchor for e in build_anchor_index(pkg)]
        pkg.set_part("word/styles.xml", b"<w:styles/>")
        out = tmp_path / "out.docx"
        pkg.save(out)
        assert [e.anchor for e in build_anchor_index(Package.open(out))] == before


class TestScannerOffsetMap:
    def test_paragraph_text_offsets_and_runs(self) -> None:
        data = document_xml(PARA_RSID_FRAGMENTED).encode("utf-8")
        para = next(_xml.iter_body_children(data))
        assert para.name == "w:p"
        text, pieces = _xml.paragraph_text(data, para)
        assert text == "The term is five (5) years from the Effective Date."
        assert [p.text for p in pieces] == [
            "The term is ",
            "five (5) ",
            "years from the ",
            "Effective Date.",
        ]
        # Each piece maps back into the original buffer at its w:t content.
        for piece in pieces:
            raw = data[piece.t.inner_start : piece.t.inner_end].decode("utf-8")
            assert raw == piece.text
            assert piece.run is not None
            assert piece.run.start < piece.t.start <= piece.t.end <= piece.run.end

        # Index 12 ("five…") falls in the second w:t at char offset 0.
        piece, offset = _xml.locate(pieces, 12)
        assert piece is pieces[1]
        assert offset == 0
        piece, offset = _xml.locate(pieces, 20)  # the space after "(5)"
        assert piece is pieces[1]
        assert offset == 8

    def test_deltext_not_in_offset_map(self) -> None:
        data = document_xml(PARA_TRACKED).encode("utf-8")
        para = next(_xml.iter_body_children(data))
        text, pieces = _xml.paragraph_text(data, para)
        assert text == "Payment due in 45 days"
        assert [p.text for p in pieces] == ["Payment due in ", "45", " days"]

    def test_entities_decoded_in_text(self) -> None:
        data = document_xml(
            "<w:p><w:r><w:t>Fee &amp; tax &lt; $100&#x2122;</w:t></w:r></w:p>"
        ).encode("utf-8")
        para = next(_xml.iter_body_children(data))
        text, _ = _xml.paragraph_text(data, para)
        assert text == "Fee & tax < $100™"

    def test_splice_replaces_exact_ranges(self) -> None:
        data = b"<w:t>five</w:t> and <w:t>years</w:t>"
        out = _xml.splice(data, [(5, 9, b"three"), (25, 30, b"days")])
        assert out == b"<w:t>three</w:t> and <w:t>days</w:t>"

    def test_emit_text_element_space_preserve(self) -> None:
        assert _xml.emit_text_element("plain") == "<w:t>plain</w:t>"
        assert _xml.emit_text_element(" lead") == '<w:t xml:space="preserve"> lead</w:t>'
        assert _xml.emit_text_element("trail ") == ('<w:t xml:space="preserve">trail </w:t>')
        assert _xml.emit_text_element("a & b < c") == "<w:t>a &amp; b &lt; c</w:t>"
        assert _xml.emit_text_element("gone", tag="w:delText") == "<w:delText>gone</w:delText>"
