"""Edit-surface tests: docx_replace (§4/§5), docx_edit_paragraph (§6), docx_insert,
docx_delete (§6a), plus conformance-corpus parity for the pinned edit cases."""

from __future__ import annotations

import base64
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from conftest import (
    FIXTURE_PARTS,
    PARA_RSID_FRAGMENTED,
    SECT_PR,
    build_docx,
    document_xml,
)

from docxengine import (
    Session,
    ToolError,
    docx_delete,
    docx_edit_paragraph,
    docx_insert,
    docx_open,
    docx_read,
    docx_replace,
    paragraph_anchor,
)
from docxengine._edits import diff_blocks, diff_units, word_diff

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FIXED_DATE = "2026-06-10T00:00:00Z"

# Anchors over the standard conftest fixture (P1 split heading, P2 rsid runs, P3 tracked).
A1 = "P1#515a"
OLD_P2 = "The term is five (5) years from the Effective Date."
NEW_P2 = "The term is three (3) years from the Effective Date."
A2 = paragraph_anchor(2, OLD_P2)
A2_NEW = paragraph_anchor(2, NEW_P2)
A3 = paragraph_anchor(3, "Payment due in 45 days")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def open_docx(parts: dict[str, str] | None = None) -> tuple[Session, str]:
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())
    return session, str(result["doc_id"])


def part_bytes(session: Session, doc_id: str) -> bytes:
    package = session.get(doc_id).package
    return package.part(package.main_document_part())


def single_para_doc(*paragraphs: str) -> dict[str, str]:
    parts = dict(FIXTURE_PARTS)
    parts["word/document.xml"] = document_xml(*paragraphs, SECT_PR)
    return parts


# ---------------------------------------------------------------------------
# docx_replace — plain (§4)
# ---------------------------------------------------------------------------


class TestDocxReplace:
    def test_split_run_replace(self) -> None:
        session, doc_id = open_docx()
        result = docx_replace(
            session, doc_id=doc_id, anchor=A2, old="five (5) years", new="three (3) years"
        )
        assert result == {"n_replaced": 1, "new_anchor": A2_NEW}
        data = part_bytes(session, doc_id)
        # First overlapping w:t got prefix+replacement; the next kept only its suffix.
        assert b'<w:r w:rsidR="00EF34AB"><w:t>three (3) years</w:t></w:r>' in data
        assert b'<w:r w:rsidR="00CD56EF"><w:t xml:space="preserve"> from the </w:t></w:r>' in data
        # Untouched regions survive verbatim (rsid noise intact).
        assert b'<w:r w:rsidR="00115E6B"><w:t>Effective Date.</w:t></w:r>' in data
        assert b'<w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD">' in data
        assert NEW_P2 in str(docx_read(session, doc_id=doc_id, anchor=A2_NEW)["content"])

    def test_replace_marks_dirty(self) -> None:
        session, doc_id = open_docx()
        assert session.get(doc_id).dirty is False
        docx_replace(session, doc_id=doc_id, anchor=A2, old="five", new="six")
        assert session.get(doc_id).dirty is True

    def test_whole_doc_replace_single_match(self) -> None:
        session, doc_id = open_docx()
        result = docx_replace(session, doc_id=doc_id, old="five (5) years", new="ten (10) years")
        assert result["n_replaced"] == 1
        assert result["new_anchor"] == paragraph_anchor(
            2, "The term is ten (10) years from the Effective Date."
        )

    def test_multiple_matches_without_all_is_ambiguous(self) -> None:
        session, doc_id = open_docx(
            single_para_doc(
                "<w:p><w:r><w:t>alpha one</w:t></w:r></w:p>",
                "<w:p><w:r><w:t>two alpha</w:t></w:r></w:p>",
            )
        )
        with pytest.raises(ToolError) as exc_info:
            docx_replace(session, doc_id=doc_id, old="alpha", new="beta")
        assert exc_info.value.code == "ambiguous_target"
        assert "2 times" in exc_info.value.message

    def test_replace_all_returns_fresh_anchors_ascending(self) -> None:
        session, doc_id = open_docx(
            single_para_doc(
                "<w:p><w:r><w:t>alpha one</w:t></w:r></w:p>",
                "<w:p><w:r><w:t>two alpha</w:t></w:r></w:p>",
            )
        )
        result = docx_replace(session, doc_id=doc_id, old="alpha", new="beta", all=True)
        assert result == {
            "n_replaced": 2,
            "anchors": [paragraph_anchor(1, "beta one"), paragraph_anchor(2, "two beta")],
        }

    def test_all_true_rerun_is_idempotent(self) -> None:
        session, doc_id = open_docx(
            single_para_doc(
                "<w:p><w:r><w:t>alpha one</w:t></w:r></w:p>",
                "<w:p><w:r><w:t>two alpha</w:t></w:r></w:p>",
            )
        )
        first = docx_replace(session, doc_id=doc_id, old="alpha", new="beta", all=True)
        assert first["n_replaced"] == 2
        snapshot = part_bytes(session, doc_id)
        second = docx_replace(session, doc_id=doc_id, old="alpha", new="beta", all=True)
        assert second == {"n_replaced": 0, "anchors": []}
        assert part_bytes(session, doc_id) == snapshot

    def test_all_true_multiple_matches_in_one_paragraph(self) -> None:
        session, doc_id = open_docx(
            single_para_doc("<w:p><w:r><w:t>alpha beta alpha</w:t></w:r></w:p>")
        )
        result = docx_replace(session, doc_id=doc_id, old="alpha", new="x", all=True)
        assert result["n_replaced"] == 2
        assert "x beta x" in str(docx_read(session, doc_id=doc_id)["content"])

    def test_replacement_containing_old_terminates(self) -> None:
        session, doc_id = open_docx(single_para_doc("<w:p><w:r><w:t>aa bb aa</w:t></w:r></w:p>"))
        result = docx_replace(session, doc_id=doc_id, old="aa", new="aaa", all=True)
        assert result["n_replaced"] == 2
        assert "aaa bb aaa" in str(docx_read(session, doc_id=doc_id)["content"])

    def test_empty_new_removes_emptied_run(self) -> None:
        session, doc_id = open_docx()
        result = docx_replace(session, doc_id=doc_id, anchor=A2, old="five (5) ", new="")
        assert result["n_replaced"] == 1
        data = part_bytes(session, doc_id)
        assert b"00EF34AB" not in data  # §4 rule 4: the emptied w:r is gone entirely
        assert "The term is years from the Effective Date." in str(
            docx_read(session, doc_id=doc_id)["content"]
        )

    def test_not_found_without_all(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_replace(session, doc_id=doc_id, old="liquidated damages", new="x")
        assert exc_info.value.code == "not_found"

    def test_empty_old_is_not_found(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_replace(session, doc_id=doc_id, old="", new="x", all=True)
        assert exc_info.value.code == "not_found"

    def test_stale_anchor_checked_before_matching(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_replace(session, doc_id=doc_id, anchor="P2#0000", old="no such text", new="x")
        assert exc_info.value.code == "anchor_stale"
        assert "P2#0000" in exc_info.value.message

    def test_anchor_invalid_forms(self) -> None:
        session, doc_id = open_docx()
        for bad in ("P2", "T1", "junk", "P0#abcd", "P2#XYZW"):
            with pytest.raises(ToolError) as exc_info:
                docx_replace(session, doc_id=doc_id, anchor=bad, old="five", new="six")
            assert exc_info.value.code == "anchor_invalid"

    def test_anchor_not_found(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_replace(session, doc_id=doc_id, anchor="P99#abcd", old="five", new="six")
        assert exc_info.value.code == "anchor_not_found"


# ---------------------------------------------------------------------------
# docx_replace — tracked (§5)
# ---------------------------------------------------------------------------


class TestDocxReplaceTracked:
    def test_tracked_replace_emits_parseable_redline(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = open_docx()
        result = docx_replace(
            session,
            doc_id=doc_id,
            anchor=A2,
            old="five (5) years",
            new="three (3) years",
            track_changes=True,
            author="Claude",
        )
        assert result == {"n_replaced": 1, "new_anchor": A2_NEW}  # hash sees as-if-accepted
        data = part_bytes(session, doc_id)
        # Existing revision ids are 1 and 2 (P3), so the new wrappers take 3 and 4,
        # attributes in §5 order, per-run formatting preserved inside w:del.
        assert (
            f'<w:del w:id="3" w:author="Claude" w:date="{FIXED_DATE}">'
            '<w:r><w:delText xml:space="preserve">five (5) </w:delText></w:r>'
            "<w:r><w:delText>years</w:delText></w:r></w:del>"
            f'<w:ins w:id="4" w:author="Claude" w:date="{FIXED_DATE}">'
            "<w:r><w:t>three (3) years</w:t></w:r></w:ins>"
            '<w:r><w:t xml:space="preserve"> from the </w:t></w:r>'
        ).encode() in data
        root = ET.fromstring(data)  # the spliced part stays well-formed
        assert len(root.findall(f".//{{{W_NS}}}del")) == 2  # new + the P3 original
        assert len(root.findall(f".//{{{W_NS}}}ins")) == 2
        content = str(docx_read(session, doc_id=doc_id, anchor=A2_NEW)["content"])
        assert "[del by Claude]" in content and "[ins by Claude]" in content
        assert "three (3) years" in content and "five (5)" not in content

    def test_tracked_with_empty_new_emits_no_ins(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = open_docx(single_para_doc(PARA_RSID_FRAGMENTED))
        docx_replace(
            session, doc_id=doc_id, old="five (5) ", new="", track_changes=True, author="Claude"
        )
        data = part_bytes(session, doc_id)
        assert b"<w:ins" not in data
        assert b'<w:delText xml:space="preserve">five (5) </w:delText>' in data

    def test_author_defaults_from_env_then_engine_name(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        monkeypatch.setenv("DOCXENGINE_AUTHOR", "EnvAuthor")
        session, doc_id = open_docx()
        docx_replace(session, doc_id=doc_id, anchor=A2, old="five", new="six", track_changes=True)
        assert b'w:author="EnvAuthor"' in part_bytes(session, doc_id)
        monkeypatch.delenv("DOCXENGINE_AUTHOR")
        session2, doc_id2 = open_docx()
        docx_replace(session2, doc_id=doc_id2, anchor=A2, old="five", new="six", track_changes=True)
        assert b'w:author="DocxEngine"' in part_bytes(session2, doc_id2)

    def test_fixed_date_output_is_deterministic(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)

        def run_once() -> bytes:
            session, doc_id = open_docx()
            docx_replace(
                session,
                doc_id=doc_id,
                anchor=A2,
                old="five (5) years",
                new="three (3) years",
                track_changes=True,
                author="Claude",
            )
            return part_bytes(session, doc_id)

        assert run_once() == run_once()


# ---------------------------------------------------------------------------
# docx_edit_paragraph (§6)
# ---------------------------------------------------------------------------

SINGLE_RUN_P2 = f"<w:p><w:r><w:t>{OLD_P2}</w:t></w:r></w:p>"
BOLD_SPLIT = (
    '<w:p><w:r><w:t xml:space="preserve">The term is five (5) </w:t></w:r>'
    "<w:r><w:rPr><w:b/></w:rPr><w:t>years from the Effective Date.</w:t></w:r></w:p>"
)


class TestDocxEditParagraph:
    def test_untracked_rewrite_single_run(self) -> None:
        session, doc_id = open_docx()
        result = docx_edit_paragraph(session, doc_id=doc_id, anchor=A2, text="Brand new text.")
        assert result["new_anchor"] == paragraph_anchor(2, "Brand new text.")
        data = part_bytes(session, doc_id)
        assert b"<w:r><w:t>Brand new text.</w:t></w:r>" in data
        assert b"00EF34AB" not in data  # the old fragmented runs are gone

    def test_tracked_diff_is_minimal_redline(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = open_docx(single_para_doc(SINGLE_RUN_P2))
        result = docx_edit_paragraph(
            session,
            doc_id=doc_id,
            anchor=paragraph_anchor(1, OLD_P2),
            text=NEW_P2,
            track_changes=True,
            author="Claude",
        )
        assert result == {
            "new_anchor": paragraph_anchor(1, NEW_P2),
            "diff": "~2 words changed",
        }
        data = part_bytes(session, doc_id)
        assert data.count(b"<w:del ") == 1 and data.count(b"<w:ins ") == 1
        assert (
            '<w:r><w:t xml:space="preserve">The term is </w:t></w:r>'
            f'<w:del w:id="1" w:author="Claude" w:date="{FIXED_DATE}">'
            '<w:r><w:delText xml:space="preserve">five (5) </w:delText></w:r></w:del>'
            f'<w:ins w:id="2" w:author="Claude" w:date="{FIXED_DATE}">'
            '<w:r><w:t xml:space="preserve">three (3) </w:t></w:r></w:ins>'
            "<w:r><w:t>years from the Effective Date.</w:t></w:r>"
        ).encode() in data
        assert docx_read(session, doc_id=doc_id)["content"] == (
            f"[{paragraph_anchor(1, NEW_P2)}] The term is [del by Claude] three (3)"
            " [ins by Claude] years from the Effective Date."
        )

    def test_word_diff_spec_example(self) -> None:
        ops = word_diff(diff_units("three year term"), diff_units("five year initial term"))
        assert diff_blocks(ops) == [
            ("del", "three "),
            ("ins", "five "),
            ("keep", "year "),
            ("ins", "initial "),
            ("keep", "term"),
        ]

    def test_tracked_keep_and_replace_spans_preserve_run_formatting(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = open_docx(single_para_doc(BOLD_SPLIT))
        docx_edit_paragraph(
            session,
            doc_id=doc_id,
            anchor=paragraph_anchor(1, OLD_P2),
            text="The term is five (5) years from the Start Date.",
            track_changes=True,
            author="Claude",
        )
        data = part_bytes(session, doc_id)
        # Kept spans re-emit per-run portions with each run's own rPr.
        assert b'<w:r><w:t xml:space="preserve">The term is five (5) </w:t></w:r>' in data
        assert (
            b'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">years from the </w:t></w:r>'
            in data
        )
        # The deleted span keeps the bold run's rPr; the replacement ins inherits it (§5).
        assert (
            b'<w:r><w:rPr><w:b/></w:rPr><w:delText xml:space="preserve">Effective </w:delText>'
            b"</w:r></w:del>"
        ) in data
        assert (
            f'<w:ins w:id="2" w:author="Claude" w:date="{FIXED_DATE}">'
            '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Start </w:t></w:r></w:ins>'
        ).encode() in data
        assert b"<w:r><w:rPr><w:b/></w:rPr><w:t>Date.</w:t></w:r>" in data

    def test_tracked_pure_insertion_takes_rpr_at_offset(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = open_docx(single_para_doc(BOLD_SPLIT))
        docx_edit_paragraph(
            session,
            doc_id=doc_id,
            anchor=paragraph_anchor(1, OLD_P2),
            text="The term is five (5) whole years from the Effective Date.",
            track_changes=True,
            author="Claude",
        )
        data = part_bytes(session, doc_id)
        # Insertion offset 21 is exactly where the bold run starts → bold rPr.
        assert (
            f'<w:ins w:id="1" w:author="Claude" w:date="{FIXED_DATE}">'
            '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">whole </w:t></w:r></w:ins>'
        ).encode() in data
        assert data.count(b"<w:del ") == 0

    def test_identical_text_changes_nothing_visible(self) -> None:
        session, doc_id = open_docx(single_para_doc(SINGLE_RUN_P2))
        anchor = paragraph_anchor(1, OLD_P2)
        result = docx_edit_paragraph(
            session, doc_id=doc_id, anchor=anchor, text=OLD_P2, track_changes=True
        )
        assert result == {"new_anchor": anchor, "diff": "~0 words changed"}
        data = part_bytes(session, doc_id)
        assert b"<w:del " not in data and b"<w:ins " not in data

    def test_stale_anchor(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_edit_paragraph(session, doc_id=doc_id, anchor="P2#beef", text="x")
        assert exc_info.value.code == "anchor_stale"

    def test_singular_diff_noun(self) -> None:
        session, doc_id = open_docx(single_para_doc("<w:p><w:r><w:t>alpha beta</w:t></w:r></w:p>"))
        result = docx_edit_paragraph(
            session, doc_id=doc_id, anchor=paragraph_anchor(1, "alpha beta"), text="alpha gamma"
        )
        assert result["diff"] == "~1 word changed"


# ---------------------------------------------------------------------------
# docx_insert (§6a)
# ---------------------------------------------------------------------------


class TestDocxInsert:
    def test_insert_after_shifts_following_ordinals(self) -> None:
        session, doc_id = open_docx()
        result = docx_insert(session, doc_id=doc_id, after=A2, content="New paragraph here.")
        assert result == {"new_anchors": [paragraph_anchor(3, "New paragraph here.")]}
        content = str(docx_read(session, doc_id=doc_id)["content"])
        lines = content.split("\n")
        assert lines[2].endswith("New paragraph here.")
        assert lines[3] == f"[{paragraph_anchor(4, 'Payment due in 45 days')}] " + (
            "Payment due in [del by J.Doe] 45 [ins by J.Doe] days"
        )

    def test_insert_before_first_paragraph(self) -> None:
        session, doc_id = open_docx()
        result = docx_insert(session, doc_id=doc_id, before=A1, content="Preamble")
        assert result == {"new_anchors": [paragraph_anchor(1, "Preamble")]}
        assert str(docx_read(session, doc_id=doc_id)["content"]).startswith(
            f"[{paragraph_anchor(1, 'Preamble')}] Preamble"
        )

    def test_minimal_markdown(self) -> None:
        session, doc_id = open_docx()
        content = "# Title\r\n\n#### Deep\n- item one\n* item two\n   \nplain"
        result = docx_insert(session, doc_id=doc_id, after=A2, content=content)
        anchors = result["new_anchors"]
        assert anchors == [
            paragraph_anchor(3, "Title"),
            paragraph_anchor(4, "Deep"),
            paragraph_anchor(5, "item one"),
            paragraph_anchor(6, "item two"),
            paragraph_anchor(7, "plain"),
        ]
        data = part_bytes(session, doc_id)
        assert (
            b'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>'
            in data
        )
        assert b'<w:pStyle w:val="Heading4"/>' in data
        assert (
            b'<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>'
            b"<w:r><w:t>item one</w:t></w:r></w:p>" in data
        )
        assert b"<w:p><w:r><w:t>plain</w:t></w:r></w:p>" in data
        projected = str(docx_read(session, doc_id=doc_id)["content"])
        assert f"[{paragraph_anchor(3, 'Title')} H1] Title" in projected
        assert f"[{paragraph_anchor(4, 'Deep')} H4] Deep" in projected

    def test_style_override_applies_to_every_paragraph(self) -> None:
        session, doc_id = open_docx()
        result = docx_insert(
            session, doc_id=doc_id, after=A2, content="## sub\n- item", style="Heading 1"
        )
        assert len(list(result["new_anchors"])) == 2
        data = part_bytes(session, doc_id)
        # "Heading 1" is not a styleId; whitespace removed → Heading1 (defined) wins.
        assert (
            b'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>sub</w:t></w:r></w:p>'
            b'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>item</w:t></w:r></w:p>'
        ) in data
        assert b"ListParagraph" not in data

    def test_unknown_style_is_style_unknown(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_insert(session, doc_id=doc_id, after=A2, content="x", style="Fancy")
        assert exc_info.value.code == "style_unknown"

    def test_after_and_before_are_mutually_exclusive(self) -> None:
        session, doc_id = open_docx()
        for kwargs in ({}, {"after": A2, "before": A1}):
            with pytest.raises(ToolError) as exc_info:
                docx_insert(session, doc_id=doc_id, content="x", **kwargs)
            assert exc_info.value.code == "anchor_invalid"

    def test_stale_position_anchor(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_insert(session, doc_id=doc_id, after="P2#0000", content="x")
        assert exc_info.value.code == "anchor_stale"

    def test_tracked_insert_wraps_each_paragraph_run(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = open_docx()
        result = docx_insert(
            session,
            doc_id=doc_id,
            after=A3,
            content="One\nTwo",
            track_changes=True,
            author="Bob",
        )
        assert result["new_anchors"] == [paragraph_anchor(4, "One"), paragraph_anchor(5, "Two")]
        data = part_bytes(session, doc_id)
        # Existing max id is 2 → the two new w:ins take 3 and 4 in document order.
        assert (
            f'<w:p><w:ins w:id="3" w:author="Bob" w:date="{FIXED_DATE}">'
            "<w:r><w:t>One</w:t></w:r></w:ins></w:p>"
            f'<w:p><w:ins w:id="4" w:author="Bob" w:date="{FIXED_DATE}">'
            "<w:r><w:t>Two</w:t></w:r></w:ins></w:p>"
        ).encode() in data
        assert "[ins by Bob]" in str(docx_read(session, doc_id=doc_id)["content"])

    def test_empty_content_is_a_no_op(self) -> None:
        session, doc_id = open_docx()
        assert docx_insert(session, doc_id=doc_id, after=A2, content="  \n \n") == {
            "new_anchors": []
        }
        assert session.get(doc_id).dirty is False


# ---------------------------------------------------------------------------
# docx_delete (§6a)
# ---------------------------------------------------------------------------


class TestDocxDelete:
    def test_delete_by_anchor(self) -> None:
        session, doc_id = open_docx()
        result = docx_delete(session, doc_id=doc_id, anchor=A2)
        assert result == {"ok": True, "deleted": 1}
        content = str(docx_read(session, doc_id=doc_id)["content"])
        assert "The term is" not in content
        assert content.count("\n") == 1  # two paragraphs remain
        assert f"[{paragraph_anchor(2, 'Payment due in 45 days')}]" in content

    def test_delete_range(self) -> None:
        session, doc_id = open_docx()
        result = docx_delete(session, doc_id=doc_id, range="P1..P2")
        assert result == {"ok": True, "deleted": 2}
        content = str(docx_read(session, doc_id=doc_id)["content"])
        assert content.startswith(f"[{paragraph_anchor(1, 'Payment due in 45 days')}]")

    def test_range_endpoint_hashes_validated_when_present(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_delete(session, doc_id=doc_id, range="P1#0000..P2")
        assert exc_info.value.code == "anchor_stale"
        result = docx_delete(session, doc_id=doc_id, range=f"P1#515a..{A2}")
        assert result == {"ok": True, "deleted": 2}

    def test_range_errors(self) -> None:
        session, doc_id = open_docx()
        cases = {
            "P3..P1": "anchor_invalid",
            "P2..P99": "anchor_not_found",
            "bogus": "anchor_invalid",
        }
        for bad_range, code in cases.items():
            with pytest.raises(ToolError) as exc_info:
                docx_delete(session, doc_id=doc_id, range=bad_range)
            assert exc_info.value.code == code

    def test_anchor_and_range_are_mutually_exclusive(self) -> None:
        session, doc_id = open_docx()
        for kwargs in ({}, {"anchor": A2, "range": "P1..P2"}):
            with pytest.raises(ToolError) as exc_info:
                docx_delete(session, doc_id=doc_id, **kwargs)
            assert exc_info.value.code == "anchor_invalid"

    def test_tracked_delete_wraps_content_after_ppr(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = open_docx(
            single_para_doc(
                '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
                "<w:r><w:t>Gone soon</w:t></w:r></w:p>",
                "<w:p/>",
                "<w:p><w:r><w:t>Also gone</w:t></w:r></w:p>",
            )
        )
        result = docx_delete(
            session, doc_id=doc_id, range="P1..P3", track_changes=True, author="Claude"
        )
        assert result == {"ok": True, "deleted": 3}
        data = part_bytes(session, doc_id)
        assert (
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            f'<w:del w:id="1" w:author="Claude" w:date="{FIXED_DATE}">'
            "<w:r><w:delText>Gone soon</w:delText></w:r></w:del></w:p>"
        ).encode() in data
        # The empty paragraph is counted but unchanged; ids stay sequential.
        assert b"<w:p/>" in data
        assert f'<w:del w:id="2" w:author="Claude" w:date="{FIXED_DATE}">'.encode() in data
        content = str(docx_read(session, doc_id=doc_id)["content"])
        assert content.split("\n")[0] == "[P1#e3b0 H1] [del by Claude]"

    def test_tracked_delete_keeps_paragraph_count(self) -> None:
        session, doc_id = open_docx()
        docx_delete(session, doc_id=doc_id, anchor=A2, track_changes=True)
        content = str(docx_read(session, doc_id=doc_id)["content"])
        assert content.count("\n") == 2  # still three paragraphs
        assert "[P2#e3b0]" in content  # as-if-accepted text is now empty


# ---------------------------------------------------------------------------
# Conformance corpus parity (pins the edit cases in conformance/cases/)
# ---------------------------------------------------------------------------

CORPUS = Path(__file__).resolve().parents[2] / "conformance" / "corpus"


@pytest.mark.skipif(not CORPUS.is_dir(), reason="conformance corpus not present")
class TestConformanceParity:
    def open_corpus(self, name: str) -> tuple[Session, str]:
        session = Session()
        result = docx_open(session, path=str(CORPUS / name / "input.docx"))
        return session, str(result["doc_id"])

    def test_replace_split_run(self) -> None:
        session, doc_id = self.open_corpus("split-runs")
        result = docx_replace(
            session, doc_id=doc_id, anchor="P2#d337", old="five (5) years", new="three (3) years"
        )
        assert result == {"n_replaced": 1, "new_anchor": "P2#eeb0"}

    def test_replace_tracked(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = self.open_corpus("split-runs")
        result = docx_replace(
            session,
            doc_id=doc_id,
            anchor="P2#d337",
            old="five (5) years",
            new="three (3) years",
            track_changes=True,
            author="Claude",
        )
        assert result == {"n_replaced": 1, "new_anchor": "P2#eeb0"}
        data = part_bytes(session, doc_id)
        assert b"<w:del " in data and b"<w:ins " in data
        ET.fromstring(data)

    def test_edit_paragraph_diff(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = self.open_corpus("minimal")
        text = (
            "Each party shall protect all Confidential Information"
            " with commercially reasonable care."
        )
        result = docx_edit_paragraph(
            session,
            doc_id=doc_id,
            anchor="P5#d27e",
            text=text,
            track_changes=True,
            author="Claude",
        )
        assert result["new_anchor"] == "P5#70d6"
        assert paragraph_anchor(5, text) == "P5#70d6"
        data = part_bytes(session, doc_id)
        assert data.count(b"<w:ins ") == 2 and data.count(b"<w:del ") == 0
