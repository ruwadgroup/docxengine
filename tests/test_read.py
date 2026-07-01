"""Read-surface tests: session store, §2 projection, outline, read windows, search."""

from __future__ import annotations

import base64
from pathlib import Path

import pytest
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
    Session,
    ToolError,
    docx_open,
    docx_outline,
    docx_read,
    docx_search,
    paragraph_anchor,
    project_read,
)

# ---------------------------------------------------------------------------
# Fixtures: headings + styles.xml basedOn chain, lists, a table, tracked changes
# ---------------------------------------------------------------------------

STYLES_CHAIN_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1">'
    '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading2">'
    '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="SectionHead">'
    '<w:name w:val="Section Head"/><w:basedOn w:val="Heading2"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="SubHead">'
    '<w:name w:val="Sub Head"/><w:basedOn w:val="SectionHead"/></w:style>'
    "</w:styles>"
)

NUMBERING_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="0">'
    '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>'
    '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl>'
    "</w:abstractNum>"
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    "</w:numbering>"
)

RICH_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/numbering" Target="numbering.xml"/>'
    "</Relationships>"
)

PARA_PLAIN = "<w:p><w:r><w:t>This Agreement is entered into by the parties.</w:t></w:r></w:p>"
PARA_SUBHEAD = (
    '<w:p><w:pPr><w:pStyle w:val="SubHead"/></w:pPr><w:r><w:t>Definitions</w:t></w:r></w:p>'
)
PARA_LIST_OL = (
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
    "<w:r><w:t>First obligation</w:t></w:r></w:p>"
)
PARA_LIST_UL = (
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>'
    "<w:r><w:t>Sub bullet</w:t></w:r></w:p>"
)
TABLE = (
    "<w:tbl>"
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
    "<w:tr><w:tc><w:p><w:r><w:t>Term</w:t></w:r></w:p></w:tc>"
    "<w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>"
    "<w:tr><w:tc><w:p><w:r><w:t>Fee</w:t></w:r></w:p></w:tc>"
    "<w:tc><w:p><w:r><w:t>$100</w:t></w:r></w:p></w:tc></w:tr>"
    "</w:tbl>"
)
PARA_CLOSING = "<w:p><w:r><w:t>Closing terms apply.</w:t></w:r></w:p>"

RICH_DOCUMENT_XML = document_xml(
    PARA_SPLIT_RUN,  # P1: Heading1, "Master Services Agreement" (anchor P1#515a)
    PARA_PLAIN,  # P2
    PARA_SUBHEAD,  # P3: SubHead -> SectionHead -> Heading2 via basedOn chain
    PARA_RSID_FRAGMENTED,  # P4: split runs
    PARA_TRACKED,  # P5: w:del("30") + w:ins("45")
    PARA_LIST_OL,  # P6
    PARA_LIST_UL,  # P7
    TABLE,  # T1
    PARA_CLOSING,  # P8
    SECT_PR,
)

RICH_PARTS: dict[str, str] = {
    "[Content_Types].xml": (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    "_rels/.rels": (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    ),
    "word/document.xml": RICH_DOCUMENT_XML,
    "word/_rels/document.xml.rels": RICH_RELS_XML,
    "word/styles.xml": STYLES_CHAIN_XML,
    "word/numbering.xml": NUMBERING_XML,
}

# Expected projection lines for the rich fixture.
A2 = paragraph_anchor(2, "This Agreement is entered into by the parties.")
A3 = paragraph_anchor(3, "Definitions")
A4 = paragraph_anchor(4, "The term is five (5) years from the Effective Date.")
A5 = paragraph_anchor(5, "Payment due in 45 days")
A6 = paragraph_anchor(6, "First obligation")
A7 = paragraph_anchor(7, "Sub bullet")
A8 = paragraph_anchor(8, "Closing terms apply.")

L1 = "[P1#515a H1] Master Services Agreement"
L2 = f"[{A2}] This Agreement is entered into by the parties."
L3 = f"[{A3} H2] Definitions"
L4 = f"[{A4}] The term is five (5) years from the Effective Date."
L5 = f"[{A5}] Payment due in [del by J.Doe] 45 [ins by J.Doe] days"
L6 = f"[{A6} List:ol L1] First obligation"
L7 = f"[{A7} List:ul L2] Sub bullet"
LT = f"[T1 2×2 @after:{A7}]\n| Term | Value |\n| --- | --- |\n| Fee | $100 |"
L8 = f"[{A8}] Closing terms apply."

FULL_PROJECTION = "\n".join([L1, L2, L3, L4, L5, L6, L7, LT, L8])

COMMENTS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:comment w:id="1" w:author="J.Doe" w:date="2026-03-01T10:00:00Z" w:initials="JD">'
    "<w:p><w:r><w:t>Please verify this figure.</w:t></w:r></w:p>"
    "</w:comment></w:comments>"
)

PARA_COMMENTED = (
    '<w:p><w:r><w:t xml:space="preserve">Revenue grew </w:t></w:r>'
    '<w:commentRangeStart w:id="1"/>'
    "<w:r><w:t>nine percent</w:t></w:r>"
    '<w:commentRangeEnd w:id="1"/>'
    '<w:r><w:commentReference w:id="1"/></w:r>'
    '<w:r><w:t xml:space="preserve"> year over year.</w:t></w:r></w:p>'
)

COMMENT_PARTS: dict[str, str] = {
    "[Content_Types].xml": RICH_PARTS["[Content_Types].xml"],
    "_rels/.rels": RICH_PARTS["_rels/.rels"],
    "word/document.xml": document_xml(PARA_COMMENTED, SECT_PR),
    "word/comments.xml": COMMENTS_XML,
}


def open_rich() -> tuple[Session, str]:
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx(RICH_PARTS)).decode())
    return session, str(result["doc_id"])


# ---------------------------------------------------------------------------
# Session
# ---------------------------------------------------------------------------


class TestSession:
    def test_sequential_doc_ids(self, docx_bytes: bytes) -> None:
        session = Session()
        assert session.open_doc(docx_bytes).doc_id == "d1"
        assert session.open_doc(docx_bytes).doc_id == "d2"
        assert len(session) == 2
        assert "d1" in session and "d2" in session

    def test_get_returns_same_document(self, docx_bytes: bytes) -> None:
        session = Session()
        doc = session.open_doc(docx_bytes)
        assert session.get(doc.doc_id) is doc

    def test_get_unknown_raises_doc_not_found(self) -> None:
        with pytest.raises(ToolError) as exc_info:
            Session().get("d99")
        assert exc_info.value.code == "doc_not_found"
        assert exc_info.value.suggestions == ["Call docx_open again."]

    def test_close_forgets_and_id_is_not_reused(self, docx_bytes: bytes) -> None:
        session = Session()
        doc = session.open_doc(docx_bytes)
        session.close(doc.doc_id)
        assert doc.doc_id not in session
        with pytest.raises(ToolError):
            session.get(doc.doc_id)
        assert session.open_doc(docx_bytes).doc_id == "d2"

    def test_close_unknown_raises(self) -> None:
        with pytest.raises(ToolError) as exc_info:
            Session().close("d7")
        assert exc_info.value.code == "doc_not_found"

    def test_dirty_state_lifecycle(self, docx_bytes: bytes) -> None:
        doc = Session().open_doc(docx_bytes)
        assert doc.dirty is False
        doc.mark_dirty()
        assert doc.dirty is True
        doc.mark_saved()
        assert doc.dirty is False


# ---------------------------------------------------------------------------
# docx_open
# ---------------------------------------------------------------------------


class TestDocxOpen:
    def test_open_from_path(self, docx_bytes: bytes, tmp_path: Path) -> None:
        path = tmp_path / "contract.docx"
        path.write_bytes(docx_bytes)
        result = docx_open(Session(), path=str(path))
        assert set(result) == {
            "doc_id",
            "summary",
            "n_paragraphs",
            "has_tracked_changes",
            "has_comments",
        }
        assert result["doc_id"] == "d1"
        assert result["n_paragraphs"] == 3
        assert result["has_tracked_changes"] is True
        assert result["has_comments"] is False
        assert result["summary"] == "Master Services Agreement — 3 paragraphs, 1 section, 0 tables"

    def test_open_from_base64_bytes(self, docx_bytes: bytes) -> None:
        result = docx_open(Session(), bytes=base64.b64encode(docx_bytes).decode())
        assert result["doc_id"] == "d1"
        assert result["n_paragraphs"] == 3

    def test_rich_summary_counts_sections_and_tables(self) -> None:
        session = Session()
        result = docx_open(session, bytes=base64.b64encode(build_docx(RICH_PARTS)).decode())
        assert result["n_paragraphs"] == 8
        assert result["summary"] == "Master Services Agreement — 8 paragraphs, 1 section, 1 table"

    def test_comments_detected(self) -> None:
        result = docx_open(Session(), bytes=base64.b64encode(build_docx(COMMENT_PARTS)).decode())
        assert result["has_comments"] is True
        assert result["has_tracked_changes"] is False

    def test_missing_args_is_open_failed(self) -> None:
        with pytest.raises(ToolError) as exc_info:
            docx_open(Session())
        assert exc_info.value.code == "open_failed"

    def test_invalid_base64_is_open_failed(self) -> None:
        with pytest.raises(ToolError) as exc_info:
            docx_open(Session(), bytes="not-valid-base64!!")
        assert exc_info.value.code == "open_failed"

    def test_unreadable_path_is_open_failed(self, tmp_path: Path) -> None:
        with pytest.raises(ToolError) as exc_info:
            docx_open(Session(), path=str(tmp_path / "missing.docx"))
        assert exc_info.value.code == "open_failed"


# ---------------------------------------------------------------------------
# docx_outline
# ---------------------------------------------------------------------------


class TestDocxOutline:
    def test_outline_resolves_basedon_chain(self) -> None:
        session, doc_id = open_rich()
        result = docx_outline(session, doc_id=doc_id)
        assert set(result) == {"outline", "tables"}
        assert result["outline"] == [
            {"anchor": "P1#515a", "level": 1, "text": "Master Services Agreement"},
            {"anchor": A3, "level": 2, "text": "Definitions"},
        ]
        assert result["tables"] == [{"anchor": "T1", "dims": "2×2", "after": A7}]

    def test_unknown_doc_id(self) -> None:
        with pytest.raises(ToolError) as exc_info:
            docx_outline(Session(), doc_id="d1")
        assert exc_info.value.code == "doc_not_found"

    def test_table_with_no_preceding_paragraph_omits_after(self) -> None:
        parts = dict(RICH_PARTS)
        parts["word/document.xml"] = document_xml(TABLE, PARA_CLOSING, SECT_PR)
        session = Session()
        doc_id = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())["doc_id"]
        result = docx_outline(session, doc_id=str(doc_id))
        assert result["tables"] == [{"anchor": "T1", "dims": "2×2"}]


# ---------------------------------------------------------------------------
# docx_read
# ---------------------------------------------------------------------------


class TestDocxRead:
    def test_whole_body_projection(self) -> None:
        session, doc_id = open_rich()
        result = docx_read(session, doc_id=doc_id)
        assert result == {"content": FULL_PROJECTION}

    def test_tracked_paragraph_shows_as_accepted_with_markers(self) -> None:
        session, doc_id = open_rich()
        content = docx_read(session, doc_id=doc_id, anchor=A5)["content"]
        assert content == L5
        assert "30" not in content  # w:delText never shown in the default projection
        assert "[del by J.Doe]" in content and "[ins by J.Doe]" in content

    def test_anchor_window(self) -> None:
        session, doc_id = open_rich()
        result = docx_read(session, doc_id=doc_id, anchor=A5, window=1)
        assert result["content"] == "\n".join([L4, L5, L6])

    def test_window_clamps_at_document_edges(self) -> None:
        session, doc_id = open_rich()
        result = docx_read(session, doc_id=doc_id, anchor="P1#515a", window=2)
        assert result["content"] == "\n".join([L1, L2, L3])

    def test_window_counts_blocks_so_table_fills_a_slot(self) -> None:
        # §2a: window counts body-level blocks; the block before P8 is the table.
        session, doc_id = open_rich()
        result = docx_read(session, doc_id=doc_id, anchor=A8, window=1)
        assert result["content"] == "\n".join([LT, L8])

    def test_stale_hash_still_reads(self) -> None:
        session, doc_id = open_rich()
        result = docx_read(session, doc_id=doc_id, anchor="P5#0000")
        assert result["content"] == L5

    def test_anchor_wins_over_range(self) -> None:
        session, doc_id = open_rich()
        result = docx_read(session, doc_id=doc_id, anchor=A5, range="P1..P2")
        assert result["content"] == L5

    def test_range_includes_interleaved_table(self) -> None:
        session, doc_id = open_rich()
        result = docx_read(session, doc_id=doc_id, range="P6..P8")
        assert result["content"] == "\n".join([L6, L7, LT, L8])

    def test_range_endpoints_may_carry_ignored_hashes(self) -> None:
        session, doc_id = open_rich()
        result = docx_read(session, doc_id=doc_id, range="P6#0000..P7#ffff")
        assert result["content"] == "\n".join([L6, L7])

    def test_range_endpoint_out_of_range_is_anchor_not_found(self) -> None:
        session, doc_id = open_rich()
        with pytest.raises(ToolError) as exc_info:
            docx_read(session, doc_id=doc_id, range="P8..P99")
        assert exc_info.value.code == "anchor_not_found"

    def test_malformed_anchor_is_anchor_invalid(self) -> None:
        session, doc_id = open_rich()
        for bad in ("X9", "P0#1234", "P1#XYZW", "p1#abcd"):
            with pytest.raises(ToolError) as exc_info:
                docx_read(session, doc_id=doc_id, anchor=bad)
            assert exc_info.value.code == "anchor_invalid"

    def test_out_of_range_is_anchor_not_found(self) -> None:
        session, doc_id = open_rich()
        with pytest.raises(ToolError) as exc_info:
            docx_read(session, doc_id=doc_id, anchor="P99#abcd")
        assert exc_info.value.code == "anchor_not_found"

    def test_inverted_range_is_anchor_invalid(self) -> None:
        session, doc_id = open_rich()
        with pytest.raises(ToolError) as exc_info:
            docx_read(session, doc_id=doc_id, range="P5..P2")
        assert exc_info.value.code == "anchor_invalid"

    def test_unknown_scope_is_anchor_invalid(self) -> None:
        session, doc_id = open_rich()
        with pytest.raises(ToolError) as exc_info:
            docx_read(session, doc_id=doc_id, scope="margins")
        assert exc_info.value.code == "anchor_invalid"

    def test_comment_marker_and_comments_scope(self) -> None:
        session = Session()
        doc_id = docx_open(session, bytes=base64.b64encode(build_docx(COMMENT_PARTS)).decode())[
            "doc_id"
        ]
        body = docx_read(session, doc_id=str(doc_id))["content"]
        anchor = paragraph_anchor(1, "Revenue grew nine percent year over year.")
        assert body == (
            f"[{anchor}] Revenue grew nine percent [comment:C1 by J.Doe] year over year."
        )
        comments = docx_read(session, doc_id=str(doc_id), scope="comments")["content"]
        comment_anchor = paragraph_anchor(1, "Please verify this figure.")
        assert comments == f"[{comment_anchor}] Please verify this figure."

    def test_missing_story_part_reads_empty(self) -> None:
        session, doc_id = open_rich()
        assert docx_read(session, doc_id=doc_id, scope="footnotes") == {"content": ""}

    def test_headers_concatenate_numerically_with_story_anchors(self) -> None:
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        parts = dict(RICH_PARTS)
        # Numeric order (2 before 10), not lexicographic; ordinals run across parts.
        parts["word/header10.xml"] = (
            f'<w:hdr xmlns:w="{w_ns}"><w:p><w:r><w:t>Second header</w:t></w:r></w:p></w:hdr>'
        )
        parts["word/header2.xml"] = (
            f'<w:hdr xmlns:w="{w_ns}"><w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>'
        )
        session = Session()
        doc_id = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())["doc_id"]
        content = docx_read(session, doc_id=str(doc_id), scope="headers")["content"]
        assert content == "\n".join(
            [
                f"[{paragraph_anchor(1, 'First header')}] First header",
                f"[{paragraph_anchor(2, 'Second header')}] Second header",
            ]
        )

    def test_fully_deleted_paragraph_projects_marker_only(self) -> None:
        parts = dict(RICH_PARTS)
        parts["word/document.xml"] = document_xml(
            '<w:p><w:del w:id="9" w:author="Bob" w:date="2026-01-01T00:00:00Z">'
            "<w:r><w:delText>gone entirely</w:delText></w:r></w:del></w:p>",
            "<w:p/>",
            SECT_PR,
        )
        session = Session()
        doc_id = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())["doc_id"]
        content = docx_read(session, doc_id=str(doc_id))["content"]
        assert content == "[P1#e3b0] [del by Bob]\n[P2#e3b0]"

    def test_pagination_round_trip(self) -> None:
        package = Package.open(build_docx(RICH_PARTS))
        pages: list[str] = []
        result = project_read(package, char_budget=120)
        pages.append(result["content"])
        seen_continuations: list[str] = []
        while "continuation" in result:
            seen_continuations.append(result["continuation"])
            result = project_read(package, range=result["continuation"], char_budget=120)
            pages.append(result["content"])
        assert len(pages) > 1
        assert "\n".join(pages) == FULL_PROJECTION
        # Continuation tokens are plain paragraph ranges ending at the last paragraph.
        assert all(token.endswith("..P8") for token in seen_continuations)

    def test_first_block_always_returned_even_over_budget(self) -> None:
        package = Package.open(build_docx(RICH_PARTS))
        result = project_read(package, char_budget=1)
        assert result["content"] == L1
        assert result["continuation"] == "P2..P8"


# ---------------------------------------------------------------------------
# docx_search
# ---------------------------------------------------------------------------


class TestDocxSearch:
    def test_match_across_fragmented_runs(self) -> None:
        # "five (5) years" spans three rsid-fragmented runs in P4.
        session, doc_id = open_rich()
        result = docx_search(session, doc_id=doc_id, query="five (5) years")
        assert result["n_matches"] == 1
        (match,) = list(result["matches"])
        assert match["anchor"] == A4
        assert "five (5) years" in match["snippet"]
        assert match["context"] == "Definitions"  # nearest heading above P4

    def test_match_across_tracked_change_boundary(self) -> None:
        # "in 45 days" only exists in the as-accepted text (45 lives inside w:ins).
        session, doc_id = open_rich()
        result = docx_search(session, doc_id=doc_id, query="in 45 days")
        assert result["n_matches"] == 1
        assert result["matches"][0]["anchor"] == A5

    def test_markers_and_deleted_text_not_searchable(self) -> None:
        session, doc_id = open_rich()
        assert docx_search(session, doc_id=doc_id, query="del by")["n_matches"] == 0
        assert docx_search(session, doc_id=doc_id, query="30")["n_matches"] == 0

    def test_no_match_returns_empty_result(self) -> None:
        session, doc_id = open_rich()
        result = docx_search(session, doc_id=doc_id, query="liquidated damages")
        assert result == {"matches": [], "n_matches": 0}

    def test_multiple_matches_in_one_paragraph(self) -> None:
        parts = dict(RICH_PARTS)
        parts["word/document.xml"] = document_xml(
            "<w:p><w:r><w:t>alpha beta alpha</w:t></w:r></w:p>", SECT_PR
        )
        session = Session()
        doc_id = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())["doc_id"]
        result = docx_search(session, doc_id=str(doc_id), query="alpha")
        assert result["n_matches"] == 2
        anchors = {m["anchor"] for m in list(result["matches"])}
        assert anchors == {paragraph_anchor(1, "alpha beta alpha")}

    def test_regex_search(self) -> None:
        session, doc_id = open_rich()
        result = docx_search(session, doc_id=doc_id, query=r"five \(\d\) years", regex=True)
        assert result["n_matches"] == 1
        assert result["matches"][0]["anchor"] == A4

    def test_invalid_regex_is_not_found(self) -> None:
        session, doc_id = open_rich()
        with pytest.raises(ToolError) as exc_info:
            docx_search(session, doc_id=doc_id, query="(unclosed", regex=True)
        assert exc_info.value.code == "not_found"

    def test_empty_query_is_not_found(self) -> None:
        session, doc_id = open_rich()
        with pytest.raises(ToolError) as exc_info:
            docx_search(session, doc_id=doc_id, query="")
        assert exc_info.value.code == "not_found"

    def test_scope_range_restricts_matches(self) -> None:
        session, doc_id = open_rich()
        assert docx_search(session, doc_id=doc_id, query="Definitions")["n_matches"] == 1
        assert (
            docx_search(session, doc_id=doc_id, query="Definitions", scope="P4..P8")["n_matches"]
            == 0
        )
        assert (
            docx_search(session, doc_id=doc_id, query="Definitions", scope="P1..P3")["n_matches"]
            == 1
        )

    def test_empty_heading_supplies_empty_context(self) -> None:
        # §2a: context is the nearest paragraph whose effective style is a heading —
        # the heading style alone qualifies, so an empty heading gives context "".
        parts = dict(RICH_PARTS)
        parts["word/document.xml"] = document_xml(
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            "<w:r><w:t>Background</w:t></w:r></w:p>",
            '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr></w:p>',  # empty heading
            "<w:p><w:r><w:t>Clause under the empty heading.</w:t></w:r></w:p>",
            SECT_PR,
        )
        session = Session()
        doc_id = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())["doc_id"]
        (match,) = list(docx_search(session, doc_id=str(doc_id), query="Clause")["matches"])
        assert match["context"] == ""

    def test_comments_scope_search(self) -> None:
        session = Session()
        doc_id = docx_open(session, bytes=base64.b64encode(build_docx(COMMENT_PARTS)).decode())[
            "doc_id"
        ]
        result = docx_search(session, doc_id=str(doc_id), query="verify", scope="comments")
        assert result["n_matches"] == 1

    def test_snippet_truncation_uses_ellipsis(self) -> None:
        long_text = "start " + "x" * 100 + " needle " + "y" * 100 + " end"
        parts = dict(RICH_PARTS)
        parts["word/document.xml"] = document_xml(
            f"<w:p><w:r><w:t>{long_text}</w:t></w:r></w:p>", SECT_PR
        )
        session = Session()
        doc_id = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())["doc_id"]
        (match,) = list(docx_search(session, doc_id=str(doc_id), query="needle")["matches"])
        assert match["snippet"].startswith("…") and match["snippet"].endswith("…")
        assert "needle" in match["snippet"]


# ---------------------------------------------------------------------------
# Conformance corpus checks (against fixtures in conformance/corpus/)
# ---------------------------------------------------------------------------

CORPUS = Path(__file__).resolve().parents[2] / "conformance" / "corpus"


@pytest.mark.skipif(not CORPUS.is_dir(), reason="conformance corpus not present")
class TestConformanceParity:
    def open_corpus(self, name: str) -> tuple[Session, str]:
        session = Session()
        result = docx_open(session, path=str(CORPUS / name / "input.docx"))
        return session, str(result["doc_id"])

    def test_outline_minimal(self) -> None:
        session, doc_id = self.open_corpus("minimal")
        assert docx_outline(session, doc_id=doc_id) == {
            "outline": [
                {"anchor": "P1#515a", "level": 1, "text": "Master Services Agreement"},
                {"anchor": "P3#8775", "level": 2, "text": "Definitions"},
            ],
            "tables": [],
        }

    def test_read_window_minimal(self) -> None:
        session, doc_id = self.open_corpus("minimal")
        result = docx_read(session, doc_id=doc_id, anchor="P4#a51c", window=1)
        assert result == {
            "content": (
                "[P3#8775 H2] Definitions\n"
                "[P4#a51c] Confidential Information means all information disclosed"
                " by one party to the other.\n"
                "[P5#d27e] Each party shall protect Confidential Information with"
                " reasonable care."
            )
        }

    def test_search_split_runs(self) -> None:
        session, doc_id = self.open_corpus("split-runs")
        result = docx_search(session, doc_id=doc_id, query="five (5) years")
        assert result["n_matches"] == 1
        assert result["matches"][0]["anchor"] == "P2#d337"

    def test_redlines_projection_markers(self) -> None:
        session, doc_id = self.open_corpus("redlines")
        content = docx_read(session, doc_id=doc_id)["content"]
        lines = content.split("\n")
        assert lines[1].endswith(
            "The fee is [del by Alice] twelve percent [ins by Alice] of net revenue."
        )
        assert "ten percent" not in content
