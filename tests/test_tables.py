"""Table-surface tests: docx_table (algorithms.md §14).

Covers all ops (create/set_cells/insert_row/insert_col/delete_row/delete_col/merge/
style), A1↔{r,c} addressing parity, the round-trip through projection and the save
validator, and the gridSpan/vMerge structure a merge produces.
"""

from __future__ import annotations

import base64

import pytest
from conftest import FIXTURE_PARTS, SECT_PR, build_docx, document_xml

from docxengine import (
    Session,
    ToolError,
    docx_open,
    docx_outline,
    docx_read,
    docx_table,
    docx_validate,
    paragraph_anchor,
)
from docxengine._tables import col_to_index, parse_a1

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


def main_xml(session: Session, doc_id: str) -> str:
    package = session.get(doc_id).package
    return package.part(package.main_document_part()).decode("utf-8")


# ---------------------------------------------------------------------------
# Addressing (§14)
# ---------------------------------------------------------------------------


class TestAddressing:
    @pytest.mark.parametrize(
        ("letters", "index"),
        [("A", 0), ("B", 1), ("Z", 25), ("AA", 26), ("AB", 27), ("BA", 52)],
    )
    def test_column_letters(self, letters: str, index: int) -> None:
        assert col_to_index(letters) == index

    @pytest.mark.parametrize(
        ("ref", "rc"),
        [("A1", (0, 0)), ("B2", (1, 1)), ("C3", (2, 2)), ("AA10", (9, 26))],
    )
    def test_a1_to_rc(self, ref: str, rc: tuple[int, int]) -> None:
        assert parse_a1(ref) == rc

    def test_a1_and_rc_addressing_agree(self) -> None:
        # Writing via {r,c} and via the equivalent A1 ref must hit the same cell.
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=2)
        docx_table(
            session, doc_id=doc_id, op="set_cells", anchor="T1", cells=[{"ref": "B2", "text": "X"}]
        )
        via_ref = main_xml(session, doc_id)

        session2, doc_id2 = open_docx()
        docx_table(session2, doc_id=doc_id2, op="create", after=A2, rows=2, cols=2)
        docx_table(
            session2,
            doc_id=doc_id2,
            op="set_cells",
            anchor="T1",
            cells=[{"r": 1, "c": 1, "text": "X"}],
        )
        via_rc = main_xml(session2, doc_id2)
        assert via_ref == via_rc

    def test_rc_wins_over_ref(self) -> None:
        # {r,c} is authoritative when both are present (§14).
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=2)
        docx_table(
            session,
            doc_id=doc_id,
            op="set_cells",
            anchor="T1",
            cells=[{"r": 0, "c": 0, "ref": "B2", "text": "TOPLEFT"}],
        )
        content = docx_read(session, doc_id=doc_id)["content"]
        # The text landed in the header (top-left), not B2.
        assert "| TOPLEFT |  |" in content


# ---------------------------------------------------------------------------
# create (§14)
# ---------------------------------------------------------------------------


class TestCreate:
    def test_create_returns_table_anchor_with_after_token(self) -> None:
        session, doc_id = open_docx()
        result = docx_table(
            session,
            doc_id=doc_id,
            op="create",
            after=A2,
            rows=2,
            cols=2,
            data=[["Term", "Value"], ["Fee", "$100"]],
            header=True,
        )
        assert result["new_anchor"] == f"T1@after:{A2}"

    def test_created_table_projects_as_markdown(self) -> None:
        session, doc_id = open_docx()
        docx_table(
            session,
            doc_id=doc_id,
            op="create",
            after=A2,
            rows=2,
            cols=2,
            data=[["Term", "Value"], ["Fee", "$100"]],
            header=True,
        )
        content = docx_read(session, doc_id=doc_id)["content"]
        assert "[T1 2×2 @after:P2#d337]" in content
        assert "| Term | Value |" in content
        assert "| --- | --- |" in content
        assert "| Fee | $100 |" in content

    def test_created_table_round_trips_through_outline(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=3, cols=2)
        tables = docx_outline(session, doc_id=doc_id)["tables"]
        assert tables == [{"anchor": "T1", "dims": "3×2", "after": A2}]

    def test_grid_widths_floor_with_remainder_on_last(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=1, cols=3)
        xml = main_xml(session, doc_id)
        # floor(9026/3) = 3008, last absorbs the remainder: 9026 - 3008*2 = 3010.
        assert '<w:gridCol w:w="3008"/><w:gridCol w:w="3008"/><w:gridCol w:w="3010"/>' in xml

    def test_header_row_styles_bold_and_shading(self) -> None:
        session, doc_id = open_docx()
        docx_table(
            session,
            doc_id=doc_id,
            op="create",
            after=A2,
            rows=2,
            cols=1,
            data=[["H"], ["b"]],
            header=True,
        )
        xml = main_xml(session, doc_id)
        assert '<w:tblStyle w:val="TableGrid"/>' in xml
        assert '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' in xml
        assert "<w:rPr><w:b/></w:rPr>" in xml
        # The TableGrid style was ensured.
        assert b'w:styleId="TableGrid"' in session.get(doc_id).package.part("word/styles.xml")

    def test_empty_cell_text_is_empty_paragraph(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=1, cols=2, data=[["x"]])
        xml = main_xml(session, doc_id)
        assert "<w:p/>" in xml  # the unfilled second cell

    def test_trailing_whitespace_preserves_space(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=1, cols=1, data=[["$100 "]])
        xml = main_xml(session, doc_id)
        assert '<w:t xml:space="preserve">$100 </w:t>' in xml

    def test_data_overflow_is_anchor_invalid(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_table(
                session, doc_id=doc_id, op="create", after=A2, rows=1, cols=1, data=[["a", "b"]]
            )
        assert err.value.code == "anchor_invalid"

    def test_created_table_passes_validator(self) -> None:
        session, doc_id = open_docx()
        docx_table(
            session,
            doc_id=doc_id,
            op="create",
            after=A2,
            rows=2,
            cols=2,
            data=[["a", "b"]],
            header=True,
        )
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# set_cells (§14)
# ---------------------------------------------------------------------------


class TestSetCells:
    def test_set_cells_replaces_text_keeping_tcpr(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=2)
        docx_table(
            session,
            doc_id=doc_id,
            op="set_cells",
            anchor="T1",
            cells=[{"ref": "A1", "text": "Term"}, {"r": 1, "c": 1, "text": "$100"}],
        )
        content = docx_read(session, doc_id=doc_id)["content"]
        assert "| Term |  |" in content
        assert "|  | $100 |" in content
        # tcW (the cell property) survives the text rewrite.
        assert '<w:tcW w:w="4513" w:type="dxa"/>' in main_xml(session, doc_id)

    def test_set_covered_cell_is_anchor_invalid(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=3)
        docx_table(session, doc_id=doc_id, op="merge", anchor="T1", range="A1:C1")
        with pytest.raises(ToolError) as err:
            docx_table(
                session,
                doc_id=doc_id,
                op="set_cells",
                anchor="T1",
                cells=[{"ref": "B1", "text": "x"}],
            )
        assert err.value.code == "anchor_invalid"


# ---------------------------------------------------------------------------
# insert / delete row & col (§14)
# ---------------------------------------------------------------------------


class TestRowColOps:
    def test_insert_row_appends_blank_row(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=2)
        docx_table(session, doc_id=doc_id, op="insert_row", anchor="T1", at=2)
        assert docx_outline(session, doc_id=doc_id)["tables"][0]["dims"] == "3×2"

    def test_insert_row_at_zero(self) -> None:
        session, doc_id = open_docx()
        docx_table(
            session, doc_id=doc_id, op="create", after=A2, rows=2, cols=1, data=[["top"], ["bot"]]
        )
        docx_table(session, doc_id=doc_id, op="insert_row", anchor="T1", at=0)
        content = docx_read(session, doc_id=doc_id)["content"]
        # The new blank row became the header (row 1).
        lines = [ln for ln in content.splitlines() if ln.startswith("|")]
        assert lines[0] == "|  |"

    def test_insert_col_widens_grid(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=2)
        docx_table(session, doc_id=doc_id, op="insert_col", anchor="T1", at=1)
        assert docx_outline(session, doc_id=doc_id)["tables"][0]["dims"] == "2×3"
        assert docx_validate(session, doc_id=doc_id)["valid"] is True

    def test_delete_row(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=3, cols=2)
        docx_table(session, doc_id=doc_id, op="delete_row", anchor="T1", at=1)
        assert docx_outline(session, doc_id=doc_id)["tables"][0]["dims"] == "2×2"

    def test_delete_col(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=3)
        docx_table(session, doc_id=doc_id, op="delete_col", anchor="T1", at=1)
        assert docx_outline(session, doc_id=doc_id)["tables"][0]["dims"] == "2×2"
        assert docx_validate(session, doc_id=doc_id)["valid"] is True

    def test_delete_vmerge_origin_promotes_continuation(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=3, cols=2)
        docx_table(session, doc_id=doc_id, op="merge", anchor="T1", range="A1:A3")
        docx_table(session, doc_id=doc_id, op="delete_row", anchor="T1", at=0)
        xml = main_xml(session, doc_id)
        # The promoted origin no longer carries a continuation w:vMerge.
        assert "<w:vMerge/>" in xml  # the third row still continues
        # The first remaining row's left cell lost its plain continue marker (promoted).
        first_row = xml.split("<w:tr>")[1]
        assert "<w:vMerge/>" not in first_row.split("</w:tr>")[0]


# ---------------------------------------------------------------------------
# merge (§14)
# ---------------------------------------------------------------------------


class TestMerge:
    def test_horizontal_merge_sets_gridspan_and_removes_cells(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=3)
        docx_table(session, doc_id=doc_id, op="merge", anchor="T1", range="A1:C1")
        xml = main_xml(session, doc_id)
        first_row = xml.split("<w:tr>")[1].split("</w:tr>")[0]
        assert '<w:gridSpan w:val="3"/>' in first_row
        assert first_row.count("<w:tc>") == 1  # the two covered cells were removed
        assert docx_validate(session, doc_id=doc_id)["valid"] is True

    def test_vertical_merge_restart_then_continue(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=3, cols=2)
        docx_table(session, doc_id=doc_id, op="merge", anchor="T1", range="A1:A3")
        xml = main_xml(session, doc_id)
        assert '<w:vMerge w:val="restart"/>' in xml
        assert xml.count("<w:vMerge/>") == 2  # two continuation rows
        assert docx_validate(session, doc_id=doc_id)["valid"] is True

    def test_rectangular_merge(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=3, cols=3)
        docx_table(session, doc_id=doc_id, op="merge", anchor="T1", range="A1:B2")
        xml = main_xml(session, doc_id)
        assert '<w:gridSpan w:val="2"/>' in xml
        assert '<w:vMerge w:val="restart"/>' in xml
        # §14 pin: vMerge (written second) precedes gridSpan in the left cell's tcPr.
        assert '<w:vMerge w:val="restart"/><w:gridSpan w:val="2"/>' in xml
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# style (§14)
# ---------------------------------------------------------------------------


class TestStyleOp:
    def test_style_adds_table_grid(self) -> None:
        session, doc_id = open_docx()
        docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=2)
        docx_table(session, doc_id=doc_id, op="style", anchor="T1", style="Table Grid")
        xml = main_xml(session, doc_id)
        assert '<w:tblStyle w:val="TableGrid"/>' in xml
        assert b'w:styleId="TableGrid"' in session.get(doc_id).package.part("word/styles.xml")
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


class TestErrors:
    def test_op_on_missing_table_is_anchor_not_found(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_table(
                session,
                doc_id=doc_id,
                op="set_cells",
                anchor="T9",
                cells=[{"ref": "A1", "text": "x"}],
            )
        assert err.value.code == "anchor_not_found"

    def test_malformed_table_anchor_is_anchor_invalid(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_table(session, doc_id=doc_id, op="delete_row", anchor="X1", at=0)
        assert err.value.code == "anchor_invalid"

    def test_create_on_stale_after_anchor_is_stale(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_table(session, doc_id=doc_id, op="create", after="P2#0000", rows=1, cols=1)
        assert err.value.code == "anchor_stale"


# ---------------------------------------------------------------------------
# Independent anchor sequences (§13)
# ---------------------------------------------------------------------------


def test_table_create_does_not_shift_paragraph_anchors() -> None:
    parts = dict(FIXTURE_PARTS)
    parts["word/document.xml"] = document_xml(
        "<w:p><w:r><w:t>One</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Two</w:t></w:r></w:p>",
        SECT_PR,
    )
    session, doc_id = open_docx(parts)
    before = paragraph_anchor(2, "Two")
    docx_table(session, doc_id=doc_id, op="create", after=before, rows=1, cols=1)
    # P2 still resolves to the same hash; the new table is T1 (independent sequence).
    content = docx_read(session, doc_id=doc_id, anchor=before, window=0)["content"]
    assert content.startswith(f"[{before}]")
