"""Tables (``docx_table``) — algorithms.md §14.

A table is ``w:tbl`` → ``w:tblPr`` → ``w:tblGrid`` (``w:gridCol`` per column) →
``w:tr*`` → ``w:tc*`` → ``w:tcPr?`` + ``w:p+``. Every op locates the body-level
``w:tbl`` by its ``T{n}`` anchor, then splices raw bytes (§3): create, set_cells,
insert/delete row/col, merge, and style. Cell addressing is ``{r,c}`` (0-based) or
A1 (base-26 column letters, 1-based row). Nested tables are excluded by scoping each
structural scan to a parent's direct children (``max_depth=1``).
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from . import _edits, _parts, _xml
from ._anchors import build_anchor_index
from ._errors import ToolError
from ._opc import Package
from ._session import Session

#: §15 default content width in twips (A4, 1440 margins): 11906 − 1440 − 1440.
DEFAULT_CONTENT_WIDTH = 9026
HEADER_FILL = "D9D9D9"

_T_ANCHOR_RE = re.compile(r"^T([1-9][0-9]*)$")
_A1_RE = re.compile(r"^([A-Za-z]+)([1-9][0-9]*)$")
_RANGE_RE = re.compile(r"^([A-Za-z]+[1-9][0-9]*):([A-Za-z]+[1-9][0-9]*)$")


def _table_invalid(detail: str) -> ToolError:
    return ToolError(
        "anchor_invalid", detail, ["Check the table anchor 'T{n}' and cell addressing."]
    )


# ---------------------------------------------------------------------------
# Addressing (§14)
# ---------------------------------------------------------------------------


def col_to_index(letters: str) -> int:
    """Base-26 column letters → 0-based index (``A``=0 … ``Z``=25, ``AA``=26)."""
    index = 0
    for ch in letters.upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def parse_a1(ref: str) -> tuple[int, int]:
    """A1 cell ref → ``(r, c)`` 0-based (``A1`` = ``(0, 0)``, ``B2`` = ``(1, 1)``)."""
    m = _A1_RE.match(ref)
    if not m:
        raise _table_invalid(f"Malformed cell reference: {ref}.")
    return int(m.group(2)) - 1, col_to_index(m.group(1))


def cell_address(cell: Mapping[str, object]) -> tuple[int, int]:
    """``{r,c}`` wins over ``ref`` when both are present (§14)."""
    if "r" in cell and "c" in cell:
        return int(cell["r"]), int(cell["c"])  # type: ignore[call-overload]
    ref = cell.get("ref")
    if isinstance(ref, str):
        return parse_a1(ref)
    raise _table_invalid("Each cell needs {r,c} or a ref.")


# ---------------------------------------------------------------------------
# Table model (raw-byte spans, recomputed per splice)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class TableModel:
    """The current spans of a table: grid columns and rows-of-cells (direct children)."""

    tbl: _xml.Span
    grid: _xml.Span | None
    grid_cols: list[_xml.Span]
    rows: list[list[_xml.Span]]


def _direct(data: bytes, parent: _xml.Span, name: str) -> list[_xml.Span]:
    if parent.empty:
        return []
    return sorted(
        _xml.iter_elements(data, parent.inner_start, parent.inner_end, names=(name,), max_depth=1),
        key=lambda s: s.start,
    )


def read_table(data: bytes, tbl: _xml.Span) -> TableModel:
    grid = next(
        _xml.iter_elements(data, tbl.inner_start, tbl.inner_end, names=("w:tblGrid",), max_depth=1),
        None,
    )
    grid_cols = _direct(data, grid, "w:gridCol") if grid is not None else []
    rows = [_direct(data, tr, "w:tc") for tr in _direct(data, tbl, "w:tr")]
    return TableModel(tbl, grid, grid_cols, rows)


def _table_ordinal(anchor: str | None) -> int:
    if anchor is None:
        raise _table_invalid("This op requires a table anchor like 'T4'.")
    m = _T_ANCHOR_RE.match(anchor)
    if not m:
        raise _table_invalid(f"Malformed table anchor: {anchor}.")
    return int(m.group(1))


def locate_table_in(data: bytes, anchor: str | None) -> _xml.Span:
    """The body-level ``w:tbl`` span for ``T{n}`` within a specific ``data`` buffer.

    Used between splices on one in-memory buffer (where the stored package part may
    lag the local edits), counting body-level ``w:tbl`` in document order (§13).
    """
    ordinal = _table_ordinal(anchor)
    count = 0
    for child in _xml.iter_body_children(data):
        if child.name == "w:tbl":
            count += 1
            if count == ordinal:
                return child
    raise ToolError(
        "anchor_not_found",
        f"Table {anchor} not found: index out of range.",
        ["Call docx_outline to re-map table anchors."],
    )


def locate_table(package: Package, anchor: str | None) -> _xml.Span:
    """The body-level ``w:tbl`` span for a ``T{n}`` anchor (against the stored part)."""
    return locate_table_in(package.part(package.main_document_part()), anchor)


# ---------------------------------------------------------------------------
# Cell properties / merge state
# ---------------------------------------------------------------------------


def _first(data: bytes, parent: _xml.Span | None, name: str) -> _xml.Span | None:
    """The first direct child named ``name`` of ``parent`` (``None`` when absent/empty)."""
    if parent is None or parent.empty:
        return None
    return next(
        _xml.iter_elements(data, parent.inner_start, parent.inner_end, names=(name,), max_depth=1),
        None,
    )


def _tc_pr(data: bytes, tc: _xml.Span) -> _xml.Span | None:
    return _first(data, tc, "w:tcPr")


def _is_vmerge_continue(data: bytes, tc: _xml.Span) -> bool:
    """True for a ``<w:vMerge/>`` (continue) cell with no ``w:val="restart"``."""
    vmerge = _first(data, _tc_pr(data, tc), "w:vMerge")
    if vmerge is None:
        return False
    end = vmerge.end if vmerge.empty else vmerge.inner_start
    return b'w:val="restart"' not in data[vmerge.start : end]


def _grid_span(data: bytes, tc: _xml.Span) -> int:
    gs = _first(data, _tc_pr(data, tc), "w:gridSpan")
    if gs is None:
        return 1
    end = gs.end if gs.empty else gs.inner_start
    m = re.search(rb'w:val="(\d+)"', data[gs.start : end])
    return int(m.group(1)) if m else 1


# ---------------------------------------------------------------------------
# Emission helpers
# ---------------------------------------------------------------------------


def _cell_xml(width: int, text: str, *, header: bool = False) -> str:
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{HEADER_FILL}"/>' if header else ""
    tcpr = f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}</w:tcPr>'
    if not text:
        body = "<w:p/>"
    elif header:
        body = f"<w:p><w:r><w:rPr><w:b/></w:rPr>{_xml.emit_text_element(text)}</w:r></w:p>"
    else:
        body = f"<w:p><w:r>{_xml.emit_text_element(text)}</w:r></w:p>"
    return f"<w:tc>{tcpr}{body}</w:tc>"


def _grid_widths(cols: int) -> list[int]:
    """Equal integer widths summing to the §15 default; the last absorbs the remainder."""
    if cols <= 0:
        return []
    base = DEFAULT_CONTENT_WIDTH // cols
    widths = [base] * cols
    widths[-1] = DEFAULT_CONTENT_WIDTH - base * (cols - 1)
    return widths


# ---------------------------------------------------------------------------
# create (§14)
# ---------------------------------------------------------------------------


def _create_table(
    package: Package,
    after: str | None,
    rows: int,
    cols: int,
    data_rows: Sequence[Sequence[str]] | None,
    header: bool,
    style: str | None,
) -> str:
    if rows <= 0 or cols <= 0:
        raise _table_invalid("create needs positive rows and cols.")
    if after is None:
        raise _table_invalid("create requires an 'after' paragraph anchor.")
    main = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    entry = _edits.require_paragraph(entries, after)
    styled = header or style is not None
    if styled:
        _parts.ensure_style(package, "TableGrid")
    data_rows = data_rows or []
    if any(len(r) > cols for r in data_rows):
        raise _table_invalid("A data row has more cells than cols.")
    widths = _grid_widths(cols)
    tbl_pr = "".join(
        [
            "<w:tblPr>",
            '<w:tblStyle w:val="TableGrid"/>' if styled else "",
            '<w:tblW w:w="0" w:type="auto"/>',
            "</w:tblPr>",
        ]
    )
    grid = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>"
    tr_xml: list[str] = []
    for r in range(rows):
        is_header = header and r == 0
        cells = []
        row_data = data_rows[r] if r < len(data_rows) else []
        for c in range(cols):
            text = row_data[c] if c < len(row_data) else ""
            cells.append(_cell_xml(widths[c], text, header=is_header))
        tr_xml.append("<w:tr>" + "".join(cells) + "</w:tr>")
    table = f"<w:tbl>{tbl_pr}{grid}{''.join(tr_xml)}</w:tbl>"
    body = package.part(main)
    position = entry.span.end
    package.set_part(main, _xml.splice(body, [(position, position, table.encode("utf-8"))]))
    # New anchor: T{k}@after:{prev}, k = body-table ordinal in document order.
    ordinal = 0
    prev_anchor: str | None = None
    for blk in build_anchor_index(package):
        if blk.kind == "table":
            ordinal += 1
            if blk.span.start == position:
                where = f"@after:{prev_anchor}" if prev_anchor is not None else "@start"
                return f"T{ordinal}{where}"
        else:
            prev_anchor = blk.anchor
    return f"T{ordinal}"


# ---------------------------------------------------------------------------
# set_cells (§14)
# ---------------------------------------------------------------------------


def _logical_columns(data: bytes, row: list[_xml.Span]) -> list[tuple[int, _xml.Span]]:
    """``(grid-column, cell span)`` for a row, honoring gridSpan widths."""
    out: list[tuple[int, _xml.Span]] = []
    col = 0
    for tc in row:
        out.append((col, tc))
        col += _grid_span(data, tc)
    return out


def _cell_at(data: bytes, model: TableModel, r: int, c: int) -> _xml.Span:
    if r < 0 or r >= len(model.rows):
        raise _table_invalid(f"Cell row {r} is out of range.")
    row = model.rows[r]
    for start_col, tc in _logical_columns(data, row):
        span = _grid_span(data, tc)
        if start_col <= c < start_col + span:
            if c != start_col:  # inside a gridSpan: a covered cell
                raise _table_invalid(f"Cell {r},{c} is covered by a horizontal merge.")
            if _is_vmerge_continue(data, tc):
                raise _table_invalid(f"Cell {r},{c} is covered by a vertical merge.")
            return tc
    raise _table_invalid(f"Cell column {c} is out of range.")


def _set_cells(package: Package, anchor: str | None, cells: Sequence[Mapping[str, object]]) -> int:
    main = package.main_document_part()
    affected = 0
    for cell in cells:
        r, c = cell_address(cell)
        text = str(cell.get("text", ""))
        data = package.part(main)
        tbl = locate_table(package, anchor)
        model = read_table(data, tbl)
        tc = _cell_at(data, model, r, c)
        # Keep w:tcPr; replace the cell's paragraphs with one carrying the first p's pPr.
        tcpr = _tc_pr(data, tc)
        first_p = next(
            _xml.iter_elements(data, tc.inner_start, tc.inner_end, names=("w:p",), max_depth=1),
            None,
        )
        ppr_xml = ""
        if first_p is not None and not first_p.empty:
            ppr = next(
                _xml.iter_elements(
                    data, first_p.inner_start, first_p.inner_end, names=("w:pPr",), max_depth=1
                ),
                None,
            )
            if ppr is not None:
                ppr_xml = data[ppr.start : ppr.end].decode("utf-8")
        run = f"<w:r>{_xml.emit_text_element(text)}</w:r>" if text else ""
        new_p = f"<w:p>{ppr_xml}{run}</w:p>"
        tcpr_xml = data[tcpr.start : tcpr.end].decode("utf-8") if tcpr is not None else ""
        new_inner = f"{tcpr_xml}{new_p}".encode()
        package.set_part(main, _xml.splice(data, [(tc.inner_start, tc.inner_end, new_inner)]))
        affected += 1
    return affected


# ---------------------------------------------------------------------------
# insert_row / insert_col / delete_row / delete_col (§14)
# ---------------------------------------------------------------------------


def _cell_props_clone(data: bytes, tc: _xml.Span) -> str:
    """A blank-text clone of ``tc``: its ``w:tcPr`` verbatim plus an empty paragraph."""
    tcpr = _tc_pr(data, tc)
    tcpr_xml = data[tcpr.start : tcpr.end].decode("utf-8") if tcpr is not None else ""
    return f"<w:tc>{tcpr_xml}<w:p/></w:tc>"


def _insert_row(package: Package, anchor: str | None, at: int) -> None:
    main = package.main_document_part()
    data = package.part(main)
    tbl = locate_table(package, anchor)
    model = read_table(data, tbl)
    n_rows = len(model.rows)
    clone_idx = min(at, n_rows - 1)
    if clone_idx < 0:
        raise _table_invalid("Cannot insert a row into a table with no rows.")
    template = model.rows[clone_idx]
    cells = "".join(_cell_props_clone(data, tc) for tc in template)
    new_row = f"<w:tr>{cells}</w:tr>".encode()
    tr_spans = _direct(data, tbl, "w:tr")
    pos = tr_spans[-1].end if at >= n_rows else tr_spans[at].start  # at == rows appends
    package.set_part(main, _xml.splice(data, [(pos, pos, new_row)]))


def _refloor_grid(package: Package, anchor: str | None) -> None:
    """Re-floor all gridCol widths and per-row tcW across the current grid."""
    main = package.main_document_part()
    data = package.part(main)
    tbl = locate_table(package, anchor)
    model = read_table(data, tbl)
    cols = len(model.grid_cols)
    if cols == 0:
        return
    widths = _grid_widths(cols)
    edits: list[tuple[int, int, bytes]] = []
    for i, gc in enumerate(model.grid_cols):
        end = gc.end if gc.empty else gc.inner_start
        m = re.search(rb'(w:w=")(\d+)(")', data[gc.start : end])
        if m:
            base = gc.start
            edits.append((base + m.start(2), base + m.end(2), str(widths[i]).encode()))
    if edits:
        package.set_part(main, _xml.splice(data, edits))


def _insert_col(package: Package, anchor: str | None, at: int) -> None:
    main = package.main_document_part()
    data = package.part(main)
    tbl = locate_table(package, anchor)
    model = read_table(data, tbl)
    cols = len(model.grid_cols)
    new_cols = cols + 1
    widths = _grid_widths(new_cols)
    # 1) add a gridCol at index `at`.
    if model.grid is not None:
        if at >= cols:
            pos = model.grid_cols[-1].end if model.grid_cols else model.grid.inner_start
        else:
            pos = model.grid_cols[at].start
        data = _xml.splice(data, [(pos, pos, f'<w:gridCol w:w="{widths[at]}"/>'.encode())])
    package.set_part(main, data)
    # 2) add one blank w:tc per row at column `at` (right-to-left rows to keep offsets).
    data = package.part(main)
    tbl = locate_table(package, anchor)
    model = read_table(data, tbl)
    edits: list[tuple[int, int, bytes]] = []
    cell_xml = _cell_xml(widths[at], "").encode("utf-8")
    for row in model.rows:
        cell_pos = (row[-1].end if row else None) if at >= len(row) else row[at].start
        if cell_pos is None:
            continue
        edits.append((cell_pos, cell_pos, cell_xml))
    package.set_part(main, _xml.splice(data, edits))
    _refloor_grid(package, anchor)


def _delete_row(package: Package, anchor: str | None, at: int) -> None:
    main = package.main_document_part()
    data = package.part(main)
    tbl = locate_table(package, anchor)
    tr_spans = _direct(data, tbl, "w:tr")
    if at < 0 or at >= len(tr_spans):
        raise _table_invalid(f"Row {at} is out of range.")
    target = tr_spans[at]
    model = read_table(data, tbl)
    # vMerge-origin promotion: if a cell in this row is a vMerge restart, the next row's
    # continuation cell at the same column becomes the new origin (drop its w:vMerge).
    promote: list[tuple[int, int, bytes]] = []
    if at + 1 < len(tr_spans):
        row = model.rows[at]
        next_row = model.rows[at + 1]
        next_cols = {col: tc for col, tc in _logical_columns(data, next_row)}
        for col, tc in _logical_columns(data, row):
            tcpr = _tc_pr(data, tc)
            if tcpr is None or tcpr.empty:
                continue
            vmerge = next(
                _xml.iter_elements(
                    data, tcpr.inner_start, tcpr.inner_end, names=("w:vMerge",), max_depth=1
                ),
                None,
            )
            if vmerge is None:
                continue
            end = vmerge.end if vmerge.empty else vmerge.inner_start
            if b'w:val="restart"' not in data[vmerge.start : end]:
                continue
            cont = next_cols.get(col)
            if cont is not None and _is_vmerge_continue(data, cont):
                cvm = next(
                    _xml.iter_elements(
                        data,
                        _tc_pr(data, cont).inner_start,  # type: ignore[union-attr]
                        _tc_pr(data, cont).inner_end,  # type: ignore[union-attr]
                        names=("w:vMerge",),
                        max_depth=1,
                    ),
                    None,
                )
                if cvm is not None:
                    promote.append((cvm.start, cvm.end, b""))
    edits = [(target.start, target.end, b""), *promote]
    package.set_part(main, _xml.splice(data, edits))


def _delete_col(package: Package, anchor: str | None, at: int) -> None:
    main = package.main_document_part()
    data = package.part(main)
    tbl = locate_table(package, anchor)
    model = read_table(data, tbl)
    cols = len(model.grid_cols)
    if at < 0 or at >= cols:
        raise _table_invalid(f"Column {at} is out of range.")
    edits: list[tuple[int, int, bytes]] = [
        (model.grid_cols[at].start, model.grid_cols[at].end, b"")
    ]
    for row in model.rows:
        for start_col, tc in _logical_columns(data, row):
            span = _grid_span(data, tc)
            if start_col <= at < start_col + span:
                if span == 1:
                    edits.append((tc.start, tc.end, b""))
                # a spanned cell shrinks: leave it (gridSpan re-floor handled elsewhere);
                # MVP removes only single-column cells at `at`.
                break
    package.set_part(main, _xml.splice(data, edits))
    _refloor_grid(package, anchor)


# ---------------------------------------------------------------------------
# merge (§14)
# ---------------------------------------------------------------------------


def _ensure_tcpr(data: bytes, tc: _xml.Span) -> tuple[bytes, int, bool]:
    """``(data, insert-point, has_tcpr)`` ensuring a ``w:tcPr`` to splice merge marks into.

    When the cell has no ``w:tcPr`` an empty one is created as the first child; the
    returned insert-point is just inside it.
    """
    tcpr = _tc_pr(data, tc)
    if tcpr is not None and not tcpr.empty:
        return data, tcpr.inner_start, True
    if tcpr is not None and tcpr.empty:
        # <w:tcPr/> → <w:tcPr></w:tcPr>
        data = _xml.splice(data, [(tcpr.start, tcpr.end, b"<w:tcPr></w:tcPr>")])
        tcpr = _tc_pr(data, _re_tc(data, tc.start))
        return data, tcpr.inner_start, True  # type: ignore[union-attr]
    data = _xml.splice(data, [(tc.inner_start, tc.inner_start, b"<w:tcPr></w:tcPr>")])
    tcpr = _tc_pr(data, _re_tc(data, tc.start))
    return data, tcpr.inner_start, True  # type: ignore[union-attr]


def _re_tc(data: bytes, start: int) -> _xml.Span:
    for tc in _xml.iter_elements(data, start, names=("w:tc",)):
        if tc.start == start:
            return tc
    raise KeyError(start)


def _merge(package: Package, anchor: str | None, range_ref: str | None) -> None:
    if range_ref is None:
        raise _table_invalid("merge requires a range like 'A1:C1'.")
    m = _RANGE_RE.match(range_ref)
    if not m:
        raise _table_invalid(f"Malformed merge range: {range_ref}.")
    (r0, c0), (r1, c1) = parse_a1(m.group(1)), parse_a1(m.group(2))
    r0, r1 = sorted((r0, r1))
    c0, c1 = sorted((c0, c1))
    main = package.main_document_part()
    horizontal = c1 > c0
    vertical = r1 > r0
    # Horizontal: per spanned row, set gridSpan on the left cell, remove covered cells.
    if horizontal:
        span = c1 - c0 + 1
        for r in range(r0, r1 + 1):
            data = package.part(main)
            tbl = locate_table_in(data, anchor)
            model = read_table(data, tbl)
            if r >= len(model.rows):
                continue
            logical = _logical_columns(data, model.rows[r])
            left_phys = next((i for i, (col, _) in enumerate(logical) if col == c0), None)
            if left_phys is None:
                continue
            covered_phys = [i for i, (col, _) in enumerate(logical) if c0 < col <= c1]
            left = logical[left_phys][1]
            data, insert_at, _ = _ensure_tcpr(data, left)
            data = _xml.splice(
                data, [(insert_at, insert_at, f'<w:gridSpan w:val="{span}"/>'.encode())]
            )
            # Re-read by physical index after the gridSpan splice shifted offsets.
            tbl = locate_table_in(data, anchor)
            model = read_table(data, tbl)
            row = model.rows[r]
            removes = [(row[i].start, row[i].end, b"") for i in covered_phys if i < len(row)]
            if removes:
                data = _xml.splice(data, removes)
            package.set_part(main, data)
    # Vertical: restart on the top cell (left column), continue below.
    if vertical:
        for r in range(r0, r1 + 1):
            data = package.part(main)
            tbl = locate_table_in(data, anchor)
            model = read_table(data, tbl)
            if r >= len(model.rows):
                continue
            logical = _logical_columns(data, model.rows[r])
            top: _xml.Span | None = next((tc for col, tc in logical if col == c0), None)
            if top is None:
                continue
            data, insert_at, _ = _ensure_tcpr(data, top)
            vmerge = '<w:vMerge w:val="restart"/>' if r == r0 else "<w:vMerge/>"
            data = _xml.splice(data, [(insert_at, insert_at, vmerge.encode("utf-8"))])
            package.set_part(main, data)


# ---------------------------------------------------------------------------
# style (§14)
# ---------------------------------------------------------------------------


def _style_table(package: Package, anchor: str | None, style: str | None) -> None:
    _parts.ensure_style(package, "TableGrid")
    main = package.main_document_part()
    data = package.part(main)
    tbl = locate_table(package, anchor)
    tbl_pr = next(
        _xml.iter_elements(data, tbl.inner_start, tbl.inner_end, names=("w:tblPr",), max_depth=1),
        None,
    )
    style_xml = '<w:tblStyle w:val="TableGrid"/>'
    if tbl_pr is None:
        data = _xml.splice(
            data, [(tbl.inner_start, tbl.inner_start, f"<w:tblPr>{style_xml}</w:tblPr>".encode())]
        )
    else:
        existing = next(
            _xml.iter_elements(
                data, tbl_pr.inner_start, tbl_pr.inner_end, names=("w:tblStyle",), max_depth=1
            ),
            None,
        )
        if existing is not None:
            data = _xml.splice(data, [(existing.start, existing.end, style_xml.encode("utf-8"))])
        else:
            data = _xml.splice(
                data, [(tbl_pr.inner_start, tbl_pr.inner_start, style_xml.encode("utf-8"))]
            )
    package.set_part(main, data)


# ---------------------------------------------------------------------------
# docx_table
# ---------------------------------------------------------------------------


def docx_table(
    session: Session,
    *,
    doc_id: str,
    op: str,
    anchor: str | None = None,
    after: str | None = None,
    rows: int | None = None,
    cols: int | None = None,
    data: Sequence[Sequence[str]] | None = None,
    header: bool = False,
    cells: Sequence[Mapping[str, object]] | None = None,
    at: int | None = None,
    range: str | None = None,  # noqa: A002 - wire name pinned by the tool schema
    style: str | None = None,
    track_changes: bool = False,
    author: str | None = None,
) -> dict[str, object]:
    """All table operations (§14): create, set_cells, insert/delete row/col, merge, style."""
    doc = session.get(doc_id)
    package = doc.package
    if op == "create":
        new_anchor = _create_table(
            package, after, rows or 0, cols or 0, data, header, style
        )
        doc.mark_dirty()
        return {"new_anchor": new_anchor, "note": "Created table."}
    if op == "set_cells":
        n = _set_cells(package, anchor, cells or [])
        doc.mark_dirty()
        return {"note": f"Wrote {n} cell(s)."}
    if op == "insert_row":
        _insert_row(package, anchor, at if at is not None else 0)
        doc.mark_dirty()
        return {"note": "Inserted row."}
    if op == "insert_col":
        _insert_col(package, anchor, at if at is not None else 0)
        doc.mark_dirty()
        return {"note": "Inserted column."}
    if op == "delete_row":
        _delete_row(package, anchor, at if at is not None else 0)
        doc.mark_dirty()
        return {"note": "Deleted row."}
    if op == "delete_col":
        _delete_col(package, anchor, at if at is not None else 0)
        doc.mark_dirty()
        return {"note": "Deleted column."}
    if op == "merge":
        _merge(package, anchor, range)
        doc.mark_dirty()
        return {"note": "Merged cells."}
    if op == "style":
        _style_table(package, anchor, style)
        doc.mark_dirty()
        return {"note": "Styled table."}
    raise _table_invalid(f"Unknown table op: {op}.")
