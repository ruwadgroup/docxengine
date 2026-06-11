/**
 * Tables (`docx_table`) per spec/algorithms.md §14.
 *
 * A table is `w:tbl → w:tblPr → w:tblGrid → w:tr* → w:tc*`. Ops: create,
 * set_cells, insert_row, insert_col, delete_row, delete_col, merge, style.
 * Cells address by zero-based {r,c} or A1 (base-26 columns, 1-based rows).
 * create distributes the §15 default content width (9026 twips) across columns,
 * the last column absorbing the remainder.
 *
 * All emission obeys §3 (splice; §3 escaping; xml:space rule). The Python twin
 * (`_tables.py`) is the byte-parity reference.
 */
import { type AnchorEntry } from "./anchors.js";
import { ToolError } from "./errors.js";
import {
  TABLE_GRID_STYLE,
  anchorInvalid,
  anchorNotFound,
  anchorStale,
  ensureStyle,
} from "./phase2common.js";
import type { DocHandle, Session } from "./session.js";
import {
  type ElementSlice,
  type SpliceEdit,
  attrs,
  childElements,
  elementExtent,
  emitTextElement,
  emitTextRuns,
  escapeAttr,
  findElement,
  getAttr,
  nextTag,
  splice,
  spliceAll,
} from "./xmlscan.js";

type ResponseFormat = "concise" | "detailed";

/** §15 default content width (A4, 1440 margins) in twips. */
const DEFAULT_CONTENT_WIDTH = 9026;
const HEADER_SHADING = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>';

const T_ANCHOR_RE = /^T([1-9][0-9]*)$/;
const A1_RE = /^([A-Za-z]+)([1-9][0-9]*)$/;
const RANGE_RE = /^([A-Za-z]+[1-9][0-9]*):([A-Za-z]+[1-9][0-9]*)$/;

// ---------------------------------------------------------------------------
// A1 addressing (§14)
// ---------------------------------------------------------------------------

/** Base-26 column letters → 0-based index (A=0 … Z=25, AA=26). */
export function colLettersToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64); // 'A' = 65 → 1
  }
  return n - 1;
}

/** A1 ref → {r, c}; `B2` → {r:1, c:1}. */
export function parseA1(ref: string): { r: number; c: number } {
  const m = A1_RE.exec(ref);
  if (!m) throw anchorInvalid(`Malformed cell ref: ${ref}.`);
  return { c: colLettersToIndex(m[1] as string), r: Number(m[2]) - 1 };
}

// ---------------------------------------------------------------------------
// Table model
// ---------------------------------------------------------------------------

interface TableModel {
  /** The full `w:tbl` slice in the current document text. */
  tbl: ElementSlice;
  tblPr: ElementSlice | null;
  grid: ElementSlice | null;
  rows: ElementSlice[]; // w:tr
}

function parseTable(xml: string, tbl: ElementSlice): TableModel {
  const kids = tbl.selfClosed ? [] : childElements(xml, tbl.contentStart, tbl.contentEnd);
  return {
    tbl,
    tblPr: kids.find((k) => k.name === "w:tblPr") ?? null,
    grid: kids.find((k) => k.name === "w:tblGrid") ?? null,
    rows: kids.filter((k) => k.name === "w:tr"),
  };
}

/** Direct `w:tc` children of a row, in document order. */
function rowCells(xml: string, tr: ElementSlice): ElementSlice[] {
  if (tr.selfClosed) return [];
  return childElements(xml, tr.contentStart, tr.contentEnd).filter((k) => k.name === "w:tc");
}

/** Grid column widths (twips) from w:tblGrid, in order. */
function gridWidths(xml: string, grid: ElementSlice | null): number[] {
  if (!grid || grid.selfClosed) return [];
  return childElements(xml, grid.contentStart, grid.contentEnd)
    .filter((k) => k.name === "w:gridCol")
    .map((g) => {
      const w = getAttr(xml, tagOf(g), "w:w");
      return w !== undefined && /^[0-9]+$/.test(w) ? Number(w) : 0;
    });
}

function tagOf(el: ElementSlice) {
  return {
    kind: el.selfClosed ? ("empty" as const) : ("start" as const),
    name: el.name,
    start: el.start,
    end: el.startTagEnd,
    nameEnd: el.nameEnd,
  };
}

/** Distribute `total` across `cols` columns; last absorbs the remainder (§14). */
function distributeWidths(total: number, cols: number): number[] {
  if (cols <= 0) return [];
  const base = Math.floor(total / cols);
  const widths = new Array(cols).fill(base);
  widths[cols - 1] = total - base * (cols - 1);
  return widths;
}

// ---------------------------------------------------------------------------
// Cell emission
// ---------------------------------------------------------------------------

/** A cell with a single text paragraph; empty text → an empty `<w:p/>`. */
function cellXml(width: number, text: string, header: boolean): string {
  const tcPr = `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${header ? HEADER_SHADING : ""}</w:tcPr>`;
  const para = cellParagraph(text, header);
  return `<w:tc>${tcPr}${para}</w:tc>`;
}

function cellParagraph(text: string, header: boolean): string {
  if (text === "") return "<w:p/>";
  const rPr = header ? "<w:rPr><w:b/></w:rPr>" : "";
  return `<w:p>${emitTextRuns(text, rPr)}</w:p>`;
}

// ---------------------------------------------------------------------------
// docx_table
// ---------------------------------------------------------------------------

export interface DocxTableCell {
  r?: number | undefined;
  c?: number | undefined;
  ref?: string | undefined;
  text: string;
}

export interface DocxTableProps {
  borders?: boolean | undefined;
  border_weight?: number | undefined;
  border_color?: string | undefined;
  width_pct?: number | undefined;
  align?: string | undefined;
  col_widths?: number[] | undefined;
  shading?: { header?: string | null; all?: string | null } | undefined;
}

export interface DocxTableArgs {
  doc_id: string;
  op:
    | "create"
    | "set_cells"
    | "insert_row"
    | "insert_col"
    | "delete_row"
    | "delete_col"
    | "merge"
    | "style"
    | "delete";
  anchor?: string | undefined;
  after?: string | undefined;
  rows?: number | undefined;
  cols?: number | undefined;
  data?: string[][] | undefined;
  header?: boolean | undefined;
  cells?: DocxTableCell[] | undefined;
  at?: number | undefined;
  range?: string | undefined;
  style?: string | undefined;
  props?: DocxTableProps | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
}

export interface DocxTableResult {
  new_anchor?: string;
  note?: string;
}

function bodyEntries(doc: DocHandle): AnchorEntry[] {
  return doc.anchorIndex();
}

function paragraphEntries(doc: DocHandle): AnchorEntry[] {
  return doc.anchorIndex().filter((e) => e.kind === "p");
}

const FULL_PARA_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;

function requireParagraph(doc: DocHandle, anchor: string): AnchorEntry {
  const m = FULL_PARA_RE.exec(anchor);
  if (!m) throw anchorInvalid(`Malformed anchor string: ${anchor}.`);
  const entry = paragraphEntries(doc)[Number(m[1]) - 1];
  if (entry === undefined) throw anchorNotFound(anchor);
  if (entry.anchor !== anchor) throw anchorStale(anchor);
  return entry;
}

/** Resolve a `T{n}` anchor to its body-level table entry. */
function requireTable(doc: DocHandle, anchor: string): AnchorEntry {
  const m = T_ANCHOR_RE.exec(anchor);
  if (!m) throw anchorInvalid(`Malformed table anchor: ${anchor}.`);
  const ord = Number(m[1]);
  const entry = bodyEntries(doc).find((e) => e.kind === "tbl" && e.ordinal === ord);
  if (entry === undefined) throw anchorNotFound(anchor);
  return entry;
}

export function docxTable(session: Session, args: DocxTableArgs): DocxTableResult {
  const doc = session.get(args.doc_id);
  switch (args.op) {
    case "create":
      return tableCreate(doc, args);
    case "set_cells":
      return tableSetCells(doc, args);
    case "insert_row":
      return tableInsertRow(doc, args);
    case "insert_col":
      return tableInsertCol(doc, args);
    case "delete_row":
      return tableDeleteRow(doc, args);
    case "delete_col":
      return tableDeleteCol(doc, args);
    case "merge":
      return tableMerge(doc, args);
    case "style":
      return tableStyle(doc, args);
    case "delete":
      return tableDelete(doc, args);
    default:
      throw new ToolError("invalid_args", `docx_table: unknown op ${String(args.op)}.`, []);
  }
}

/** §2 `@after:` token for a table that follows the paragraph `prev`. */
function afterToken(prevAnchor: string | null): string {
  return prevAnchor === null ? "@start" : `@after:${prevAnchor}`;
}

// --- create -----------------------------------------------------------------

function tableCreate(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  if (args.after === undefined) throw anchorInvalid("op 'create' requires after.");
  const entry = requireParagraph(doc, args.after);
  const rows = Math.max(0, Math.trunc(args.rows ?? 0));
  const cols = Math.max(0, Math.trunc(args.cols ?? 0));
  if (rows <= 0 || cols <= 0) {
    throw anchorInvalid("op 'create' requires positive rows and cols.");
  }
  const data = args.data ?? [];
  if (data.length > rows) throw anchorInvalid("data has more rows than the table.");
  for (const row of data) {
    if (row.length > cols) throw anchorInvalid("data row has more cells than columns.");
  }
  const header = args.header === true;
  const styled = header || args.style !== undefined;
  if (styled) ensureStyle(doc.pkg, "TableGrid", TABLE_GRID_STYLE);

  const widths = distributeWidths(DEFAULT_CONTENT_WIDTH, cols);
  const gridCols = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const tblPr =
    `<w:tblPr>${styled ? '<w:tblStyle w:val="TableGrid"/>' : ""}` +
    `<w:tblW w:w="0" w:type="auto"/></w:tblPr>`;

  const trs: string[] = [];
  for (let r = 0; r < rows; r++) {
    const isHeader = header && r === 0;
    const cells: string[] = [];
    for (let c = 0; c < cols; c++) {
      const text = data[r]?.[c] ?? "";
      cells.push(cellXml(widths[c] as number, text, isHeader));
    }
    trs.push(`<w:tr>${cells.join("")}</w:tr>`);
  }
  const tbl = `<w:tbl>${tblPr}<w:tblGrid>${gridCols}</w:tblGrid>${trs.join("")}</w:tbl>`;

  const xml = doc.documentXml();
  const insertAt = entry.block.end;
  doc.pkg.setPart(doc.documentPartName, splice(xml, insertAt, insertAt, tbl));
  doc.invalidate();

  // The new table is the body table whose slice begins exactly at the insert
  // point; its ordinal gives the T-anchor (§13: independent T sequence).
  const created = bodyEntries(doc).find((e) => e.kind === "tbl" && e.start === insertAt);
  const tAnchor = created ? created.anchor : `T${countTablesUpTo(doc, insertAt) + 1}`;
  const newAnchor = `${tAnchor}${afterToken(args.after)}`;
  return {
    new_anchor: newAnchor,
    note: `${rows}×${cols} table inserted${styled ? "; header row styled with 'Table Grid'." : "."}`,
  };
}

function countTablesUpTo(doc: DocHandle, offset: number): number {
  return bodyEntries(doc).filter((e) => e.kind === "tbl" && e.start < offset).length;
}

// --- set_cells --------------------------------------------------------------

interface CellAddr {
  r: number;
  c: number;
  text: string;
}

function resolveCellAddr(cell: DocxTableCell): CellAddr {
  if (cell.r !== undefined && cell.c !== undefined) {
    return { r: cell.r, c: cell.c, text: cell.text };
  }
  if (cell.ref !== undefined) {
    const { r, c } = parseA1(cell.ref);
    return { r, c, text: cell.text };
  }
  throw anchorInvalid("Cell needs {r,c} or ref.");
}

function tableSetCells(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  const entry = requireTableForEdit(doc, args);
  const cells = (args.cells ?? []).map(resolveCellAddr);
  const xml = doc.documentXml();
  const model = parseTable(xml, entry.block);
  const edits: SpliceEdit[] = [];
  for (const addr of cells) {
    const tr = model.rows[addr.r];
    if (tr === undefined) throw anchorInvalid(`Cell row ${addr.r} out of range.`);
    const tcs = rowCells(xml, tr);
    const tc = tcs[addr.c];
    if (tc === undefined) throw anchorInvalid(`Cell column ${addr.c} out of range.`);
    if (isCovered(xml, tc)) throw anchorInvalid(`Cell is part of a merge and cannot be set.`);
    edits.push(setCellTextEdit(xml, tc, addr.text));
  }
  if (edits.length > 0) {
    doc.pkg.setPart(doc.documentPartName, spliceAll(xml, edits));
    doc.invalidate();
  }
  return { note: `Set ${cells.length} cell${cells.length === 1 ? "" : "s"}.` };
}

/** True iff a cell is covered (a gridSpan-removed cell or a vMerge continuation). */
function isCovered(xml: string, tc: ElementSlice): boolean {
  if (tc.selfClosed) return false;
  const tcPr = childElements(xml, tc.contentStart, tc.contentEnd).find((k) => k.name === "w:tcPr");
  if (!tcPr || tcPr.selfClosed) return false;
  const vMerge = findElement(xml, "w:vMerge", tcPr.contentStart, tcPr.contentEnd);
  if (vMerge) {
    const val = getAttr(xml, tagOf(vMerge), "w:val");
    // A vMerge with no val (or val=continue) is a continuation cell.
    if (val === undefined || val === "continue") return true;
  }
  return false;
}

/**
 * Replace a cell's paragraphs with one `<w:p>` carrying the cell's first
 * paragraph's `w:pPr` verbatim (when present) + one run with the new text,
 * preserving `w:tcPr` (§14).
 */
function setCellTextEdit(xml: string, tc: ElementSlice, text: string): SpliceEdit {
  const kids = childElements(xml, tc.contentStart, tc.contentEnd);
  const tcPr = kids.find((k) => k.name === "w:tcPr");
  const firstP = kids.find((k) => k.name === "w:p");
  let pPr = "";
  if (firstP && !firstP.selfClosed) {
    const inner = childElements(xml, firstP.contentStart, firstP.contentEnd).find(
      (k) => k.name === "w:pPr",
    );
    if (inner) pPr = xml.slice(inner.start, inner.end);
  }
  const run = text !== "" ? emitTextRuns(text) : "";
  const para = `<w:p>${pPr}${run}</w:p>`;
  // Replace from after tcPr (or content start) to content end.
  const contentStart = tcPr ? tcPr.end : tc.contentStart;
  return { start: contentStart, end: tc.contentEnd, text: para };
}

// --- insert_row -------------------------------------------------------------

function tableInsertRow(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  const entry = requireTableForEdit(doc, args);
  const xml = doc.documentXml();
  const model = parseTable(xml, entry.block);
  const last = model.rows.length - 1;
  const at = clampInsert(args.at, model.rows.length);
  // Clone the structure (cell props, not text) of row min(at, last).
  const templateIdx = Math.min(at, last);
  const template = model.rows[templateIdx];
  if (template === undefined) throw anchorInvalid("Table has no rows to clone.");
  const newRow = cloneRowStructure(xml, template);
  // Insert before index `at` (at == rows appends).
  const insertAt =
    at >= model.rows.length
      ? (model.rows[model.rows.length - 1] as ElementSlice).end
      : (model.rows[at] as ElementSlice).start;
  doc.pkg.setPart(doc.documentPartName, splice(xml, insertAt, insertAt, newRow));
  doc.invalidate();
  return { note: `Inserted a row at index ${at}.` };
}

/** Clone a row's cell structure (tcPr) with blank text. */
function cloneRowStructure(xml: string, tr: ElementSlice): string {
  const cells = rowCells(xml, tr);
  const newCells = cells.map((tc) => {
    const tcPr = tc.selfClosed
      ? null
      : childElements(xml, tc.contentStart, tc.contentEnd).find((k) => k.name === "w:tcPr");
    const tcPrXml = tcPr ? xml.slice(tcPr.start, tcPr.end) : "";
    return `<w:tc>${tcPrXml}<w:p/></w:tc>`;
  });
  return `<w:tr>${newCells.join("")}</w:tr>`;
}

// --- insert_col -------------------------------------------------------------

function tableInsertCol(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  const entry = requireTableForEdit(doc, args);
  const xml = doc.documentXml();
  const model = parseTable(xml, entry.block);
  const widths = gridWidths(xml, model.grid);
  const nCols = widths.length || maxRowCells(xml, model.rows);
  const at = clampInsert(args.at, nCols);

  const edits: SpliceEdit[] = [];
  // New gridCol: split the neighbor's width.
  if (model.grid && !model.grid.selfClosed) {
    const gridCols = childElements(xml, model.grid.contentStart, model.grid.contentEnd).filter(
      (k) => k.name === "w:gridCol",
    );
    const neighborIdx = Math.min(at, gridCols.length - 1);
    const neighbor = gridCols[neighborIdx];
    const neighborW = neighbor ? Number(getAttr(xml, tagOf(neighbor), "w:w") ?? "0") : 0;
    const half = Math.floor(neighborW / 2);
    const newColW = neighborW - half;
    // Insert the new gridCol at position `at`.
    const insertAt =
      at >= gridCols.length
        ? (gridCols[gridCols.length - 1] as ElementSlice).end
        : (gridCols[at] as ElementSlice).start;
    edits.push({ start: insertAt, end: insertAt, text: `<w:gridCol w:w="${newColW}"/>` });
    // Shrink the neighbor.
    if (neighbor && half > 0) {
      edits.push({
        start: neighbor.start,
        end: neighbor.end,
        text: `<w:gridCol w:w="${half}"/>`,
      });
    }
  }
  // One blank w:tc per row at column `at`.
  for (const tr of model.rows) {
    const tcs = rowCells(xml, tr);
    const insertAt =
      at >= tcs.length
        ? tcs.length > 0
          ? (tcs[tcs.length - 1] as ElementSlice).end
          : tr.contentStart
        : (tcs[at] as ElementSlice).start;
    edits.push({ start: insertAt, end: insertAt, text: "<w:tc><w:tcPr/><w:p/></w:tc>" });
  }
  doc.pkg.setPart(doc.documentPartName, spliceAll(xml, edits));
  doc.invalidate();
  return { note: `Inserted a column at index ${at}.` };
}

function maxRowCells(xml: string, rows: ElementSlice[]): number {
  let max = 0;
  for (const tr of rows) max = Math.max(max, rowCells(xml, tr).length);
  return max;
}

// --- delete_row / delete_col ------------------------------------------------

function tableDeleteRow(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  const entry = requireTableForEdit(doc, args);
  const xml = doc.documentXml();
  const model = parseTable(xml, entry.block);
  const at = args.at ?? 0;
  const tr = model.rows[at];
  if (tr === undefined) throw anchorInvalid(`Row ${at} out of range.`);
  // A vMerge-origin row deletion promotes the next continuation to origin.
  const edits: SpliceEdit[] = [{ start: tr.start, end: tr.end, text: "" }];
  promoteVMergeAfterRow(xml, model, at, edits);
  doc.pkg.setPart(doc.documentPartName, spliceAll(xml, edits));
  doc.invalidate();
  return { note: `Deleted row ${at}.` };
}

/** If a deleted origin row had vMerge restart cells, promote the next row's continuation. */
function promoteVMergeAfterRow(
  xml: string,
  model: TableModel,
  at: number,
  edits: SpliceEdit[],
): void {
  const originRow = model.rows[at];
  const nextRow = model.rows[at + 1];
  if (!originRow || !nextRow) return;
  const originCells = rowCells(xml, originRow);
  const nextCells = rowCells(xml, nextRow);
  originCells.forEach((tc, idx) => {
    if (vMergeKind(xml, tc) !== "restart") return;
    const below = nextCells[idx];
    if (below && vMergeKind(xml, below) === "continue") {
      // Promote the continuation cell to restart.
      const tcPr = childElements(xml, below.contentStart, below.contentEnd).find(
        (k) => k.name === "w:tcPr",
      );
      if (!tcPr || tcPr.selfClosed) return;
      const vMerge = findElement(xml, "w:vMerge", tcPr.contentStart, tcPr.contentEnd);
      if (vMerge) {
        edits.push({ start: vMerge.start, end: vMerge.end, text: '<w:vMerge w:val="restart"/>' });
      }
    }
  });
}

function vMergeKind(xml: string, tc: ElementSlice): "restart" | "continue" | null {
  if (tc.selfClosed) return null;
  const tcPr = childElements(xml, tc.contentStart, tc.contentEnd).find((k) => k.name === "w:tcPr");
  if (!tcPr || tcPr.selfClosed) return null;
  const vMerge = findElement(xml, "w:vMerge", tcPr.contentStart, tcPr.contentEnd);
  if (!vMerge) return null;
  const val = getAttr(xml, tagOf(vMerge), "w:val");
  return val === "restart" ? "restart" : "continue";
}

function tableDeleteCol(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  const entry = requireTableForEdit(doc, args);
  const xml = doc.documentXml();
  const model = parseTable(xml, entry.block);
  const at = args.at ?? 0;
  const edits: SpliceEdit[] = [];
  // Remove the w:gridCol at `at`.
  if (model.grid && !model.grid.selfClosed) {
    const gridCols = childElements(xml, model.grid.contentStart, model.grid.contentEnd).filter(
      (k) => k.name === "w:gridCol",
    );
    const gc = gridCols[at];
    if (gc) edits.push({ start: gc.start, end: gc.end, text: "" });
  }
  // Remove that cell index in each row.
  for (const tr of model.rows) {
    const tcs = rowCells(xml, tr);
    const tc = tcs[at];
    if (tc) edits.push({ start: tc.start, end: tc.end, text: "" });
  }
  if (edits.length === 0) throw anchorInvalid(`Column ${at} out of range.`);
  doc.pkg.setPart(doc.documentPartName, spliceAll(xml, edits));
  doc.invalidate();
  return { note: `Deleted column ${at}.` };
}

// --- merge ------------------------------------------------------------------

function tableMerge(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  const entry = requireTableForEdit(doc, args);
  if (args.range === undefined) throw anchorInvalid("op 'merge' requires range.");
  const m = RANGE_RE.exec(args.range);
  if (!m) throw anchorInvalid(`Malformed range: ${args.range}.`);
  const a = parseA1(m[1] as string);
  const b = parseA1(m[2] as string);
  const r0 = Math.min(a.r, b.r);
  const r1 = Math.max(a.r, b.r);
  const c0 = Math.min(a.c, b.c);
  const c1 = Math.max(a.c, b.c);
  const span = c1 - c0 + 1;

  const xml = doc.documentXml();
  const model = parseTable(xml, entry.block);
  const edits: SpliceEdit[] = [];
  const vertical = r1 > r0;

  // One coherent edit per spanned row: rewrite the left cell (gridSpan when the
  // span is horizontal, vMerge restart/continue when vertical), then remove the
  // covered cells in that row (§14).
  for (let r = r0; r <= r1; r++) {
    const tr = model.rows[r];
    if (tr === undefined) throw anchorInvalid(`Merge row ${r} out of range.`);
    const tcs = rowCells(xml, tr);
    const left = tcs[c0];
    if (left === undefined) throw anchorInvalid(`Merge column ${c0} out of range.`);

    // Each mark is spliced as the FIRST child of the cell's tcPr (§14); gridSpan
    // is written first, then vMerge prepended — so the order is vMerge, gridSpan,
    // then the base props (tcW, …).
    let leadProps = "";
    if (span > 1) leadProps = `<w:gridSpan w:val="${span}"/>${leadProps}`;
    if (vertical) {
      leadProps = `${r === r0 ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>"}${leadProps}`;
    }

    if (vertical && r > r0) {
      // A continuation cell keeps only an empty <w:p/> (§14).
      edits.push(rebuildCell(xml, left, leadProps, true));
    } else if (leadProps !== "") {
      edits.push(rebuildCell(xml, left, leadProps, false));
    }
    // Remove the covered cells in this row.
    for (let c = c0 + 1; c <= c1; c++) {
      const covered = tcs[c];
      if (covered) edits.push({ start: covered.start, end: covered.end, text: "" });
    }
  }
  doc.pkg.setPart(doc.documentPartName, spliceAll(xml, edits));
  doc.invalidate();
  return { note: `Merged ${args.range}.` };
}

/**
 * Rebuild a cell, prepending the merge marks (`leadProps`: gridSpan/vMerge) as
 * the FIRST children of the cell's tcPr (§14), dropping any prior gridSpan/vMerge
 * and keeping the remaining base props (tcW, shd, …) and content. `blank`
 * replaces the content with a single empty `<w:p/>` (continuation cell).
 */
function rebuildCell(xml: string, tc: ElementSlice, leadProps: string, blank: boolean): SpliceEdit {
  const kids = tc.selfClosed ? [] : childElements(xml, tc.contentStart, tc.contentEnd);
  const tcPr = kids.find((k) => k.name === "w:tcPr");
  let baseProps = "";
  if (tcPr && !tcPr.selfClosed) {
    baseProps = childElements(xml, tcPr.contentStart, tcPr.contentEnd)
      .filter((k) => k.name !== "w:gridSpan" && k.name !== "w:vMerge")
      .map((k) => xml.slice(k.start, k.end))
      .join("");
  }
  const tcPrXml = `<w:tcPr>${leadProps}${baseProps}</w:tcPr>`;
  if (blank) {
    return { start: tc.start, end: tc.end, text: `<w:tc>${tcPrXml}<w:p/></w:tc>` };
  }
  // Keep the existing content (everything after tcPr, or all content if none).
  const content = tc.selfClosed
    ? "<w:p/>"
    : xml.slice(tcPr ? tcPr.end : tc.contentStart, tc.contentEnd);
  return { start: tc.start, end: tc.end, text: `<w:tc>${tcPrXml}${content}</w:tc>` };
}

// --- style ------------------------------------------------------------------

/** Legacy `op:"style"` (no props): stamp the Table Grid named style. */
function applyNamedStyle(doc: DocHandle, entry: AnchorEntry): void {
  ensureStyle(doc.pkg, "TableGrid", TABLE_GRID_STYLE);
  const xml = doc.documentXml();
  const model = parseTable(xml, entry.block);
  const tblStyle = '<w:tblStyle w:val="TableGrid"/>';
  if (model.tblPr && !model.tblPr.selfClosed) {
    const existing = findElement(
      xml,
      "w:tblStyle",
      model.tblPr.contentStart,
      model.tblPr.contentEnd,
    );
    if (existing) {
      doc.pkg.setPart(doc.documentPartName, splice(xml, existing.start, existing.end, tblStyle));
    } else {
      doc.pkg.setPart(
        doc.documentPartName,
        splice(xml, model.tblPr.contentStart, model.tblPr.contentStart, tblStyle),
      );
    }
  } else if (model.tblPr && model.tblPr.selfClosed) {
    doc.pkg.setPart(
      doc.documentPartName,
      splice(xml, model.tblPr.start, model.tblPr.end, `<w:tblPr>${tblStyle}</w:tblPr>`),
    );
  } else {
    doc.pkg.setPart(
      doc.documentPartName,
      splice(
        xml,
        entry.block.contentStart,
        entry.block.contentStart,
        `<w:tblPr>${tblStyle}</w:tblPr>`,
      ),
    );
  }
  doc.invalidate();
}

// -- comprehensive table formatting (op:"style" with props, §14) --------------

const BORDER_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV"] as const;

function hexColor(c: unknown): string {
  return typeof c === "string" && c.length > 0 ? (c.startsWith("#") ? c.slice(1) : c) : "auto";
}

function tblBordersXml(on: unknown, weight: unknown, color: unknown): string {
  const attrs = on
    ? `w:val="single" w:sz="${typeof weight === "number" ? Math.trunc(weight) : 4}" w:space="0" w:color="${hexColor(color)}"`
    : 'w:val="none" w:sz="0" w:space="0" w:color="auto"';
  return `<w:tblBorders>${BORDER_SIDES.map((s) => `<w:${s} ${attrs}/>`).join("")}</w:tblBorders>`;
}

function shdXml(fill: unknown): string {
  return `<w:shd w:val="clear" w:color="auto" w:fill="${hexColor(fill)}"/>`;
}

function childXml(xml: string, parent: ElementSlice | undefined, name: string): string | null {
  if (!parent || parent.selfClosed) return null;
  const el = findElement(xml, name, parent.contentStart, parent.contentEnd);
  return el ? xml.slice(el.start, el.end) : null;
}

function buildTblPr(
  xml: string,
  tblPr: ElementSlice | undefined,
  style: string | undefined,
  props: DocxTableProps,
): string {
  const parts: string[] = [];
  const styleXml =
    style !== undefined ? '<w:tblStyle w:val="TableGrid"/>' : childXml(xml, tblPr, "w:tblStyle");
  if (styleXml) parts.push(styleXml);
  if (typeof props.width_pct === "number") {
    parts.push(`<w:tblW w:w="${Math.round(props.width_pct * 50)}" w:type="pct"/>`);
  } else {
    parts.push('<w:tblW w:w="0" w:type="auto"/>');
  }
  if (typeof props.align === "string" && props.align) parts.push(`<w:jc w:val="${props.align}"/>`);
  if ("borders" in props) {
    parts.push(tblBordersXml(props.borders, props.border_weight, props.border_color));
  }
  if (props.col_widths && props.col_widths.length > 0) parts.push('<w:tblLayout w:type="fixed"/>');
  const lookXml = childXml(xml, tblPr, "w:tblLook");
  if (lookXml) parts.push(lookXml);
  return `<w:tblPr>${parts.join("")}</w:tblPr>`;
}

function cellShdEdit(xml: string, tc: ElementSlice, fill: unknown): SpliceEdit | null {
  const tcPr = tc.selfClosed
    ? undefined
    : childElements(xml, tc.contentStart, tc.contentEnd).find((k) => k.name === "w:tcPr");
  const shd =
    tcPr && !tcPr.selfClosed
      ? childElements(xml, tcPr.contentStart, tcPr.contentEnd).find((k) => k.name === "w:shd")
      : undefined;
  if (fill === null) {
    return shd ? { start: shd.start, end: shd.end, text: "" } : null;
  }
  const text = shdXml(fill);
  if (shd) return { start: shd.start, end: shd.end, text };
  if (tcPr && !tcPr.selfClosed) {
    const tcW = childElements(xml, tcPr.contentStart, tcPr.contentEnd).find(
      (k) => k.name === "w:tcW",
    );
    const pos = tcW ? tcW.end : tcPr.contentStart;
    return { start: pos, end: pos, text };
  }
  return { start: tc.contentStart, end: tc.contentStart, text: `<w:tcPr>${shdXml(fill)}</w:tcPr>` };
}

function shadingEdits(
  xml: string,
  block: ElementSlice,
  shading: { header?: string | null; all?: string | null },
): SpliceEdit[] {
  const rows = childElements(xml, block.contentStart, block.contentEnd).filter(
    (k) => k.name === "w:tr",
  );
  const cellsOf = (tr: ElementSlice): ElementSlice[] =>
    tr.selfClosed
      ? []
      : childElements(xml, tr.contentStart, tr.contentEnd).filter((k) => k.name === "w:tc");
  const byCell = new Map<number, { tc: ElementSlice; fill: unknown }>();
  if ("all" in shading) {
    for (const tr of rows)
      for (const tc of cellsOf(tr)) byCell.set(tc.start, { tc, fill: shading.all });
  }
  if ("header" in shading && rows[0]) {
    for (const tc of cellsOf(rows[0])) byCell.set(tc.start, { tc, fill: shading.header });
  }
  const edits: SpliceEdit[] = [];
  for (const { tc, fill } of byCell.values()) {
    const e = cellShdEdit(xml, tc, fill);
    if (e) edits.push(e);
  }
  return edits;
}

function tableStyle(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  const entry = requireTableForEdit(doc, args);
  if (args.props === undefined) {
    applyNamedStyle(doc, entry);
    return { note: "Styled table." };
  }
  const xml = doc.documentXml();
  const block = entry.block;
  const kids = childElements(xml, block.contentStart, block.contentEnd);
  const tblPr = kids.find((k) => k.name === "w:tblPr");
  const edits: SpliceEdit[] = [];
  const newTblPr = buildTblPr(xml, tblPr, args.style, args.props);
  edits.push(
    tblPr
      ? { start: tblPr.start, end: tblPr.end, text: newTblPr }
      : { start: block.contentStart, end: block.contentStart, text: newTblPr },
  );
  const colWidths = args.props.col_widths;
  if (Array.isArray(colWidths) && colWidths.length > 0) {
    const grid = kids.find((k) => k.name === "w:tblGrid");
    const cols = colWidths.map((w) => `<w:gridCol w:w="${Math.trunc(w)}"/>`).join("");
    if (grid)
      edits.push({ start: grid.start, end: grid.end, text: `<w:tblGrid>${cols}</w:tblGrid>` });
  }
  if (args.props.shading && typeof args.props.shading === "object") {
    edits.push(...shadingEdits(xml, block, args.props.shading));
  }
  doc.pkg.setPart(doc.documentPartName, spliceAll(xml, edits));
  doc.invalidate();
  return { note: "Styled table." };
}

function tableDelete(doc: DocHandle, args: DocxTableArgs): DocxTableResult {
  const entry = requireTableForEdit(doc, args);
  const xml = doc.documentXml();
  doc.pkg.setPart(doc.documentPartName, splice(xml, entry.block.start, entry.block.end, ""));
  doc.invalidate();
  return { note: `Deleted table ${args.anchor}.` };
}

// ---------------------------------------------------------------------------
// Anchor requirement for editing ops
// ---------------------------------------------------------------------------

function requireTableForEdit(doc: DocHandle, args: DocxTableArgs): AnchorEntry {
  if (args.anchor === undefined) throw anchorInvalid(`op '${args.op}' requires a table anchor.`);
  return requireTable(doc, args.anchor);
}

/** Clamp an insert index into `[0, count]`. */
function clampInsert(at: number | undefined, count: number): number {
  const v = at === undefined ? count : Math.trunc(at);
  if (v < 0) return 0;
  if (v > count) return count;
  return v;
}
