/**
 * Create a document from Markdown (`docx_create content_md`) per §22.
 *
 * Deterministic skeleton parts are emitted in the §22 creation order:
 * word/document.xml, word/styles.xml, [Content_Types].xml, _rels/.rels,
 * word/_rels/document.xml.rels, docProps/core.xml (created/modified =
 * DOCXENGINE_FIXED_DATE or its default). Block mapping: ATX headings, quotes,
 * `---`/`***` rules, `-`/`*`/`1.` list items (via §17 numbering.xml), GitHub
 * pipe tables (§14), else plain paragraphs. Inline `**bold**`/`*italic*`/
 * `` `code` `` split the text into runs at marker boundaries (§3 escaping).
 *
 * The Python twin (`_create.py`) is the byte-parity reference.
 */
import { strToU8, zipSync, type Zippable } from "fflate";

import { ToolError } from "./errors.js";
import type { Session } from "./session.js";
import { isValid, validateDoc } from "./validate.js";
import { emitTextElement, emitTextRuns, escapeAttr, escapeText } from "./xmlscan.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

const DEFAULT_DATE = "2026-01-01T00:00:00Z";

const DEFAULT_CONTENT_WIDTH = 9026; // §15 A4 default content width (twips)

/** The created/modified timestamp: DOCXENGINE_FIXED_DATE or the §22 default. */
function coreDate(): string {
  const fixed = process.env["DOCXENGINE_FIXED_DATE"];
  return fixed !== undefined && fixed !== "" ? fixed : DEFAULT_DATE;
}

// ---------------------------------------------------------------------------
// Inline markdown → runs (§22)
// ---------------------------------------------------------------------------

interface InlineRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  code?: boolean;
}

/**
 * Split inline markdown into runs at marker boundaries. Supports `**x**`/
 * `__x__` (bold), `*x*`/`_x_` (italic), `` `x` `` (code). Markers do not nest
 * in MVP; an unmatched marker is literal text.
 */
export function parseInline(text: string): InlineRun[] {
  const runs: InlineRun[] = [];
  let buf = "";
  let i = 0;
  const flush = (extra?: Partial<InlineRun>): void => {
    if (buf !== "") {
      runs.push({ text: buf });
      buf = "";
    }
    if (extra) runs.push({ text: extra.text ?? "", ...extra });
  };
  while (i < text.length) {
    const two = text.slice(i, i + 2);
    if (two === "**" || two === "__") {
      const close = text.indexOf(two, i + 2);
      if (close >= 0) {
        flush();
        runs.push({ text: text.slice(i + 2, close), bold: true });
        i = close + 2;
        continue;
      }
    }
    const ch = text[i] as string;
    if (ch === "*" || ch === "_") {
      const close = text.indexOf(ch, i + 1);
      if (close >= 0 && close > i + 1) {
        flush();
        runs.push({ text: text.slice(i + 1, close), italic: true });
        i = close + 1;
        continue;
      }
    }
    if (ch === "`") {
      const close = text.indexOf("`", i + 1);
      if (close >= 0) {
        flush();
        runs.push({ text: text.slice(i + 1, close), code: true });
        i = close + 1;
        continue;
      }
    }
    buf += ch;
    i += 1;
  }
  flush();
  return runs.filter((r) => r.text !== "" || runs.length === 1);
}

/** Emit one `w:r` for an inline run with §22 run-property markup. */
function emitInlineRun(run: InlineRun): string {
  const rprParts: string[] = [];
  if (run.bold) rprParts.push("<w:b/>");
  if (run.italic) rprParts.push("<w:i/>");
  if (run.code) {
    rprParts.push('<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>');
  }
  const rPr = rprParts.length > 0 ? `<w:rPr>${rprParts.join("")}</w:rPr>` : "";
  return `<w:r>${rPr}${emitTextElement("w:t", run.text)}</w:r>`;
}

/** HTML line breaks accepted in markdown (`<br>`/`<br/>`/`<br />`) → `<w:br/>`. */
const BR_RE = /<br\s*\/?>/gi;
/** A break-only run — valid between styled runs (a bare `<w:br/>` may not sit in `w:p`). */
const LINE_BREAK_RUN = "<w:r><w:br/></w:r>";

/** Emit the run sequence for an inline string; `<br>`/`\n` become `<w:br/>` soft breaks. */
export function emitInline(text: string): string {
  return text
    .replace(BR_RE, "\n")
    .split("\n")
    .map((line) => parseInline(line).map(emitInlineRun).join(""))
    .join(LINE_BREAK_RUN);
}

// ---------------------------------------------------------------------------
// Block parsing (§22)
// ---------------------------------------------------------------------------

const HEADING_RE = /^(#{1,6})\s+(.*)$/;
const QUOTE_RE = /^>\s?(.*)$/;
const HR_RE = /^(?:-{3,}|\*{3,})$/;
const UL_RE = /^[-*]\s+(.*)$/;
const OL_RE = /^\d+\.\s+(.*)$/;
const TABLE_ROW_RE = /^\|(.+)\|\s*$/;
const TABLE_SEP_RE = /^\|?[\s:|-]+\|?$/;
/** GitHub task-list item (`- [ ] x` / `* [x] x`): the checkbox glyph replaces the bullet. */
const TASK_RE = /^[-*]\s+\[([ xX])\]\s?(.*)$/;
const BALLOT_EMPTY = "☐"; // ☐
const BALLOT_CHECKED = "☒"; // ☒

/** A single markdown line classified into a block kind (tables are handled separately). */
export type LineBlock =
  | { kind: "heading"; level: number; text: string }
  | { kind: "quote"; text: string }
  | { kind: "rule" }
  | { kind: "task"; text: string }
  | { kind: "ul"; text: string }
  | { kind: "ol"; text: string }
  | { kind: "plain"; text: string };

/**
 * Classify one markdown line into a §22/§6a block. Task items carry their
 * checkbox glyph folded into `text` (`☐ `/`☒ `) so the bullet is dropped.
 * Shared by `docx_create` (§22) and `docx_insert` (§6a).
 */
export function classifyLine(line: string): LineBlock {
  const heading = HEADING_RE.exec(line);
  if (heading)
    return { kind: "heading", level: (heading[1] as string).length, text: heading[2] as string };
  const quote = QUOTE_RE.exec(line);
  if (quote) return { kind: "quote", text: quote[1] as string };
  if (HR_RE.test(line.trim())) return { kind: "rule" };
  const task = TASK_RE.exec(line);
  if (task) {
    const glyph = (task[1] as string).toLowerCase() === "x" ? BALLOT_CHECKED : BALLOT_EMPTY;
    return { kind: "task", text: `${glyph} ${task[2] as string}` };
  }
  const ul = UL_RE.exec(line);
  if (ul) return { kind: "ul", text: ul[1] as string };
  const ol = OL_RE.exec(line);
  if (ol) return { kind: "ol", text: ol[1] as string };
  return { kind: "plain", text: line };
}

interface NumberingPlan {
  /** abstractNumId/numId for ol (allocated lazily). */
  ol?: number;
  ul?: number;
}

/** Split a GitHub table row into trimmed cell strings. */
function splitRow(line: string): string[] {
  const m = TABLE_ROW_RE.exec(line);
  const inner = m ? (m[1] as string) : line.replace(/^\|/, "").replace(/\|$/, "");
  return inner.split("|").map((c) => c.trim());
}

function isSeparatorRow(line: string): boolean {
  if (!line.includes("-")) return false;
  return TABLE_SEP_RE.test(line.trim()) && line.includes("|");
}

interface BuildState {
  paragraphs: string[];
  /** numbering.xml abstractNum + num entries, in creation order. */
  numbering: string[];
  nextNumId: number;
  plan: NumberingPlan;
  /** Whether a TableGrid style was used (always shipped by §22, but tracked). */
  usedTableGrid: boolean;
  bodyParagraphCount: number;
}

/** Build the §17 abstractNum + num markup for ol/ul; returns the numId. */
function ensureNumbering(state: BuildState, kind: "ol" | "ul"): number {
  if (kind === "ol" && state.plan.ol !== undefined) return state.plan.ol;
  if (kind === "ul" && state.plan.ul !== undefined) return state.plan.ul;
  const id = state.nextNumId++;
  const levels: string[] = [];
  const OL_FORMATS = ["decimal", "lowerLetter", "lowerRoman"];
  const UL_GLYPHS = ["•", "◦", "▪"];
  for (let ilvl = 0; ilvl < 9; ilvl++) {
    const left = 720 * (ilvl + 1);
    const ind = `<w:ind w:left="${left}" w:hanging="360"/>`;
    if (kind === "ol") {
      const fmt = OL_FORMATS[ilvl % OL_FORMATS.length] as string;
      levels.push(
        `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>` +
          `<w:lvlText w:val="${escapeAttr(`%${ilvl + 1}.`)}"/><w:pPr>${ind}</w:pPr></w:lvl>`,
      );
    } else {
      const glyph = UL_GLYPHS[ilvl % UL_GLYPHS.length] as string;
      levels.push(
        `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
          `<w:lvlText w:val="${escapeAttr(glyph)}"/><w:pPr>${ind}</w:pPr></w:lvl>`,
      );
    }
  }
  state.numbering.push(
    `<w:abstractNum w:abstractNumId="${id}">${levels.join("")}</w:abstractNum>`,
    `<w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num>`,
  );
  if (kind === "ol") state.plan.ol = id;
  else state.plan.ul = id;
  return id;
}

/** Emit a list-item paragraph (numPr + ListParagraph style) per §17. */
function emitListItem(state: BuildState, text: string, kind: "ol" | "ul"): void {
  const numId = ensureNumbering(state, kind);
  const numPr = `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`;
  const pPr = `<w:pPr><w:pStyle w:val="ListParagraph"/>${numPr}</w:pPr>`;
  state.paragraphs.push(`<w:p>${pPr}${emitInline(text)}</w:p>`);
  state.bodyParagraphCount += 1;
}

/** Emit a §14 table from parsed rows (header iff a separator row follows row 1). */
function emitTable(state: BuildState, rows: string[][], header: boolean): void {
  const cols = Math.max(...rows.map((r) => r.length), 1);
  state.usedTableGrid = header;
  const widths = distributeWidths(DEFAULT_CONTENT_WIDTH, cols);
  const gridCols = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const tblPr =
    `<w:tblPr>${header ? '<w:tblStyle w:val="TableGrid"/>' : ""}` +
    `<w:tblW w:w="0" w:type="auto"/></w:tblPr>`;
  const trs: string[] = [];
  rows.forEach((row, r) => {
    const isHeader = header && r === 0;
    const cells: string[] = [];
    for (let c = 0; c < cols; c++) {
      const text = row[c] ?? "";
      const shd = isHeader ? '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' : "";
      const tcPr = `<w:tcPr><w:tcW w:w="${widths[c]}" w:type="dxa"/>${shd}</w:tcPr>`;
      const para = cellParagraph(text, isHeader);
      cells.push(`<w:tc>${tcPr}${para}</w:tc>`);
    }
    trs.push(`<w:tr>${cells.join("")}</w:tr>`);
  });
  state.paragraphs.push(`<w:tbl>${tblPr}<w:tblGrid>${gridCols}</w:tblGrid>${trs.join("")}</w:tbl>`);
}

function cellParagraph(text: string, header: boolean): string {
  if (text === "") return "<w:p/>";
  if (header) {
    return `<w:p>${emitTextRuns(text.replace(BR_RE, "\n"), "<w:rPr><w:b/></w:rPr>")}</w:p>`;
  }
  // Cell text supports inline markdown (consistent with §22 inline parsing).
  return `<w:p>${emitInline(text)}</w:p>`;
}

function distributeWidths(total: number, cols: number): number[] {
  if (cols <= 0) return [];
  const base = Math.floor(total / cols);
  const widths = new Array<number>(cols).fill(base);
  widths[cols - 1] = total - base * (cols - 1);
  return widths;
}

/** Emit a horizontal-rule paragraph (§22). */
function emitRule(state: BuildState): void {
  state.paragraphs.push(
    "<w:p><w:pPr><w:pBdr>" +
      '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>' +
      "</w:pBdr></w:pPr></w:p>",
  );
  state.bodyParagraphCount += 1;
}

/** Emit a heading/quote/plain paragraph with an optional pStyle. */
function emitParagraph(state: BuildState, text: string, style: string | null): void {
  const pPr = style !== null ? `<w:pPr><w:pStyle w:val="${escapeAttr(style)}"/></w:pPr>` : "";
  const inner = emitInline(text);
  state.paragraphs.push(`<w:p>${pPr}${inner}</w:p>`);
  state.bodyParagraphCount += 1;
}

/** Parse the markdown body into the §22 block sequence. */
function buildBody(md: string): BuildState {
  const state: BuildState = {
    paragraphs: [],
    numbering: [],
    nextNumId: 1,
    plan: {},
    usedTableGrid: false,
    bodyParagraphCount: 0,
  };
  // One trailing \r stripped per line (§6a `docx_insert` shares this).
  const lines = md.split("\n").map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] as string;
    if (line.trim() === "") continue;

    // GitHub table: a pipe row optionally followed by a separator row.
    if (TABLE_ROW_RE.test(line)) {
      const rows: string[][] = [];
      let j = i;
      let header = false;
      // Collect the first row.
      rows.push(splitRow(lines[j] as string));
      j += 1;
      // A separator row right after row 1 → header table.
      if (j < lines.length && isSeparatorRow(lines[j] as string)) {
        header = true;
        j += 1;
      }
      // Collect subsequent body rows.
      while (
        j < lines.length &&
        TABLE_ROW_RE.test(lines[j] as string) &&
        !isSeparatorRow(lines[j] as string)
      ) {
        rows.push(splitRow(lines[j] as string));
        j += 1;
      }
      emitTable(state, rows, header);
      i = j - 1;
      continue;
    }

    const block = classifyLine(line);
    switch (block.kind) {
      case "heading":
        emitParagraph(state, block.text, `Heading${block.level}`);
        break;
      case "quote":
        emitParagraph(state, block.text, "Quote");
        break;
      case "rule":
        emitRule(state);
        break;
      case "task":
        emitParagraph(state, block.text, "ListParagraph");
        break;
      case "ul":
        emitListItem(state, block.text, "ul");
        break;
      case "ol":
        emitListItem(state, block.text, "ol");
        break;
      case "plain":
        emitParagraph(state, block.text, null);
        break;
    }
  }
  return state;
}

// ---------------------------------------------------------------------------
// Skeleton parts (§22)
// ---------------------------------------------------------------------------

/** The base style set §22 ships (Normal, Heading1-6, ListParagraph, TableGrid, Quote). */
function stylesXml(): string {
  const headings: string[] = [];
  for (let n = 1; n <= 6; n++) {
    headings.push(
      `<w:style w:type="paragraph" w:styleId="Heading${n}">` +
        `<w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/>` +
        `<w:pPr><w:outlineLvl w:val="${n - 1}"/></w:pPr>` +
        `<w:rPr><w:b/></w:rPr></w:style>`,
    );
  }
  const normal =
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
  const listParagraph =
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>' +
    '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>';
  const tableGrid =
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>' +
    "<w:tblPr><w:tblBorders>" +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    "</w:tblBorders></w:tblPr></w:style>";
  const quote =
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>' +
    '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr>' +
    "<w:rPr><w:i/></w:rPr></w:style>";
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:styles xmlns:w="${W_NS}">` +
    normal +
    headings.join("") +
    listParagraph +
    tableGrid +
    quote +
    "</w:styles>"
  );
}

function documentXml(bodyParagraphs: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:document xmlns:w="${W_NS}">` +
    `<w:body>${bodyParagraphs}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"' +
    ' w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
  );
}

function numberingXml(entries: string[]): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:numbering xmlns:w="${W_NS}">${entries.join("")}</w:numbering>`
  );
}

function contentTypesXml(hasNumbering: boolean): string {
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
  ];
  if (hasNumbering) {
    overrides.push(
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    );
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<Types xmlns="${CT_NS}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    overrides.join("") +
    "</Types>"
  );
}

function rootRelsXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<Relationships xmlns="${REL_NS}">` +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    "</Relationships>"
  );
}

function documentRelsXml(hasNumbering: boolean): string {
  const rels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
  ];
  if (hasNumbering) {
    rels.push(
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    );
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<Relationships xmlns="${REL_NS}">${rels.join("")}</Relationships>`
  );
}

function corePropsXml(date: string): string {
  const d = escapeText(date);
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/"' +
    ' xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"' +
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${d}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${d}</dcterms:modified>` +
    "</cp:coreProperties>"
  );
}

// ---------------------------------------------------------------------------
// docx_create
// ---------------------------------------------------------------------------

export interface DocxCreateArgs {
  content_md?: string | undefined;
  spec?: Record<string, unknown> | undefined;
}

export interface DocxCreateResult {
  doc_id: string;
  n_paragraphs: number;
}

/** Create a new document from Markdown (or a structured spec) and register it. */
export function docxCreate(session: Session, args: DocxCreateArgs): DocxCreateResult {
  if (args.content_md !== undefined && args.spec !== undefined) {
    throw new ToolError("invalid_args", "Provide exactly one of content_md or spec.", [
      "content_md and spec are mutually exclusive.",
    ]);
  }
  const md = args.content_md ?? (args.spec !== undefined ? specToMarkdown(args.spec) : "");
  const built = buildBody(md);
  const hasNumbering = built.numbering.length > 0;
  const date = coreDate();

  // §22 creation order: document, styles, [Content_Types], _rels/.rels,
  // word/_rels/document.xml.rels, docProps/core.xml; numbering after styles
  // when present (it must precede the rels that reference it for zip order, but
  // the spec only pins the six base parts — numbering rides alongside styles).
  const parts: Record<string, string> = {
    "word/document.xml": documentXml(built.paragraphs.join("")),
    "word/styles.xml": stylesXml(),
  };
  if (hasNumbering) parts["word/numbering.xml"] = numberingXml(built.numbering);
  parts["[Content_Types].xml"] = contentTypesXml(hasNumbering);
  parts["_rels/.rels"] = rootRelsXml();
  parts["word/_rels/document.xml.rels"] = documentRelsXml(hasNumbering);
  parts["docProps/core.xml"] = corePropsXml(date);

  const bytes = zipPackage(parts);
  const doc = session.open(bytes);

  // §22/Phase-2 invariant: every create ends by passing the §8 validator.
  const issues = validateDoc(doc);
  if (!isValid(issues)) {
    const errors = issues.filter((i) => i.severity === "error");
    throw new ToolError(
      "validation_failed",
      "Created document failed validation.",
      errors.map((e) => e.message),
    );
  }

  return { doc_id: doc.id, n_paragraphs: built.bodyParagraphCount };
}

/** Build the zip bytes for the skeleton parts (level 0 — re-stored on save). */
function zipPackage(parts: Record<string, string>): Uint8Array {
  const zippable: Zippable = {};
  for (const [name, xml] of Object.entries(parts)) zippable[name] = strToU8(xml);
  return zipSync(zippable, { level: 0 });
}

/**
 * Minimal structured-spec → markdown shim. A `spec` object with a `blocks`
 * array of `{type, text, level}` is lowered to markdown so the same block
 * mapping applies. Unknown specs degrade to an empty document.
 */
function specToMarkdown(spec: Record<string, unknown>): string {
  const blocks = spec["blocks"];
  if (!Array.isArray(blocks)) return "";
  const lines: string[] = [];
  for (const b of blocks) {
    if (b == null || typeof b !== "object") continue;
    const block = b as Record<string, unknown>;
    const text = typeof block["text"] === "string" ? (block["text"] as string) : "";
    const type = typeof block["type"] === "string" ? (block["type"] as string) : "paragraph";
    if (type === "heading") {
      const level =
        typeof block["level"] === "number" ? Math.min(6, Math.max(1, block["level"] as number)) : 1;
      lines.push(`${"#".repeat(level)} ${text}`);
    } else if (type === "list_item" || type === "bullet") {
      lines.push(`- ${text}`);
    } else {
      lines.push(text);
    }
    lines.push("");
  }
  return lines.join("\n");
}
