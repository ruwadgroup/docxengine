/**
 * Projection per spec/algorithms.md §2 and §2a: the outline (pStyle resolved
 * through the styles.xml basedOn chain), the agent-facing line format with
 * heading/list annotations and ins/del/comment markers, table projection,
 * windowed/ranged reads with pagination, and search over the §4 coalesced
 * text with snippet + nearest-heading context.
 */
import { normalizedText, paragraphAnchor, parseAnchor } from "./anchors.js";
import { ToolError } from "./errors.js";
import type { DocHandle } from "./session.js";
import {
  type ElementSlice,
  type Tag,
  attrs,
  childElements,
  decodeEntities,
  elementExtent,
  findElement,
  getAttr,
  nextTag,
  scopeText,
} from "./xmlscan.js";

/** Pagination cap on `content` characters (§2a). */
export const READ_CHAR_CAP = 24_000;
const SNIPPET_RADIUS = 40;
const ELLIPSIS = "…";
const TIMES = "×";

const STORY_SCOPES: ReadonlySet<string> = new Set([
  "body",
  "footnotes",
  "comments",
  "headers",
  "footers",
]);

interface ContentRange {
  contentStart: number;
  contentEnd: number;
}

/** Attribute map of an already-resolved element slice. */
function sliceAttrs(xml: string, el: ElementSlice): Record<string, string> {
  const tag: Tag = {
    kind: el.selfClosed ? "empty" : "start",
    name: el.name,
    start: el.start,
    end: el.startTagEnd,
    nameEnd: el.nameEnd,
  };
  return attrs(xml, tag);
}

// ---------------------------------------------------------------------------
// Part-level context: styles.xml, numbering.xml, comments.xml
// ---------------------------------------------------------------------------

export interface NumberingMap {
  /** `w:num` numId → abstractNumId. */
  numToAbstract: Map<string, string>;
  /** abstractNumId → (ilvl → numFmt). */
  abstractFormats: Map<string, Map<number, string>>;
}

export interface ProjectionContext {
  /** styleId → basedOn styleId (or null). */
  styles: Map<string, string | null>;
  numbering: NumberingMap;
  /** comment id → author. */
  commentAuthors: Map<string, string>;
}

/** styleId → basedOn from styles.xml (absent part → empty map). */
export function parseStyles(xml: string | undefined): Map<string, string | null> {
  const map = new Map<string, string | null>();
  if (xml === undefined) return map;
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return map;
    if (t.name === "w:style" && t.kind !== "end") {
      const id = getAttr(xml, t, "w:styleId");
      const el = elementExtent(xml, t);
      if (id !== undefined) {
        const basedOn = el.selfClosed
          ? null
          : findElement(xml, "w:basedOn", el.contentStart, el.contentEnd);
        map.set(id, basedOn ? (sliceAttrs(xml, basedOn)["w:val"] ?? null) : null);
      }
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

/** Numbering chains from numbering.xml (absent part → empty maps). */
export function parseNumbering(xml: string | undefined): NumberingMap {
  const numToAbstract = new Map<string, string>();
  const abstractFormats = new Map<string, Map<number, string>>();
  if (xml === undefined) return { numToAbstract, abstractFormats };
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return { numToAbstract, abstractFormats };
    if (t.name === "w:abstractNum" && t.kind === "start") {
      const el = elementExtent(xml, t);
      const id = getAttr(xml, t, "w:abstractNumId");
      if (id !== undefined) {
        const formats = new Map<number, string>();
        for (const lvl of childElements(xml, el.contentStart, el.contentEnd)) {
          if (lvl.name !== "w:lvl" || lvl.selfClosed) continue;
          const ilvl = sliceAttrs(xml, lvl)["w:ilvl"];
          const fmtEl = findElement(xml, "w:numFmt", lvl.contentStart, lvl.contentEnd);
          const fmt = fmtEl ? sliceAttrs(xml, fmtEl)["w:val"] : undefined;
          if (ilvl !== undefined && fmt !== undefined) formats.set(Number(ilvl), fmt);
        }
        abstractFormats.set(id, formats);
      }
      i = el.end;
      continue;
    }
    if (t.name === "w:num" && t.kind === "start") {
      const el = elementExtent(xml, t);
      const numId = getAttr(xml, t, "w:numId");
      const absEl = findElement(xml, "w:abstractNumId", el.contentStart, el.contentEnd);
      const abs = absEl ? sliceAttrs(xml, absEl)["w:val"] : undefined;
      if (numId !== undefined && abs !== undefined) numToAbstract.set(numId, abs);
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

/** comment id → author from comments.xml (absent part → empty map). */
export function parseCommentAuthors(xml: string | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (xml === undefined) return map;
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return map;
    if (t.name === "w:comment" && t.kind !== "end") {
      const a = attrs(xml, t);
      const id = a["w:id"];
      if (id !== undefined) map.set(id, a["w:author"] ?? "unknown");
    }
    i = t.end;
  }
}

/** Build the projection context for a document (styles/numbering/comments). */
export function projectionContext(doc: DocHandle): ProjectionContext {
  const text = (name: string): string | undefined =>
    doc.pkg.has(name) ? doc.pkg.partText(name) : undefined;
  return {
    styles: parseStyles(text("word/styles.xml")),
    numbering: parseNumbering(text("word/numbering.xml")),
    commentAuthors: parseCommentAuthors(text("word/comments.xml")),
  };
}

// ---------------------------------------------------------------------------
// Paragraph properties: pStyle → heading level, numPr → list annotation
// ---------------------------------------------------------------------------

const HEADING_STYLE_RE = /^Heading([1-9])$/;

/**
 * Effective heading level of a styleId: `Heading{n}` itself, or reached via
 * the basedOn chain (§2). Cycles terminate; unknown styles stop the walk.
 */
export function headingLevel(
  styleId: string | undefined,
  styles: Map<string, string | null>,
): number | null {
  const seen = new Set<string>();
  let cur = styleId;
  while (cur !== undefined && !seen.has(cur)) {
    const m = HEADING_STYLE_RE.exec(cur);
    if (m) return Number(m[1]);
    seen.add(cur);
    const next = styles.get(cur);
    cur = next ?? undefined;
  }
  return null;
}

interface ParaProps {
  styleId: string | undefined;
  ilvl: number;
  numId: string | undefined;
}

function paraProps(xml: string, p: ContentRange): ParaProps {
  const pPr = findElement(xml, "w:pPr", p.contentStart, p.contentEnd);
  if (!pPr || pPr.selfClosed) return { styleId: undefined, ilvl: 0, numId: undefined };
  const pStyle = findElement(xml, "w:pStyle", pPr.contentStart, pPr.contentEnd);
  const styleId = pStyle ? sliceAttrs(xml, pStyle)["w:val"] : undefined;
  let ilvl = 0;
  let numId: string | undefined;
  const numPr = findElement(xml, "w:numPr", pPr.contentStart, pPr.contentEnd);
  if (numPr && !numPr.selfClosed) {
    const ilvlEl = findElement(xml, "w:ilvl", numPr.contentStart, numPr.contentEnd);
    if (ilvlEl) {
      const v = Number(sliceAttrs(xml, ilvlEl)["w:val"]);
      if (Number.isInteger(v) && v >= 0) ilvl = v;
    }
    const numIdEl = findElement(xml, "w:numId", numPr.contentStart, numPr.contentEnd);
    if (numIdEl) numId = sliceAttrs(xml, numIdEl)["w:val"];
  }
  return { styleId, ilvl, numId };
}

/** Effective heading level of a paragraph element. */
export function paragraphHeadingLevel(
  xml: string,
  p: ContentRange,
  styles: Map<string, string | null>,
): number | null {
  return headingLevel(paraProps(xml, p).styleId, styles);
}

/** §2 annotation tokens, in the fixed order: heading first, then list. */
export function annotationTokens(xml: string, p: ContentRange, ctx: ProjectionContext): string[] {
  const props = paraProps(xml, p);
  const tokens: string[] = [];
  const lvl = headingLevel(props.styleId, ctx.styles);
  if (lvl !== null) tokens.push(`H${lvl}`);
  if (props.numId !== undefined && props.numId !== "0") {
    const abstract = ctx.numbering.numToAbstract.get(props.numId);
    const fmt =
      abstract === undefined
        ? undefined
        : ctx.numbering.abstractFormats.get(abstract)?.get(props.ilvl);
    tokens.push(`List:${fmt === "bullet" ? "ul" : "ol"} L${props.ilvl + 1}`);
  }
  return tokens;
}

// ---------------------------------------------------------------------------
// Decorated paragraph text: as-if-accepted, with §2 markers
// ---------------------------------------------------------------------------

/**
 * The paragraph's normalized text with `[ins by …]` / `[del by …]` /
 * `[comment:C{id} by …]` markers inserted at span ends (§2: markers go into
 * the raw concatenation with one space on each side; normalization then
 * absorbs doubled/edge spaces). `w:delText` content is not shown.
 */
export function decoratedText(
  xml: string,
  p: ContentRange,
  commentAuthors: Map<string, string>,
): string {
  let out = "";
  const stack: string[] = [];
  let i = p.contentStart;
  for (;;) {
    const t = nextTag(xml, i, p.contentEnd);
    if (!t) break;
    if (t.name === "w:ins" || t.name === "w:del") {
      const word = t.name === "w:ins" ? "ins" : "del";
      if (t.kind === "end") {
        const marker = stack.pop();
        if (marker !== undefined) out += marker;
      } else {
        const author = getAttr(xml, t, "w:author") ?? "unknown";
        const marker = ` [${word} by ${author}] `;
        if (t.kind === "start") stack.push(marker);
        else out += marker; // empty wrapper (e.g. paragraph-mark revision)
      }
      i = t.end;
      continue;
    }
    if (t.name === "w:commentReference" && t.kind !== "end") {
      const id = getAttr(xml, t, "w:id") ?? "?";
      const author = commentAuthors.get(id) ?? "unknown";
      out += ` [comment:C${id} by ${author}] `;
      i = t.kind === "empty" ? t.end : elementExtent(xml, t, p.contentEnd).end;
      continue;
    }
    if (t.name === "w:t" && t.kind !== "end") {
      const el = elementExtent(xml, t, p.contentEnd);
      if (!el.selfClosed) out += decodeEntities(xml.slice(el.contentStart, el.contentEnd));
      i = el.end;
      continue;
    }
    if (t.name === "w:delText" && t.kind !== "end") {
      i = elementExtent(xml, t, p.contentEnd).end; // shown as-if-accepted
      continue;
    }
    i = t.end;
  }
  return normalizedText(out);
}

// ---------------------------------------------------------------------------
// Projection blocks (body or story)
// ---------------------------------------------------------------------------

export interface ProjBlock {
  kind: "p" | "tbl";
  /** 1-based among same-kind blocks of the scope. */
  ordinal: number;
  anchor: string;
  /** The part text this block lives in. */
  xml: string;
  block: ElementSlice;
  /** Normalized text (paragraphs; null for tables). */
  normalized: string | null;
}

/** Body-level blocks with their §1 anchors. */
export function bodyProjBlocks(doc: DocHandle): ProjBlock[] {
  const xml = doc.documentXml();
  return doc.anchorIndex().map((e) => ({
    kind: e.kind,
    ordinal: e.ordinal,
    anchor: e.anchor,
    xml,
    block: e.block,
    normalized: e.normalized,
  }));
}

function allParagraphs(xml: string): ElementSlice[] {
  const out: ElementSlice[] = [];
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return out;
    if (t.name === "w:p" && t.kind !== "end") {
      const el = elementExtent(xml, t);
      out.push(el);
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

function storyPartNames(doc: DocHandle, scope: string): string[] {
  if (scope === "footnotes") return ["word/footnotes.xml"];
  if (scope === "comments") return ["word/comments.xml"];
  const re = new RegExp(`^word/${scope === "headers" ? "header" : "footer"}([0-9]*)\\.xml$`);
  return doc.pkg
    .entryNames()
    .map((name) => ({ name, m: re.exec(name) }))
    .filter((x): x is { name: string; m: RegExpExecArray } => x.m !== null)
    .sort((a, b) => Number(a.m[1] || "0") - Number(b.m[1] || "0"))
    .map((x) => x.name);
}

/**
 * Blocks of a story scope (§2a): `body` → body-level blocks; other stories →
 * every `w:p` of the story part(s), with anchors over the story's own
 * paragraph sequence. Missing story parts contribute nothing.
 */
export function storyProjBlocks(doc: DocHandle, scope: string): ProjBlock[] {
  if (scope === "body") return bodyProjBlocks(doc);
  const out: ProjBlock[] = [];
  let ordinal = 0;
  for (const name of storyPartNames(doc, scope)) {
    if (!doc.pkg.has(name)) continue;
    const xml = doc.pkg.partText(name);
    for (const p of allParagraphs(xml)) {
      ordinal++;
      const normalized = normalizedText(scopeText(xml, p));
      out.push({
        kind: "p",
        ordinal,
        anchor: paragraphAnchor(ordinal, normalized),
        xml,
        block: p,
        normalized,
      });
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Rendering: paragraph lines and table projections (§2)
// ---------------------------------------------------------------------------

/** One §2 projection line: `[{anchor}{annotations}] {text}`. */
export function paragraphLine(b: ProjBlock, ctx: ProjectionContext): string {
  const tokens = annotationTokens(b.xml, b.block, ctx);
  const head = `[${b.anchor}${tokens.length > 0 ? ` ${tokens.join(" ")}` : ""}]`;
  const text = decoratedText(b.xml, b.block, ctx.commentAuthors);
  return text === "" ? head : `${head} ${text}`;
}

/** rows = `w:tr` count; cols = `w:gridCol` count (fallback: max `w:tc` per row). */
export function tableDims(xml: string, tbl: ContentRange): { rows: number; cols: number } {
  const kids = childElements(xml, tbl.contentStart, tbl.contentEnd);
  const trs = kids.filter((k) => k.name === "w:tr");
  let cols = 0;
  const grid = kids.find((k) => k.name === "w:tblGrid");
  if (grid && !grid.selfClosed) {
    cols = childElements(xml, grid.contentStart, grid.contentEnd).filter(
      (k) => k.name === "w:gridCol",
    ).length;
  }
  if (cols === 0) {
    for (const tr of trs) {
      if (tr.selfClosed) continue;
      const n = childElements(xml, tr.contentStart, tr.contentEnd).filter(
        (k) => k.name === "w:tc",
      ).length;
      if (n > cols) cols = n;
    }
  }
  return { rows: trs.length, cols };
}

function cellText(xml: string, tc: ElementSlice): string {
  if (tc.selfClosed) return "";
  const texts = childElements(xml, tc.contentStart, tc.contentEnd)
    .filter((k) => k.name === "w:p")
    .map((p) => normalizedText(scopeText(xml, p)));
  return texts.join(" ").replace(/\|/g, "\\|");
}

/** §2 table projection: header line + GitHub-style markdown rows. */
export function tableProjection(b: ProjBlock, prevParaAnchor: string | null): string {
  const { rows, cols } = tableDims(b.xml, b.block);
  const at = prevParaAnchor === null ? "@start" : `@after:${prevParaAnchor}`;
  const lines = [`[${b.anchor} ${rows}${TIMES}${cols} ${at}]`];
  const trs = childElements(b.xml, b.block.contentStart, b.block.contentEnd).filter(
    (k) => k.name === "w:tr",
  );
  trs.forEach((tr, idx) => {
    const cells = tr.selfClosed
      ? []
      : childElements(b.xml, tr.contentStart, tr.contentEnd).filter((k) => k.name === "w:tc");
    lines.push(`| ${cells.map((tc) => cellText(b.xml, tc)).join(" | ")} |`);
    if (idx === 0) lines.push(`|${" --- |".repeat(Math.max(cols, 1))}`);
  });
  return lines.join("\n");
}

function prevParagraphAnchor(blocks: ProjBlock[], idx: number): string | null {
  for (let j = idx - 1; j >= 0; j--) {
    const b = blocks[j] as ProjBlock;
    if (b.kind === "p") return b.anchor;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Outline (docx_outline)
// ---------------------------------------------------------------------------

export interface OutlineHeading {
  anchor: string;
  level: number;
  text: string;
}

export interface OutlineTable {
  anchor: string;
  dims: string;
  after?: string;
}

/** Heading tree + table list per spec/tools/docx_outline.json. */
export function outlineOf(doc: DocHandle): { outline: OutlineHeading[]; tables: OutlineTable[] } {
  const ctx = projectionContext(doc);
  const outline: OutlineHeading[] = [];
  const tables: OutlineTable[] = [];
  let prevPara: string | null = null;
  for (const b of bodyProjBlocks(doc)) {
    if (b.kind === "p") {
      const level = paragraphHeadingLevel(b.xml, b.block, ctx.styles);
      if (level !== null) outline.push({ anchor: b.anchor, level, text: b.normalized ?? "" });
      prevPara = b.anchor;
    } else {
      const { rows, cols } = tableDims(b.xml, b.block);
      const dims = `${rows}${TIMES}${cols}`;
      tables.push(
        prevPara === null
          ? { anchor: b.anchor, dims }
          : { anchor: b.anchor, dims, after: prevPara },
      );
    }
  }
  return { outline, tables };
}

// ---------------------------------------------------------------------------
// Read (docx_read): anchor+window / range / whole story, with pagination
// ---------------------------------------------------------------------------

export interface ReadOpts {
  anchor?: string | undefined;
  range?: string | undefined;
  window?: number | undefined;
  scope?: string | undefined;
}

export interface ReadResult {
  content: string;
  continuation?: string;
}

const RANGE_END_RE = /^P([1-9][0-9]*)(?:#[0-9a-f]{4})?$/;

/** Parse `P{a}..P{b}` (optional `#hhhh` suffixes ignored) → [a, b]. */
export function parseParagraphRange(range: string): [number, number] {
  const sep = range.indexOf("..");
  const ma = sep < 0 ? null : RANGE_END_RE.exec(range.slice(0, sep));
  const mb = sep < 0 ? null : RANGE_END_RE.exec(range.slice(sep + 2));
  if (!ma || !mb) {
    throw new ToolError("anchor_invalid", `Malformed range string: ${range}.`, [
      "Use the form 'P10..P24'.",
    ]);
  }
  const a = Number(ma[1]);
  const b = Number(mb[1]);
  if (a > b) {
    throw new ToolError("anchor_invalid", `Malformed range string: ${range} (start after end).`, [
      "Use the form 'P10..P24'.",
    ]);
  }
  return [a, b];
}

function requireScope(scope: string): void {
  if (!STORY_SCOPES.has(scope)) {
    throw new ToolError("anchor_invalid", `Unknown scope: ${scope}.`, [
      "Use body, footnotes, comments, headers or footers.",
    ]);
  }
}

/** The §2 projection of a window/range/story, paginated per §2a. */
export function readProjection(doc: DocHandle, opts: ReadOpts): ReadResult {
  const scope = opts.scope ?? "body";
  requireScope(scope);
  const ctx = projectionContext(doc);
  const blocks = storyProjBlocks(doc, scope);
  let lo = 0;
  let hi = blocks.length - 1;
  if (opts.anchor !== undefined) {
    const parsed = parseAnchor(opts.anchor);
    const idx = blocks.findIndex((b) => b.kind === parsed.kind && b.ordinal === parsed.ordinal);
    if (idx < 0) {
      throw new ToolError(
        "anchor_not_found",
        `Anchor ${opts.anchor} not found: index out of range or table anchor missing.`,
        ["Call docx_outline to re-map anchors."],
      );
    }
    const w = Math.max(0, Math.trunc(opts.window ?? 0));
    lo = Math.max(0, idx - w);
    hi = Math.min(blocks.length - 1, idx + w);
  } else if (opts.range !== undefined) {
    const [a, b] = parseParagraphRange(opts.range);
    const i = blocks.findIndex((x) => x.kind === "p" && x.ordinal === a);
    const j = blocks.findIndex((x) => x.kind === "p" && x.ordinal === b);
    if (i < 0 || j < 0) {
      throw new ToolError(
        "anchor_not_found",
        `Anchor P${i < 0 ? a : b} not found: index out of range or table anchor missing.`,
        ["Call docx_outline to re-map anchors."],
      );
    }
    lo = i;
    hi = j;
  }
  let lastPara = -1;
  for (let k = hi; k >= lo; k--) {
    const b = blocks[k] as ProjBlock;
    if (b.kind === "p") {
      lastPara = b.ordinal;
      break;
    }
  }
  let content = "";
  for (let k = lo; k <= hi; k++) {
    const b = blocks[k] as ProjBlock;
    const piece =
      b.kind === "p" ? paragraphLine(b, ctx) : tableProjection(b, prevParagraphAnchor(blocks, k));
    const sep = content === "" ? "" : "\n";
    if (
      b.kind === "p" &&
      content !== "" &&
      content.length + sep.length + piece.length > READ_CHAR_CAP
    ) {
      return { content, continuation: `P${b.ordinal}..P${lastPara}` };
    }
    content += sep + piece;
  }
  return { content };
}

// ---------------------------------------------------------------------------
// Search (docx_search): §4 coalesced text, snippets, nearest-heading context
// ---------------------------------------------------------------------------

export interface SearchOpts {
  query: string;
  regex?: boolean | undefined;
  scope?: string | undefined;
}

export interface SearchMatch {
  anchor: string;
  snippet: string;
  context?: string;
}

function makeSnippet(raw: string, start: number, end: number): string {
  const from = Math.max(0, start - SNIPPET_RADIUS);
  const to = Math.min(raw.length, end + SNIPPET_RADIUS);
  let snippet = normalizedText(raw.slice(from, to));
  if (from > 0) snippet = ELLIPSIS + snippet;
  if (to < raw.length) snippet += ELLIPSIS;
  return snippet;
}

function occurrences(raw: string, query: string, re: RegExp | null): [number, number][] {
  const out: [number, number][] = [];
  if (re !== null) {
    for (const m of raw.matchAll(re)) {
      const hit = m[0];
      if (hit === "") continue; // zero-length matches are skipped (§2a)
      const at = m.index ?? 0;
      out.push([at, at + hit.length]);
    }
    return out;
  }
  let i = raw.indexOf(query);
  while (i >= 0) {
    out.push([i, i + query.length]);
    i = raw.indexOf(query, i + query.length);
  }
  return out;
}

/** Search per spec/tools/docx_search.json and §2a. */
export function searchProjection(
  doc: DocHandle,
  opts: SearchOpts,
): { matches: SearchMatch[]; n_matches: number } {
  const query = opts.query;
  if (typeof query !== "string" || query === "") {
    throw new ToolError("not_found", "Text not found: empty query.", [
      "Provide a non-empty query string.",
    ]);
  }
  let re: RegExp | null = null;
  if (opts.regex === true) {
    try {
      re = new RegExp(query, "g");
    } catch (e) {
      throw new ToolError("not_found", `Invalid regex: ${query} (${(e as Error).message}).`, [
        "Escape regex metacharacters or set regex: false.",
      ]);
    }
  }
  const scope = opts.scope ?? "body";
  let blocks: ProjBlock[];
  let inScope: (b: ProjBlock) => boolean;
  if (STORY_SCOPES.has(scope)) {
    blocks = storyProjBlocks(doc, scope);
    inScope = () => true;
  } else if (scope.includes("..")) {
    const [a, b] = parseParagraphRange(scope);
    blocks = bodyProjBlocks(doc);
    inScope = (x) => x.ordinal >= a && x.ordinal <= b;
  } else {
    throw new ToolError("not_found", `Unknown scope: ${scope}.`, [
      "Use a story name or a range like 'P10..P24'.",
    ]);
  }
  const ctx = projectionContext(doc);
  const matches: SearchMatch[] = [];
  let lastHeading: string | null = null;
  for (const b of blocks) {
    if (b.kind !== "p") continue;
    if (paragraphHeadingLevel(b.xml, b.block, ctx.styles) !== null) {
      lastHeading = b.normalized ?? "";
    }
    if (!inScope(b)) continue;
    const raw = scopeText(b.xml, b.block);
    for (const [s, e] of occurrences(raw, query, re)) {
      const snippet = makeSnippet(raw, s, e);
      matches.push(
        lastHeading === null
          ? { anchor: b.anchor, snippet }
          : { anchor: b.anchor, snippet, context: lastHeading },
      );
    }
  }
  return { matches, n_matches: matches.length };
}
