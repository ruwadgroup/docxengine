/**
 * Templates (`docx_template_fill`) per spec/algorithms.md §21.
 *
 * A mustache subset — `{{var}}`, `{{#s}}…{{/s}}` (loop / render-once on truthy),
 * `{{^s}}…{{/s}}` (inverted), `{{!c}}` (comment, dropped) — is matched against
 * the §4 *coalesced* paragraph text so split-run placeholders resolve. A
 * `{{var}}` is written into the first overlapping `w:t` and the rest are
 * trimmed (§4 first-overlap). Loop/inverted sections whose open and close tags
 * sit in whole paragraphs of one body region clone the spanned paragraphs per
 * array element (substituting `{{.}}`/`{{key}}`); when both tags sit in cells of
 * exactly one table row, the row is cloned. Missing vars stay verbatim and are
 * listed in `unfilled` (dedup, document order); `strict:true` raises
 * `placeholder_unfilled`. Emission is XML-escaping only (§3), never HTML.
 *
 * The Python twin (`_template.py`) is the byte-parity reference. The engine
 * operates purely on the document-part text (no throwaway packages) so the
 * splice output is deterministic across languages.
 */
import { ToolError } from "./errors.js";
import type { DocHandle, Session } from "./session.js";
import {
  type ElementSlice,
  childElements,
  elementExtent,
  escapeText,
  findBody,
  needsSpacePreserve,
  nextTag,
  splice,
  textPieces,
} from "./xmlscan.js";

// ---------------------------------------------------------------------------
// Tag tokens
// ---------------------------------------------------------------------------

type TokenKind = "var" | "section" | "inverted" | "close" | "comment";

interface TagToken {
  kind: TokenKind;
  key: string;
  /** Start of `{{` within the scanned string. */
  start: number;
  /** One past `}}` within the scanned string. */
  end: number;
}

const TAG_RE = /\{\{([#^/!]?)\s*([^{}]*?)\s*\}\}/g;

function classify(sigil: string): TokenKind {
  switch (sigil) {
    case "#":
      return "section";
    case "^":
      return "inverted";
    case "/":
      return "close";
    case "!":
      return "comment";
    default:
      return "var";
  }
}

/** Scan a coalesced string for mustache tags, in order. */
function scanTags(text: string): TagToken[] {
  const out: TagToken[] = [];
  TAG_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = TAG_RE.exec(text)) !== null) {
    out.push({
      kind: classify(m[1] ?? ""),
      key: m[2] ?? "",
      start: m.index,
      end: m.index + m[0].length,
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Value model
// ---------------------------------------------------------------------------

type Scope = Record<string, unknown>;

function lookup(scopes: Scope[], key: string): unknown {
  if (key === ".") {
    const top = scopes.length > 0 ? scopes[scopes.length - 1] : undefined;
    // Scalar loop elements are wrapped as `{".": value}` by asScope.
    if (top != null && typeof top === "object" && "." in top) return (top as Scope)["."];
    return top;
  }
  for (let i = scopes.length - 1; i >= 0; i--) {
    const s = scopes[i];
    if (s != null && typeof s === "object" && !Array.isArray(s) && key in s) {
      return (s as Scope)[key];
    }
  }
  return undefined;
}

function isTruthy(value: unknown): boolean {
  if (value === undefined || value === null || value === false || value === "") return false;
  if (Array.isArray(value)) return value.length > 0;
  return true;
}

function stringify(value: unknown): string {
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "number") return String(value);
  if (typeof value === "string") return value;
  return "";
}

/** Wrap a loop element so `{{.}}` resolves to a scalar element. */
function asScope(item: unknown): Scope {
  if (item != null && typeof item === "object" && !Array.isArray(item)) return item as Scope;
  return { ".": item } as unknown as Scope;
}

// ---------------------------------------------------------------------------
// Fill state
// ---------------------------------------------------------------------------

interface FillState {
  filled: number;
  loopsExpanded: Map<string, number>;
  unfilled: string[];
  unfilledSet: Set<string>;
  strict: boolean;
}

function noteUnfilled(state: FillState, key: string): void {
  if (!state.unfilledSet.has(key)) {
    state.unfilledSet.add(key);
    state.unfilled.push(key);
  }
}

// ---------------------------------------------------------------------------
// Paragraph enumeration (within an arbitrary markup string)
// ---------------------------------------------------------------------------

function paragraphsIn(xml: string, from = 0, to: number = xml.length): ElementSlice[] {
  const out: ElementSlice[] = [];
  let i = from;
  for (;;) {
    const t = nextTag(xml, i, to);
    if (!t) return out;
    if (t.name === "w:p" && t.kind !== "end") {
      const el = elementExtent(xml, t, to);
      out.push(el);
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

function tablesIn(xml: string, from = 0, to: number = xml.length): ElementSlice[] {
  const out: ElementSlice[] = [];
  let i = from;
  for (;;) {
    const t = nextTag(xml, i, to);
    if (!t) return out;
    if (t.kind === "end") {
      i = t.end;
      continue;
    }
    if (t.name === "w:tbl") {
      const el = elementExtent(xml, t, to);
      out.push(el);
      i = el.end;
      continue;
    }
    i = elementExtent(xml, t, to).end;
  }
}

/** Coalesced `w:t` text of a scope (paragraph or row). */
function coalescedText(xml: string, scope: { contentStart: number; contentEnd: number }): string {
  let s = "";
  for (const p of textPieces(xml, scope)) if (p.kind === "t") s += p.text;
  return s;
}

// ---------------------------------------------------------------------------
// Var substitution within one paragraph (§4 first-overlap), pure string in/out
// ---------------------------------------------------------------------------

/**
 * Substitute `{{var}}`/`{{!comment}}` in one paragraph slice of `xml`. Section
 * tokens are left in place. Returns the new full `xml`. Variable values are
 * written into the first overlapping `w:t`; later overlapping `w:t` lose the
 * covered span (§4). `pStart` is the paragraph's `<` offset.
 */
function fillParagraphVars(
  xml: string,
  para: ElementSlice,
  scopes: Scope[],
  state: FillState,
): string {
  const pieces = textPieces(xml, para).filter((p) => p.kind === "t");
  if (pieces.length === 0) return xml;
  let text = "";
  const ranges: { start: number; end: number }[] = [];
  for (const piece of pieces) {
    ranges.push({ start: text.length, end: text.length + piece.text.length });
    text += piece.text;
  }
  const tags = scanTags(text).filter((t) => t.kind === "var" || t.kind === "comment");
  if (tags.length === 0) return xml;

  // Resolve each tag → either a replacement value (drop covered span, emit value
  // into the first overlapping piece) or "leave verbatim".
  interface Resolved {
    start: number;
    end: number;
    value: string;
  }
  const resolved: Resolved[] = [];
  for (const t of tags) {
    if (t.kind === "comment") {
      resolved.push({ start: t.start, end: t.end, value: "" });
      continue;
    }
    const v = lookup(scopes, t.key);
    if (v === undefined) {
      noteUnfilled(state, t.key);
      continue; // leave verbatim
    }
    resolved.push({ start: t.start, end: t.end, value: stringify(v) });
    state.filled += 1;
  }
  if (resolved.length === 0) return xml;
  resolved.sort((a, b) => a.start - b.start);

  // Compute the new text of each piece.
  const newTexts: string[] = [];
  for (let pi = 0; pi < pieces.length; pi++) {
    const range = ranges[pi] as { start: number; end: number };
    let out = "";
    let cursor = range.start;
    while (cursor < range.end) {
      const r = resolved.find((e) => e.start <= cursor && cursor < e.end);
      if (r) {
        // Emit the value only at the position where the tag starts, and only in
        // the first piece that reaches it.
        if (cursor === r.start) out += r.value;
        cursor = Math.min(r.end, range.end);
        continue;
      }
      out += (pieces[pi] as { text: string }).text[cursor - range.start] as string;
      cursor += 1;
    }
    newTexts.push(out);
  }

  // Splice each piece's w:t content (descending offset order to keep positions).
  const edits = pieces.map((piece, pi) => {
    const el = (piece as { el: ElementSlice }).el;
    const t = newTexts[pi] as string;
    const space = needsSpacePreserve(t) ? ' xml:space="preserve"' : "";
    return { start: el.start, end: el.end, text: `<w:t${space}>${escapeText(t)}</w:t>` };
  });
  let result = xml;
  for (const e of [...edits].sort((a, b) => b.start - a.start)) {
    result = splice(result, e.start, e.end, e.text);
  }
  return result;
}

/** Erase residual section/close/comment tags from a paragraph's `w:t` pieces. */
function dropSectionTags(xml: string, para: ElementSlice): string {
  const pieces = textPieces(xml, para).filter((p) => p.kind === "t");
  if (pieces.length === 0) return xml;
  let text = "";
  const ranges: { start: number; end: number }[] = [];
  for (const piece of pieces) {
    ranges.push({ start: text.length, end: text.length + piece.text.length });
    text += piece.text;
  }
  const tags = scanTags(text).filter(
    (t) => t.kind === "section" || t.kind === "inverted" || t.kind === "close",
  );
  if (tags.length === 0) return xml;
  const newTexts: string[] = [];
  for (let pi = 0; pi < pieces.length; pi++) {
    const range = ranges[pi] as { start: number; end: number };
    let out = "";
    let cursor = range.start;
    while (cursor < range.end) {
      const tag = tags.find((t) => t.start <= cursor && cursor < t.end);
      if (tag) {
        cursor = Math.min(tag.end, range.end);
        continue;
      }
      out += (pieces[pi] as { text: string }).text[cursor - range.start] as string;
      cursor += 1;
    }
    newTexts.push(out);
  }
  const edits = pieces.map((piece, pi) => {
    const el = (piece as { el: ElementSlice }).el;
    const t = newTexts[pi] as string;
    const space = needsSpacePreserve(t) ? ' xml:space="preserve"' : "";
    return { start: el.start, end: el.end, text: `<w:t${space}>${escapeText(t)}</w:t>` };
  });
  let result = xml;
  for (const e of [...edits].sort((a, b) => b.start - a.start)) {
    result = splice(result, e.start, e.end, e.text);
  }
  return result;
}

// ---------------------------------------------------------------------------
// Region expansion: returns rewritten markup for a region [from, to) of `xml`
// ---------------------------------------------------------------------------

/**
 * Fully render a markup fragment under the scope chain. The fragment is
 * processed left-to-right in **segments**: section spans are expanded (each
 * element rendered recursively, already-final), and the markup outside any
 * section has its `{{var}}` filled exactly once. Returns the rendered string.
 *
 * Single-pass discipline guarantees no value is filled twice — already-rendered
 * clone output is never re-scanned.
 */
function renderFragment(fragment: string, scopes: Scope[], state: FillState): string {
  const span = firstSectionSpan(fragment);
  if (span === null) {
    // No section: fill vars in every paragraph (descending offset order).
    let work = fragment;
    for (const p of [...paragraphsIn(work)].reverse()) {
      const para = findParaAt(work, p.start) ?? p;
      work = dropSectionTags(work, para);
      const para2 = findParaAt(work, p.start) ?? para;
      work = fillParagraphVars(work, para2, scopes, state);
    }
    return work;
  }
  // Render: [before] + [expanded section] + [recurse on the rest].
  const before = renderPlain(fragment.slice(0, span.start), scopes, state);
  const expanded = expandSpan(span, scopes, state);
  const rest = renderFragment(fragment.slice(span.end), scopes, state);
  return before + expanded + rest;
}

/** Fill vars in a section-free markup slice (paragraphs only). */
function renderPlain(fragment: string, scopes: Scope[], state: FillState): string {
  let work = fragment;
  for (const p of [...paragraphsIn(work)].reverse()) {
    const para = findParaAt(work, p.start) ?? p;
    work = fillParagraphVars(work, para, scopes, state);
  }
  return work;
}

function findParaAt(xml: string, start: number): ElementSlice | null {
  const t = nextTag(xml, start);
  if (!t || t.name !== "w:p" || t.kind === "end") return null;
  return elementExtent(xml, t);
}

interface SectionSpan {
  kind: "paragraph" | "row";
  token: TagToken;
  /** Offset of the whole region within the fragment. */
  start: number;
  end: number;
  /** Inner markup to clone (paragraphs between tag-only open/close, or the row). */
  inner: string;
}

/**
 * Locate the first section span in `fragment` — a table-row span (open/close in
 * one row's cells) or a paragraph span (open/close on whole paragraphs),
 * whichever begins earlier. Returns null when no section exists.
 */
function firstSectionSpan(fragment: string): SectionSpan | null {
  const rowSpan = firstRowSpan(fragment);
  const paraSpan = firstParagraphSpan(fragment);
  if (rowSpan && paraSpan) return rowSpan.start <= paraSpan.start ? rowSpan : paraSpan;
  return rowSpan ?? paraSpan;
}

function firstParagraphSpan(fragment: string): SectionSpan | null {
  const paras = paragraphsIn(fragment);
  let openIdx = -1;
  let openTok: TagToken | null = null;
  for (let i = 0; i < paras.length; i++) {
    const tags = scanTags(coalescedText(fragment, paras[i] as ElementSlice));
    const open = tags.find((t) => t.kind === "section" || t.kind === "inverted");
    if (open) {
      openIdx = i;
      openTok = open;
      break;
    }
  }
  if (openIdx < 0 || openTok === null) return null;

  let depth = 0;
  let closeIdx = -1;
  for (let i = openIdx; i < paras.length; i++) {
    for (const t of scanTags(coalescedText(fragment, paras[i] as ElementSlice))) {
      if ((t.kind === "section" || t.kind === "inverted") && t.key === openTok.key) {
        if (i === openIdx && t.start === openTok.start) continue;
        depth += 1;
      } else if (t.kind === "close" && t.key === openTok.key) {
        if (depth === 0) {
          closeIdx = i;
          break;
        }
        depth -= 1;
      }
    }
    if (closeIdx >= 0) break;
  }
  if (closeIdx < 0) {
    throw new ToolError("template_syntax", `Unclosed section {{#${openTok.key}}}.`, [
      "Every {{#section}} needs a matching {{/section}}.",
    ]);
  }
  const openPara = paras[openIdx] as ElementSlice;
  const closePara = paras[closeIdx] as ElementSlice;
  const hasInner = closeIdx - openIdx >= 2;
  const inner = hasInner
    ? fragment.slice(
        (paras[openIdx + 1] as ElementSlice).start,
        (paras[closeIdx - 1] as ElementSlice).end,
      )
    : "";
  return { kind: "paragraph", token: openTok, start: openPara.start, end: closePara.end, inner };
}

function firstRowSpan(fragment: string): SectionSpan | null {
  for (const tbl of tablesIn(fragment)) {
    const rows = childElements(fragment, tbl.contentStart, tbl.contentEnd).filter(
      (k) => k.name === "w:tr",
    );
    for (const row of rows) {
      const tags = scanTags(coalescedText(fragment, row));
      const open = tags.find((t) => t.kind === "section" || t.kind === "inverted");
      const close = tags.find((t) => t.kind === "close");
      if (open && close && open.key === close.key) {
        return {
          kind: "row",
          token: open,
          start: row.start,
          end: row.end,
          inner: fragment.slice(row.start, row.end),
        };
      }
    }
  }
  return null;
}

/** Expand one located section span into its rendered replacement markup. */
function expandSpan(span: SectionSpan, scopes: Scope[], state: FillState): string {
  const render = span.kind === "row" ? renderRow : renderFragment;
  const value = lookup(scopes, span.token.key);
  if (span.token.kind === "inverted") {
    // Inverted sections are conditions, not loops: never recorded in
    // loops_expanded (only `{{#section}}` array/truthy expansions are §21).
    return isTruthy(value) ? "" : render(span.inner, scopes, state);
  }
  if (Array.isArray(value)) {
    const items = value as unknown[];
    let out = "";
    for (const item of items) out += render(span.inner, [...scopes, asScope(item)], state);
    state.loopsExpanded.set(span.token.key, items.length);
    return out;
  }
  if (isTruthy(value)) {
    state.loopsExpanded.set(span.token.key, 1);
    return render(span.inner, [...scopes, asScope(value)], state);
  }
  state.loopsExpanded.set(span.token.key, 0);
  return "";
}

/** Render one `w:tr`: drop section tags + fill vars in its cell paragraphs. */
function renderRow(rowXml: string, scopes: Scope[], state: FillState): string {
  let work = rowXml;
  for (const p of [...paragraphsIn(work)].reverse()) {
    const para = findParaAt(work, p.start) ?? p;
    work = dropSectionTags(work, para);
    const para2 = findParaAt(work, p.start) ?? para;
    work = fillParagraphVars(work, para2, scopes, state);
  }
  return work;
}

// ---------------------------------------------------------------------------
// Top-level body processing
// ---------------------------------------------------------------------------

function processDocument(xml: string, data: Scope, state: FillState): string {
  const body = findBody(xml);
  const before = xml.slice(0, body.contentStart);
  const after = xml.slice(body.contentEnd);
  const rendered = renderFragment(xml.slice(body.contentStart, body.contentEnd), [data], state);
  return before + rendered + after;
}

// ---------------------------------------------------------------------------
// docx_template_fill
// ---------------------------------------------------------------------------

export interface DocxTemplateFillArgs {
  template: string;
  data: Record<string, unknown>;
  syntax?: "mustache" | undefined;
  strict?: boolean | undefined;
}

export interface DocxTemplateFillResult {
  doc_id: string;
  filled: number;
  loops_expanded: Record<string, number>;
  unfilled: string[];
  note?: string;
}

/** Fill a mustache template and register the filled doc as the next `d{n}`. */
export function docxTemplateFill(
  session: Session,
  args: DocxTemplateFillArgs,
): DocxTemplateFillResult {
  const syntax = args.syntax ?? "mustache";
  if (syntax !== "mustache") {
    throw new ToolError("template_syntax", `Unsupported template syntax: ${syntax}.`, [
      "Only the mustache subset is supported.",
    ]);
  }
  const doc = session.open(args.template);
  return fillDoc(doc, args.data, { strict: args.strict === true });
}

/**
 * Fill `doc` in place from mustache template `data` (§21); returns fill stats.
 *
 * The in-language core shared by {@link docxTemplateFill} and
 * `Document.fillTemplate` — operates on an already-open handle so the caller
 * controls the source (path *or* bytes) and the session.
 */
export function fillDoc(
  doc: DocHandle,
  data: Record<string, unknown> | undefined,
  opts: { strict?: boolean } = {},
): DocxTemplateFillResult {
  const scope = (data ?? {}) as Scope;
  const state: FillState = {
    filled: 0,
    loopsExpanded: new Map(),
    unfilled: [],
    unfilledSet: new Set(),
    strict: opts.strict === true,
  };

  const filled = processDocument(doc.documentXml(), scope, state);
  doc.pkg.setPart(doc.documentPartName, filled);
  doc.invalidate();

  if (state.strict && state.unfilled.length > 0) {
    throw new ToolError(
      "placeholder_unfilled",
      `Unfilled placeholders: ${state.unfilled.join(", ")}.`,
      ["Supply data for every placeholder or set strict: false."],
    );
  }

  const loops_expanded: Record<string, number> = {};
  for (const [k, v] of state.loopsExpanded) loops_expanded[k] = v;

  return {
    doc_id: doc.id,
    filled: state.filled,
    loops_expanded,
    unfilled: state.unfilled,
    note:
      state.unfilled.length === 0
        ? "All placeholders resolved."
        : `${state.unfilled.length} placeholder${state.unfilled.length === 1 ? "" : "s"} left unfilled.`,
  };
}
