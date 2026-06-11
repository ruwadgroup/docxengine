/**
 * Styles (`docx_style`) and formatting (`docx_format`) per spec/algorithms.md
 * §16. `docx_style` lists / defines / applies named styles; `docx_format`
 * either edits a style definition (style_selector — one document-wide change)
 * or applies direct formatting to an anchor/range.
 *
 * All emission obeys §3 (splice; §3 escaping) and the §16 closed prop set and
 * child order. The Python twin (`_styles.py`) is the byte-parity reference.
 */
import { type AnchorEntry, buildAnchorIndex } from "./anchors.js";
import { ToolError } from "./errors.js";
import { headingLevel, parseStyles } from "./projector.js";
import {
  type CanonProps,
  canonicalizeProps,
  mergeChildren,
  paraPropChildren,
  paraPropsInner,
  paraPropsEdit,
  runPropChildren,
  runPropsInner,
  runPropsEdit,
  RPR_ORDER,
  PPR_ORDER,
} from "./props.js";
import type { DocHandle, Session } from "./session.js";
import { anchorInvalid, anchorNotFound, anchorStale } from "./phase2common.js";
import {
  type ElementSlice,
  type SpliceEdit,
  WS_RUN_RE,
  attrs,
  childElements,
  elementExtent,
  escapeAttr,
  findElement,
  getAttr,
  nextTag,
  splice,
  spliceAll,
} from "./xmlscan.js";

type ResponseFormat = "concise" | "detailed";

const STYLES_PART = "word/styles.xml";

const FULL_ANCHOR_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;
const RANGE_RE = /^P([1-9][0-9]*)(?:#([0-9a-f]{4}))?\.\.P([1-9][0-9]*)(?:#([0-9a-f]{4}))?$/;

// ---------------------------------------------------------------------------
// Style records
// ---------------------------------------------------------------------------

export interface StyleRecord {
  id: string;
  name: string;
  type: string;
  based_on?: string;
  in_use: number;
}

interface RawStyle {
  id: string;
  name: string;
  type: string;
  basedOn: string | null;
  el: ElementSlice;
}

/** All `w:style` elements of styles.xml, in document order (raw, with extents). */
function rawStyles(xml: string): RawStyle[] {
  const out: RawStyle[] = [];
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return out;
    if (t.name === "w:style" && t.kind !== "end") {
      const el = elementExtent(xml, t);
      const a = attrs(xml, t);
      const id = a["w:styleId"] ?? "";
      const type = a["w:type"] ?? "paragraph";
      let name = id;
      let basedOn: string | null = null;
      if (!el.selfClosed) {
        const nameEl = findElement(xml, "w:name", el.contentStart, el.contentEnd);
        if (nameEl) name = getAttr(xml, slcTag(xml, nameEl), "w:val") ?? id;
        const basedEl = findElement(xml, "w:basedOn", el.contentStart, el.contentEnd);
        if (basedEl) basedOn = getAttr(xml, slcTag(xml, basedEl), "w:val") ?? null;
      }
      out.push({ id, name, type, basedOn, el });
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

/** A pseudo-Tag for an already-resolved slice (to reuse `getAttr`). */
function slcTag(_xml: string, el: ElementSlice) {
  return {
    kind: el.selfClosed ? ("empty" as const) : ("start" as const),
    name: el.name,
    start: el.start,
    end: el.startTagEnd,
    nameEnd: el.nameEnd,
  };
}

/**
 * Resolve a style name-or-id to its styleId (§16 resolution). Match order: the
 * styleId verbatim, then the styleId with §1 whitespace removed (`"Heading 2"`
 * → `Heading2`), then a style whose `w:name` equals the argument; otherwise
 * `style_unknown`.
 */
export function resolveStyleId(doc: DocHandle, nameOrId: string): string {
  const xml = doc.pkg.has(STYLES_PART) ? doc.pkg.partText(STYLES_PART) : "";
  const styles = rawStyles(xml);
  const ids = new Set(styles.map((s) => s.id));
  if (ids.has(nameOrId)) return nameOrId;
  const compact = nameOrId.replace(WS_RUN_RE, "");
  if (ids.has(compact)) return compact;
  for (const s of styles) if (s.name === nameOrId) return s.id;
  throw new ToolError("style_unknown", `Named style ${nameOrId} does not exist.`, [
    'Call docx_style {op: "list"} to see available styles.',
  ]);
}

// ---------------------------------------------------------------------------
// docx_style
// ---------------------------------------------------------------------------

export interface DocxStyleArgs {
  doc_id: string;
  op: "list" | "define" | "apply";
  anchor?: string | undefined;
  style?: string | undefined;
  name?: string | undefined;
  based_on?: string | undefined;
  props?: Record<string, unknown> | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxStyleResult {
  styles?: StyleRecord[];
  style_id?: string;
  new_anchor?: string;
  note?: string;
}

function bodyParagraphs(doc: DocHandle): AnchorEntry[] {
  return doc.anchorIndex().filter((e) => e.kind === "p");
}

/** Count body paragraphs whose effective style (§2 cascade) resolves to `id`. */
function styleUsage(doc: DocHandle, styleMap: Map<string, string | null>): Map<string, number> {
  const counts = new Map<string, number>();
  const xml = doc.documentXml();
  for (const e of bodyParagraphs(doc)) {
    const effective = effectiveStyleId(xml, e.block, styleMap);
    if (effective !== null) counts.set(effective, (counts.get(effective) ?? 0) + 1);
  }
  return counts;
}

/**
 * The paragraph's effective styleId for in-use counting: the declared pStyle.
 * (basedOn does not change the styleId — in-use counts the style itself, not
 * its ancestors; the §2 cascade only affects heading-level resolution.)
 */
function effectiveStyleId(
  xml: string,
  p: ElementSlice,
  _styleMap: Map<string, string | null>,
): string | null {
  if (p.selfClosed) return null;
  const pPr = childElements(xml, p.contentStart, p.contentEnd).find((k) => k.name === "w:pPr");
  if (!pPr || pPr.selfClosed) return null;
  const pStyle = findElement(xml, "w:pStyle", pPr.contentStart, pPr.contentEnd);
  if (!pStyle) return null;
  return getAttr(xml, slcTag(xml, pStyle), "w:val") ?? null;
}

export function docxStyle(session: Session, args: DocxStyleArgs): DocxStyleResult {
  const doc = session.get(args.doc_id);
  if (args.op === "list") return styleList(doc);
  if (args.op === "define") return styleDefine(doc, args);
  if (args.op === "apply") return styleApply(doc, args);
  throw new ToolError("invalid_args", `docx_style: unknown op ${String(args.op)}.`, [
    "Use list, define or apply.",
  ]);
}

function styleList(doc: DocHandle): DocxStyleResult {
  const xml = doc.pkg.has(STYLES_PART) ? doc.pkg.partText(STYLES_PART) : "";
  const styles = rawStyles(xml);
  const styleMap = parseStyles(xml || undefined);
  const usage = styleUsage(doc, styleMap);
  const records: StyleRecord[] = styles.map((s) => ({
    id: s.id,
    name: s.name,
    type: s.type,
    ...(s.basedOn !== null ? { based_on: s.basedOn } : {}),
    in_use: usage.get(s.id) ?? 0,
  }));
  return { styles: records };
}

/** §16 define: id = name minus whitespace; collisions take suffix 2,3,…. */
function styleDefine(doc: DocHandle, args: DocxStyleArgs): DocxStyleResult {
  const name = args.name;
  if (name === undefined) {
    throw new ToolError("invalid_args", "docx_style define: name is required.", []);
  }
  const xml = doc.pkg.has(STYLES_PART) ? doc.pkg.partText(STYLES_PART) : emptyStylesXml();
  const existing = new Set(rawStyles(xml).map((s) => s.id));
  const base = name.replace(WS_RUN_RE, "");
  let id = base;
  let n = 2;
  while (existing.has(id)) id = `${base}${n++}`;

  const props = canonicalizeProps(args.props);
  const pPrInner = paraPropsInner(props);
  const rPrInner = runPropsInner(props);
  // §16 child order: w:name, w:basedOn?, w:pPr?, w:rPr?
  let body = `<w:name w:val="${escapeAttr(name)}"/>`;
  if (args.based_on !== undefined) body += `<w:basedOn w:val="${escapeAttr(args.based_on)}"/>`;
  if (pPrInner !== "") body += `<w:pPr>${pPrInner}</w:pPr>`;
  if (rPrInner !== "") body += `<w:rPr>${rPrInner}</w:rPr>`;
  const styleEl = `<w:style w:type="paragraph" w:styleId="${escapeAttr(id)}">${body}</w:style>`;

  if (!doc.pkg.has(STYLES_PART)) {
    doc.pkg.setPart(STYLES_PART, emptyStylesXml().replace("</w:styles>", `${styleEl}</w:styles>`));
    ensureStylesContentType(doc);
  } else {
    doc.pkg.setPart(STYLES_PART, insertBeforeStylesClose(xml, styleEl));
  }
  doc.markDirty();
  return { style_id: id, note: `Defined style '${name}' (${id}).` };
}

function styleApply(doc: DocHandle, args: DocxStyleArgs): DocxStyleResult {
  if (args.anchor === undefined || args.style === undefined) {
    throw new ToolError("invalid_args", "docx_style apply: anchor and style are required.", []);
  }
  const entry = requireParagraph(doc, args.anchor);
  const styleId = resolveStyleId(doc, args.style);
  const xml = doc.documentXml();
  const newXml = applyPStyle(xml, entry.block, styleId);
  doc.pkg.setPart(doc.documentPartName, newXml);
  doc.invalidate();
  const fresh = bodyParagraphs(doc)[entry.ordinal - 1] as AnchorEntry;
  return { new_anchor: fresh.anchor, note: `Applied style '${styleId}' to ${args.anchor}.` };
}

/** Splice `<w:pStyle w:val="{id}"/>` as the first `w:pPr` child (creating pPr). */
export function applyPStyle(xml: string, p: ElementSlice, styleId: string): string {
  const pStyleEl = `<w:pStyle w:val="${escapeAttr(styleId)}"/>`;
  if (p.selfClosed) {
    const open = xml.slice(p.start, p.end - 2);
    return splice(xml, p.start, p.end, `${open}><w:pPr>${pStyleEl}</w:pPr></w:p>`);
  }
  const pPr = childElements(xml, p.contentStart, p.contentEnd).find((k) => k.name === "w:pPr");
  if (!pPr) {
    return splice(xml, p.contentStart, p.contentStart, `<w:pPr>${pStyleEl}</w:pPr>`);
  }
  if (pPr.selfClosed) {
    return splice(xml, pPr.start, pPr.end, `<w:pPr>${pStyleEl}</w:pPr>`);
  }
  // Replace an existing leading pStyle, else insert before the first child.
  const existing = childElements(xml, pPr.contentStart, pPr.contentEnd).find(
    (k) => k.name === "w:pStyle",
  );
  if (existing) {
    return splice(xml, existing.start, existing.end, pStyleEl);
  }
  return splice(xml, pPr.contentStart, pPr.contentStart, pStyleEl);
}

function emptyStylesXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>'
  );
}

/** Splice markup before `</w:styles>` (or expand a self-closed root). */
function insertBeforeStylesClose(xml: string, markup: string): string {
  const close = xml.lastIndexOf("</w:styles>");
  if (close >= 0) return splice(xml, close, close, markup);
  const selfClose = xml.lastIndexOf("/>");
  if (selfClose < 0) return xml;
  return splice(xml, selfClose, selfClose + 2, `>${markup}</w:styles>`);
}

function ensureStylesContentType(doc: DocHandle): void {
  const ct = doc.pkg.partText("[Content_Types].xml");
  if (ct.includes('PartName="/word/styles.xml"')) return;
  const close = ct.lastIndexOf("</Types>");
  if (close < 0) return;
  doc.pkg.setPart(
    "[Content_Types].xml",
    splice(
      ct,
      close,
      close,
      '<Override PartName="/word/styles.xml" ' +
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    ),
  );
}

// ---------------------------------------------------------------------------
// docx_format (§16)
// ---------------------------------------------------------------------------

export interface DocxFormatArgs {
  doc_id: string;
  anchor?: string | undefined;
  range?: string | undefined;
  style_selector?: { style: string } | undefined;
  props: Record<string, unknown>;
  track_changes?: boolean | undefined;
  author?: string | undefined;
}

export interface DocxFormatResult {
  affected: number;
  anchors?: string[];
  note?: string;
}

export function docxFormat(session: Session, args: DocxFormatArgs): DocxFormatResult {
  const doc = session.get(args.doc_id);
  const props = canonicalizeProps(args.props);
  if (args.style_selector !== undefined) {
    return formatStyle(doc, args.style_selector.style, props);
  }
  if (args.anchor === undefined && args.range === undefined) {
    throw anchorInvalid("Provide one of anchor, range, or style_selector.");
  }
  return formatDirect(doc, args, props);
}

/** style_selector: merge props into the style's rPr/pPr — one document-wide edit. */
function formatStyle(doc: DocHandle, styleName: string, props: CanonProps): DocxFormatResult {
  const styleId = resolveStyleId(doc, styleName);
  const xml = doc.pkg.partText(STYLES_PART);
  const style = rawStyles(xml).find((s) => s.id === styleId);
  if (!style) throw new ToolError("style_unknown", `Named style ${styleName} does not exist.`, []);
  const newXml = mergeIntoStyle(xml, style.el, props);
  doc.pkg.setPart(STYLES_PART, newXml);
  doc.markDirty();
  // One document-wide edit to the style definition: no paragraphs are touched
  // directly (§16), so affected is 0 and anchors is empty — matching the Python
  // engine and the spec, which reserves affected/anchors for the direct branch.
  return {
    affected: 0,
    anchors: [],
    note: `Edited style ${styleId}.`,
  };
}

/** Merge §16 props into a `w:style` element's pPr/rPr, creating them in order. */
export function mergeIntoStyle(xml: string, style: ElementSlice, props: CanonProps): string {
  const runChildren = runPropChildren(props);
  const paraChildren = paraPropChildren(props);
  if (runChildren.length === 0 && paraChildren.length === 0) return xml;
  // §16 child order: w:name, w:basedOn?, w:pPr?, w:rPr?.
  const kids = style.selfClosed ? [] : childElements(xml, style.contentStart, style.contentEnd);
  const pPr = kids.find((k) => k.name === "w:pPr");
  const rPr = kids.find((k) => k.name === "w:rPr");
  const edits: SpliceEdit[] = [];

  if (paraChildren.length > 0) {
    if (pPr) {
      const inner = pPr.selfClosed ? "" : xml.slice(pPr.contentStart, pPr.contentEnd);
      edits.push({
        start: pPr.start,
        end: pPr.end,
        text: `<w:pPr>${mergeChildren(inner, paraChildren, PPR_ORDER)}</w:pPr>`,
      });
    } else {
      // Insert pPr after w:basedOn (or w:name), before w:rPr.
      const at = insertSlotForPPr(kids, style);
      edits.push({ start: at, end: at, text: `<w:pPr>${paraPropsInner(props)}</w:pPr>` });
    }
  }
  if (runChildren.length > 0) {
    if (rPr) {
      const inner = rPr.selfClosed ? "" : xml.slice(rPr.contentStart, rPr.contentEnd);
      edits.push({
        start: rPr.start,
        end: rPr.end,
        text: `<w:rPr>${mergeChildren(inner, runChildren, RPR_ORDER)}</w:rPr>`,
      });
    } else {
      const at = style.selfClosed ? -1 : insertSlotForRPr(kids, style);
      if (style.selfClosed) {
        const open = xml.slice(style.start, style.end - 2);
        return splice(
          xml,
          style.start,
          style.end,
          `${open}><w:rPr>${runPropsInner(props)}</w:rPr></w:style>`,
        );
      }
      edits.push({ start: at, end: at, text: `<w:rPr>${runPropsInner(props)}</w:rPr>` });
    }
  }
  return spliceAll(xml, edits);
}

/** Offset to insert a new w:pPr: after w:name/w:basedOn, before w:pPr/w:rPr. */
function insertSlotForPPr(kids: ElementSlice[], style: ElementSlice): number {
  const rPr = kids.find((k) => k.name === "w:rPr");
  if (rPr) return rPr.start;
  return style.contentEnd;
}

/** Offset to insert a new w:rPr: after everything (last §16 child). */
function insertSlotForRPr(_kids: ElementSlice[], style: ElementSlice): number {
  return style.contentEnd;
}

/** Direct formatting on an anchor/range: rPr into every run, pPr into each para. */
function formatDirect(doc: DocHandle, args: DocxFormatArgs, props: CanonProps): DocxFormatResult {
  const entries = bodyParagraphs(doc);
  let targets: AnchorEntry[];
  if (args.anchor !== undefined) {
    targets = [requireParagraph(doc, args.anchor)];
  } else {
    targets = resolveRange(entries, args.range as string);
  }
  const xml = doc.documentXml();
  const edits: SpliceEdit[] = [];
  for (const t of targets) {
    const paraEdit = paraPropsEdit(xml, t.block, props);
    if (paraEdit) edits.push(paraEdit);
    for (const run of childElements(xml, t.block.contentStart, t.block.contentEnd)) {
      if (run.name !== "w:r") continue;
      const runEdit = runPropsEdit(xml, run, props);
      if (runEdit) edits.push(runEdit);
    }
  }
  if (edits.length > 0) {
    doc.pkg.setPart(doc.documentPartName, spliceAll(xml, edits));
    doc.invalidate();
  }
  const fresh = bodyParagraphs(doc);
  const ordinals = [...new Set(targets.map((t) => t.ordinal))].sort((a, b) => a - b);
  const anchors = ordinals.map((o) => (fresh[o - 1] as AnchorEntry).anchor);
  return { affected: ordinals.length, anchors };
}

// ---------------------------------------------------------------------------
// Anchor / range validation (§6a order: parse → range → hash)
// ---------------------------------------------------------------------------

function requireParagraph(doc: DocHandle, anchor: string): AnchorEntry {
  const m = FULL_ANCHOR_RE.exec(anchor);
  if (!m) throw anchorInvalid(`Malformed anchor string: ${anchor}.`);
  const entries = bodyParagraphs(doc);
  const entry = entries[Number(m[1]) - 1];
  if (entry === undefined) throw anchorNotFound(anchor);
  if (entry.anchor !== anchor) throw anchorStale(anchor);
  return entry;
}

function resolveRange(entries: readonly AnchorEntry[], range: string): AnchorEntry[] {
  const m = RANGE_RE.exec(range);
  if (!m) throw anchorInvalid(`Malformed range string: ${range}.`);
  const start = Number(m[1]);
  const end = Number(m[3]);
  if (start > end) throw anchorInvalid(`Inverted range: ${range}.`);
  for (const [ordinal, hash] of [
    [start, m[2]],
    [end, m[4]],
  ] as [number, string | undefined][]) {
    const entry = entries[ordinal - 1];
    if (entry === undefined) throw anchorNotFound(`P${ordinal}`);
    if (hash !== undefined && entry.anchor !== `P${ordinal}#${hash}`) {
      throw anchorStale(`P${ordinal}#${hash}`);
    }
  }
  return entries.slice(start - 1, end) as AnchorEntry[];
}

void buildAnchorIndex;
void headingLevel;
