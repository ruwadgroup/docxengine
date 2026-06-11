/**
 * Lists (`docx_list`) and numbering per spec/algorithms.md §17.
 *
 * Creates `word/numbering.xml` on demand (content-type Override + document rel),
 * allocates `abstractNum`/`num` ids (max + 1), and wires each item's
 * `<w:numPr>` plus the ensured `ListParagraph` style. ol cascades
 * decimal/lowerLetter/lowerRoman; ul cascades •/◦/▪ as `numFmt="bullet"`.
 *
 * The Python twin (`_lists.py`) is the byte-parity reference.
 */
import { type AnchorEntry } from "./anchors.js";
import { ToolError } from "./errors.js";
import {
  LIST_PARAGRAPH_STYLE,
  addRelationship,
  anchorInvalid,
  anchorNotFound,
  anchorStale,
  ensureContentOverride,
  ensureStyle,
} from "./phase2common.js";
import type { DocHandle, Session } from "./session.js";
import {
  type ElementSlice,
  attrs,
  childElements,
  elementExtent,
  emitTextElement,
  escapeAttr,
  findElement,
  getAttr,
  nextTag,
  splice,
} from "./xmlscan.js";

type ResponseFormat = "concise" | "detailed";

const NUMBERING_PART = "word/numbering.xml";
const NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const NUMBERING_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

const FULL_ANCHOR_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;
const RANGE_RE = /^P([1-9][0-9]*)(?:#([0-9a-f]{4}))?\.\.P([1-9][0-9]*)(?:#([0-9a-f]{4}))?$/;

const OL_FORMATS = ["decimal", "lowerLetter", "lowerRoman"];
const UL_GLYPHS = ["•", "◦", "▪"];

// ---------------------------------------------------------------------------
// numbering.xml model
// ---------------------------------------------------------------------------

interface NumberingState {
  xml: string;
  /** Existing abstractNumIds (numeric). */
  abstractIds: number[];
  /** Existing numIds (numeric). */
  numIds: number[];
}

function emptyNumberingXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>'
  );
}

function readNumbering(doc: DocHandle): NumberingState {
  const xml = doc.pkg.has(NUMBERING_PART) ? doc.pkg.partText(NUMBERING_PART) : emptyNumberingXml();
  const abstractIds: number[] = [];
  const numIds: number[] = [];
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) break;
    if (t.name === "w:abstractNum" && t.kind !== "end") {
      const v = attrs(xml, t)["w:abstractNumId"];
      if (v !== undefined && /^[0-9]+$/.test(v)) abstractIds.push(Number(v));
    } else if (t.name === "w:num" && t.kind !== "end") {
      const v = attrs(xml, t)["w:numId"];
      if (v !== undefined && /^[0-9]+$/.test(v)) numIds.push(Number(v));
    }
    i = t.end;
  }
  return { xml, abstractIds, numIds };
}

/** Ensure numbering.xml exists with content-type Override + document rel. */
function ensureNumberingPart(doc: DocHandle): void {
  if (doc.pkg.has(NUMBERING_PART)) return;
  doc.pkg.setPart(NUMBERING_PART, emptyNumberingXml());
  ensureContentOverride(doc.pkg, NUMBERING_PART, NUMBERING_CT);
  addRelationship(doc.pkg, doc.documentPartName, NUMBERING_REL, "numbering.xml");
}

/** §17 abstractNum markup for `kind`, 9 cascading levels. */
function abstractNumXml(abstractNumId: number, kind: "ol" | "ul"): string {
  const levels: string[] = [];
  for (let ilvl = 0; ilvl < 9; ilvl++) {
    const left = 720 * (ilvl + 1);
    const ind = `<w:ind w:left="${left}" w:hanging="360"/>`;
    if (kind === "ol") {
      const fmt = OL_FORMATS[ilvl % OL_FORMATS.length] as string;
      const lvlText = `%${ilvl + 1}.`;
      levels.push(
        `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>` +
          `<w:lvlText w:val="${escapeAttr(lvlText)}"/><w:pPr>${ind}</w:pPr></w:lvl>`,
      );
    } else {
      const glyph = UL_GLYPHS[ilvl % UL_GLYPHS.length] as string;
      levels.push(
        `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
          `<w:lvlText w:val="${escapeAttr(glyph)}"/><w:pPr>${ind}</w:pPr></w:lvl>`,
      );
    }
  }
  return `<w:abstractNum w:abstractNumId="${abstractNumId}">${levels.join("")}</w:abstractNum>`;
}

/**
 * Allocate a fresh abstractNum + num for `kind`; splice both into numbering.xml.
 * Returns the new numId. abstractNum precedes num (Word requires abstractNum
 * definitions before the num that reference them — both spliced before the
 * close tag, abstractNum first).
 */
function allocateList(doc: DocHandle, kind: "ol" | "ul"): number {
  ensureNumberingPart(doc);
  const state = readNumbering(doc);
  const abstractNumId = (state.abstractIds.length ? Math.max(...state.abstractIds) : 0) + 1;
  const numId = (state.numIds.length ? Math.max(...state.numIds) : 0) + 1;
  const abs = abstractNumXml(abstractNumId, kind);
  const num = `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractNumId}"/></w:num>`;
  const xml = doc.pkg.partText(NUMBERING_PART);
  const close = xml.lastIndexOf("</w:numbering>");
  // abstractNum entries precede num entries; insert abstractNum first, then num.
  doc.pkg.setPart(NUMBERING_PART, splice(xml, close, close, abs + num));
  return numId;
}

/** Allocate a restart num referencing an existing abstractNum (§17). */
function allocateRestart(doc: DocHandle, abstractNumId: string, at: number): number {
  ensureNumberingPart(doc);
  const state = readNumbering(doc);
  const numId = (state.numIds.length ? Math.max(...state.numIds) : 0) + 1;
  const num =
    `<w:num w:numId="${numId}"><w:abstractNumId w:val="${escapeAttr(abstractNumId)}"/>` +
    `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${at}"/></w:lvlOverride></w:num>`;
  const xml = doc.pkg.partText(NUMBERING_PART);
  const close = xml.lastIndexOf("</w:numbering>");
  doc.pkg.setPart(NUMBERING_PART, splice(xml, close, close, num));
  return numId;
}

// ---------------------------------------------------------------------------
// Paragraph numPr / pStyle wiring
// ---------------------------------------------------------------------------

const NUMPR_RE = /<w:numPr\b/;

/** Build the `<w:numPr>` markup for a level + numId. */
function numPrXml(level: number, numId: number): string {
  return `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`;
}

/**
 * Set a paragraph's numPr (as the first w:pPr child, after any pStyle) and
 * ensure pStyle ListParagraph. Replaces any existing numPr. The pPr inner
 * markup is rebuilt deterministically from its existing children. Returns the
 * new document text.
 */
function setNumPr(xml: string, p: ElementSlice, level: number, numId: number): string {
  const numPr = numPrXml(level, numId);
  const pStyleEl = '<w:pStyle w:val="ListParagraph"/>';
  if (p.selfClosed) {
    const open = xml.slice(p.start, p.end - 2);
    return splice(xml, p.start, p.end, `${open}><w:pPr>${pStyleEl}${numPr}</w:pPr></w:p>`);
  }
  const pPr = childElements(xml, p.contentStart, p.contentEnd).find((k) => k.name === "w:pPr");
  if (!pPr) {
    return splice(xml, p.contentStart, p.contentStart, `<w:pPr>${pStyleEl}${numPr}</w:pPr>`);
  }
  if (pPr.selfClosed) {
    return splice(xml, pPr.start, pPr.end, `<w:pPr>${pStyleEl}${numPr}</w:pPr>`);
  }
  // Rebuild the pPr inner markup: keep every child except an existing numPr,
  // ensure a pStyle (ListParagraph) is present, and place numPr right after it.
  const kids = childElements(xml, pPr.contentStart, pPr.contentEnd);
  let inner = "";
  let hasPStyle = false;
  let placedNum = false;
  for (const k of kids) {
    if (k.name === "w:numPr") continue; // drop any existing numbering
    inner += xml.slice(k.start, k.end);
    if (k.name === "w:pStyle") {
      hasPStyle = true;
      inner += numPr; // numPr immediately follows pStyle
      placedNum = true;
    }
  }
  if (!hasPStyle) inner = pStyleEl + numPr + inner;
  else if (!placedNum) inner += numPr;
  return splice(xml, pPr.start, pPr.end, `<w:pPr>${inner}</w:pPr>`);
}

/** Remove a paragraph's numPr (convert to:paragraphs). */
function removeNumPr(xml: string, p: ElementSlice): string {
  if (p.selfClosed) return xml;
  const pPr = childElements(xml, p.contentStart, p.contentEnd).find((k) => k.name === "w:pPr");
  if (!pPr || pPr.selfClosed) return xml;
  const numPr = childElements(xml, pPr.contentStart, pPr.contentEnd).find(
    (k) => k.name === "w:numPr",
  );
  if (!numPr) return xml;
  return splice(xml, numPr.start, numPr.end, "");
}

/** Rewrite a paragraph's w:ilvl (set_level); no-op when no numPr. */
function setIlvl(xml: string, p: ElementSlice, level: number): string {
  if (p.selfClosed) return xml;
  const pPr = childElements(xml, p.contentStart, p.contentEnd).find((k) => k.name === "w:pPr");
  if (!pPr || pPr.selfClosed) return xml;
  const numPr = findElement(xml, "w:numPr", pPr.contentStart, pPr.contentEnd);
  if (!numPr || numPr.selfClosed) return xml;
  const ilvl = findElement(xml, "w:ilvl", numPr.contentStart, numPr.contentEnd);
  if (ilvl) return splice(xml, ilvl.start, ilvl.end, `<w:ilvl w:val="${level}"/>`);
  return splice(xml, numPr.contentStart, numPr.contentStart, `<w:ilvl w:val="${level}"/>`);
}

/** Read a paragraph's numId (or null). */
function paragraphNumId(xml: string, p: ElementSlice): string | null {
  if (p.selfClosed) return null;
  const pPr = childElements(xml, p.contentStart, p.contentEnd).find((k) => k.name === "w:pPr");
  if (!pPr || pPr.selfClosed) return null;
  const numPr = findElement(xml, "w:numPr", pPr.contentStart, pPr.contentEnd);
  if (!numPr || numPr.selfClosed) return null;
  const numIdEl = findElement(xml, "w:numId", numPr.contentStart, numPr.contentEnd);
  return numIdEl ? (getAttr(xml, slcTag(numIdEl), "w:val") ?? null) : null;
}

function slcTag(el: ElementSlice) {
  return {
    kind: el.selfClosed ? ("empty" as const) : ("start" as const),
    name: el.name,
    start: el.start,
    end: el.startTagEnd,
    nameEnd: el.nameEnd,
  };
}

/** numId → abstractNumId from numbering.xml. */
function numToAbstract(doc: DocHandle, numId: string): string | null {
  if (!doc.pkg.has(NUMBERING_PART)) return null;
  const xml = doc.pkg.partText(NUMBERING_PART);
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return null;
    if (t.name === "w:num" && t.kind !== "end") {
      const el = elementExtent(xml, t);
      if (getAttr(xml, t, "w:numId") === numId) {
        const absEl = findElement(xml, "w:abstractNumId", el.contentStart, el.contentEnd);
        return absEl ? (getAttr(xml, slcTag(absEl), "w:val") ?? null) : null;
      }
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

// ---------------------------------------------------------------------------
// docx_list
// ---------------------------------------------------------------------------

export interface DocxListItem {
  text: string;
  level?: number | undefined;
}

export interface DocxListArgs {
  doc_id: string;
  op: "create" | "restart" | "set_level" | "convert";
  anchor?: string | undefined;
  range?: string | undefined;
  after?: string | undefined;
  kind?: "ol" | "ul" | undefined;
  items?: DocxListItem[] | undefined;
  at?: number | undefined;
  level?: number | undefined;
  to?: "ol" | "ul" | "paragraphs" | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxListResult {
  new_anchors?: string[];
  n_affected?: number;
  note?: string;
}

function bodyParagraphs(doc: DocHandle): AnchorEntry[] {
  return doc.anchorIndex().filter((e) => e.kind === "p");
}

function requireParagraph(doc: DocHandle, anchor: string): AnchorEntry {
  const m = FULL_ANCHOR_RE.exec(anchor);
  if (!m) throw anchorInvalid(`Malformed anchor string: ${anchor}.`);
  const entry = bodyParagraphs(doc)[Number(m[1]) - 1];
  if (entry === undefined) throw anchorNotFound(anchor);
  if (entry.anchor !== anchor) throw anchorStale(anchor);
  return entry;
}

function resolveTargets(doc: DocHandle, args: DocxListArgs): AnchorEntry[] {
  if (args.anchor !== undefined) return [requireParagraph(doc, args.anchor)];
  if (args.range !== undefined) {
    const m = RANGE_RE.exec(args.range);
    if (!m) throw anchorInvalid(`Malformed range string: ${args.range}.`);
    const start = Number(m[1]);
    const end = Number(m[3]);
    if (start > end) throw anchorInvalid(`Inverted range: ${args.range}.`);
    const entries = bodyParagraphs(doc);
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
  throw anchorInvalid("Provide anchor or range.");
}

export function docxList(session: Session, args: DocxListArgs): DocxListResult {
  const doc = session.get(args.doc_id);
  switch (args.op) {
    case "create":
      return listCreate(doc, args);
    case "convert":
      return listConvert(doc, args);
    case "set_level":
      return listSetLevel(doc, args);
    case "restart":
      return listRestart(doc, args);
    default:
      throw new ToolError("invalid_args", `docx_list: unknown op ${String(args.op)}.`, [
        "Use create, restart, set_level or convert.",
      ]);
  }
}

/** create: insert N list-item paragraphs after the anchor; new abstractNum/num. */
function listCreate(doc: DocHandle, args: DocxListArgs): DocxListResult {
  if (args.after === undefined) throw anchorInvalid("op 'create' requires after.");
  const entry = requireParagraph(doc, args.after);
  const items = args.items ?? [];
  const kind = args.kind ?? "ul";
  ensureStyle(doc.pkg, "ListParagraph", LIST_PARAGRAPH_STYLE);
  if (items.length === 0) return { new_anchors: [], n_affected: 0, note: "No items." };
  const numId = allocateList(doc, kind);

  const xml = doc.documentXml();
  const pieces: string[] = [];
  for (const item of items) {
    const level = Math.max(0, Math.trunc(item.level ?? 0));
    const pPr = `<w:pPr><w:pStyle w:val="ListParagraph"/>${numPrXml(level, numId)}</w:pPr>`;
    const run = item.text !== "" ? `<w:r>${emitTextElement("w:t", item.text)}</w:r>` : "";
    pieces.push(`<w:p>${pPr}${run}</w:p>`);
  }
  doc.pkg.setPart(
    doc.documentPartName,
    splice(xml, entry.block.end, entry.block.end, pieces.join("")),
  );
  doc.invalidate();
  const fresh = bodyParagraphs(doc);
  const base = entry.ordinal + 1;
  const anchors = items.map((_, i) => (fresh[base - 1 + i] as AnchorEntry).anchor);
  return {
    new_anchors: anchors,
    n_affected: items.length,
    note: `Created ${kind} list with ${items.length} item${items.length === 1 ? "" : "s"}.`,
  };
}

/** convert: ol/ul → set numPr (reuse/create one list for the run); paragraphs → remove. */
function listConvert(doc: DocHandle, args: DocxListArgs): DocxListResult {
  const to = args.to;
  if (to === undefined) throw new ToolError("invalid_args", "op 'convert' requires to.", []);
  const targets = resolveTargets(doc, args);
  let xml = doc.documentXml();

  if (to === "paragraphs") {
    // Apply in reverse document order so offsets stay valid.
    for (const t of [...targets].reverse()) {
      const block = currentBlock(doc, xml, t.ordinal);
      xml = removeNumPr(xml, block);
    }
    doc.pkg.setPart(doc.documentPartName, xml);
    doc.invalidate();
    const fresh = bodyParagraphs(doc);
    return {
      new_anchors: targets.map((t) => (fresh[t.ordinal - 1] as AnchorEntry).anchor),
      n_affected: targets.length,
      note: `Converted ${targets.length} paragraph${targets.length === 1 ? "" : "s"} to plain.`,
    };
  }

  ensureStyle(doc.pkg, "ListParagraph", LIST_PARAGRAPH_STYLE);
  const numId = allocateList(doc, to);
  xml = doc.documentXml();
  for (const t of [...targets].reverse()) {
    const block = currentBlock(doc, xml, t.ordinal);
    xml = setNumPr(xml, block, 0, numId);
  }
  doc.pkg.setPart(doc.documentPartName, xml);
  doc.invalidate();
  const fresh = bodyParagraphs(doc);
  return {
    new_anchors: targets.map((t) => (fresh[t.ordinal - 1] as AnchorEntry).anchor),
    n_affected: targets.length,
    note: `Converted ${targets.length} paragraph${targets.length === 1 ? "" : "s"} to ${to}.`,
  };
}

/** set_level: rewrite w:ilvl on the target(s). */
function listSetLevel(doc: DocHandle, args: DocxListArgs): DocxListResult {
  if (args.level === undefined)
    throw new ToolError("invalid_args", "op 'set_level' requires level.", []);
  const level = Math.max(0, Math.trunc(args.level));
  const targets = resolveTargets(doc, args);
  let xml = doc.documentXml();
  for (const t of [...targets].reverse()) {
    const block = currentBlock(doc, xml, t.ordinal);
    xml = setIlvl(xml, block, level);
  }
  doc.pkg.setPart(doc.documentPartName, xml);
  doc.invalidate();
  const fresh = bodyParagraphs(doc);
  return {
    new_anchors: targets.map((t) => (fresh[t.ordinal - 1] as AnchorEntry).anchor),
    n_affected: targets.length,
    note: `Set level ${level} on ${targets.length} item${targets.length === 1 ? "" : "s"}.`,
  };
}

/** restart: new num referencing the same abstractNum; repoint the paragraph. */
function listRestart(doc: DocHandle, args: DocxListArgs): DocxListResult {
  if (args.anchor === undefined) throw anchorInvalid("op 'restart' requires anchor.");
  const entry = requireParagraph(doc, args.anchor);
  const at = args.at ?? 1;
  const xml0 = doc.documentXml();
  const numId = paragraphNumId(xml0, entry.block);
  if (numId === null) {
    throw anchorInvalid("Target paragraph is not a numbered list item.");
  }
  const abstractId = numToAbstract(doc, numId);
  if (abstractId === null) {
    throw anchorInvalid("Target list has no abstract numbering to restart.");
  }
  const newNumId = allocateRestart(doc, abstractId, at);
  // Repoint the paragraph's w:numId to the new num.
  const xml = doc.documentXml();
  const block = currentBlock(doc, xml, entry.ordinal);
  const pPr = childElements(xml, block.contentStart, block.contentEnd).find(
    (k) => k.name === "w:pPr",
  );
  const numPr = pPr ? findElement(xml, "w:numPr", pPr.contentStart, pPr.contentEnd) : null;
  const numIdEl = numPr ? findElement(xml, "w:numId", numPr.contentStart, numPr.contentEnd) : null;
  if (numIdEl) {
    doc.pkg.setPart(
      doc.documentPartName,
      splice(xml, numIdEl.start, numIdEl.end, `<w:numId w:val="${newNumId}"/>`),
    );
    doc.invalidate();
  }
  const fresh = bodyParagraphs(doc);
  return {
    new_anchors: [(fresh[entry.ordinal - 1] as AnchorEntry).anchor],
    n_affected: 1,
    note: `Restarted numbering at ${at}.`,
  };
}

/** The body paragraph block at `ordinal` in the *given* (possibly edited) xml. */
function currentBlock(doc: DocHandle, xml: string, ordinal: number): ElementSlice {
  let i = 0;
  let n = 0;
  // Walk body-level w:p in document order.
  const body = findElement(xml, "w:body");
  if (!body) throw new Error("no w:body");
  i = body.contentStart;
  for (;;) {
    const t = nextTag(xml, i, body.contentEnd);
    if (!t) break;
    if (t.kind === "end") {
      i = t.end;
      continue;
    }
    const el = elementExtent(xml, t, body.contentEnd);
    if (el.name === "w:p") {
      n++;
      if (n === ordinal) return el;
    }
    i = el.end;
  }
  void doc;
  throw anchorNotFound(`P${ordinal}`);
}

void NUMPR_RE;
