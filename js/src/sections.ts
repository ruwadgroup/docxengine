/**
 * Sections (`docx_section`) per spec/algorithms.md §15.
 *
 * `w:sectPr` carries `w:pgSz`/`w:pgMar` (twips) plus optional `w:cols` and
 * `w:headerReference`/`w:footerReference`. Ops: list, set_geometry (page size,
 * orientation, margins, columns), set_header/set_footer (create a header/footer
 * part, register its content-type + rel, splice the reference), insert_break
 * (clone the body sectPr onto a paragraph's pPr with a `w:type`).
 *
 * Sections are `S{n}` over `w:sectPr` in document order; the trailing body
 * sectPr is the last S. The Python twin (`_sections.py`) is the byte-parity
 * reference.
 */
import { ToolError } from "./errors.js";
import {
  addRelationship,
  anchorInvalid,
  anchorNotFound,
  anchorStale,
  ensureContentOverride,
} from "./phase2common.js";
import type { DocHandle, Session } from "./session.js";
import {
  type ElementSlice,
  attrs,
  childElements,
  elementExtent,
  emitTextElement,
  findElement,
  nextTag,
  splice,
} from "./xmlscan.js";

type ResponseFormat = "concise" | "detailed";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HEADER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const HEADER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

const FULL_ANCHOR_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;
const S_ANCHOR_RE = /^S([1-9][0-9]*)$/;

const TWIPS_PER_CM = 567;

/** §15 page-size presets (portrait twips). */
const PAGE_SIZES: Readonly<Record<string, { w: number; h: number }>> = {
  A3: { w: 16838, h: 23811 },
  A4: { w: 11906, h: 16838 },
  A5: { w: 8391, h: 11906 },
  Letter: { w: 12240, h: 15840 },
  Legal: { w: 12240, h: 20160 },
  Tabloid: { w: 15840, h: 24480 },
};

// ---------------------------------------------------------------------------
// Section model
// ---------------------------------------------------------------------------

interface SectionRecord {
  /** 1-based S ordinal. */
  ordinal: number;
  /** The w:sectPr slice in the document text. */
  sectPr: ElementSlice;
  /** True for the trailing body sectPr (the only one not in a w:pPr). */
  isBody: boolean;
}

/** Every w:sectPr in document order — paragraph-level breaks then the body one. */
function sections(xml: string): SectionRecord[] {
  const out: SectionRecord[] = [];
  let ordinal = 0;
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) break;
    if (t.name === "w:sectPr" && t.kind !== "end") {
      const el = elementExtent(xml, t);
      ordinal++;
      out.push({ ordinal, sectPr: el, isBody: false });
      i = el.end;
      continue;
    }
    i = t.end;
  }
  // The last sectPr that is a direct child of w:body is the body sectPr.
  const body = findElement(xml, "w:body");
  if (body) {
    const kids = childElements(xml, body.contentStart, body.contentEnd);
    const bodySect = kids.find((k) => k.name === "w:sectPr");
    if (bodySect) {
      const rec = out.find((r) => r.sectPr.start === bodySect.start);
      if (rec) rec.isBody = true;
    }
  }
  return out;
}

function requireSection(xml: string, sectionId: string): SectionRecord {
  const m = S_ANCHOR_RE.exec(sectionId);
  if (!m) throw anchorInvalid(`Malformed section id: ${sectionId}.`);
  const ord = Number(m[1]);
  const rec = sections(xml).find((r) => r.ordinal === ord);
  if (!rec) throw anchorNotFound(sectionId);
  return rec;
}

// ---------------------------------------------------------------------------
// Geometry parsing
// ---------------------------------------------------------------------------

function childOf(xml: string, sect: ElementSlice, name: string): ElementSlice | null {
  if (sect.selfClosed) return null;
  return (
    childElements(xml, sect.contentStart, sect.contentEnd).find((k) => k.name === name) ?? null
  );
}

function pgSzOf(xml: string, sect: ElementSlice): { w: number; h: number; orient: string } {
  const el = childOf(xml, sect, "w:pgSz");
  if (!el) return { w: 12240, h: 15840, orient: "portrait" };
  const a = attrs(xml, tagOf(el));
  return {
    w: Number(a["w:w"] ?? "12240"),
    h: Number(a["w:h"] ?? "15840"),
    orient: a["w:orient"] ?? "portrait",
  };
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

function pageSizeName(w: number, h: number): string {
  for (const [name, dims] of Object.entries(PAGE_SIZES)) {
    if ((dims.w === w && dims.h === h) || (dims.w === h && dims.h === w)) return name;
  }
  return "custom";
}

// ---------------------------------------------------------------------------
// docx_section
// ---------------------------------------------------------------------------

export interface DocxSectionMargins {
  top?: number | undefined;
  bottom?: number | undefined;
  left?: number | undefined;
  right?: number | undefined;
}

export interface DocxSectionInfo {
  id: string;
  break_type: string;
  page_size: string;
  orientation: string;
  columns: number;
  has_header: boolean;
  has_footer: boolean;
}

export interface DocxSectionArgs {
  doc_id: string;
  op: "list" | "set_geometry" | "set_header" | "set_footer" | "insert_break";
  section?: string | undefined;
  page_size?: string | undefined;
  orientation?: string | undefined;
  margins?: DocxSectionMargins | undefined;
  columns?: number | undefined;
  content?: string | undefined;
  variant?: "default" | "first" | "even" | undefined;
  after?: string | undefined;
  break_type?: "nextPage" | "continuous" | "evenPage" | "oddPage" | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxSectionResult {
  sections?: DocxSectionInfo[];
  section?: string;
  new_anchor?: string;
  note?: string;
}

export function docxSection(session: Session, args: DocxSectionArgs): DocxSectionResult {
  const doc = session.get(args.doc_id);
  switch (args.op) {
    case "list":
      return sectionList(doc);
    case "set_geometry":
      return sectionSetGeometry(doc, args);
    case "set_header":
      return sectionSetHeaderFooter(doc, args, "header");
    case "set_footer":
      return sectionSetHeaderFooter(doc, args, "footer");
    case "insert_break":
      return sectionInsertBreak(doc, args);
    default:
      throw new ToolError("invalid_args", `docx_section: unknown op ${String(args.op)}.`, []);
  }
}

// --- list --------------------------------------------------------------------

function sectionList(doc: DocHandle): DocxSectionResult {
  const xml = doc.documentXml();
  const recs = sections(xml);
  const out: DocxSectionInfo[] = recs.map((rec) => {
    const sz = pgSzOf(xml, rec.sectPr);
    const colsEl = childOf(xml, rec.sectPr, "w:cols");
    const columns = colsEl ? Number(attrs(xml, tagOf(colsEl))["w:num"] ?? "1") : 1;
    const typeEl = childOf(xml, rec.sectPr, "w:type");
    const breakType = typeEl ? (attrs(xml, tagOf(typeEl))["w:val"] ?? "nextPage") : "nextPage";
    let hasHeader = false;
    let hasFooter = false;
    if (!rec.sectPr.selfClosed) {
      for (const k of childElements(xml, rec.sectPr.contentStart, rec.sectPr.contentEnd)) {
        if (k.name === "w:headerReference") hasHeader = true;
        if (k.name === "w:footerReference") hasFooter = true;
      }
    }
    return {
      id: `S${rec.ordinal}`,
      break_type: breakType,
      page_size: pageSizeName(sz.w, sz.h),
      orientation: sz.orient === "landscape" ? "landscape" : "portrait",
      columns,
      has_header: hasHeader,
      has_footer: hasFooter,
    };
  });
  return { sections: out };
}

// --- set_geometry ------------------------------------------------------------

function cmToTwips(cm: number): number {
  return Math.round(cm * TWIPS_PER_CM);
}

function sectionSetGeometry(doc: DocHandle, args: DocxSectionArgs): DocxSectionResult {
  const sectionId = args.section ?? "S1";
  const xml = doc.documentXml();
  const rec = requireSection(xml, sectionId);

  // Resolve target page size/orientation, keeping unspecified attributes.
  const cur = pgSzOf(xml, rec.sectPr);
  let w = cur.w;
  let h = cur.h;
  let orient = cur.orient;
  if (args.page_size !== undefined) {
    const preset = PAGE_SIZES[args.page_size];
    if (!preset) throw anchorInvalid(`Unknown page size: ${args.page_size}.`);
    w = preset.w;
    h = preset.h;
    // Re-apply the current/requested orientation on the preset.
  }
  if (args.orientation !== undefined) orient = args.orientation;
  // Apply orientation: landscape → w/h swapped from portrait preset values.
  let outW = w;
  let outH = h;
  if (orient === "landscape") {
    // Ensure landscape means the wider dimension is width.
    const portraitW = Math.min(w, h);
    const portraitH = Math.max(w, h);
    outW = portraitH;
    outH = portraitW;
  } else {
    const portraitW = Math.min(w, h);
    const portraitH = Math.max(w, h);
    outW = portraitW;
    outH = portraitH;
  }
  const orientAttr = orient === "landscape" ? ' w:orient="landscape"' : "";
  const pgSzXml = `<w:pgSz w:w="${outW}" w:h="${outH}"${orientAttr}/>`;

  // Margins (cm → twips); keep defaults for unspecified.
  let pgMarXml: string | null = null;
  if (args.margins !== undefined) {
    const m = args.margins;
    const existing = childOf(xml, rec.sectPr, "w:pgMar");
    const ea = existing ? attrs(xml, tagOf(existing)) : {};
    const top = m.top !== undefined ? cmToTwips(m.top) : Number(ea["w:top"] ?? "1440");
    const right = m.right !== undefined ? cmToTwips(m.right) : Number(ea["w:right"] ?? "1440");
    const bottom = m.bottom !== undefined ? cmToTwips(m.bottom) : Number(ea["w:bottom"] ?? "1440");
    const left = m.left !== undefined ? cmToTwips(m.left) : Number(ea["w:left"] ?? "1440");
    const header = Number(ea["w:header"] ?? "708");
    const footer = Number(ea["w:footer"] ?? "708");
    const gutter = Number(ea["w:gutter"] ?? "0");
    pgMarXml =
      `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}"` +
      ` w:header="${header}" w:footer="${footer}" w:gutter="${gutter}"/>`;
  }

  let colsXml: string | null = null;
  if (args.columns !== undefined && args.columns > 1) {
    colsXml = `<w:cols w:num="${Math.trunc(args.columns)}" w:space="708"/>`;
  }

  // Splice each child in place (replace existing or insert), keeping the
  // sectPr's other children verbatim. We rebuild the sectPr content.
  const next = rewriteSectPr(xml, rec.sectPr, { pgSzXml, pgMarXml, colsXml });
  doc.pkg.setPart(doc.documentPartName, next);
  doc.invalidate();
  return { section: sectionId, note: `Updated geometry of ${sectionId}.` };
}

/**
 * Replace/insert pgSz, pgMar, cols inside a w:sectPr, splicing in place so
 * untouched children survive. pgSz/pgMar sit after any header/footer refs;
 * cols after pgMar.
 */
function rewriteSectPr(
  xml: string,
  sect: ElementSlice,
  parts: { pgSzXml: string; pgMarXml: string | null; colsXml: string | null },
): string {
  // Expand a self-closed sectPr into an open/close pair first.
  let work = xml;
  let sectEl = sect;
  if (sectEl.selfClosed) {
    work = splice(work, sectEl.start, sectEl.end, "<w:sectPr></w:sectPr>");
    const reopened = findElement(work, "w:sectPr", sectEl.start);
    if (!reopened) throw new Error("sectPr re-expansion failed");
    sectEl = reopened;
  }

  const apply = (cur: string, el: ElementSlice, name: string, value: string | null): string => {
    if (value === null) return cur;
    const existing = childOf(cur, el, name);
    if (existing) {
      return splice(cur, existing.start, existing.end, value);
    }
    // Insert at the right position: references precede pgSz; pgSz precedes
    // pgMar; pgMar precedes cols. We splice before the first child that should
    // follow, else at content end.
    const order = ["w:headerReference", "w:footerReference", "w:pgSz", "w:pgMar", "w:cols"];
    const myIdx = order.indexOf(name);
    const kids = el.selfClosed ? [] : childElements(cur, el.contentStart, el.contentEnd);
    let insertAt = el.contentEnd;
    for (const k of kids) {
      const kIdx = order.indexOf(k.name);
      if (kIdx > myIdx || kIdx === -1) {
        insertAt = k.start;
        break;
      }
    }
    return splice(cur, insertAt, insertAt, value);
  };

  // Each apply may shift offsets, so re-resolve the sectPr each time.
  work = apply(work, sectEl, "w:pgSz", parts.pgSzXml);
  sectEl = findElement(work, "w:sectPr", sect.start)!;
  if (parts.pgMarXml !== null) {
    work = apply(work, sectEl, "w:pgMar", parts.pgMarXml);
    sectEl = findElement(work, "w:sectPr", sect.start)!;
  }
  if (parts.colsXml !== null) {
    work = apply(work, sectEl, "w:cols", parts.colsXml);
  }
  return work;
}

// --- set_header / set_footer -------------------------------------------------

function variantType(variant: string | undefined): string {
  if (variant === "first") return "first";
  if (variant === "even") return "even";
  return "default";
}

/** Next free header{N}/footer{N} index across both kinds (max existing + 1). */
function nextHeaderFooterIndex(doc: DocHandle): number {
  let max = 0;
  for (const name of doc.pkg.entryNames()) {
    const m = /^word\/(?:header|footer)([0-9]+)\.xml$/.exec(name);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max + 1;
}

function sectionSetHeaderFooter(
  doc: DocHandle,
  args: DocxSectionArgs,
  kind: "header" | "footer",
): DocxSectionResult {
  const sectionId = args.section ?? "S1";
  const xml = doc.documentXml();
  const rec = requireSection(xml, sectionId);
  const variant = variantType(args.variant);

  const idx = nextHeaderFooterIndex(doc);
  const partName = `word/${kind}${idx}.xml`;
  const rootTag = kind === "header" ? "w:hdr" : "w:ftr";
  const body = headerFooterBody(args.content ?? "");
  const partXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<${rootTag} xmlns:w="${W_NS}">${body}</${rootTag}>`;
  doc.pkg.setPart(partName, partXml);
  ensureContentOverride(doc.pkg, partName, kind === "header" ? HEADER_CT : FOOTER_CT);
  const relType = kind === "header" ? HEADER_REL : FOOTER_REL;
  const rId = addRelationship(doc.pkg, doc.documentPartName, relType, `${kind}${idx}.xml`);

  // Splice the reference into the sectPr (references precede pgSz).
  const refTag =
    kind === "header"
      ? `<w:headerReference w:type="${variant}" r:id="${rId}"/>`
      : `<w:footerReference w:type="${variant}" r:id="${rId}"/>`;
  const next = insertReference(doc.documentXml(), rec.sectPr.start, refTag);
  doc.pkg.setPart(doc.documentPartName, next);
  doc.invalidate();

  return {
    section: sectionId,
    note: `Set ${variant} ${kind} on ${sectionId} (word/${kind}${idx}.xml).`,
  };
}

/** Splice a header/footer reference as the first child of the sectPr at `sectStart`. */
function insertReference(xml: string, sectStart: number, refTag: string): string {
  const sect = findElement(xml, "w:sectPr", sectStart);
  if (!sect) throw new Error("sectPr not found for reference insertion");
  if (sect.selfClosed) {
    const expanded = splice(xml, sect.start, sect.end, `<w:sectPr>${refTag}</w:sectPr>`);
    return expanded;
  }
  // First child position is contentStart; references precede everything.
  return splice(xml, sect.contentStart, sect.contentStart, refTag);
}

/** §22 markdown→paragraph (plain paragraphs only in headers/footers). */
export function headerFooterBody(content: string): string {
  const paras: string[] = [];
  for (const raw of content.split("\n")) {
    const line = raw.endsWith("\r") ? raw.slice(0, -1) : raw;
    if (line.trim() === "") continue;
    paras.push(`<w:p><w:r>${emitTextElement("w:t", line)}</w:r></w:p>`);
  }
  if (paras.length === 0) paras.push("<w:p/>");
  return paras.join("");
}

// --- insert_break ------------------------------------------------------------

function paragraphEntries(doc: DocHandle) {
  return doc.anchorIndex().filter((e) => e.kind === "p");
}

function requireParagraph(doc: DocHandle, anchor: string) {
  const m = FULL_ANCHOR_RE.exec(anchor);
  if (!m) throw anchorInvalid(`Malformed anchor string: ${anchor}.`);
  const entry = paragraphEntries(doc)[Number(m[1]) - 1];
  if (entry === undefined) throw anchorNotFound(anchor);
  if (entry.anchor !== anchor) throw anchorStale(anchor);
  return entry;
}

function sectionInsertBreak(doc: DocHandle, args: DocxSectionArgs): DocxSectionResult {
  if (args.after == null) throw anchorInvalid("op 'insert_break' requires after.");
  const entry = requireParagraph(doc, args.after); // hash FIRST
  const breakType = args.break_type ?? "nextPage";

  const xml = doc.documentXml();
  // Clone the body sectPr's geometry children into a paragraph-level sectPr.
  const recs = sections(xml);
  const bodySect = recs.find((r) => r.isBody) ?? recs[recs.length - 1];
  const inner =
    bodySect && !bodySect.sectPr.selfClosed
      ? xml.slice(bodySect.sectPr.contentStart, bodySect.sectPr.contentEnd)
      : "";
  // Strip any existing w:type, then prepend the requested break type.
  const innerNoType = stripChild(inner, "w:type");
  const sectInner = `<w:type w:val="${breakType}"/>${innerNoType}`;
  const sectPrXml = `<w:sectPr>${sectInner}</w:sectPr>`;

  // Splice the sectPr into the target paragraph's pPr (creating pPr if absent),
  // as its first child (a section break's sectPr precedes other pPr content).
  const p = entry.block;
  let next: string;
  const kids = p.selfClosed ? [] : childElements(xml, p.contentStart, p.contentEnd);
  const pPr = kids.find((k) => k.name === "w:pPr");
  if (pPr) {
    next = splice(xml, pPr.contentStart, pPr.contentStart, sectPrXml);
  } else if (p.selfClosed) {
    next = splice(xml, p.start, p.end, `<w:p><w:pPr>${sectPrXml}</w:pPr></w:p>`);
  } else {
    next = splice(xml, p.contentStart, p.contentStart, `<w:pPr>${sectPrXml}</w:pPr>`);
  }
  doc.pkg.setPart(doc.documentPartName, next);
  doc.invalidate();

  // The fresh anchor of the paragraph (its text is unchanged → same anchor).
  const fresh = paragraphEntries(doc)[entry.ordinal - 1];
  return {
    new_anchor: fresh ? fresh.anchor : args.after,
    note: `Inserted a ${breakType} section break after ${args.after}.`,
  };
}

/** Remove a single named child element from a sectPr inner fragment. */
function stripChild(inner: string, name: string): string {
  let i = 0;
  for (;;) {
    const t = nextTag(inner, i);
    if (!t) return inner;
    if (t.name === name && t.kind !== "end") {
      const el = elementExtent(inner, t);
      return inner.slice(0, el.start) + inner.slice(el.end);
    }
    i = t.end;
  }
}
