/**
 * Fields (`docx_field`) per spec/algorithms.md §20.
 *
 * `insert_toc` adds, after the anchor, a paragraph holding the run-triple TOC
 * field (`TOC \o "1-{levels}" \h \z \u`). `insert_page_number` ensures the
 * section's footer exists (§15 machinery) and appends a `PAGE` field run-triple
 * to it. `update` sets `<w:updateFields w:val="true"/>` in word/settings.xml
 * (created on demand). Values materialize only at render — results never report
 * computed page/TOC numbers. The Python twin (`_fields.py`) is byte-parity.
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
import { type ElementSlice, attrs, childElements, findElement, splice } from "./xmlscan.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const FOOTER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const FOOTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const HEADER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const HEADER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const SETTINGS_PART = "word/settings.xml";
const SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const SETTINGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";

const FULL_ANCHOR_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;

// ---------------------------------------------------------------------------
// Field run-triples (§20)
// ---------------------------------------------------------------------------

/** A field run-triple: begin / instrText / separate / placeholder / end. */
function fieldRunTriple(instr: string, placeholder: string): string {
  const space = /^\s|\s$/.test(instr) ? ' xml:space="preserve"' : "";
  return (
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText${space}>${instr}</w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    `<w:r><w:t>${placeholder}</w:t></w:r>` +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  );
}

// ---------------------------------------------------------------------------
// docx_field
// ---------------------------------------------------------------------------

export interface DocxFieldArgs {
  doc_id: string;
  op: "insert_toc" | "insert_page_number" | "update";
  after?: string | undefined;
  levels?: number | undefined;
  scope?: "header" | "footer" | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
}

export interface DocxFieldResult {
  new_anchor?: string;
  updated?: number;
  note?: string;
}

export function docxField(session: Session, args: DocxFieldArgs): DocxFieldResult {
  const doc = session.get(args.doc_id);
  switch (args.op) {
    case "insert_toc":
      return fieldInsertToc(doc, args);
    case "insert_page_number":
      return fieldInsertPageNumber(doc, args);
    case "update":
      return fieldUpdate(doc);
    default:
      throw new ToolError("invalid_args", `docx_field: unknown op ${String(args.op)}.`, []);
  }
}

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

// --- insert_toc --------------------------------------------------------------

function fieldInsertToc(doc: DocHandle, args: DocxFieldArgs): DocxFieldResult {
  if (args.after == null) throw anchorInvalid("op 'insert_toc' requires after.");
  const entry = requireParagraph(doc, args.after); // hash FIRST
  const levels = Math.max(1, Math.trunc(args.levels ?? 3));
  const instr = ` TOC \\o "1-${levels}" \\h \\z \\u `;
  const para = `<w:p>${fieldRunTriple(instr, "Right-click to update field.")}</w:p>`;

  const xml = doc.documentXml();
  const position = entry.block.end;
  doc.pkg.setPart(doc.documentPartName, splice(xml, position, position, para));
  doc.invalidate();

  const fresh = paragraphEntries(doc)[entry.ordinal]; // the inserted paragraph
  return {
    new_anchor: fresh ? fresh.anchor : args.after,
    note: `Inserted a TOC field (levels 1-${levels}) after ${args.after}.`,
  };
}

// --- insert_page_number ------------------------------------------------------

function fieldInsertPageNumber(doc: DocHandle, args: DocxFieldArgs): DocxFieldResult {
  const scope = args.scope ?? "footer";
  const partName = ensureScopePart(doc, scope);
  // Append a PAGE field paragraph to the header/footer part.
  const xml = doc.pkg.partText(partName);
  const rootTag = scope === "header" ? "w:hdr" : "w:ftr";
  const para = `<w:p>${fieldRunTriple(" PAGE ", "1")}</w:p>`;
  const close = xml.lastIndexOf(`</${rootTag}>`);
  let next: string;
  if (close >= 0) {
    next = splice(xml, close, close, para);
  } else {
    const selfClose = xml.lastIndexOf("/>");
    next = splice(xml, selfClose, selfClose + 2, `>${para}</${rootTag}>`);
  }
  doc.pkg.setPart(partName, next);
  return { note: `Inserted a PAGE field in the ${scope} (${partName}).` };
}

/**
 * Ensure the body section has a default header/footer of the given scope,
 * creating the part + content-type + rel + sectPr reference when absent.
 * Returns the part name to append the field paragraph into.
 */
function ensureScopePart(doc: DocHandle, scope: "header" | "footer"): string {
  const refName = scope === "header" ? "w:headerReference" : "w:footerReference";
  const xml = doc.documentXml();
  const bodySect = bodySectPr(xml);

  // If a default-variant reference already exists, reuse its part.
  if (bodySect && !bodySect.selfClosed) {
    for (const k of childElements(xml, bodySect.contentStart, bodySect.contentEnd)) {
      if (k.name === refName) {
        const a = attrs(xml, tagOf(k));
        if ((a["w:type"] ?? "default") === "default") {
          const rId = a["r:id"];
          if (rId) {
            const rel = doc.pkg.rels(doc.documentPartName).find((r) => r.id === rId);
            if (rel) {
              const target = rel.target.startsWith("/")
                ? rel.target.slice(1)
                : `word/${rel.target}`;
              if (doc.pkg.has(target)) return target;
            }
          }
        }
      }
    }
  }

  // Create a fresh part.
  const idx = nextHeaderFooterIndex(doc);
  const partName = `word/${scope}${idx}.xml`;
  const rootTag = scope === "header" ? "w:hdr" : "w:ftr";
  doc.pkg.setPart(
    partName,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + `<${rootTag} xmlns:w="${W_NS}"/>`,
  );
  ensureContentOverride(doc.pkg, partName, scope === "header" ? HEADER_CT : FOOTER_CT);
  const rId = addRelationship(
    doc.pkg,
    doc.documentPartName,
    scope === "header" ? HEADER_REL : FOOTER_REL,
    `${scope}${idx}.xml`,
  );

  // Splice the reference as the first child of the body sectPr.
  const refTag = `<${refName} w:type="default" r:id="${rId}"/>`;
  const next = insertBodyReference(doc.documentXml(), refTag);
  doc.pkg.setPart(doc.documentPartName, next);
  doc.invalidate();
  return partName;
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

/** The trailing body-level w:sectPr, or null. */
function bodySectPr(xml: string): ElementSlice | null {
  const body = findElement(xml, "w:body");
  if (!body) return null;
  const kids = childElements(xml, body.contentStart, body.contentEnd);
  return kids.find((k) => k.name === "w:sectPr") ?? null;
}

function nextHeaderFooterIndex(doc: DocHandle): number {
  let max = 0;
  for (const name of doc.pkg.entryNames()) {
    const m = /^word\/(?:header|footer)([0-9]+)\.xml$/.exec(name);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max + 1;
}

/** Splice a reference as the first child of the trailing body sectPr. */
function insertBodyReference(xml: string, refTag: string): string {
  const sect = bodySectPr(xml);
  if (!sect) throw new ToolError("anchor_not_found", "Document has no body section.", []);
  if (sect.selfClosed) {
    return splice(xml, sect.start, sect.end, `<w:sectPr>${refTag}</w:sectPr>`);
  }
  return splice(xml, sect.contentStart, sect.contentStart, refTag);
}

// --- update ------------------------------------------------------------------

function fieldUpdate(doc: DocHandle): DocxFieldResult {
  ensureSettingsPart(doc);
  const xml = doc.pkg.partText(SETTINGS_PART);
  // Idempotent: if updateFields is already present, leave it.
  if (findElement(xml, "w:updateFields")) {
    return { updated: 1, note: "Fields already flagged for update." };
  }
  const entry = '<w:updateFields w:val="true"/>';
  const settings = findElement(xml, "w:settings");
  let next: string;
  if (settings && !settings.selfClosed) {
    // updateFields is the first child of w:settings.
    next = splice(xml, settings.contentStart, settings.contentStart, entry);
  } else if (settings && settings.selfClosed) {
    next = splice(
      xml,
      settings.start,
      settings.end,
      `<w:settings xmlns:w="${W_NS}">${entry}</w:settings>`,
    );
  } else {
    next = xml;
  }
  doc.pkg.setPart(SETTINGS_PART, next);
  return { updated: 1, note: "Flagged all fields for update on next render." };
}

function ensureSettingsPart(doc: DocHandle): void {
  if (doc.pkg.has(SETTINGS_PART)) return;
  doc.pkg.setPart(
    SETTINGS_PART,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + `<w:settings xmlns:w="${W_NS}"/>`,
  );
  ensureContentOverride(doc.pkg, SETTINGS_PART, SETTINGS_CT);
  addRelationship(doc.pkg, doc.documentPartName, SETTINGS_REL, "settings.xml");
}
