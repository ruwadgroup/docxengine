/**
 * Shared Phase-2 package-mutation helpers per spec/algorithms.md §13–§22.
 *
 * These wire new parts into the package the §3/§8 way: ensure a content-type
 * Default/Override, allocate the next relationship id and append a document
 * relationship, and ensure a named style exists in word/styles.xml. Every
 * helper splices (never re-serializes) untouched regions. The Python twin is
 * `python/src/docxengine/_phase2_common.py`-equivalent — byte parity required.
 */
import type { Package } from "./opc.js";
import { ToolError } from "./errors.js";
import {
  type ElementSlice,
  type Tag,
  attrs,
  childElements,
  elementExtent,
  escapeAttr,
  nextTag,
  splice,
} from "./xmlscan.js";

const CONTENT_TYPES_PART = "[Content_Types].xml";

/** §8a repair content-type map; anything else gets octet-stream. */
export const EXT_CONTENT_TYPES: Readonly<Record<string, string>> = {
  rels: "application/vnd.openxmlformats-package.relationships+xml",
  xml: "application/xml",
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  gif: "image/gif",
};

/** Lowercased extension of a part name's basename ("" when none). */
export function extensionOf(partName: string): string {
  const slash = partName.lastIndexOf("/");
  const base = slash < 0 ? partName : partName.slice(slash + 1);
  const dot = base.lastIndexOf(".");
  return dot < 0 ? "" : base.slice(dot + 1).toLowerCase();
}

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

/** Splice a `Default Extension=…` before `</Types>` iff the extension is uncovered. */
export function ensureContentDefault(pkg: Package, ext: string, contentType: string): void {
  const lower = ext.toLowerCase();
  const xml = pkg.partText(CONTENT_TYPES_PART);
  // Already covered by a Default → no-op (idempotent).
  for (const tag of scan(xml, "Default")) {
    const a = attrs(xml, tag);
    if ((a["Extension"] ?? "").toLowerCase() === lower) return;
  }
  const close = xml.lastIndexOf("</Types>");
  if (close < 0) return;
  const entry = `<Default Extension="${escapeAttr(lower)}" ContentType="${escapeAttr(contentType)}"/>`;
  pkg.setPart(CONTENT_TYPES_PART, splice(xml, close, close, entry));
}

/** Splice an `Override PartName=…` before `</Types>` iff the part has no Override. */
export function ensureContentOverride(pkg: Package, partName: string, contentType: string): void {
  const want = partName.startsWith("/") ? partName : `/${partName}`;
  const xml = pkg.partText(CONTENT_TYPES_PART);
  for (const tag of scan(xml, "Override")) {
    if ((attrs(xml, tag)["PartName"] ?? "") === want) return;
  }
  const close = xml.lastIndexOf("</Types>");
  if (close < 0) return;
  const entry = `<Override PartName="${escapeAttr(want)}" ContentType="${escapeAttr(contentType)}"/>`;
  pkg.setPart(CONTENT_TYPES_PART, splice(xml, close, close, entry));
}

// ---------------------------------------------------------------------------
// Relationships
// ---------------------------------------------------------------------------

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

/** Next free `rId{n}` for a part's rels (max existing numeric suffix + 1). */
export function nextRelId(pkg: Package, partName: string): string {
  let max = 0;
  for (const rel of pkg.rels(partName)) {
    const m = /^rId([0-9]+)$/.exec(rel.id);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return `rId${max + 1}`;
}

/**
 * Append a relationship to a part's rels (creating the rels part if absent),
 * returning the allocated rId. Targets are package-relative to the source part.
 */
export function addRelationship(
  pkg: Package,
  sourcePart: string,
  relType: string,
  target: string,
  opts: { id?: string; targetMode?: "Internal" | "External" } = {},
): string {
  const relsName = relsPartFor(sourcePart);
  const xml = pkg.has(relsName) ? pkg.partText(relsName) : EMPTY_RELS;
  const id = opts.id ?? nextRelId(pkg, sourcePart);
  const mode = opts.targetMode === "External" ? ' TargetMode="External"' : "";
  const entry =
    `<Relationship Id="${escapeAttr(id)}" Type="${escapeAttr(relType)}" ` +
    `Target="${escapeAttr(target)}"${mode}/>`;
  const close = xml.lastIndexOf("</Relationships>");
  let next: string;
  if (close < 0) {
    // Self-closed root `<Relationships .../>` → expand it.
    const selfClose = xml.lastIndexOf("/>");
    next = splice(xml, selfClose, selfClose + 2, `>${entry}</Relationships>`);
  } else {
    next = splice(xml, close, close, entry);
  }
  pkg.setPart(relsName, next);
  return id;
}

/** The rels part name for a part (`word/document.xml` → `word/_rels/document.xml.rels`). */
export function relsPartFor(partName: string): string {
  const slash = partName.lastIndexOf("/");
  const dir = slash < 0 ? "" : partName.slice(0, slash + 1);
  const base = slash < 0 ? partName : partName.slice(slash + 1);
  return `${dir}_rels/${base}.rels`;
}

// ---------------------------------------------------------------------------
// Styles (§16 ensure-style): create a paragraph/table/character style on demand
// ---------------------------------------------------------------------------

/** True iff a `w:styleId="{id}"` style is present in styles.xml. */
export function styleExists(pkg: Package, id: string): boolean {
  if (!pkg.has("word/styles.xml")) return false;
  const xml = pkg.partText("word/styles.xml");
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return false;
    if (t.name === "w:style" && t.kind !== "end") {
      if (attrs(xml, t)["w:styleId"] === id) return true;
      i = elementExtent(xml, t).end;
      continue;
    }
    i = t.end;
  }
}

/**
 * Ensure a named style exists in word/styles.xml, splicing the supplied
 * definition before `</w:styles>` when absent (idempotent). Creates a minimal
 * styles part if missing.
 */
export function ensureStyle(pkg: Package, id: string, definition: string): void {
  if (!pkg.has("word/styles.xml")) {
    pkg.setPart(
      "word/styles.xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        `${definition}</w:styles>`,
    );
    ensureContentOverride(
      pkg,
      "word/styles.xml",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    );
    return;
  }
  if (styleExists(pkg, id)) return;
  const xml = pkg.partText("word/styles.xml");
  const close = xml.lastIndexOf("</w:styles>");
  if (close >= 0) {
    pkg.setPart("word/styles.xml", splice(xml, close, close, definition));
    return;
  }
  // Self-closed root `<w:styles .../>` → expand it around the definition.
  const selfClose = xml.lastIndexOf("/>");
  if (selfClose < 0) return;
  pkg.setPart(
    "word/styles.xml",
    splice(xml, selfClose, selfClose + 2, `>${definition}</w:styles>`),
  );
}

/** The canonical `TableGrid` definition (§14/§22). */
export const TABLE_GRID_STYLE =
  '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>' +
  "<w:tblPr><w:tblBorders>" +
  '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  "</w:tblBorders></w:tblPr></w:style>";

/** The canonical `ListParagraph` definition (§17/§22). */
export const LIST_PARAGRAPH_STYLE =
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>';

/** The canonical `CommentReference` definition (§18). */
export const COMMENT_REFERENCE_STYLE =
  '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/>' +
  '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>';

// ---------------------------------------------------------------------------
// Internal scan helper
// ---------------------------------------------------------------------------

function scan(xml: string, name: string): Tag[] {
  const out: Tag[] = [];
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return out;
    if (t.name === name && t.kind !== "end") out.push(t);
    i = t.end;
  }
}

// ---------------------------------------------------------------------------
// Anchor helpers shared by Phase-2 tools (the §6a edit-grade validation order)
// ---------------------------------------------------------------------------

export function anchorInvalid(detail: string): ToolError {
  return new ToolError("anchor_invalid", detail, [
    "Check the format 'P{index}#{hash}' (ranges: 'P10..P24'); tables are 'T{n}'.",
  ]);
}

export function anchorNotFound(label: string): ToolError {
  return new ToolError("anchor_not_found", `Anchor ${label} not found: index out of range.`, [
    "Call docx_outline to re-map anchors.",
  ]);
}

export function anchorStale(anchor: string): ToolError {
  return new ToolError(
    "anchor_stale",
    `Anchor ${anchor} is stale: the hash no longer matches the paragraph content.`,
    ["Call docx_read {anchor, window} and retry with the fresh anchor."],
  );
}

/** The `w:pPr` element of a paragraph/cell-paragraph, or null. */
export function findPPr(xml: string, p: ElementSlice): ElementSlice | null {
  if (p.selfClosed) return null;
  const kids = childElements(xml, p.contentStart, p.contentEnd);
  return kids.find((k) => k.name === "w:pPr") ?? null;
}
