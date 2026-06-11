/**
 * Media (`docx_media`) per spec/algorithms.md §19.
 *
 * `insert` writes `word/media/image{k}.{ext}` (k = max existing + 1), a
 * document image rel, a content-type Default for the extension, and an inline
 * drawing run after/before the anchor paragraph. EMU = round(cm × 360000);
 * when only one of width/height is given, the source's pixel dimensions are
 * parsed (PNG IHDR / JPEG SOFn) and the other side scaled by aspect ratio.
 * `extract` copies the M{id} part's bytes to a path; `replace` overwrites the
 * part's bytes keeping the rel/rId. M{ordinal} = document order of drawing
 * references. The Python twin (`_media.py`) is the byte-parity reference.
 */
import { ToolError } from "./errors.js";
import { nodeFs } from "./nodeenv.js";
import {
  EXT_CONTENT_TYPES,
  addRelationship,
  anchorInvalid,
  anchorNotFound,
  anchorStale,
  ensureContentDefault,
  extensionOf,
} from "./phase2common.js";
import type { DocHandle, Session } from "./session.js";
import { attrs, bodyBlocks, nextTag, splice } from "./xmlscan.js";

const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const FALLBACK_CT = "application/octet-stream";

const FULL_ANCHOR_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;
const M_ID_RE = /^M([1-9][0-9]*)$/;

const EMU_PER_CM = 360000;

// ---------------------------------------------------------------------------
// Pixel-dimension parsing (§19)
// ---------------------------------------------------------------------------

const PNG_SIG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function readU16BE(b: Uint8Array, off: number): number {
  return (b[off]! << 8) | b[off + 1]!;
}

function readU32BE(b: Uint8Array, off: number): number {
  return b[off]! * 0x1000000 + (b[off + 1]! << 16) + (b[off + 2]! << 8) + b[off + 3]!;
}

/** PNG width/height from bytes 16–24 (big-endian) after the IHDR signature. */
export function pngDimensions(b: Uint8Array): { width: number; height: number } | null {
  if (b.length < 24) return null;
  for (let i = 0; i < 8; i++) if (b[i] !== PNG_SIG[i]) return null;
  // After the 8-byte signature: 4-byte length, "IHDR", then width(4), height(4).
  if (!(b[12] === 0x49 && b[13] === 0x48 && b[14] === 0x44 && b[15] === 0x52)) return null;
  return { width: readU32BE(b, 16), height: readU32BE(b, 20) };
}

/** JPEG width/height from the first SOF0/1/2 marker (big-endian height,width). */
export function jpegDimensions(b: Uint8Array): { width: number; height: number } | null {
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return null;
  let i = 2;
  while (i + 9 < b.length) {
    if (b[i] !== 0xff) {
      i++;
      continue;
    }
    const marker = b[i + 1]!;
    // Standalone markers without a length payload.
    if (marker === 0xd8 || marker === 0xd9 || (marker >= 0xd0 && marker <= 0xd7)) {
      i += 2;
      continue;
    }
    const len = readU16BE(b, i + 2);
    if (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) {
      // SOFn: precision(1), height(2), width(2) follow the 2-byte length.
      const height = readU16BE(b, i + 5);
      const width = readU16BE(b, i + 7);
      return { width, height };
    }
    i += 2 + len;
  }
  return null;
}

function parseDimensions(bytes: Uint8Array, ext: string): { width: number; height: number } | null {
  if (ext === "png") return pngDimensions(bytes);
  if (ext === "jpg" || ext === "jpeg") return jpegDimensions(bytes);
  return null;
}

/** EMU = round(cm × 360000). */
export function cmToEmu(cm: number): number {
  return Math.round(cm * EMU_PER_CM);
}

// ---------------------------------------------------------------------------
// Drawing emission
// ---------------------------------------------------------------------------

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/** The §19 inline drawing run for an image rel at the given EMU extent. */
export function drawingRun(rId: string, cx: number, cy: number, name: string): string {
  return (
    '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="${name}"/>` +
    `<a:graphic xmlns:a="${A_NS}">` +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic xmlns:pic="${PIC_NS}"><pic:nvPicPr><pic:cNvPr id="1" name="${name}"/>` +
    "<pic:cNvPicPr/></pic:nvPicPr>" +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
  );
}

// ---------------------------------------------------------------------------
// Media id resolution (M{ordinal})
// ---------------------------------------------------------------------------

interface MediaRef {
  /** 1-based document order. */
  ordinal: number;
  /** The r:embed rId of the a:blip. */
  rId: string;
}

/** Every drawing reference (`a:blip` r:embed) in document order. */
function mediaRefs(xml: string): MediaRef[] {
  const out: MediaRef[] = [];
  let i = 0;
  let ordinal = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) break;
    if (t.name === "a:blip" && t.kind !== "end") {
      const rId = attrs(xml, t)["r:embed"] ?? attrs(xml, t)["r:link"] ?? "";
      ordinal++;
      out.push({ ordinal, rId });
    }
    i = t.end;
  }
  return out;
}

function resolveMediaPart(doc: DocHandle, mediaId: string): { partName: string; rId: string } {
  const m = M_ID_RE.exec(mediaId);
  if (!m) throw anchorInvalid(`Malformed media id: ${mediaId}.`);
  const ord = Number(m[1]);
  const ref = mediaRefs(doc.documentXml()).find((r) => r.ordinal === ord);
  if (!ref) {
    throw new ToolError("not_found", `Media ${mediaId} does not exist.`, [
      "Call docx_outline to see media ids (M1, M2, …).",
    ]);
  }
  // Resolve the rId to the target part via document rels.
  const rel = doc.pkg.rels(doc.documentPartName).find((r) => r.id === ref.rId);
  if (!rel) {
    throw new ToolError("not_found", `Media ${mediaId} relationship is missing.`, [
      "The document references an image with no matching relationship.",
    ]);
  }
  const target = rel.target.startsWith("/") ? rel.target.slice(1) : `word/${rel.target}`;
  return { partName: target, rId: ref.rId };
}

// ---------------------------------------------------------------------------
// docx_media
// ---------------------------------------------------------------------------

export interface DocxMediaArgs {
  doc_id: string;
  op: "insert" | "extract" | "replace";
  after?: string | undefined;
  before?: string | undefined;
  image?: string | undefined;
  width_cm?: number | undefined;
  height_cm?: number | undefined;
  media_id?: string | undefined;
  path?: string | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
}

export interface DocxMediaResult {
  media_id?: string;
  new_anchor?: string;
  path?: string;
  note?: string;
}

export function docxMedia(session: Session, args: DocxMediaArgs): DocxMediaResult {
  const doc = session.get(args.doc_id);
  switch (args.op) {
    case "insert":
      return mediaInsert(doc, args);
    case "extract":
      return mediaExtract(doc, args);
    case "replace":
      return mediaReplace(doc, args);
    default:
      throw new ToolError("invalid_args", `docx_media: unknown op ${String(args.op)}.`, []);
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

/** Next free media index `k` (max existing word/media/image{k}.* + 1). */
function nextMediaIndex(doc: DocHandle): number {
  let max = 0;
  for (const name of doc.pkg.entryNames()) {
    const m = /^word\/media\/image([0-9]+)\./.exec(name);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max + 1;
}

function readImageBytes(path: string): Uint8Array {
  try {
    return nodeFs().readFileSync(path);
  } catch (e) {
    throw new ToolError("not_found", `Cannot read image ${path}: ${(e as Error).message}.`, [
      "Check the path to the source image.",
    ]);
  }
}

// --- insert ------------------------------------------------------------------

function mediaInsert(doc: DocHandle, args: DocxMediaArgs): DocxMediaResult {
  if ((args.after == null) === (args.before == null)) {
    throw anchorInvalid("op 'insert' requires exactly one of after or before.");
  }
  if (args.image == null) throw anchorInvalid("op 'insert' requires image.");
  const entry = requireParagraph(doc, (args.after ?? args.before) as string); // hash FIRST

  const bytes = readImageBytes(args.image);
  const ext = extensionOf(args.image);
  const k = nextMediaIndex(doc);
  const partName = `word/media/image${k}.${ext}`;
  const name = `image${k}`;

  // EMU extent from width/height (aspect-preserving when one is omitted).
  const { cx, cy } = resolveExtent(bytes, ext, args.width_cm, args.height_cm);

  // Write the part, its content-type Default, and the document image rel.
  doc.pkg.setPart(partName, bytes);
  ensureContentDefault(doc.pkg, ext, EXT_CONTENT_TYPES[ext] ?? FALLBACK_CT);
  const rId = addRelationship(doc.pkg, doc.documentPartName, IMAGE_REL, `media/image${k}.${ext}`);

  // Splice an inline drawing paragraph after/before the anchor paragraph.
  const xml = doc.documentXml();
  const run = drawingRun(rId, cx, cy, name);
  const para = `<w:p>${run}</w:p>`;
  const position = args.after != null ? entry.block.end : entry.block.start;
  doc.pkg.setPart(doc.documentPartName, splice(xml, position, position, para));
  doc.invalidate();

  // The new media's M-ordinal = its document order among drawing references.
  const refs = mediaRefs(doc.documentXml());
  const newRef = refs.find((r) => r.rId === rId);
  const mediaId = newRef ? `M${newRef.ordinal}` : `M${refs.length}`;
  const base = args.after != null ? entry.ordinal + 1 : entry.ordinal;
  const fresh = paragraphEntries(doc)[base - 1];
  const newAnchor = fresh ? fresh.anchor : (args.after ?? args.before);
  const result: DocxMediaResult = {
    media_id: mediaId,
    note: `Inserted ${name}.${ext} (${cx}×${cy} EMU).`,
  };
  if (newAnchor !== undefined) result.new_anchor = newAnchor;
  return result;
}

/** Resolve the EMU extent from cm args, parsing pixel dims when one is omitted. */
function resolveExtent(
  bytes: Uint8Array,
  ext: string,
  widthCm?: number,
  heightCm?: number,
): { cx: number; cy: number } {
  if (widthCm != null && heightCm != null) {
    return { cx: cmToEmu(widthCm), cy: cmToEmu(heightCm) };
  }
  if (widthCm != null) {
    const dims = parseDimensions(bytes, ext);
    const cx = cmToEmu(widthCm);
    if (dims && dims.width > 0) {
      return { cx, cy: Math.round(cx * (dims.height / dims.width)) };
    }
    return { cx, cy: cx };
  }
  if (heightCm != null) {
    const dims = parseDimensions(bytes, ext);
    const cy = cmToEmu(heightCm);
    if (dims && dims.height > 0) {
      return { cx: Math.round(cy * (dims.width / dims.height)), cy };
    }
    return { cx: cy, cy };
  }
  // Neither given → fall back to the pixel size at 96 DPI, else a 4cm square.
  const dims = parseDimensions(bytes, ext);
  if (dims && dims.width > 0 && dims.height > 0) {
    const cx = Math.round((dims.width / 96) * 914400);
    const cy = Math.round((dims.height / 96) * 914400);
    return { cx, cy };
  }
  const fallback = cmToEmu(4);
  return { cx: fallback, cy: fallback };
}

// --- extract -----------------------------------------------------------------

function mediaExtract(doc: DocHandle, args: DocxMediaArgs): DocxMediaResult {
  if (args.media_id == null) throw anchorInvalid("op 'extract' requires media_id.");
  if (args.path == null) throw anchorInvalid("op 'extract' requires path.");
  const { partName } = resolveMediaPart(doc, args.media_id);
  const bytes = doc.pkg.part(partName);
  try {
    nodeFs().writeFileSync(args.path, bytes);
  } catch (e) {
    throw new ToolError(
      "save_failed",
      `I/O failure writing media to ${args.path}: ${(e as Error).message}.`,
      ["Check the path and permissions."],
    );
  }
  return { media_id: args.media_id, path: args.path, note: `Extracted ${args.media_id}.` };
}

// --- replace -----------------------------------------------------------------

function mediaReplace(doc: DocHandle, args: DocxMediaArgs): DocxMediaResult {
  if (args.media_id == null) throw anchorInvalid("op 'replace' requires media_id.");
  if (args.image == null) throw anchorInvalid("op 'replace' requires image.");
  const { partName } = resolveMediaPart(doc, args.media_id);
  const newBytes = readImageBytes(args.image);
  const newExt = extensionOf(args.image);
  const oldExt = extensionOf(partName);
  // Overwrite in place keeping the rel and rId; the part name is unchanged.
  doc.pkg.setPart(partName, newBytes);
  if (newExt !== oldExt) {
    // Add a content-type Default for the new extension when it differs.
    ensureContentDefault(doc.pkg, newExt, EXT_CONTENT_TYPES[newExt] ?? FALLBACK_CT);
  }
  doc.invalidate();
  return { media_id: args.media_id, note: `Replaced ${args.media_id} bytes.` };
}
