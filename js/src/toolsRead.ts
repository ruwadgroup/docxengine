/**
 * Read-surface tools — docx_open / docx_outline / docx_read / docx_search —
 * with result shapes exactly per spec/tools/*.json and the §2a pins.
 * Errors are thrown as ToolError (spec/errors.json codes).
 */
import { ToolError } from "./errors.js";
import {
  type OutlineHeading,
  type OutlineTable,
  type ReadResult,
  type SearchMatch,
  outlineOf,
  paragraphHeadingLevel,
  projectionContext,
  readProjection,
  searchProjection,
} from "./projector.js";
import type { Session } from "./session.js";
import { nextTag } from "./xmlscan.js";

type ResponseFormat = "concise" | "detailed";

// ---------------------------------------------------------------------------
// docx_open
// ---------------------------------------------------------------------------

export interface DocxOpenArgs {
  path?: string | undefined;
  /** Base64-encoded .docx content (alternative to path). */
  bytes?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxOpenResult {
  doc_id: string;
  summary: string;
  n_paragraphs: number;
  has_tracked_changes: boolean;
  has_comments: boolean;
}

function countNamed(xml: string, names: readonly string[]): Map<string, number> {
  const counts = new Map<string, number>(names.map((n) => [n, 0]));
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return counts;
    if (t.kind !== "end") {
      const c = counts.get(t.name);
      if (c !== undefined) counts.set(t.name, c + 1);
    }
    i = t.end;
  }
}

function plural(n: number, noun: string): string {
  return `${n} ${noun}${n === 1 ? "" : "s"}`;
}

/** Decode standard base64 via WHATWG `atob` (browsers and Node alike). */
function base64Bytes(b64: string): Uint8Array {
  let binary: string;
  try {
    binary = atob(b64);
  } catch (e) {
    throw new ToolError(
      "open_failed",
      `Cannot open: bytes is not valid base64 (${(e as Error).message}).`,
      ["Pass standard base64 (RFC 4648) of the .docx bytes."],
    );
  }
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

export function docxOpen(session: Session, args: DocxOpenArgs): DocxOpenResult {
  if (args.path === undefined && args.bytes === undefined) {
    throw new ToolError("open_failed", "Cannot open: provide path or bytes.", [
      "Pass a filesystem path or base64-encoded bytes.",
    ]);
  }
  const doc =
    args.path !== undefined
      ? session.open(args.path)
      : session.open(base64Bytes(args.bytes as string));
  const xml = doc.documentXml();
  const ctx = projectionContext(doc);
  const index = doc.anchorIndex();
  let nParagraphs = 0;
  let nTables = 0;
  for (const e of index) {
    if (e.kind === "p") nParagraphs++;
    else nTables++;
  }
  const counts = countNamed(xml, ["w:sectPr", "w:ins", "w:del", "w:commentReference"]);
  // §2a title: first non-empty heading paragraph, else first non-empty
  // paragraph, else "Untitled".
  let title: string | null = null;
  let firstNonEmpty: string | null = null;
  for (const e of index) {
    if (e.kind !== "p" || !e.normalized) continue;
    if (paragraphHeadingLevel(xml, e.block, ctx.styles) !== null) {
      title = e.normalized;
      break;
    }
    if (firstNonEmpty === null) firstNonEmpty = e.normalized;
  }
  const summary =
    `${title ?? firstNonEmpty ?? "Untitled"} — ${plural(nParagraphs, "paragraph")}, ` +
    `${plural(counts.get("w:sectPr") ?? 0, "section")}, ${plural(nTables, "table")}`;
  return {
    doc_id: doc.id,
    summary,
    n_paragraphs: nParagraphs,
    has_tracked_changes: (counts.get("w:ins") ?? 0) + (counts.get("w:del") ?? 0) > 0,
    has_comments: (counts.get("w:commentReference") ?? 0) > 0,
  };
}

// ---------------------------------------------------------------------------
// docx_outline
// ---------------------------------------------------------------------------

export interface DocxOutlineArgs {
  doc_id: string;
  response_format?: ResponseFormat | undefined;
}

export interface DocxOutlineResult {
  outline: OutlineHeading[];
  tables: OutlineTable[];
}

export function docxOutline(session: Session, args: DocxOutlineArgs): DocxOutlineResult {
  return outlineOf(session.get(args.doc_id));
}

// ---------------------------------------------------------------------------
// docx_read
// ---------------------------------------------------------------------------

export interface DocxReadArgs {
  doc_id: string;
  anchor?: string | undefined;
  range?: string | undefined;
  window?: number | undefined;
  scope?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export function docxRead(session: Session, args: DocxReadArgs): ReadResult {
  const doc = session.get(args.doc_id);
  return readProjection(doc, {
    anchor: args.anchor,
    range: args.range,
    window: args.window,
    scope: args.scope,
  });
}

// ---------------------------------------------------------------------------
// docx_search
// ---------------------------------------------------------------------------

export interface DocxSearchArgs {
  doc_id: string;
  query: string;
  regex?: boolean | undefined;
  scope?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxSearchResult {
  matches: SearchMatch[];
  n_matches: number;
}

export function docxSearch(session: Session, args: DocxSearchArgs): DocxSearchResult {
  const doc = session.get(args.doc_id);
  return searchProjection(doc, { query: args.query, regex: args.regex, scope: args.scope });
}
