/**
 * Anchors per spec/algorithms.md §1: `P{ordinal}#{hash}` / `T{ordinal}`.
 *
 * The hash is the first 4 lowercase hex chars of SHA-256 over the UTF-8
 * encoding of the paragraph's normalized text; ordinals are 1-based positions
 * among body-level `w:p` / `w:tbl` elements.
 */
import { ToolError } from "./errors.js";
import { sha256Hex } from "./sha256.js";
import { type ElementSlice, WS_RUN_RE, bodyBlocks, scopeText } from "./xmlscan.js";

const utf8 = new TextEncoder();

/**
 * Normalized text (algorithms.md §1 steps 2–4), given the already-concatenated
 * `w:t` character data (step 1 — see `xmlscan.scopeText`):
 * NFC → collapse every maximal §1-whitespace run to one ASCII space → strip
 * leading/trailing spaces.
 */
export function normalizedText(raw: string): string {
  const collapsed = raw.normalize("NFC").replace(WS_RUN_RE, " ");
  let start = 0;
  let end = collapsed.length;
  while (start < end && collapsed[start] === " ") start++;
  while (end > start && collapsed[end - 1] === " ") end--;
  return collapsed.slice(start, end);
}

/** First 4 lowercase hex chars of SHA-256 over the UTF-8 encoding of `text`. */
export function anchorHash(text: string): string {
  return sha256Hex(utf8.encode(text)).slice(0, 4);
}

export function paragraphAnchor(ordinal: number, normalized: string): string {
  return `P${ordinal}#${anchorHash(normalized)}`;
}

export function tableAnchor(ordinal: number): string {
  return `T${ordinal}`;
}

// ---------------------------------------------------------------------------
// Anchor index
// ---------------------------------------------------------------------------

export interface AnchorEntry {
  anchor: string;
  kind: "p" | "tbl";
  /** 1-based among body-level elements of the same kind. */
  ordinal: number;
  /** Element slice in the document part text. */
  start: number;
  end: number;
  /** Normalized paragraph text (null for tables — internals are Phase 2). */
  normalized: string | null;
  /** The underlying block element. */
  block: ElementSlice;
}

/**
 * Build the anchor index over a document part's text: body-level `w:p` get
 * `P{n}#{hash}`, body-level `w:tbl` get `T{n}`. The trailing `w:sectPr` is not
 * a paragraph; paragraphs nested inside tables get no body ordinal in MVP.
 */
export function buildAnchorIndex(xml: string): AnchorEntry[] {
  const out: AnchorEntry[] = [];
  let pOrd = 0;
  let tOrd = 0;
  for (const block of bodyBlocks(xml)) {
    if (block.name === "w:p") {
      pOrd++;
      const normalized = normalizedText(scopeText(xml, block));
      out.push({
        anchor: paragraphAnchor(pOrd, normalized),
        kind: "p",
        ordinal: pOrd,
        start: block.start,
        end: block.end,
        normalized,
        block,
      });
    } else if (block.name === "w:tbl") {
      tOrd++;
      out.push({
        anchor: tableAnchor(tOrd),
        kind: "tbl",
        ordinal: tOrd,
        start: block.start,
        end: block.end,
        normalized: null,
        block,
      });
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Anchor parsing / validation (algorithms.md §1 "Validation before every edit")
// ---------------------------------------------------------------------------

export type ParsedAnchor =
  | { kind: "p"; ordinal: number; hash: string }
  | { kind: "tbl"; ordinal: number };

const P_ANCHOR_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;
const T_ANCHOR_RE = /^T([1-9][0-9]*)$/;

/** Parse an anchor string; unparseable → `anchor_invalid`. */
export function parseAnchor(anchor: string): ParsedAnchor {
  const p = P_ANCHOR_RE.exec(anchor);
  if (p) return { kind: "p", ordinal: Number(p[1]), hash: p[2] as string };
  const t = T_ANCHOR_RE.exec(anchor);
  if (t) return { kind: "tbl", ordinal: Number(t[1]) };
  throw new ToolError("anchor_invalid", `Malformed anchor string: ${anchor}.`, [
    "Check the format 'P{index}#{hash}'.",
  ]);
}

/**
 * Resolve an anchor against the current document text: recompute the hash at
 * the given ordinal. Ordinal out of range → `anchor_not_found`; hash mismatch
 * → `anchor_stale`.
 */
export function resolveAnchor(xml: string, anchor: string): AnchorEntry {
  const parsed = parseAnchor(anchor);
  const index = buildAnchorIndex(xml);
  const entry = index.find((e) => e.kind === parsed.kind && e.ordinal === parsed.ordinal);
  if (!entry) {
    throw new ToolError(
      "anchor_not_found",
      `Anchor ${anchor} not found: index out of range or table anchor missing.`,
      ["Call docx_outline to re-map anchors."],
    );
  }
  if (parsed.kind === "p" && anchorHash(entry.normalized ?? "") !== parsed.hash) {
    throw new ToolError(
      "anchor_stale",
      `Anchor ${anchor} is stale: the hash no longer matches the paragraph content.`,
      [`docx_read(window:P${parsed.ordinal})`],
    );
  }
  return entry;
}
