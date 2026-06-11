/**
 * Edit-surface algorithms per spec/algorithms.md §4–§7: run coalescing and
 * splice-based replacement (§4), tracked-changes emission (§5), the LCS word
 * diff and paragraph rebuild (§6/§6a), and revision scan/accept/reject with
 * the run-merge post-pass (§7).
 *
 * Everything here is a pure function over the document part's text; the tool
 * layer (toolsEdit.ts) owns sessions, anchors, and result shapes. The Python
 * implementation (`python/src/docxengine/_edits.py`) is the byte-parity twin —
 * any change here must keep both emitting identical XML.
 */
import type { AnchorEntry } from "./anchors.js";
import {
  type ElementSlice,
  type SpliceEdit,
  type TextPiece,
  attrs,
  bodyBlocks,
  childElements,
  decodeEntities,
  elementExtent,
  emitTextElement,
  escapeAttr,
  findElement,
  isWhitespaceChar,
  nextTag,
  splice,
  spliceAll,
  textPieces,
} from "./xmlscan.js";

// ---------------------------------------------------------------------------
// Revision metadata (§5)
// ---------------------------------------------------------------------------

/** §5 date: DOCXENGINE_FIXED_DATE verbatim, else current UTC ISO-8601 seconds + Z. */
export function revisionDate(): string {
  const fixed = process.env["DOCXENGINE_FIXED_DATE"];
  if (fixed !== undefined && fixed !== "") return fixed;
  return new Date().toISOString().replace(/\.[0-9]{3}Z$/, "Z");
}

/** §5 author: the argument verbatim, else env DOCXENGINE_AUTHOR, else "DocxEngine". */
export function revisionAuthor(author?: string | null): string {
  if (author != null) return author;
  return process.env["DOCXENGINE_AUTHOR"] || "DocxEngine";
}

/** §5 id allocation base: max existing `w:ins`/`w:del` `w:id` in the part (0 if none). */
export function maxRevisionId(xml: string): number {
  let max = 0;
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return max;
    if ((t.name === "w:ins" || t.name === "w:del") && t.kind !== "end") {
      const v = attrs(xml, t)["w:id"];
      if (v !== undefined && /^[0-9]+$/.test(v)) max = Math.max(max, Number(v));
    }
    i = t.end;
  }
}

/** A `w:ins`/`w:del` start tag with attributes in §5 order: id, author, date. */
export function revisionOpen(
  kind: "ins" | "del",
  id: number,
  author: string,
  date: string,
): string {
  return `<w:${kind} w:id="${id}" w:author="${escapeAttr(author)}" w:date="${escapeAttr(date)}">`;
}

/** Emit one minimal run: `<w:r>{rPr}<w:t…>text</w:t></w:r>` (§3 xml:space rule). */
export function emitRun(rPr: string, text: string, tag: "w:t" | "w:delText" = "w:t"): string {
  return `<w:r>${rPr}${emitTextElement(tag, text)}</w:r>`;
}

// ---------------------------------------------------------------------------
// Matching (§4 step 2 / §2a): literal, case-sensitive, non-overlapping, LTR
// ---------------------------------------------------------------------------

export function findOccurrences(text: string, old: string): [number, number][] {
  const out: [number, number][] = [];
  if (old === "") return out;
  let i = text.indexOf(old);
  while (i >= 0) {
    out.push([i, i + old.length]);
    i = text.indexOf(old, i + old.length);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Containers: the enclosing `w:r` of a text piece (or the `w:t` itself)
// ---------------------------------------------------------------------------

interface Container {
  start: number;
  end: number;
  /** The run's direct-child `w:rPr` verbatim ("" when absent or not a run). */
  rPr: string;
}

function containerOf(xml: string, piece: TextPiece): Container {
  if (piece.runStart < 0) {
    // Malformed-but-tolerated: a w:t with no enclosing run is its own container.
    return { start: piece.el.start, end: piece.el.end, rPr: "" };
  }
  const tag = nextTag(xml, piece.runStart);
  if (!tag || tag.start !== piece.runStart) throw new Error("run boundary mismatch");
  const run = elementExtent(xml, tag);
  let rPr = "";
  if (!run.selfClosed) {
    const child = childElements(xml, run.contentStart, run.contentEnd).find(
      (k) => k.name === "w:rPr",
    );
    if (child) rPr = xml.slice(child.start, child.end);
  }
  return { start: run.start, end: run.end, rPr };
}

function containerKey(piece: TextPiece): number {
  return piece.runStart >= 0 ? piece.runStart : piece.el.start;
}

/** The §4 searchable pieces of a paragraph: its `w:t` descendants in order. */
function tPieces(xml: string, scope: { contentStart: number; contentEnd: number }): TextPiece[] {
  return textPieces(xml, scope).filter((p) => p.kind === "t");
}

// ---------------------------------------------------------------------------
// §4 untracked replacement: splice one match into the paragraph
// ---------------------------------------------------------------------------

/**
 * Apply one untracked replacement of searchable-text range `[s, e)` with
 * `newText` (§4 rules 3–5): the first overlapping `w:t` takes prefix +
 * replacement (+ its suffix when the match ends inside it); subsequent
 * overlapping `w:t` keep only their suffix; a run whose only `w:t` is left
 * empty is removed entirely, an emptied `w:t` among siblings is dropped alone.
 */
export function applyPlainMatch(
  xml: string,
  para: ElementSlice,
  s: number,
  e: number,
  newText: string,
): string {
  const pieces = tPieces(xml, para);
  const hits = pieces.filter((p) => p.textOffset < e && p.textOffset + p.text.length > s);
  const runTCounts = new Map<number, number>();
  for (const p of pieces) {
    if (p.runStart >= 0) runTCounts.set(p.runStart, (runTCounts.get(p.runStart) ?? 0) + 1);
  }
  const edits: SpliceEdit[] = [];
  hits.forEach((p, i) => {
    const lo = Math.max(0, s - p.textOffset);
    const hi = Math.min(p.text.length, Math.max(0, e - p.textOffset));
    const text = i === 0 ? p.text.slice(0, lo) + newText + p.text.slice(hi) : p.text.slice(hi);
    if (text !== "") {
      edits.push({ start: p.el.start, end: p.el.end, text: emitTextElement("w:t", text) });
    } else if (p.runStart >= 0 && runTCounts.get(p.runStart) === 1) {
      edits.push({ start: p.runStart, end: p.runEnd, text: "" }); // §4 rule 4
    } else {
      edits.push({ start: p.el.start, end: p.el.end, text: "" });
    }
  });
  return spliceAll(xml, edits);
}

// ---------------------------------------------------------------------------
// §5 tracked replacement: rebuild the matched run region as a redline
// ---------------------------------------------------------------------------

export interface TrackedIds {
  delId: number;
  /** null when the replacement text is empty (no `w:ins` is emitted). */
  insId: number | null;
}

/**
 * Apply one tracked replacement at searchable-text range `[s, e)` (§5): the
 * matched region becomes prefix-run + `w:del` (one `w:delText` run per
 * overlapped run, each keeping its own `rPr`) + `w:ins` (first matched run's
 * `rPr`) + suffix-run.
 */
export function applyTrackedReplace(
  xml: string,
  para: ElementSlice,
  s: number,
  e: number,
  newText: string,
  author: string,
  date: string,
  ids: TrackedIds,
): string {
  const pieces = tPieces(xml, para);
  const hits = pieces.filter((p) => p.textOffset < e && p.textOffset + p.text.length > s);
  const containers: Container[] = [];
  for (const p of hits) {
    const c = containerOf(xml, p);
    const last = containers[containers.length - 1];
    if (last === undefined || last.start !== c.start) containers.push(c);
  }
  const first = containers[0];
  const last = containers[containers.length - 1];
  if (first === undefined || last === undefined) return xml;
  const starts = new Set(containers.map((c) => c.start));
  const grouped = new Map<number, TextPiece[]>();
  for (const p of pieces) {
    const key = containerKey(p);
    if (!starts.has(key)) continue;
    const list = grouped.get(key);
    if (list) list.push(p);
    else grouped.set(key, [p]);
  }
  const clamp = (p: TextPiece, index: number): number =>
    Math.max(0, Math.min(p.text.length, index - p.textOffset));
  const groupOf = (c: Container): TextPiece[] => grouped.get(c.start) ?? [];
  const prefix = groupOf(first)
    .map((p) => p.text.slice(0, clamp(p, s)))
    .join("");
  const suffix = groupOf(last)
    .map((p) => p.text.slice(clamp(p, e)))
    .join("");
  let out = "";
  if (prefix !== "") out += emitRun(first.rPr, prefix);
  out += revisionOpen("del", ids.delId, author, date);
  for (const c of containers) {
    const matched = groupOf(c)
      .map((p) => p.text.slice(clamp(p, s), clamp(p, e)))
      .join("");
    out += emitRun(c.rPr, matched, "w:delText");
  }
  out += "</w:del>";
  if (newText !== "" && ids.insId !== null) {
    out += revisionOpen("ins", ids.insId, author, date) + emitRun(first.rPr, newText) + "</w:ins>";
  }
  if (suffix !== "") out += emitRun(last.rPr, suffix);
  return splice(xml, first.start, last.end, out);
}

// ---------------------------------------------------------------------------
// §6 word-level diff
// ---------------------------------------------------------------------------

/** §6 step 1: units of word + following whitespace; leading whitespace → unit 1. */
export function diffUnits(text: string): string[] {
  const units: string[] = [];
  let leading = "";
  let i = 0;
  while (i < text.length) {
    const ws = isWhitespaceChar(text[i] as string);
    let j = i + 1;
    while (j < text.length && isWhitespaceChar(text[j] as string) === ws) j++;
    const token = text.slice(i, j);
    if (ws) {
      if (units.length > 0) units[units.length - 1] += token;
      else leading += token;
    } else {
      units.push(leading + token);
      leading = "";
    }
    i = j;
  }
  if (leading !== "") units.push(leading);
  return units;
}

export interface DiffOp {
  op: "keep" | "del" | "ins";
  unit: string;
}

/** §6 step 2: LCS lengths of `old[i:]` vs `new[j:]` (the pinned table shape). */
export function lcsOps(oldU: readonly string[], newU: readonly string[]): Int32Array[] {
  const n = oldU.length;
  const m = newU.length;
  const L: Int32Array[] = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    const row = L[i] as Int32Array;
    const below = L[i + 1] as Int32Array;
    for (let j = m - 1; j >= 0; j--) {
      row[j] =
        oldU[i] === newU[j]
          ? (below[j + 1] as number) + 1
          : Math.max(below[j] as number, row[j + 1] as number);
    }
  }
  return L;
}

/** §6 step 2: LCS over units with the pinned deterministic forward backtrack. */
export function wordDiff(oldU: readonly string[], newU: readonly string[]): DiffOp[] {
  const n = oldU.length;
  const m = newU.length;
  const L = lcsOps(oldU, newU);
  const ops: DiffOp[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (
      oldU[i] === newU[j] &&
      (L[i] as Int32Array)[j] === ((L[i + 1] as Int32Array)[j + 1] as number) + 1
    ) {
      ops.push({ op: "keep", unit: oldU[i] as string });
      i++;
      j++;
    } else if (((L[i + 1] as Int32Array)[j] as number) >= ((L[i] as Int32Array)[j + 1] as number)) {
      ops.push({ op: "del", unit: oldU[i] as string });
      i++;
    } else {
      ops.push({ op: "ins", unit: newU[j] as string });
      j++;
    }
  }
  for (; i < n; i++) ops.push({ op: "del", unit: oldU[i] as string });
  for (; j < m; j++) ops.push({ op: "ins", unit: newU[j] as string });
  return ops;
}

export interface DiffBlock {
  kind: "keep" | "del" | "ins";
  text: string;
}

/** §6 step 3: maximal del/ins runs as one span each; del precedes ins at a position. */
export function diffBlocks(ops: readonly DiffOp[]): DiffBlock[] {
  const blocks: DiffBlock[] = [];
  let dels: string[] = [];
  let inss: string[] = [];
  const flush = (): void => {
    if (dels.length > 0) {
      blocks.push({ kind: "del", text: dels.join("") });
      dels = [];
    }
    if (inss.length > 0) {
      blocks.push({ kind: "ins", text: inss.join("") });
      inss = [];
    }
  };
  for (const { op, unit } of ops) {
    if (op === "keep") {
      flush();
      const lastBlock = blocks[blocks.length - 1];
      if (lastBlock !== undefined && lastBlock.kind === "keep") lastBlock.text += unit;
      else blocks.push({ kind: "keep", text: unit });
    } else if (op === "del") {
      dels.push(unit);
    } else {
      inss.push(unit);
    }
  }
  flush();
  return blocks;
}

/**
 * `(rPr, portion)` per overlapped run for `[start, end)` of the §4 text;
 * consecutive `w:t` pieces of the same run concatenate into one portion (§6a).
 */
function runPortions(
  xml: string,
  pieces: readonly TextPiece[],
  start: number,
  end: number,
): { rPr: string; portion: string }[] {
  const out: { key: number; rPr: string; portion: string }[] = [];
  for (const p of pieces) {
    if (p.textOffset >= end || p.textOffset + p.text.length <= start) continue;
    const lo = Math.max(0, start - p.textOffset);
    const hi = Math.min(p.text.length, end - p.textOffset);
    const key = containerKey(p);
    const lastEntry = out[out.length - 1];
    if (lastEntry !== undefined && lastEntry.key === key) {
      lastEntry.portion += p.text.slice(lo, hi);
    } else {
      out.push({ key, rPr: containerOf(xml, p).rPr, portion: p.text.slice(lo, hi) });
    }
  }
  return out.map(({ rPr, portion }) => ({ rPr, portion }));
}

/** §6a insert-only spans: the rPr of the run containing the insertion offset. */
function rprAtOffset(xml: string, pieces: readonly TextPiece[], offset: number): string {
  if (pieces.length === 0) return ""; // empty paragraph yields no rPr
  for (const p of pieces) {
    if (p.textOffset <= offset && offset < p.textOffset + p.text.length) {
      return containerOf(xml, p).rPr;
    }
  }
  // End-of-paragraph insertion takes the last run's rPr.
  return containerOf(xml, pieces[pieces.length - 1] as TextPiece).rPr;
}

/**
 * Replace a paragraph's content after `w:pPr` from §6 diff blocks (§6a):
 * untracked → one run with the first existing run's rPr and the new text;
 * tracked → kept spans re-emitted per overlapped run, del/ins wrappers with
 * §5 metadata, ids allocated in emission order.
 */
export function rebuildParagraph(
  xml: string,
  p: ElementSlice,
  blocks: readonly DiffBlock[],
  opts: { tracked: boolean; author: string; date: string },
): string {
  let pPrXml = "";
  let firstRun: ElementSlice | null = null;
  if (!p.selfClosed) {
    const pPr = childElements(xml, p.contentStart, p.contentEnd).find((k) => k.name === "w:pPr");
    if (pPr) pPrXml = xml.slice(pPr.start, pPr.end);
    firstRun = findElement(xml, "w:r", p.contentStart, p.contentEnd);
  }
  const parts: string[] = [pPrXml];
  if (!opts.tracked) {
    const text = blocks
      .filter((b) => b.kind !== "del")
      .map((b) => b.text)
      .join("");
    if (text !== "") {
      let rPr = "";
      if (firstRun !== null && !firstRun.selfClosed) {
        const child = childElements(xml, firstRun.contentStart, firstRun.contentEnd).find(
          (k) => k.name === "w:rPr",
        );
        if (child) rPr = xml.slice(child.start, child.end);
      }
      parts.push(emitRun(rPr, text));
    }
  } else {
    const pieces = p.selfClosed ? [] : tPieces(xml, p);
    let revId = maxRevisionId(xml) + 1;
    let pos = 0; // offset into the old §4 concatenated text
    let replaceRPr: string | null = null; // first deleted run's rPr, for del→ins pairs
    for (const block of blocks) {
      if (block.kind === "keep") {
        for (const { rPr, portion } of runPortions(xml, pieces, pos, pos + block.text.length)) {
          parts.push(emitRun(rPr, portion));
        }
        pos += block.text.length;
        replaceRPr = null;
      } else if (block.kind === "del") {
        const portions = runPortions(xml, pieces, pos, pos + block.text.length);
        parts.push(revisionOpen("del", revId++, opts.author, opts.date));
        for (const { rPr, portion } of portions) parts.push(emitRun(rPr, portion, "w:delText"));
        parts.push("</w:del>");
        pos += block.text.length;
        replaceRPr = portions.length > 0 ? (portions[0] as { rPr: string }).rPr : "";
      } else {
        const rPr = replaceRPr ?? rprAtOffset(xml, pieces, pos);
        parts.push(
          revisionOpen("ins", revId++, opts.author, opts.date) +
            emitRun(rPr, block.text) +
            "</w:ins>",
        );
        replaceRPr = null;
      }
    }
  }
  const inner = parts.join("");
  if (p.selfClosed) {
    const openTag = xml.slice(p.start, p.end - 2) + ">"; // reopen "<w:p …/>"
    return splice(xml, p.start, p.end, openTag + inner + "</w:p>");
  }
  return splice(xml, p.contentStart, p.contentEnd, inner);
}

// ---------------------------------------------------------------------------
// §7 revisions
// ---------------------------------------------------------------------------

/** Rename `w:t` → `w:delText` in raw XML (tracked paragraph deletion, §6a). */
export function tToDelText(content: string): string {
  return content
    .replaceAll("<w:t>", "<w:delText>")
    .replaceAll("<w:t ", "<w:delText ")
    .replaceAll("<w:t/>", "<w:delText/>")
    .replaceAll("</w:t>", "</w:delText>");
}

/** Rename `w:delText` → `w:t` in raw XML (reject of a `w:del`, §7). */
export function delTextToT(content: string): string {
  return content
    .replaceAll("<w:delText>", "<w:t>")
    .replaceAll("<w:delText ", "<w:t ")
    .replaceAll("<w:delText/>", "<w:t/>")
    .replaceAll("</w:delText>", "</w:t>");
}

/** One `w:ins`/`w:del` element of the document part. */
export interface Revision {
  el: ElementSlice;
  kind: "ins" | "del";
  /** `R{w:id}` (the attribute verbatim). */
  id: string;
  author: string;
  date: string;
  /** Containing body block's anchor, when there is one. */
  anchor: string | null;
  /** Containing body *paragraph* ordinal (merge post-pass), when there is one. */
  ordinal: number | null;
  /** The wrapper's own raw text: `w:t` for ins, `w:delText` for del (§6a). */
  text: string;
}

/** Every `w:ins`/`w:del` of the part in document order, nested included (§6a). */
export function scanRevisions(xml: string, blocks: readonly AnchorEntry[]): Revision[] {
  const out: Revision[] = [];
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return out;
    if ((t.name === "w:ins" || t.name === "w:del") && t.kind !== "end") {
      const a = attrs(xml, t);
      const el = elementExtent(xml, t);
      const kind: Revision["kind"] = t.name === "w:ins" ? "ins" : "del";
      const pieceKind = kind === "ins" ? "t" : "delText";
      let text = "";
      if (!el.selfClosed) {
        for (const p of textPieces(xml, el)) if (p.kind === pieceKind) text += p.text;
      }
      const block = blocks.find((b) => b.start < el.start && el.end <= b.end);
      out.push({
        el,
        kind,
        id: `R${a["w:id"] ?? ""}`,
        author: a["w:author"] ?? "unknown",
        date: a["w:date"] ?? "",
        anchor: block !== undefined ? block.anchor : null,
        ordinal: block !== undefined && block.kind === "p" ? block.ordinal : null,
        text,
      });
      i = t.end; // descend into content so nested wrappers are seen
      continue;
    }
    i = t.end;
  }
}

export interface RevisionFilter {
  author?: string | undefined;
  /** §7 ISO date-prefix match against w:date. */
  date?: string | undefined;
  /** Only revisions with w:date on or after this ISO date (§6a). */
  after?: string | undefined;
  /** Only revisions with w:date strictly before this ISO date (§6a). */
  before?: string | undefined;
}

/** §7/§6a filters: author exact; date prefix; after ≤ w:date < before (strings). */
export function revisionMatches(rev: Revision, flt: RevisionFilter): boolean {
  if (flt.author !== undefined && rev.author !== flt.author) return false;
  if (flt.date !== undefined && !rev.date.startsWith(flt.date)) return false;
  if (flt.after !== undefined && rev.date < flt.after) return false;
  return flt.before === undefined || rev.date < flt.before;
}

/** §7: accept ins / reject del → unwrap; accept del / reject ins → remove. */
export function resolveRevisions(
  xml: string,
  candidates: readonly Revision[],
  accept: boolean,
): string {
  const edits: SpliceEdit[] = [];
  for (const rev of candidates) {
    const { el } = rev;
    if ((rev.kind === "ins") === accept) {
      let inner = el.selfClosed ? "" : xml.slice(el.contentStart, el.contentEnd);
      if (rev.kind === "del") inner = delTextToT(inner);
      edits.push({ start: el.start, end: el.end, text: inner });
    } else {
      edits.push({ start: el.start, end: el.end, text: "" });
    }
  }
  return spliceAll(xml, edits);
}

// ---------------------------------------------------------------------------
// §7 post-pass: merge adjacent sibling runs with rsid-insensitive equal rPr
// ---------------------------------------------------------------------------

const RSID_ATTR_RE = /[ \t\r\n]+w:rsid[A-Za-z]*="[^"]*"/g;

/** §7 post-pass comparison key: drop `rsid*` attributes; empty rPr ≡ absent. */
function rPrKey(rPr: string): string {
  const stripped = rPr.replace(RSID_ATTR_RE, "");
  return stripped === "<w:rPr/>" || stripped === "<w:rPr></w:rPr>" ? "" : stripped;
}

/** `(raw rPr, concatenated text)` iff the run is `rPr? + w:t*` only. */
function mergeableParts(xml: string, run: ElementSlice): { rPr: string; text: string } | null {
  let rPr = "";
  let sawText = false;
  let text = "";
  if (!run.selfClosed) {
    for (const child of childElements(xml, run.contentStart, run.contentEnd)) {
      if (child.name === "w:rPr" && rPr === "" && !sawText) {
        rPr = xml.slice(child.start, child.end);
      } else if (child.name === "w:t") {
        sawText = true;
        if (!child.selfClosed) {
          text += decodeEntities(xml.slice(child.contentStart, child.contentEnd));
        }
      } else {
        return null;
      }
    }
  }
  return { rPr, text };
}

/** §7 post-pass: merge adjacent sibling runs with identical rPr (rsid-blind). */
export function mergeParagraphRuns(xml: string, ordinal: number): string {
  for (;;) {
    const p = paragraphBlock(xml, ordinal);
    if (p.selfClosed) return xml;
    const runs = childElements(xml, p.contentStart, p.contentEnd).filter((c) => c.name === "w:r");
    let merged: SpliceEdit | null = null;
    for (let k = 0; k + 1 < runs.length; k++) {
      const first = runs[k] as ElementSlice;
      const second = runs[k + 1] as ElementSlice;
      if (first.end !== second.start) continue;
      const a = mergeableParts(xml, first);
      const b = mergeableParts(xml, second);
      if (a === null || b === null || rPrKey(a.rPr) !== rPrKey(b.rPr)) continue;
      const openTag = first.selfClosed
        ? xml.slice(first.start, first.end - 2) + ">"
        : xml.slice(first.start, first.startTagEnd);
      merged = {
        start: first.start,
        end: second.end,
        text: openTag + a.rPr + emitTextElement("w:t", a.text + b.text) + "</w:r>",
      };
      break;
    }
    if (merged === null) return xml;
    xml = splice(xml, merged.start, merged.end, merged.text);
  }
}

// ---------------------------------------------------------------------------
// Body navigation
// ---------------------------------------------------------------------------

/** The body-level `w:p` at a 1-based ordinal in the current part text. */
export function paragraphBlock(xml: string, ordinal: number): ElementSlice {
  let n = 0;
  for (const b of bodyBlocks(xml)) {
    if (b.name === "w:p" && ++n === ordinal) return b;
  }
  throw new Error(`paragraph ordinal ${ordinal} out of range`);
}
