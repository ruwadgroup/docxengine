/**
 * Edit-surface tools — docx_replace / docx_edit_paragraph / docx_insert /
 * docx_delete / docx_revision — with result shapes per spec/tools/*.json.
 *
 * Every edit validates its anchor hash FIRST (algorithms.md §1): hash mismatch
 * → anchor_stale, ordinal out of range → anchor_not_found, unparseable →
 * anchor_invalid. Errors are thrown as ToolError (spec/errors.json codes).
 * Behavior mirrors `python/src/docxengine/_tools_edit.py` exactly (§6a).
 */
import type { AnchorEntry } from "./anchors.js";
import {
  type Revision,
  type RevisionFilter,
  applyPlainMatch,
  applyTrackedReplace,
  diffBlocks,
  diffUnits,
  findOccurrences,
  maxRevisionId,
  mergeParagraphRuns,
  paragraphBlock,
  rebuildParagraph,
  resolveRevisions,
  revisionAuthor,
  revisionDate,
  revisionMatches,
  revisionOpen,
  scanRevisions,
  tToDelText,
  wordDiff,
} from "./edits.js";
import { ToolError } from "./errors.js";
import { parseStyles } from "./projector.js";
import type { DocHandle, Session } from "./session.js";
import {
  type SpliceEdit,
  WS_RUN_RE,
  childElements,
  emitTextElement,
  escapeAttr,
  splice,
  spliceAll,
  textWithMap,
} from "./xmlscan.js";

type ResponseFormat = "concise" | "detailed";

// ---------------------------------------------------------------------------
// Shared helpers (anchor grammar + §6a validation order: parse → range → hash)
// ---------------------------------------------------------------------------

/** Full edit-grade anchor: edits always validate the hash (§6a). */
const FULL_ANCHOR_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;
/** Paragraph range; endpoint hashes are validated when present (§6a). */
const RANGE_RE = /^P([1-9][0-9]*)(?:#([0-9a-f]{4}))?\.\.P([1-9][0-9]*)(?:#([0-9a-f]{4}))?$/;

function plural(n: number, noun: string): string {
  return n === 1 ? `${n} ${noun}` : `${n} ${noun}s`;
}

function anchorInvalidError(detail: string): ToolError {
  return new ToolError("anchor_invalid", detail, [
    "Check the format 'P{index}#{hash}' (ranges: 'P10..P24').",
  ]);
}

function anchorNotFoundError(label: string): ToolError {
  return new ToolError("anchor_not_found", `Anchor ${label} not found: index out of range.`, [
    "Call docx_outline to re-map anchors.",
  ]);
}

function anchorStaleError(anchor: string): ToolError {
  return new ToolError(
    "anchor_stale",
    `Anchor ${anchor} is stale: the hash no longer matches the paragraph content.`,
    ["Call docx_read {anchor, window} and retry with the fresh anchor."],
  );
}

/** Body paragraphs only, indexable by `ordinal - 1`. */
function paragraphEntries(doc: DocHandle): AnchorEntry[] {
  return doc.anchorIndex().filter((e) => e.kind === "p");
}

function entryAt(entries: readonly AnchorEntry[], ordinal: number, label: string): AnchorEntry {
  const entry = entries[ordinal - 1];
  if (entry === undefined) throw anchorNotFoundError(label);
  return entry;
}

/** §6a validation order: parse → ordinal in range → hash match. */
function requireParagraph(entries: readonly AnchorEntry[], anchor: string): AnchorEntry {
  const m = FULL_ANCHOR_RE.exec(anchor);
  if (!m) throw anchorInvalidError(`Malformed anchor string: ${anchor}.`);
  const entry = entryAt(entries, Number(m[1]), anchor);
  if (entry.anchor !== anchor) throw anchorStaleError(anchor);
  return entry;
}

function commit(doc: DocHandle, xml: string): void {
  doc.pkg.setPart(doc.documentPartName, xml);
  doc.invalidate();
}

function freshParagraphAnchor(doc: DocHandle, ordinal: number): string {
  return entryAt(paragraphEntries(doc), ordinal, `P${ordinal}`).anchor;
}

// ---------------------------------------------------------------------------
// docx_replace (§4 coalescing, §5 tracking, §6a result shape)
// ---------------------------------------------------------------------------

export interface DocxReplaceArgs {
  doc_id: string;
  anchor?: string | undefined;
  old: string;
  new: string;
  all?: boolean | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxReplaceResult {
  n_replaced: number;
  new_anchor?: string;
  anchors?: string[];
}

export function docxReplace(session: Session, args: DocxReplaceArgs): DocxReplaceResult {
  const doc = session.get(args.doc_id);
  let xml = doc.documentXml();
  const entries = paragraphEntries(doc);
  // Anchor hash validated FIRST (§1) — before any matching or argument checks.
  const target = args.anchor != null ? requireParagraph(entries, args.anchor) : null;
  if (!args.old) {
    throw new ToolError("not_found", "Text not found: the search text is empty.", [
      "Provide non-empty old text.",
    ]);
  }
  const all = args.all === true;
  const tracked = args.track_changes === true;
  const ordinals = target !== null ? [target.ordinal] : entries.map((e) => e.ordinal);

  // §6a: matches are non-overlapping, left-to-right, found per paragraph
  // before any splice. §5 ids are allocated in emission (document) order.
  interface Job {
    ordinal: number;
    s: number;
    e: number;
    delId: number;
    insId: number | null;
  }
  const jobs: Job[] = [];
  let next = maxRevisionId(xml);
  for (const ordinal of ordinals) {
    const entry = entries[ordinal - 1] as AnchorEntry;
    const { text } = textWithMap(xml, entry.block);
    for (const [s, e] of findOccurrences(text, args.old)) {
      jobs.push({
        ordinal,
        s,
        e,
        delId: tracked ? ++next : 0,
        insId: tracked && args.new !== "" ? ++next : null,
      });
    }
  }
  if (!all) {
    if (jobs.length === 0) {
      throw new ToolError("not_found", `Text not found: ${args.old}.`, [
        "Broaden the query; check the projection for the exact text.",
      ]);
    }
    if (jobs.length > 1) {
      throw new ToolError(
        "ambiguous_target",
        `${args.old} matches ${jobs.length} times without all: true.`,
        ["Add all: true or narrow with an anchor."],
      );
    }
  }
  if (jobs.length === 0) return { n_replaced: 0, anchors: [] }; // idempotent (§6a)

  const author = revisionAuthor(args.author);
  const date = revisionDate();
  // Apply matches in reverse document order so earlier offsets stay valid.
  for (const j of [...jobs].reverse()) {
    const block = paragraphBlock(xml, j.ordinal);
    xml = tracked
      ? applyTrackedReplace(xml, block, j.s, j.e, args.new, author, date, {
          delId: j.delId,
          insId: j.insId,
        })
      : applyPlainMatch(xml, block, j.s, j.e, args.new);
  }
  commit(doc, xml);
  const affected = [...new Set(jobs.map((j) => j.ordinal))].sort((a, b) => a - b);
  if (all) {
    // §6a: all → anchors of affected paragraphs ascending; otherwise new_anchor.
    return { n_replaced: jobs.length, anchors: affected.map((o) => freshParagraphAnchor(doc, o)) };
  }
  const ordinal = target !== null ? target.ordinal : (affected[0] as number);
  return { n_replaced: jobs.length, new_anchor: freshParagraphAnchor(doc, ordinal) };
}

// ---------------------------------------------------------------------------
// docx_edit_paragraph (§6 LCS word diff, §6a)
// ---------------------------------------------------------------------------

export interface DocxEditParagraphArgs {
  doc_id: string;
  anchor: string;
  text: string;
  track_changes?: boolean | undefined;
  author?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxEditParagraphResult {
  new_anchor: string;
  diff: string;
}

export function docxEditParagraph(
  session: Session,
  args: DocxEditParagraphArgs,
): DocxEditParagraphResult {
  const doc = session.get(args.doc_id);
  const xml = doc.documentXml();
  const entry = requireParagraph(paragraphEntries(doc), args.anchor); // hash validated FIRST
  const { text: raw } = textWithMap(xml, entry.block);
  const ops = wordDiff(diffUnits(raw), diffUnits(args.text));
  const blocks = diffBlocks(ops);
  const newXml = rebuildParagraph(xml, entry.block, blocks, {
    tracked: args.track_changes === true,
    author: revisionAuthor(args.author),
    date: revisionDate(),
  });
  commit(doc, newXml);
  // §6a: n = max(#deleted units, #inserted units); noun singular when n is 1.
  const changed = Math.max(
    ops.filter((o) => o.op === "del").length,
    ops.filter((o) => o.op === "ins").length,
  );
  return {
    new_anchor: freshParagraphAnchor(doc, entry.ordinal),
    diff: `~${plural(changed, "word")} changed`,
  };
}

// ---------------------------------------------------------------------------
// docx_insert (§6a minimal markdown)
// ---------------------------------------------------------------------------

export interface DocxInsertArgs {
  doc_id: string;
  after?: string | undefined;
  before?: string | undefined;
  content: string;
  style?: string | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxInsertResult {
  new_anchors: string[];
}

interface MdBlock {
  text: string;
  style?: string | undefined;
}

const MD_HEADING_RE = /^(#{1,9}) /;

/** §6a minimal markdown: one paragraph per non-blank line. */
export function parseMinimalMarkdown(content: string): MdBlock[] {
  const out: MdBlock[] = [];
  for (const raw of content.split("\n")) {
    const line = raw.endsWith("\r") ? raw.slice(0, -1) : raw;
    if (line.replace(WS_RUN_RE, "") === "") continue;
    const h = MD_HEADING_RE.exec(line);
    if (h) {
      const level = (h[1] as string).length;
      out.push({ text: line.slice(level + 1), style: `Heading${level}` });
    } else if (line.startsWith("- ") || line.startsWith("* ")) {
      out.push({ text: line.slice(2), style: "ListParagraph" }); // numPr wiring is Phase 2
    } else {
      out.push({ text: line });
    }
  }
  return out;
}

/** §6a: the styleId verbatim if defined, else with whitespace removed, else error. */
function resolveStyleId(doc: DocHandle, style: string): string {
  const styles = parseStyles(
    doc.pkg.has("word/styles.xml") ? doc.pkg.partText("word/styles.xml") : undefined,
  );
  if (styles.has(style)) return style;
  const compact = style.replace(WS_RUN_RE, "");
  if (styles.has(compact)) return compact;
  throw new ToolError("style_unknown", `Named style ${style} does not exist.`, [
    'Call docx_style {op: "list"} to see available styles.',
  ]);
}

export function docxInsert(session: Session, args: DocxInsertArgs): DocxInsertResult {
  const doc = session.get(args.doc_id);
  if ((args.after == null) === (args.before == null)) {
    throw anchorInvalidError("Provide exactly one of after or before.");
  }
  const entries = paragraphEntries(doc);
  const entry = requireParagraph(entries, (args.after ?? args.before) as string); // hash FIRST
  const paragraphs = parseMinimalMarkdown(args.content);
  const styleId = args.style != null ? resolveStyleId(doc, args.style) : null;
  if (paragraphs.length === 0) return { new_anchors: [] };

  const xml = doc.documentXml();
  const tracked = args.track_changes === true;
  const author = revisionAuthor(args.author);
  const date = revisionDate();
  let revId = tracked ? maxRevisionId(xml) + 1 : 0;
  const pieces: string[] = [];
  for (const b of paragraphs) {
    const effective = styleId ?? b.style;
    const pPr =
      effective != null ? `<w:pPr><w:pStyle w:val="${escapeAttr(effective)}"/></w:pPr>` : "";
    let run = b.text !== "" ? `<w:r>${emitTextElement("w:t", b.text)}</w:r>` : "";
    if (tracked && run !== "") {
      run = revisionOpen("ins", revId++, author, date) + run + "</w:ins>";
    }
    pieces.push(`<w:p>${pPr}${run}</w:p>`);
  }
  const position = args.after != null ? entry.block.end : entry.block.start;
  commit(doc, splice(xml, position, position, pieces.join("")));
  const base = args.after != null ? entry.ordinal + 1 : entry.ordinal;
  const fresh = paragraphEntries(doc);
  return {
    new_anchors: paragraphs.map((_, i) => entryAt(fresh, base + i, `P${base + i}`).anchor),
  };
}

// ---------------------------------------------------------------------------
// docx_delete (§6a)
// ---------------------------------------------------------------------------

export interface DocxDeleteArgs {
  doc_id: string;
  anchor?: string | undefined;
  range?: string | undefined;
  track_changes?: boolean | undefined;
  author?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxDeleteResult {
  ok: boolean;
  deleted: number;
}

export function docxDelete(session: Session, args: DocxDeleteArgs): DocxDeleteResult {
  const doc = session.get(args.doc_id);
  if ((args.anchor == null) === (args.range == null)) {
    throw anchorInvalidError("Provide exactly one of anchor or range.");
  }
  const entries = paragraphEntries(doc);
  let targets: AnchorEntry[];
  if (args.anchor != null) {
    targets = [requireParagraph(entries, args.anchor)]; // hash validated FIRST
  } else {
    const range = args.range as string;
    const m = RANGE_RE.exec(range);
    if (!m) throw anchorInvalidError(`Malformed range string: ${range}.`);
    const start = Number(m[1]);
    const end = Number(m[3]);
    if (start > end) throw anchorInvalidError(`Inverted range: ${range}.`);
    const endpoints: [number, string | undefined][] = [
      [start, m[2]],
      [end, m[4]],
    ];
    for (const [ordinal, hash] of endpoints) {
      const entry = entryAt(entries, ordinal, `P${ordinal}`);
      if (hash !== undefined && entry.anchor !== `P${ordinal}#${hash}`) {
        throw anchorStaleError(`P${ordinal}#${hash}`);
      }
    }
    // The range deletes paragraphs a..b only; body tables between them are
    // untouched (table ops are Phase 2).
    targets = entries.slice(start - 1, end);
  }
  const xml = doc.documentXml();
  let newXml: string;
  if (args.track_changes !== true) {
    newXml = spliceAll(
      xml,
      targets.map((t) => ({ start: t.block.start, end: t.block.end, text: "" })),
    );
  } else {
    // §6a: one w:del per non-empty paragraph wrapping the full run content
    // after w:pPr (w:t renamed to w:delText; the paragraph mark survives).
    const author = revisionAuthor(args.author);
    const date = revisionDate();
    let revId = maxRevisionId(xml) + 1;
    const edits: SpliceEdit[] = [];
    for (const t of targets) {
      const p = t.block;
      if (p.selfClosed) continue; // counted, but nothing to wrap
      const pPr = childElements(xml, p.contentStart, p.contentEnd).find((k) => k.name === "w:pPr");
      const contentStart = pPr !== undefined ? pPr.end : p.contentStart;
      if (contentStart >= p.contentEnd) continue;
      edits.push({
        start: contentStart,
        end: p.contentEnd,
        text:
          revisionOpen("del", revId++, author, date) +
          tToDelText(xml.slice(contentStart, p.contentEnd)) +
          "</w:del>",
      });
    }
    newXml = spliceAll(xml, edits);
  }
  commit(doc, newXml);
  return { ok: true, deleted: targets.length };
}

// ---------------------------------------------------------------------------
// docx_revision (§7, §6a)
// ---------------------------------------------------------------------------

const REVISION_OPS = new Set(["list", "accept", "reject", "accept_all", "reject_all"]);

export interface DocxRevisionFilter extends RevisionFilter {}

export interface DocxRevisionArgs {
  doc_id: string;
  op: "list" | "accept" | "reject" | "accept_all" | "reject_all";
  id?: string | undefined;
  filter?: DocxRevisionFilter | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface RevisionEntry {
  id: string;
  type: "ins" | "del";
  author: string;
  date: string;
  anchor?: string;
  text: string;
}

export interface DocxRevisionResult {
  revisions?: RevisionEntry[];
  accepted?: number;
  rejected?: number;
  remaining_by_author?: Record<string, number>;
  anchors?: string[];
}

export function docxRevision(session: Session, args: DocxRevisionArgs): DocxRevisionResult {
  const doc = session.get(args.doc_id);
  if (!REVISION_OPS.has(args.op)) {
    throw new ToolError("not_found", `Unknown revision op: ${String(args.op)}.`, [
      "Use list, accept, reject, accept_all, or reject_all.",
    ]);
  }
  const blocks = doc.anchorIndex();
  let xml = doc.documentXml();
  const revisions = scanRevisions(xml, blocks);
  const flt: RevisionFilter = args.filter ?? {};

  if (args.op === "list") {
    const listed: RevisionEntry[] = [];
    for (const rev of revisions) {
      if (!revisionMatches(rev, flt)) continue;
      listed.push({
        id: rev.id,
        type: rev.kind,
        author: rev.author,
        date: rev.date,
        ...(rev.anchor !== null ? { anchor: rev.anchor } : {}),
        text: rev.text,
      });
    }
    return { revisions: listed };
  }

  const accept = args.op === "accept" || args.op === "accept_all";
  let selected: Revision[];
  if (args.op === "accept_all" || args.op === "reject_all") {
    selected = [...revisions]; // §6a: _all ops ignore id/filter
  } else if (args.id != null) {
    // An id selecting nothing resolves nothing (§7 idempotency), not an error.
    selected = revisions.filter((rev) => rev.id === args.id);
  } else {
    selected = revisions.filter((rev) => revisionMatches(rev, flt));
  }
  // A candidate nested inside another candidate resolves with its container (§6a).
  const candidates = selected.filter(
    (rev) => !selected.some((other) => other.el.start < rev.el.start && rev.el.end < other.el.end),
  );
  const ordinals = [
    ...new Set(candidates.map((rev) => rev.ordinal).filter((o): o is number => o !== null)),
  ].sort((a, b) => a - b);
  if (candidates.length > 0) {
    xml = resolveRevisions(xml, candidates, accept);
    for (const ordinal of ordinals) xml = mergeParagraphRuns(xml, ordinal); // §7 post-pass
    commit(doc, xml);
  }
  const fresh = paragraphEntries(doc);
  const remaining = scanRevisions(doc.documentXml(), doc.anchorIndex());
  const byAuthor = new Map<string, number>();
  for (const rev of remaining) byAuthor.set(rev.author, (byAuthor.get(rev.author) ?? 0) + 1);
  const remainingByAuthor: Record<string, number> = {};
  for (const [author, n] of [...byAuthor].sort((a, b) =>
    a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0,
  )) {
    remainingByAuthor[author] = n;
  }
  const anchors = ordinals.map((o) => entryAt(fresh, o, `P${o}`).anchor);
  return accept
    ? { accepted: candidates.length, remaining_by_author: remainingByAuthor, anchors }
    : { rejected: candidates.length, remaining_by_author: remainingByAuthor, anchors };
}
