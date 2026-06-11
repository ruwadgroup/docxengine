/**
 * Comments (`docx_comment`) per spec/algorithms.md §18.
 *
 * `add` wires all five package places: the `w:commentRangeStart`/`End`
 * markers and the `w:commentReference` run around the anchor paragraph's
 * runs, the `w:comment` in `word/comments.xml` (created on demand with its
 * content-type Override + document rel), and the ensured `CommentReference`
 * style. `reply`/`resolve` use the modern w15 `word/commentsExtended.xml`
 * thread metadata (`commentEx` + `w14:paraId`); `delete` removes all five
 * places for an id and its replies; `list` returns one entry per thread root.
 *
 * The Python twin (`_comments.py`) is the byte-parity reference.
 */
import { ToolError } from "./errors.js";
import { revisionAuthor, revisionDate } from "./edits.js";
import {
  COMMENT_REFERENCE_STYLE,
  addRelationship,
  anchorInvalid,
  ensureContentOverride,
  ensureStyle,
} from "./phase2common.js";
import type { DocHandle, Session } from "./session.js";
import {
  type ElementSlice,
  attrs,
  bodyBlocks,
  childElements,
  elementExtent,
  emitTextElement,
  escapeAttr,
  nextTag,
  scopeText,
  splice,
} from "./xmlscan.js";
import { normalizedText, paragraphAnchor } from "./anchors.js";
import { sha256Hex } from "./sha256.js";

const utf8 = new TextEncoder();

type ResponseFormat = "concise" | "detailed";

const COMMENTS_PART = "word/comments.xml";
const COMMENTS_EXT_PART = "word/commentsExtended.xml";
const COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXT_CT =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
const COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_EXT_REL = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";

const FULL_ANCHOR_RE = /^P([1-9][0-9]*)#([0-9a-f]{4})$/;
const COMMENT_ID_RE = /^C([0-9]+)$/;

// ---------------------------------------------------------------------------
// Anchor resolution (paragraph; §6a validation order)
// ---------------------------------------------------------------------------

function paragraphEntries(doc: DocHandle) {
  return doc.anchorIndex().filter((e) => e.kind === "p");
}

function requireParagraph(doc: DocHandle, anchor: string) {
  const m = FULL_ANCHOR_RE.exec(anchor);
  if (!m) throw anchorInvalid(`Malformed anchor string: ${anchor}.`);
  const entry = paragraphEntries(doc)[Number(m[1]) - 1];
  if (entry === undefined) {
    throw new ToolError("anchor_stale", `Anchor ${anchor} not found: index out of range.`, [
      "Call docx_outline to re-map anchors.",
    ]);
  }
  if (entry.anchor !== anchor) {
    throw new ToolError(
      "anchor_stale",
      `Anchor ${anchor} is stale: the hash no longer matches the paragraph content.`,
      ["Call docx_read {anchor, window} and retry with the fresh anchor."],
    );
  }
  return entry;
}

/** §18 initials: uppercased first letter of each whitespace-separated word. */
export function commentInitials(author: string): string {
  let out = "";
  for (const word of author.split(/\s+/)) {
    if (word.length > 0) out += word[0]!.toUpperCase();
  }
  return out;
}

// ---------------------------------------------------------------------------
// comments.xml model
// ---------------------------------------------------------------------------

interface CommentRecord {
  id: number;
  author: string;
  date: string;
  text: string;
  /** w14:paraId of the comment's first w:p (for threading), or "". */
  paraId: string;
  el: ElementSlice;
}

function emptyCommentsXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"/>`
  );
}

function emptyCommentsExtXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w15:commentsEx xmlns:w15="${W15_NS}"/>`
  );
}

/** Concatenated w:t text of a w:comment element (raw, document order). */
function commentText(xml: string, el: ElementSlice): string {
  return scopeText(xml, el);
}

/** w14:paraId of the first w:p inside a w:comment element ("" when absent). */
function firstParaId(xml: string, el: ElementSlice): string {
  if (el.selfClosed) return "";
  let i = el.contentStart;
  for (;;) {
    const tag = nextTag(xml, i, el.contentEnd);
    if (!tag) return "";
    if (tag.name === "w:p" && tag.kind !== "end") {
      return attrs(xml, tag)["w14:paraId"] ?? "";
    }
    i = tag.end;
  }
}

function parseComments(xml: string | undefined): CommentRecord[] {
  const out: CommentRecord[] = [];
  if (xml === undefined) return out;
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return out;
    if (t.name === "w:comment" && t.kind !== "end") {
      const el = elementExtent(xml, t);
      const a = attrs(xml, t);
      out.push({
        id: Number(a["w:id"] ?? "0"),
        author: a["w:author"] ?? "unknown",
        date: a["w:date"] ?? "",
        text: commentText(xml, el),
        paraId: firstParaId(xml, el),
        el,
      });
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

function maxCommentId(records: CommentRecord[]): number {
  let max = -1;
  for (const r of records) max = Math.max(max, r.id);
  return max;
}

/**
 * Allocate a fresh w14:paraId (§18). Word uses random ids, but for
 * cross-language byte parity DocxEngine derives one deterministically: the
 * first 8 uppercase hex chars of SHA-256 over the UTF-8 of `paraId:{id}:{text}`
 * (algorithms.md §18). The Python twin must use the identical derivation.
 */
export function makeParaId(id: number, text: string): string {
  return sha256Hex(utf8.encode(`paraId:${id}:${text}`))
    .slice(0, 8)
    .toUpperCase();
}

// ---------------------------------------------------------------------------
// commentsExtended.xml model (w15)
// ---------------------------------------------------------------------------

interface CommentExRecord {
  paraId: string;
  paraIdParent: string | null;
  done: boolean;
  el: ElementSlice;
}

function parseCommentsEx(xml: string | undefined): CommentExRecord[] {
  const out: CommentExRecord[] = [];
  if (xml === undefined) return out;
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return out;
    if (t.name === "w15:commentEx" && t.kind !== "end") {
      const el = elementExtent(xml, t);
      const a = attrs(xml, t);
      out.push({
        paraId: a["w15:paraId"] ?? "",
        paraIdParent: a["w15:paraIdParent"] ?? null,
        done: (a["w15:done"] ?? "0") === "1",
        el,
      });
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

function ensureCommentsExtPart(doc: DocHandle): void {
  if (doc.pkg.has(COMMENTS_EXT_PART)) return;
  doc.pkg.setPart(COMMENTS_EXT_PART, emptyCommentsExtXml());
  ensureContentOverride(doc.pkg, COMMENTS_EXT_PART, COMMENTS_EXT_CT);
  addRelationship(doc.pkg, doc.documentPartName, COMMENTS_EXT_REL, "commentsExtended.xml");
}

function ensureCommentsPart(doc: DocHandle): void {
  if (doc.pkg.has(COMMENTS_PART)) return;
  doc.pkg.setPart(COMMENTS_PART, emptyCommentsXml());
  ensureContentOverride(doc.pkg, COMMENTS_PART, COMMENTS_CT);
  addRelationship(doc.pkg, doc.documentPartName, COMMENTS_REL, "comments.xml");
}

// ---------------------------------------------------------------------------
// docx_comment
// ---------------------------------------------------------------------------

export interface DocxCommentReply {
  author: string;
  date: string;
  text: string;
}

export interface DocxCommentEntry {
  id: string;
  anchor: string;
  author: string;
  date: string;
  text: string;
  resolved: boolean;
  replies: DocxCommentReply[];
}

export interface DocxCommentArgs {
  doc_id: string;
  op: "add" | "reply" | "resolve" | "list" | "delete";
  anchor?: string | undefined;
  comment_id?: string | undefined;
  text?: string | undefined;
  author?: string | undefined;
  response_format?: ResponseFormat | undefined;
}

export interface DocxCommentResult {
  comment_id?: string;
  anchor?: string;
  comments?: DocxCommentEntry[];
  note?: string;
}

export function docxComment(session: Session, args: DocxCommentArgs): DocxCommentResult {
  const doc = session.get(args.doc_id);
  switch (args.op) {
    case "add":
      return commentAdd(doc, args);
    case "reply":
      return commentReply(doc, args);
    case "resolve":
      return commentResolve(doc, args);
    case "delete":
      return commentDelete(doc, args);
    case "list":
      return commentList(doc);
    default:
      throw new ToolError("invalid_args", `docx_comment: unknown op ${String(args.op)}.`, []);
  }
}

function parseCommentId(value: string): number {
  const m = COMMENT_ID_RE.exec(value);
  if (!m) throw anchorInvalid(`Malformed comment id: ${value}.`);
  return Number(m[1]);
}

// --- add ---------------------------------------------------------------------

function emitComment(
  id: number,
  author: string,
  date: string,
  text: string,
  paraId: string,
): string {
  const initials = commentInitials(author);
  return (
    `<w:comment w:id="${id}" w:author="${escapeAttr(author)}" w:date="${escapeAttr(date)}"` +
    ` w:initials="${escapeAttr(initials)}">` +
    `<w:p w14:paraId="${paraId}"><w:r>${emitTextElement("w:t", text)}</w:r></w:p>` +
    `</w:comment>`
  );
}

/** Append a w:comment before </w:comments> (or expand a self-closed root). */
function appendComment(doc: DocHandle, commentXml: string): void {
  const xml = doc.pkg.partText(COMMENTS_PART);
  const close = xml.lastIndexOf("</w:comments>");
  if (close >= 0) {
    doc.pkg.setPart(COMMENTS_PART, splice(xml, close, close, commentXml));
    return;
  }
  const selfClose = xml.lastIndexOf("/>");
  doc.pkg.setPart(
    COMMENTS_PART,
    splice(xml, selfClose, selfClose + 2, `>${commentXml}</w:comments>`),
  );
}

function appendCommentEx(doc: DocHandle, entry: string): void {
  const xml = doc.pkg.partText(COMMENTS_EXT_PART);
  const close = xml.lastIndexOf("</w15:commentsEx>");
  if (close >= 0) {
    doc.pkg.setPart(COMMENTS_EXT_PART, splice(xml, close, close, entry));
    return;
  }
  const selfClose = xml.lastIndexOf("/>");
  doc.pkg.setPart(
    COMMENTS_EXT_PART,
    splice(xml, selfClose, selfClose + 2, `>${entry}</w15:commentsEx>`),
  );
}

function commentAdd(doc: DocHandle, args: DocxCommentArgs): DocxCommentResult {
  if (args.anchor == null) throw anchorInvalid("op 'add' requires anchor.");
  const entry = requireParagraph(doc, args.anchor); // hash FIRST
  const author = revisionAuthor(args.author);
  const date = revisionDate();
  const text = args.text ?? "";

  ensureCommentsPart(doc);
  ensureStyle(doc.pkg, "CommentReference", COMMENT_REFERENCE_STYLE);
  const id = maxCommentId(parseComments(doc.pkg.partText(COMMENTS_PART))) + 1;
  const paraId = makeParaId(id, text);

  // (4) the comment part entry.
  appendComment(doc, emitComment(id, author, date, text, paraId));

  // (1)/(2)/(3) the in-document markers around the paragraph's runs: start
  // before the runs (after any w:pPr), end + reference after them.
  const xml = doc.documentXml();
  const insertStart = runRegion(xml, entry.block).start;
  const insertEnd = runRegion(xml, entry.block).end;
  const start = `<w:commentRangeStart w:id="${id}"/>`;
  const end =
    `<w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
    `<w:commentReference w:id="${id}"/></w:r>`;
  // Splice end first (higher offset) so the start offset stays valid.
  let next = splice(xml, insertEnd, insertEnd, end);
  next = splice(next, insertStart, insertStart, start);
  doc.pkg.setPart(doc.documentPartName, next);
  doc.invalidate();

  return { comment_id: `C${id}`, anchor: args.anchor };
}

/**
 * The run region of a paragraph: [after any leading w:pPr, before the end
 * tag). The comment range brackets exactly the paragraph's run content.
 */
function runRegion(xml: string, p: ElementSlice): { start: number; end: number } {
  if (p.selfClosed) return { start: p.startTagEnd, end: p.startTagEnd };
  const kids = childElements(xml, p.contentStart, p.contentEnd);
  const pPr = kids.find((k) => k.name === "w:pPr");
  const start = pPr ? pPr.end : p.contentStart;
  return { start, end: p.contentEnd };
}

// --- reply -------------------------------------------------------------------

function commentReply(doc: DocHandle, args: DocxCommentArgs): DocxCommentResult {
  if (args.comment_id == null) throw anchorInvalid("op 'reply' requires comment_id.");
  const targetId = parseCommentId(args.comment_id);
  if (!doc.pkg.has(COMMENTS_PART)) throw notFound(args.comment_id);
  const records = parseComments(doc.pkg.partText(COMMENTS_PART));
  const root = records.find((r) => r.id === targetId);
  if (root === undefined) throw notFound(args.comment_id);

  const author = revisionAuthor(args.author);
  const date = revisionDate();
  const text = args.text ?? "";
  const id = maxCommentId(records) + 1;
  const paraId = makeParaId(id, text);

  appendComment(doc, emitComment(id, author, date, text, paraId));
  ensureCommentsExtPart(doc);
  appendCommentEx(
    doc,
    `<w15:commentEx w15:paraId="${paraId}" w15:paraIdParent="${root.paraId}" w15:done="0"/>`,
  );

  return { comment_id: `C${id}` };
}

// --- resolve -----------------------------------------------------------------

function commentResolve(doc: DocHandle, args: DocxCommentArgs): DocxCommentResult {
  if (args.comment_id == null) throw anchorInvalid("op 'resolve' requires comment_id.");
  const targetId = parseCommentId(args.comment_id);
  if (!doc.pkg.has(COMMENTS_PART)) throw notFound(args.comment_id);
  const records = parseComments(doc.pkg.partText(COMMENTS_PART));
  const root = records.find((r) => r.id === targetId);
  if (root === undefined) throw notFound(args.comment_id);

  ensureCommentsExtPart(doc);
  const exXml = doc.pkg.partText(COMMENTS_EXT_PART);
  const exRecords = parseCommentsEx(exXml);
  const existing = exRecords.find((e) => e.paraId === root.paraId);
  if (existing) {
    // Set w15:done="1" on the root's commentEx (replace the attr or add it).
    const tagXml = exXml.slice(existing.el.start, existing.el.end);
    let updated: string;
    if (/w15:done\s*=\s*"[^"]*"/.test(tagXml)) {
      updated = tagXml.replace(/w15:done\s*=\s*"[^"]*"/, 'w15:done="1"');
    } else {
      updated = tagXml.replace(/\/?>$/, ' w15:done="1"/>');
    }
    doc.pkg.setPart(COMMENTS_EXT_PART, splice(exXml, existing.el.start, existing.el.end, updated));
  } else {
    appendCommentEx(doc, `<w15:commentEx w15:paraId="${root.paraId}" w15:done="1"/>`);
  }
  return { comment_id: args.comment_id, note: "Comment thread resolved." };
}

// --- delete ------------------------------------------------------------------

function commentDelete(doc: DocHandle, args: DocxCommentArgs): DocxCommentResult {
  if (args.comment_id == null) throw anchorInvalid("op 'delete' requires comment_id.");
  const targetId = parseCommentId(args.comment_id);
  if (!doc.pkg.has(COMMENTS_PART)) throw notFound(args.comment_id);
  const commentsXml = doc.pkg.partText(COMMENTS_PART);
  const records = parseComments(commentsXml);
  const root = records.find((r) => r.id === targetId);
  if (root === undefined) throw notFound(args.comment_id);

  // Collect the thread: the root plus its replies (w15 paraIdParent chain).
  const exRecords = doc.pkg.has(COMMENTS_EXT_PART)
    ? parseCommentsEx(doc.pkg.partText(COMMENTS_EXT_PART))
    : [];
  const threadParaIds = new Set<string>([root.paraId]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const ex of exRecords) {
      if (ex.paraIdParent && threadParaIds.has(ex.paraIdParent) && !threadParaIds.has(ex.paraId)) {
        threadParaIds.add(ex.paraId);
        changed = true;
      }
    }
  }
  const removeIds = new Set<number>();
  for (const r of records) {
    if (r.id === targetId || threadParaIds.has(r.paraId)) removeIds.add(r.id);
  }

  // (4) Remove the w:comment entries (descending offset to keep offsets valid).
  let cx = commentsXml;
  const toRemove = records
    .filter((r) => removeIds.has(r.id))
    .sort((a, b) => b.el.start - a.el.start);
  for (const r of toRemove) cx = splice(cx, r.el.start, r.el.end, "");
  doc.pkg.setPart(COMMENTS_PART, cx);

  // commentsExtended: drop the thread's commentEx entries.
  if (doc.pkg.has(COMMENTS_EXT_PART)) {
    let ex = doc.pkg.partText(COMMENTS_EXT_PART);
    const exToRemove = parseCommentsEx(ex)
      .filter((e) => threadParaIds.has(e.paraId))
      .sort((a, b) => b.el.start - a.el.start);
    for (const e of exToRemove) ex = splice(ex, e.el.start, e.el.end, "");
    doc.pkg.setPart(COMMENTS_EXT_PART, ex);
  }

  // (1)/(2)/(3) Remove the in-document markers for every removed id.
  const docXml = doc.documentXml();
  const removed = removeDocumentMarkers(docXml, removeIds);
  doc.pkg.setPart(doc.documentPartName, removed);
  doc.invalidate();

  return {
    comment_id: args.comment_id,
    note: `Deleted comment ${args.comment_id} and its replies.`,
  };
}

/** Remove commentRangeStart/End and the reference run for the given ids. */
function removeDocumentMarkers(xml: string, ids: Set<number>): string {
  const cuts: { start: number; end: number }[] = [];
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) break;
    if ((t.name === "w:commentRangeStart" || t.name === "w:commentRangeEnd") && t.kind !== "end") {
      const id = Number(attrs(xml, t)["w:id"] ?? "-1");
      if (ids.has(id)) cuts.push({ start: t.start, end: t.end });
      i = t.end;
      continue;
    }
    if (t.name === "w:r" && t.kind === "start") {
      const el = elementExtent(xml, t);
      const ref = findCommentReference(xml, el);
      if (ref !== null && ids.has(ref)) {
        cuts.push({ start: el.start, end: el.end });
        i = el.end;
        continue;
      }
    }
    i = t.end;
  }
  cuts.sort((a, b) => b.start - a.start);
  let out = xml;
  for (const c of cuts) out = splice(out, c.start, c.end, "");
  return out;
}

/** The w:id of a w:commentReference inside a run, or null. */
function findCommentReference(xml: string, run: ElementSlice): number | null {
  let i = run.contentStart;
  for (;;) {
    const t = nextTag(xml, i, run.contentEnd);
    if (!t) return null;
    if (t.name === "w:commentReference" && t.kind !== "end") {
      const id = attrs(xml, t)["w:id"];
      return id !== undefined ? Number(id) : null;
    }
    i = t.end;
  }
}

// --- list --------------------------------------------------------------------

function commentList(doc: DocHandle): DocxCommentResult {
  const records = parseComments(
    doc.pkg.has(COMMENTS_PART) ? doc.pkg.partText(COMMENTS_PART) : undefined,
  );
  const exRecords = doc.pkg.has(COMMENTS_EXT_PART)
    ? parseCommentsEx(doc.pkg.partText(COMMENTS_EXT_PART))
    : [];
  const exByParaId = new Map<string, CommentExRecord>();
  for (const e of exRecords) exByParaId.set(e.paraId, e);

  // A reply is any comment whose commentEx has a paraIdParent; everything else
  // is a thread root.
  const replyParaIds = new Set<string>();
  const parentOf = new Map<string, string>();
  for (const e of exRecords) {
    if (e.paraIdParent) {
      replyParaIds.add(e.paraId);
      parentOf.set(e.paraId, e.paraIdParent);
    }
  }
  const byParaId = new Map<string, CommentRecord>();
  for (const r of records) if (r.paraId) byParaId.set(r.paraId, r);

  // Map each comment to its thread root by walking paraIdParent.
  function rootOf(r: CommentRecord): CommentRecord {
    let cur = r;
    const seen = new Set<string>();
    for (;;) {
      const parent = parentOf.get(cur.paraId);
      if (parent === undefined || seen.has(cur.paraId)) return cur;
      seen.add(cur.paraId);
      const next = byParaId.get(parent);
      if (next === undefined) return cur;
      cur = next;
    }
  }

  const rangeAnchors = commentRangeAnchors(doc);
  const entries: DocxCommentEntry[] = [];
  for (const r of records) {
    if (r.paraId && replyParaIds.has(r.paraId)) continue; // a reply, not a root
    const replies: DocxCommentReply[] = [];
    for (const cand of records) {
      if (cand.id === r.id) continue;
      if (cand.paraId && replyParaIds.has(cand.paraId) && rootOf(cand).id === r.id) {
        replies.push({ author: cand.author, date: cand.date, text: cand.text });
      }
    }
    const resolved = r.paraId ? (exByParaId.get(r.paraId)?.done ?? false) : false;
    entries.push({
      id: `C${r.id}`,
      anchor: rangeAnchors.get(r.id) ?? "",
      author: r.author,
      date: r.date,
      text: r.text,
      resolved,
      replies,
    });
  }
  return { comments: entries };
}

/** comment id → body anchor of its commentRangeStart's containing paragraph. */
function commentRangeAnchors(doc: DocHandle): Map<number, string> {
  const out = new Map<number, string>();
  const xml = doc.documentXml();
  let pOrd = 0;
  for (const block of bodyBlocks(xml)) {
    if (block.name !== "w:p") continue;
    pOrd++;
    // Scan for a commentRangeStart inside this paragraph.
    let i = block.contentStart;
    for (;;) {
      const t = nextTag(xml, i, block.contentEnd);
      if (!t) break;
      if (t.name === "w:commentRangeStart" && t.kind !== "end") {
        const id = Number(attrs(xml, t)["w:id"] ?? "-1");
        if (!out.has(id)) {
          const normalized = normalizedText(scopeText(xml, block));
          out.set(id, paragraphAnchor(pOrd, normalized));
        }
      }
      i = t.end;
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

function notFound(commentId: string): ToolError {
  return new ToolError("not_found", `Comment ${commentId} does not exist.`, [
    'Call docx_comment {op: "list"} to see available comment ids.',
  ]);
}
