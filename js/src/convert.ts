/**
 * Convert (`docx_convert`) per spec/algorithms.md §23.
 *
 * `md`/`html` are produced in-engine from the §2 projection model: headings
 * `#`×level, ordered/unordered list items indented per ilvl, GitHub tables,
 * `**bold**`/`*italic*` reconstructed from run `w:b`/`w:i`, comments inline as
 * `<!-- comment:{author}: {text} -->`, revisions in accepted view (ins shown,
 * del omitted) with `[ins]`/`[del]` markers configurable off. `md` keeps
 * `&`,`<`,`>` literal; `html` HTML-escapes and adds inline styles for alignment
 * and color. `pdf`/`png` go through the §24 render adapter.
 *
 * Cross-language parity is conformance-tested on `md`/`html` content — this
 * module mirrors `_convert.py` to the byte. `pdf`/`png` parity is not required.
 */
import { normalizedText } from "./anchors.js";
import { ToolError } from "./errors.js";
import {
  bodyProjBlocks,
  headingLevel,
  projectionContext,
  tableDims,
  type ProjectionContext,
} from "./projector.js";
import { renderToFile } from "./render.js";
import type { DocHandle, Session } from "./session.js";
import {
  type ElementSlice,
  childElements,
  decodeEntities,
  elementExtent,
  findElement,
  getAttr,
  nextTag,
  scopeText,
} from "./xmlscan.js";

// ---------------------------------------------------------------------------
// Run-level model for one paragraph (accepted view)
// ---------------------------------------------------------------------------

interface ConvRun {
  text: string;
  bold: boolean;
  italic: boolean;
  /** A tracked-insertion run (for the `[ins]` marker). */
  ins: boolean;
}

interface ParaModel {
  /** Heading level (1-9) or null. */
  heading: number | null;
  /** List annotation: {kind, level} or null. */
  list: { kind: "ol" | "ul"; level: number } | null;
  alignment: string | null;
  color: string | null;
  runs: ConvRun[];
  /** Inline comment notes appended at range ends: `{author}: {text}`. */
  comments: { author: string; text: string }[];
  /** True iff the paragraph contains any tracked deletion. */
  hasDeletion: boolean;
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

/** styleId of a paragraph's w:pStyle, or undefined. */
function paragraphStyleId(xml: string, p: ElementSlice): string | undefined {
  const pPr = findElement(xml, "w:pPr", p.contentStart, p.contentEnd);
  if (!pPr || pPr.selfClosed) return undefined;
  const pStyle = findElement(xml, "w:pStyle", pPr.contentStart, pPr.contentEnd);
  return pStyle ? getAttr(xml, tagOf(pStyle), "w:val") : undefined;
}

/** Parse one body paragraph into the §23 conversion model (accepted view). */
function parseParagraph(
  xml: string,
  p: ElementSlice,
  ctx: ProjectionContext,
  commentAuthors: Map<string, string>,
): ParaModel {
  const styleId = paragraphStyleId(xml, p);
  const heading = headingLevel(styleId, ctx.styles);
  const list = listAnnotation(xml, p, ctx);
  const { alignment, color } = paragraphFormatting(xml, p);

  const runs: ConvRun[] = [];
  const comments: { author: string; text: string }[] = [];
  let hasDeletion = false;
  // Walk the paragraph in document order, tracking ins/del wrappers.
  let insDepth = 0;
  let delDepth = 0;
  let i = p.contentStart;
  for (;;) {
    const t = nextTag(xml, i, p.contentEnd);
    if (!t) break;
    if (t.name === "w:ins") {
      if (t.kind === "start") insDepth++;
      else if (t.kind === "end") insDepth = Math.max(0, insDepth - 1);
      i = t.end;
      continue;
    }
    if (t.name === "w:del") {
      if (t.kind === "start") {
        delDepth++;
        hasDeletion = true;
      } else if (t.kind === "end") delDepth = Math.max(0, delDepth - 1);
      i = t.end;
      continue;
    }
    if (t.name === "w:commentReference" && t.kind !== "end") {
      const id = getAttr(xml, t, "w:id") ?? "?";
      // Comment text is resolved from comments.xml lazily by the caller's map.
      comments.push({ author: commentAuthors.get(id) ?? "unknown", text: commentTextFor(ctx, id) });
      i = t.kind === "empty" ? t.end : elementExtent(xml, t, p.contentEnd).end;
      continue;
    }
    if (t.name === "w:r" && t.kind === "start") {
      const run = elementExtent(xml, t, p.contentEnd);
      // delText is omitted in accepted view; only w:t shows.
      if (delDepth === 0) {
        const text = runText(xml, run);
        if (text !== "") {
          const { bold, italic } = runFormatting(xml, run);
          runs.push({ text, bold, italic, ins: insDepth > 0 });
        }
      }
      // A comment reference often sits inside its own run; capture it here so
      // the §23 inline note is emitted at the range end.
      const ref = findElement(xml, "w:commentReference", run.contentStart, run.contentEnd);
      if (ref) {
        const id = getAttr(xml, tagOf(ref), "w:id") ?? "?";
        comments.push({
          author: commentAuthors.get(id) ?? "unknown",
          text: commentTextFor(ctx, id),
        });
      }
      i = run.end;
      continue;
    }
    i = t.end;
  }
  return { heading, list, alignment, color, runs, comments, hasDeletion };
}

/** Concatenated w:t text of a run (decoded; delText excluded). */
function runText(xml: string, run: ElementSlice): string {
  let s = "";
  let i = run.contentStart;
  for (;;) {
    const t = nextTag(xml, i, run.contentEnd);
    if (!t) break;
    if (t.name === "w:t" && t.kind !== "end") {
      const el = elementExtent(xml, t, run.contentEnd);
      if (!el.selfClosed) s += decodeEntities(xml.slice(el.contentStart, el.contentEnd));
      i = el.end;
      continue;
    }
    i = t.end;
  }
  return s;
}

/** bold/italic from a run's rPr (toggle present and not `w:val="0"`). */
function runFormatting(xml: string, run: ElementSlice): { bold: boolean; italic: boolean } {
  const rPr = findElement(xml, "w:rPr", run.contentStart, run.contentEnd);
  if (!rPr || rPr.selfClosed || rPr.start !== run.contentStart) {
    return { bold: false, italic: false };
  }
  const kids = childElements(xml, rPr.contentStart, rPr.contentEnd);
  const on = (name: string): boolean => {
    const el = kids.find((k) => k.name === name);
    if (!el) return false;
    const val = getAttr(xml, tagOf(el), "w:val");
    return val !== "0" && val !== "false";
  };
  return { bold: on("w:b"), italic: on("w:i") };
}

/** Alignment (jc) and color from a paragraph's pPr/run rPr. */
function paragraphFormatting(
  xml: string,
  p: ElementSlice,
): {
  alignment: string | null;
  color: string | null;
} {
  let alignment: string | null = null;
  let color: string | null = null;
  const pPr = findElement(xml, "w:pPr", p.contentStart, p.contentEnd);
  if (pPr && !pPr.selfClosed) {
    const jc = findElement(xml, "w:jc", pPr.contentStart, pPr.contentEnd);
    if (jc) alignment = getAttr(xml, tagOf(jc), "w:val") ?? null;
  }
  // Color from the first run's rPr (paragraph-level approximation, §23 html).
  const firstRun = findElement(xml, "w:r", p.contentStart, p.contentEnd);
  if (firstRun) {
    const rPr = findElement(xml, "w:rPr", firstRun.contentStart, firstRun.contentEnd);
    if (rPr && !rPr.selfClosed) {
      const c = findElement(xml, "w:color", rPr.contentStart, rPr.contentEnd);
      if (c) {
        const v = getAttr(xml, tagOf(c), "w:val");
        if (v && v !== "auto") color = v;
      }
    }
  }
  return { alignment, color };
}

/** List annotation (ol/ul + level) from numPr, mirroring §2. */
function listAnnotation(
  xml: string,
  p: ElementSlice,
  ctx: ProjectionContext,
): { kind: "ol" | "ul"; level: number } | null {
  const pPr = findElement(xml, "w:pPr", p.contentStart, p.contentEnd);
  if (!pPr || pPr.selfClosed) return null;
  const numPr = findElement(xml, "w:numPr", pPr.contentStart, pPr.contentEnd);
  if (!numPr || numPr.selfClosed) return null;
  const numIdEl = findElement(xml, "w:numId", numPr.contentStart, numPr.contentEnd);
  const numId = numIdEl ? getAttr(xml, tagOf(numIdEl), "w:val") : undefined;
  if (numId === undefined || numId === "0") return null;
  let level = 0;
  const ilvlEl = findElement(xml, "w:ilvl", numPr.contentStart, numPr.contentEnd);
  if (ilvlEl) {
    const v = Number(getAttr(xml, tagOf(ilvlEl), "w:val"));
    if (Number.isInteger(v) && v >= 0) level = v;
  }
  const abstract = ctx.numbering.numToAbstract.get(numId);
  const fmt =
    abstract === undefined ? undefined : ctx.numbering.abstractFormats.get(abstract)?.get(level);
  return { kind: fmt === "bullet" ? "ul" : "ol", level };
}

// ---------------------------------------------------------------------------
// comments.xml text resolution
// ---------------------------------------------------------------------------

interface CommentCtx extends ProjectionContext {
  commentTexts: Map<string, string>;
}

function commentTextFor(ctx: ProjectionContext, id: string): string {
  const cc = ctx as CommentCtx;
  return cc.commentTexts?.get(id) ?? "";
}

/** Build the comment-id → text map from comments.xml. */
function commentTexts(doc: DocHandle): Map<string, string> {
  const map = new Map<string, string>();
  if (!doc.pkg.has("word/comments.xml")) return map;
  const xml = doc.pkg.partText("word/comments.xml");
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return map;
    if (t.name === "w:comment" && t.kind !== "end") {
      const el = elementExtent(xml, t);
      const id = getAttr(xml, t, "w:id");
      if (id !== undefined) map.set(id, normalizedText(scopeText(xml, el)));
      i = el.end;
      continue;
    }
    i = t.end;
  }
}

// ---------------------------------------------------------------------------
// Markdown rendering (§23)
// ---------------------------------------------------------------------------

interface ConvertOpts {
  /** Show `[ins]`/`[del]` markers (default true per §23). */
  markers: boolean;
}

/** Reconstruct inline markdown from a paragraph's runs (§23). */
function runsToMarkdown(runs: ConvRun[], opts: ConvertOpts): string {
  let out = "";
  for (const run of runs) {
    let text = run.text;
    if (run.bold) text = `**${text}**`;
    if (run.italic) text = `*${text}*`;
    if (run.ins && opts.markers) text = `[ins]${text}`;
    out += text;
  }
  return out;
}

function paragraphToMarkdown(model: ParaModel, opts: ConvertOpts): string {
  let text = runsToMarkdown(model.runs, opts);
  if (model.hasDeletion && opts.markers) text += "[del]";
  for (const c of model.comments) {
    text += ` <!-- comment:${c.author}: ${c.text} -->`;
  }
  if (model.heading !== null) {
    return `${"#".repeat(model.heading)} ${text}`;
  }
  if (model.list !== null) {
    const indent = "  ".repeat(model.list.level);
    const bullet = model.list.kind === "ol" ? "1. " : "- ";
    return `${indent}${bullet}${text}`;
  }
  return text;
}

/** A §2 GitHub table for md. */
function tableToMarkdown(xml: string, tbl: ElementSlice): string {
  const { cols } = tableDims(xml, tbl);
  const lines: string[] = [];
  const rows = childElements(xml, tbl.contentStart, tbl.contentEnd).filter(
    (k) => k.name === "w:tr",
  );
  rows.forEach((tr, idx) => {
    const cells = tr.selfClosed
      ? []
      : childElements(xml, tr.contentStart, tr.contentEnd).filter((k) => k.name === "w:tc");
    const cellText = cells.map((tc) => mdCellText(xml, tc));
    lines.push(`| ${cellText.join(" | ")} |`);
    if (idx === 0) lines.push(`|${" --- |".repeat(Math.max(cols, 1))}`);
  });
  return lines.join("\n");
}

function mdCellText(xml: string, tc: ElementSlice): string {
  if (tc.selfClosed) return "";
  const texts = childElements(xml, tc.contentStart, tc.contentEnd)
    .filter((k) => k.name === "w:p")
    .map((p) => normalizedText(scopeText(xml, p)));
  return texts.join(" ").replace(/\|/g, "\\|");
}

/**
 * Render the whole document to markdown. Blocks join with a blank line, except
 * consecutive list items, which join with a single newline (tight list, §23).
 */
function toMarkdown(doc: DocHandle, ctx: CommentCtx, opts: ConvertOpts): string {
  const blocks = bodyProjBlocks(doc);
  let out = "";
  let prevWasListItem = false;
  for (const b of blocks) {
    let line: string;
    let isListItem = false;
    if (b.kind === "tbl") {
      line = tableToMarkdown(b.xml, b.block);
    } else {
      const model = parseParagraph(b.xml, b.block, ctx, ctx.commentAuthors);
      isListItem = model.list !== null;
      line = paragraphToMarkdown(model, opts);
    }
    if (out === "") {
      out = line;
    } else {
      out += (prevWasListItem && isListItem ? "\n" : "\n\n") + line;
    }
    prevWasListItem = isListItem;
  }
  return out;
}

// ---------------------------------------------------------------------------
// HTML rendering (§23)
// ---------------------------------------------------------------------------

function htmlEscape(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function runsToHtml(runs: ConvRun[]): string {
  let out = "";
  for (const run of runs) {
    let text = htmlEscape(run.text);
    if (run.bold) text = `<strong>${text}</strong>`;
    if (run.italic) text = `<em>${text}</em>`;
    out += text;
  }
  return out;
}

function styleAttr(model: ParaModel): string {
  const styles: string[] = [];
  if (model.alignment !== null) {
    const css =
      model.alignment === "both"
        ? "justify"
        : model.alignment === "center" || model.alignment === "right" || model.alignment === "left"
          ? model.alignment
          : null;
    if (css) styles.push(`text-align:${css}`);
  }
  if (model.color !== null) styles.push(`color:#${model.color}`);
  return styles.length > 0 ? ` style="${styles.join(";")}"` : "";
}

function paragraphToHtml(model: ParaModel): string {
  const inner = runsToHtml(model.runs);
  const sa = styleAttr(model);
  if (model.heading !== null) {
    const h = Math.min(6, model.heading);
    return `<h${h}${sa}>${inner}</h${h}>`;
  }
  return `<p${sa}>${inner}</p>`;
}

function tableToHtml(xml: string, tbl: ElementSlice): string {
  const rows = childElements(xml, tbl.contentStart, tbl.contentEnd).filter(
    (k) => k.name === "w:tr",
  );
  const trs = rows.map((tr) => {
    const cells = tr.selfClosed
      ? []
      : childElements(xml, tr.contentStart, tr.contentEnd).filter((k) => k.name === "w:tc");
    const tds = cells.map(
      (tc) => `<td>${htmlEscape(mdCellText(xml, tc).replace(/\\\|/g, "|"))}</td>`,
    );
    return `<tr>${tds.join("")}</tr>`;
  });
  return `<table>${trs.join("")}</table>`;
}

/** Render the whole document to HTML, grouping consecutive list items. */
function toHtml(doc: DocHandle, ctx: CommentCtx, opts: ConvertOpts): string {
  const blocks = bodyProjBlocks(doc);
  const out: string[] = [];
  let listOpen: "ol" | "ul" | null = null;
  const closeList = (): void => {
    if (listOpen !== null) {
      out.push(`</${listOpen}>`);
      listOpen = null;
    }
  };
  for (const b of blocks) {
    if (b.kind === "tbl") {
      closeList();
      out.push(tableToHtml(b.xml, b.block));
      continue;
    }
    const model = parseParagraph(b.xml, b.block, ctx, ctx.commentAuthors);
    if (model.list !== null) {
      if (listOpen !== model.list.kind) {
        closeList();
        out.push(`<${model.list.kind}>`);
        listOpen = model.list.kind;
      }
      out.push(`<li>${runsToHtml(model.runs)}</li>`);
      continue;
    }
    closeList();
    out.push(paragraphToHtml(model));
  }
  closeList();
  void opts;
  return out.join("\n");
}

// ---------------------------------------------------------------------------
// docx_convert
// ---------------------------------------------------------------------------

export interface DocxConvertArgs {
  doc_id: string;
  to: "md" | "html" | "pdf" | "png";
  path?: string | undefined;
}

export interface DocxConvertResult {
  content?: string;
  path?: string;
  renderer?: string;
  note?: string;
}

/** Convert an open document to md/html (in-engine) or pdf/png (render adapter). */
export function docxConvert(session: Session, args: DocxConvertArgs): DocxConvertResult {
  const doc = session.get(args.doc_id);
  const to = args.to;
  if (to !== "md" && to !== "html" && to !== "pdf" && to !== "png") {
    throw new ToolError("unsupported_format", `Unsupported conversion target: ${String(to)}.`, [
      "Use to: md, html, pdf or png.",
    ]);
  }
  if (to === "md" || to === "html") {
    const base = projectionContext(doc);
    const ctx: CommentCtx = { ...base, commentTexts: commentTexts(doc) };
    const opts: ConvertOpts = { markers: true };
    const content = to === "md" ? toMarkdown(doc, ctx, opts) : toHtml(doc, ctx, opts);
    return { content, note: convertNote(doc) };
  }
  // pdf/png → render adapter (§24).
  if (args.path === undefined) {
    throw new ToolError("unsupported_format", `'${to}' requires an output path.`, [
      "Pass path for pdf/png targets.",
    ]);
  }
  return renderToFile(doc, to, args.path);
}

/** A free-form note counting inline annotations (masked in parity). */
function convertNote(doc: DocHandle): string {
  const xml = doc.documentXml();
  const comments = countOccurrences(xml, "<w:commentReference");
  const revisions = countOccurrences(xml, "<w:ins ") + countOccurrences(xml, "<w:del ");
  if (comments === 0 && revisions === 0) return "Converted (no comments or tracked changes).";
  const parts: string[] = [];
  if (comments > 0) parts.push(`${comments} comment${comments === 1 ? "" : "s"}`);
  if (revisions > 0) parts.push(`${revisions} tracked change${revisions === 1 ? "" : "s"}`);
  return `${parts.join(" and ")} annotated inline`;
}

function countOccurrences(s: string, sub: string): number {
  let n = 0;
  let i = s.indexOf(sub);
  while (i >= 0) {
    n += 1;
    i = s.indexOf(sub, i + sub.length);
  }
  return n;
}
