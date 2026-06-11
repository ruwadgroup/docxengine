/**
 * The native surface (docs/sdks/javascript.md §2): a Pythonic object handle over
 * the same session + tool layer as `call()`. No business logic lives here —
 * every mutation goes through the contract tools, so anchors, tracked changes,
 * and the validation gate behave identically on every surface.
 *
 * A Document owns a private {@link Session} (pass `session` to share one), so it
 * is safe to use many independently in one process — a multi-tenant server, a
 * browser tab. It is storage-agnostic: open from a path or bytes, persist with
 * {@link Document.save} (a path, Node only) or {@link Document.toBytes} (raw
 * bytes, anywhere). This is distinct from the file-first MCP server.
 */
import { dispatch } from "./dispatch.js";
import { ToolError } from "./errors.js";
import { type DocHandle, Session } from "./session.js";
import { fillDoc } from "./template.js";
import type { DocxReplaceResult } from "./toolsEdit.js";
import {
  type DocxRepairResult,
  type DocxSaveResult,
  type DocxValidateResult,
  exportBytes,
} from "./toolsLifecycle.js";
import { type ElementSlice, childElements, findElement, getAttr, nextTag } from "./xmlscan.js";

type Result = Record<string, unknown>;

export interface EditOptions {
  trackChanges?: boolean | undefined;
  author?: string | undefined;
}
export interface ReplaceOptions extends EditOptions {
  all?: boolean | undefined;
}
export interface InsertOptions extends EditOptions {
  style?: string | undefined;
}

/** The paragraph's `w:pStyle` styleId, or null when unstyled. */
function paragraphStyle(xml: string, block: ElementSlice): string | null {
  if (block.selfClosed) return null;
  const pPr = childElements(xml, block.contentStart, block.contentEnd).find(
    (k) => k.name === "w:pPr",
  );
  if (!pPr || pPr.selfClosed) return null;
  const ps = findElement(xml, "w:pStyle", pPr.contentStart, pPr.contentEnd);
  if (!ps) return null;
  const tag = nextTag(xml, ps.start);
  return tag ? (getAttr(xml, tag, "w:val") ?? null) : null;
}

export class DocumentParagraph {
  readonly anchor: string;
  /** Normalized text (spec/algorithms.md §1). */
  readonly text: string;
  /** `w:pStyle` styleId (e.g. "Heading1"), or null. */
  readonly style: string | null;
  private readonly doc: Document;

  constructor(doc: Document, anchor: string, text: string, style: string | null) {
    this.doc = doc;
    this.anchor = anchor;
    this.text = text;
    this.style = style;
  }

  /**
   * A paragraph is a throwaway view: after any edit to it the held anchor goes
   * stale (anchors are content-addressed). Re-fetch via `paragraphs()`/`find()`;
   * a stale anchor raises `anchor_stale`, the spec's normal recovery signal.
   */
  replace(oldText: string, newText: string, opts: ReplaceOptions = {}): Promise<DocxReplaceResult> {
    return this.doc.replace(oldText, newText, { ...opts, anchor: this.anchor });
  }

  edit(text: string, opts: EditOptions = {}): Promise<Result> {
    return this.doc.editParagraph(this.anchor, text, opts);
  }

  insertAfter(content: string, opts: InsertOptions = {}): Promise<Result> {
    return this.doc.insert(content, { ...opts, after: this.anchor });
  }

  insertBefore(content: string, opts: InsertOptions = {}): Promise<Result> {
    return this.doc.insert(content, { ...opts, before: this.anchor });
  }

  delete(opts: EditOptions = {}): Promise<Result> {
    return this.doc.delete({ ...opts, anchor: this.anchor });
  }
}

export class Document {
  /** The doc_id, interchangeable with the `call()`/dispatch surface. */
  readonly id: string;
  /** The backing session — share it, or hand it to an agent loop. */
  readonly session: Session;

  private constructor(session: Session, id: string) {
    this.session = session;
    this.id = id;
  }

  // -- construction ---------------------------------------------------------

  /** Open a .docx from a path or raw bytes (on a fresh session unless given). */
  static async open(
    source: string | Uint8Array,
    opts: { session?: Session } = {},
  ): Promise<Document> {
    const session = opts.session ?? new Session();
    return new Document(session, session.open(source).id);
  }

  /** Create a new document from Markdown or a structured spec (§22). */
  static async create(
    opts: { contentMd?: string; spec?: Record<string, unknown>; session?: Session } = {},
  ): Promise<Document> {
    const session = opts.session ?? new Session();
    const result = dispatch(session, "docx_create", {
      content_md: opts.contentMd,
      spec: opts.spec,
    }) as { doc_id: string };
    return new Document(session, result.doc_id);
  }

  /** Open a template (path or bytes), fill it (§21), return the filled Document. */
  static async fillTemplate(
    template: string | Uint8Array,
    data: Record<string, unknown>,
    opts: { syntax?: string; strict?: boolean; session?: Session } = {},
  ): Promise<Document> {
    const syntax = opts.syntax ?? "mustache";
    if (syntax !== "mustache") {
      throw new ToolError("template_syntax", `Unsupported template syntax: ${syntax}.`, [
        "Only the mustache subset is supported.",
      ]);
    }
    const session = opts.session ?? new Session();
    const doc = session.open(template);
    fillDoc(doc, data, { strict: opts.strict === true });
    return new Document(session, doc.id);
  }

  /** Wrap an already-open doc_id in `session` (shares its state). */
  static attach(session: Session, docId: string): Document {
    session.get(docId); // throws doc_not_found if unknown
    return new Document(session, docId);
  }

  // -- escape hatch + state -------------------------------------------------

  /** Dispatch any contract tool against this document (doc_id injected). */
  async call(tool: string, args: Record<string, unknown> = {}): Promise<unknown> {
    return dispatch(this.session, tool, { ...args, doc_id: this.id });
  }

  /** True when the document has unsaved modifications. */
  get dirty(): boolean {
    return this.handle().dirty;
  }

  // -- read -----------------------------------------------------------------

  outline(): Promise<Result> {
    return this.run("docx_outline", {});
  }

  read(
    opts: { anchor?: string; range?: string; window?: number; scope?: string } = {},
  ): Promise<Result> {
    return this.run("docx_read", { ...opts });
  }

  search(query: string, opts: { regex?: boolean; scope?: string } = {}): Promise<Result> {
    return this.run("docx_search", { query, ...opts });
  }

  /** Body paragraphs with anchor, normalized text, and styleId. */
  paragraphs(): DocumentParagraph[] {
    const handle = this.handle();
    const xml = handle.documentXml();
    return handle
      .anchorIndex()
      .filter((e) => e.kind === "p")
      .map(
        (e) =>
          new DocumentParagraph(this, e.anchor, e.normalized ?? "", paragraphStyle(xml, e.block)),
      );
  }

  /** First paragraph whose normalized text contains `text`, or null. */
  find(text: string): DocumentParagraph | null {
    return this.paragraphs().find((p) => p.text.includes(text)) ?? null;
  }

  // -- edit -----------------------------------------------------------------

  async replace(
    oldText: string,
    newText: string,
    opts: ReplaceOptions & { anchor?: string } = {},
  ): Promise<DocxReplaceResult> {
    return (await this.run("docx_replace", {
      old: oldText,
      new: newText,
      anchor: opts.anchor,
      all: opts.all,
      track_changes: opts.trackChanges,
      author: opts.author,
    })) as unknown as DocxReplaceResult;
  }

  editParagraph(anchor: string, text: string, opts: EditOptions = {}): Promise<Result> {
    return this.run("docx_edit_paragraph", {
      anchor,
      text,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  insert(
    content: string,
    opts: InsertOptions & { after?: string; before?: string } = {},
  ): Promise<Result> {
    return this.run("docx_insert", {
      content,
      after: opts.after,
      before: opts.before,
      style: opts.style,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  delete(opts: { anchor?: string; range?: string } & EditOptions = {}): Promise<Result> {
    return this.run("docx_delete", {
      anchor: opts.anchor,
      range: opts.range,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  revision(
    op: string,
    opts: { id?: string; filter?: Record<string, string> } = {},
  ): Promise<Result> {
    return this.run("docx_revision", { op, ...opts });
  }

  comment(
    op: string,
    opts: { anchor?: string; commentId?: string; text?: string; author?: string } = {},
  ): Promise<Result> {
    return this.run("docx_comment", {
      op,
      anchor: opts.anchor,
      comment_id: opts.commentId,
      text: opts.text,
      author: opts.author,
    });
  }

  table(op: string, opts: TableOptions = {}): Promise<Result> {
    return this.run("docx_table", {
      op,
      anchor: opts.anchor,
      after: opts.after,
      rows: opts.rows,
      cols: opts.cols,
      data: opts.data,
      header: opts.header,
      cells: opts.cells,
      at: opts.at,
      range: opts.range,
      style: opts.style,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  style(op: string, opts: StyleOptions = {}): Promise<Result> {
    return this.run("docx_style", {
      op,
      anchor: opts.anchor,
      style: opts.style,
      name: opts.name,
      based_on: opts.basedOn,
      props: opts.props,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  format(props: Record<string, unknown>, opts: FormatOptions = {}): Promise<Result> {
    return this.run("docx_format", {
      props,
      anchor: opts.anchor,
      range: opts.range,
      style_selector: opts.styleSelector,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  list(op: string, opts: ListOptions = {}): Promise<Result> {
    return this.run("docx_list", {
      op,
      anchor: opts.anchor,
      range: opts.range,
      after: opts.after,
      kind: opts.kind,
      items: opts.items,
      at: opts.at,
      level: opts.level,
      to: opts.to,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  section(op: string, opts: SectionOptions = {}): Promise<Result> {
    return this.run("docx_section", {
      op,
      section: opts.section,
      page_size: opts.pageSize,
      orientation: opts.orientation,
      margins: opts.margins,
      columns: opts.columns,
      content: opts.content,
      variant: opts.variant,
      after: opts.after,
      break_type: opts.breakType,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  media(op: string, opts: MediaOptions = {}): Promise<Result> {
    return this.run("docx_media", {
      op,
      after: opts.after,
      before: opts.before,
      image: opts.image,
      width_cm: opts.widthCm,
      height_cm: opts.heightCm,
      media_id: opts.mediaId,
      path: opts.path,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  field(op: string, opts: FieldOptions = {}): Promise<Result> {
    return this.run("docx_field", {
      op,
      after: opts.after,
      levels: opts.levels,
      scope: opts.scope,
      track_changes: opts.trackChanges,
      author: opts.author,
    });
  }

  // -- lifecycle ------------------------------------------------------------

  async validate(): Promise<DocxValidateResult> {
    return (await this.run("docx_validate", {})) as unknown as DocxValidateResult;
  }

  async repair(): Promise<DocxRepairResult> {
    return (await this.run("docx_repair", {})) as unknown as DocxRepairResult;
  }

  renderPreview(opts: { pages?: number[] } = {}): Promise<Result> {
    return this.run("docx_render_preview", { pages: opts.pages });
  }

  convert(to: string, opts: { path?: string } = {}): Promise<Result> {
    return this.run("docx_convert", { to, path: opts.path });
  }

  /** docx_save: validation gate + atomic write to a path (Node only). */
  async save(path: string): Promise<DocxSaveResult> {
    return (await this.run("docx_save", { path })) as unknown as DocxSaveResult;
  }

  /** The validated .docx bytes, no filesystem (e.g. a browser download). */
  toBytes(): Uint8Array {
    return exportBytes(this.session, this.id);
  }

  // -- internals ------------------------------------------------------------

  private handle(): DocHandle {
    return this.session.get(this.id);
  }

  /** Dispatch `tool` against this doc; omit undefined args so handler defaults apply. */
  private async run(tool: string, args: Record<string, unknown>): Promise<Result> {
    const clean: Record<string, unknown> = { doc_id: this.id };
    for (const [k, v] of Object.entries(args)) if (v !== undefined) clean[k] = v;
    return dispatch(this.session, tool, clean) as Result;
  }
}

export interface TableOptions extends EditOptions {
  anchor?: string;
  after?: string;
  rows?: number;
  cols?: number;
  data?: string[][];
  header?: boolean;
  cells?: unknown;
  at?: string;
  range?: string;
  style?: string;
}
export interface StyleOptions extends EditOptions {
  anchor?: string;
  style?: string;
  name?: string;
  basedOn?: string;
  props?: Record<string, unknown>;
}
export interface FormatOptions extends EditOptions {
  anchor?: string;
  range?: string;
  styleSelector?: string;
}
export interface ListOptions extends EditOptions {
  anchor?: string;
  range?: string;
  after?: string;
  kind?: string;
  items?: string[];
  at?: string;
  level?: number;
  to?: string;
}
export interface SectionOptions extends EditOptions {
  section?: number;
  pageSize?: string;
  orientation?: string;
  margins?: Record<string, unknown>;
  columns?: number;
  content?: string;
  variant?: string;
  after?: string;
  breakType?: string;
}
export interface MediaOptions extends EditOptions {
  after?: string;
  before?: string;
  image?: string;
  widthCm?: number;
  heightCm?: number;
  mediaId?: string;
  path?: string;
}
export interface FieldOptions extends EditOptions {
  after?: string;
  levels?: unknown;
  scope?: string;
}
