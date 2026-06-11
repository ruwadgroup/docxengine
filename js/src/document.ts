/**
 * The native surface (docs/sdks/javascript.md §2): a thin object wrapper over
 * the same session + tool layer as `call()`. No business logic lives here —
 * every mutation goes through the contract tools, so anchors, tracking, and
 * the validation gate behave identically on both surfaces.
 */
import { defaultSession, dispatch } from "./dispatch.js";
import type { DocHandle } from "./session.js";
import type { DocxReplaceResult } from "./toolsEdit.js";
import type { DocxRepairResult, DocxSaveResult, DocxValidateResult } from "./toolsLifecycle.js";
import { type ElementSlice, childElements, findElement, getAttr, nextTag } from "./xmlscan.js";

export interface ReplaceOptions {
  all?: boolean | undefined;
  trackChanges?: boolean | undefined;
  author?: string | undefined;
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
  /** `w:pStyle` styleId (e.g. "Heading2"), or null. */
  readonly style: string | null;
  private readonly doc: Document;

  constructor(doc: Document, anchor: string, text: string, style: string | null) {
    this.doc = doc;
    this.anchor = anchor;
    this.text = text;
    this.style = style;
  }

  /** docx_replace scoped to this paragraph. */
  async replace(
    oldText: string,
    newText: string,
    opts: ReplaceOptions = {},
  ): Promise<DocxReplaceResult> {
    return (await this.doc.call("docx_replace", {
      anchor: this.anchor,
      old: oldText,
      new: newText,
      all: opts.all,
      track_changes: opts.trackChanges,
      author: opts.author,
    })) as DocxReplaceResult;
  }
}

export class Document {
  /** The doc_id, interchangeable with the `call()` surface. */
  readonly id: string;

  private constructor(id: string) {
    this.id = id;
  }

  /** Open a .docx from a path or raw bytes on the shared session. */
  static async open(source: string | Uint8Array): Promise<Document> {
    return new Document(defaultSession.open(source).id);
  }

  /** Dispatch any contract tool against this document. */
  async call(tool: string, args: Record<string, unknown> = {}): Promise<unknown> {
    return dispatch(defaultSession, tool, { ...args, doc_id: this.id });
  }

  /** True when the document has unsaved modifications. */
  get dirty(): boolean {
    return this.handle().dirty;
  }

  /** Body paragraphs with anchor, normalized text, and styleId. */
  paragraphs(): DocumentParagraph[] {
    const xml = this.handle().documentXml();
    return this.handle()
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

  async validate(): Promise<DocxValidateResult> {
    return (await this.call("docx_validate")) as DocxValidateResult;
  }

  async repair(): Promise<DocxRepairResult> {
    return (await this.call("docx_repair")) as DocxRepairResult;
  }

  /** docx_save: validation gate + atomic write. The doc_id stays open. */
  async save(path: string): Promise<DocxSaveResult> {
    return (await this.call("docx_save", { path })) as DocxSaveResult;
  }

  private handle(): DocHandle {
    return defaultSession.get(this.id);
  }
}
