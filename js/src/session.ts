/**
 * Session/doc store per spec/algorithms.md §2a and §11: doc ids are `d1`,
 * `d2`, … in open order and persist for the process lifetime.
 */
import { type AnchorEntry, buildAnchorIndex } from "./anchors.js";
import { ToolError } from "./errors.js";
import { Package } from "./opc.js";

const OFFICE_DOCUMENT_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/** An open document: the package plus cached document-part text and anchors. */
export class DocHandle {
  readonly id: string;
  readonly pkg: Package;
  /** Main document part name (resolved from the package-level rels). */
  readonly documentPartName: string;
  private cache: { xml: string; index: AnchorEntry[] } | null = null;
  private dirtyFlag = false;

  constructor(id: string, pkg: Package) {
    this.id = id;
    this.pkg = pkg;
    this.documentPartName = resolveDocumentPart(pkg);
  }

  /** True when the document has been modified since open or the last save. */
  get dirty(): boolean {
    return this.dirtyFlag;
  }

  markDirty(): void {
    this.dirtyFlag = true;
  }

  /** docx_save calls this after a successful write (the doc_id stays open). */
  markSaved(): void {
    this.dirtyFlag = false;
  }

  /** The document part decoded as UTF-8 text (cached until `invalidate`). */
  documentXml(): string {
    return this.ensureCache().xml;
  }

  /** The §1 anchor index over the current document part (cached). */
  anchorIndex(): AnchorEntry[] {
    return this.ensureCache().index;
  }

  /** Drop caches; edit tools call this after splicing the document part. */
  invalidate(): void {
    this.cache = null;
    this.dirtyFlag = true;
  }

  private ensureCache(): { xml: string; index: AnchorEntry[] } {
    if (!this.cache) {
      const xml = this.pkg.partText(this.documentPartName);
      this.cache = { xml, index: buildAnchorIndex(xml) };
    }
    return this.cache;
  }
}

function resolveDocumentPart(pkg: Package): string {
  for (const rel of pkg.rels()) {
    if (rel.type === OFFICE_DOCUMENT_REL && rel.targetMode === "Internal") {
      const target = rel.target.startsWith("/") ? rel.target.slice(1) : rel.target;
      if (pkg.has(target)) return target;
    }
  }
  if (pkg.has("word/document.xml")) return "word/document.xml";
  throw new ToolError("open_failed", "Cannot open: package has no main document part.", [
    "Check the path; the message says what the file actually is.",
  ]);
}

/** The doc_id → document store backing one CLI/MCP process. */
export class Session {
  private nextOrdinal = 1;
  private readonly docs = new Map<string, DocHandle>();

  /** Open a .docx from a path or raw bytes; registers it as the next `d{n}`. */
  open(source: string | Uint8Array): DocHandle {
    const pkg = Package.open(source);
    const doc = new DocHandle(`d${this.nextOrdinal++}`, pkg);
    this.docs.set(doc.id, doc);
    return doc;
  }

  /** Look up a doc_id; unknown → `doc_not_found`. */
  get(docId: string): DocHandle {
    const doc = this.docs.get(docId);
    if (!doc) {
      throw new ToolError("doc_not_found", `Unknown or expired doc_id: ${docId}.`, [
        "Call docx_open again.",
      ]);
    }
    return doc;
  }
}
