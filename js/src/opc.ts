/**
 * OPC package handling per spec/algorithms.md §3 (raw bytes) and §9 (save).
 *
 * Every part is held as raw bytes; only parts actually modified are ever
 * re-encoded. Untouched parts pass through with byte-identical decompressed
 * content, in the source zip's original entry order; new parts are appended.
 */
import { unzipSync, zipSync, type Zippable } from "fflate";

import { ToolError } from "./errors.js";
import { nodeFs, nodePath } from "./nodeenv.js";
import { type Tag, attrs, nextTag } from "./xmlscan.js";

const CONTENT_TYPES_PART = "[Content_Types].xml";

/** Entry metadata is normalized: DOS timestamp 1980-01-01 00:00:00 (§9). */
const DOS_EPOCH = new Date(1980, 0, 1, 0, 0, 0);

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

export interface Relationship {
  id: string;
  type: string;
  target: string;
  /** `Internal` unless the rel carries `TargetMode="External"`. */
  targetMode: string;
}

export interface ContentTypes {
  /** lowercased extension → content type. */
  defaults: Map<string, string>;
  /** part name (leading `/` as written) → content type. */
  overrides: Map<string, string>;
}

function scanElements(xml: string, name: string): Tag[] {
  const out: Tag[] = [];
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return out;
    if (t.name === name && t.kind !== "end") out.push(t);
    i = t.end;
  }
}

export class Package {
  /** Current raw bytes per part name (zip entry name, no leading slash). */
  private readonly parts = new Map<string, Uint8Array>();
  /** Entry order of the source zip. */
  private readonly originalOrder: string[] = [];
  /** New parts appended after the originals, in creation order (§9). */
  private readonly appended: string[] = [];
  private readonly dirtySet = new Set<string>();

  private constructor() {}

  /** Open a package from a file path or raw zip bytes. */
  static open(source: string | Uint8Array): Package {
    let data: Uint8Array;
    if (typeof source === "string") {
      try {
        data = nodeFs().readFileSync(source);
      } catch (e) {
        throw new ToolError(
          "open_failed",
          `Cannot open ${source}: unreadable path (${(e as Error).message}).`,
          ["Check the path; the message says what the file actually is."],
        );
      }
    } else {
      data = source;
    }
    let entries: Record<string, Uint8Array>;
    try {
      entries = unzipSync(data);
    } catch (e) {
      throw new ToolError(
        "open_failed",
        `Cannot open: not a zip archive (${(e as Error).message}).`,
        ["Check the path; the message says what the file actually is."],
      );
    }
    const pkg = new Package();
    for (const [name, bytes] of Object.entries(entries)) {
      pkg.originalOrder.push(name);
      pkg.parts.set(name, bytes);
    }
    if (!pkg.parts.has(CONTENT_TYPES_PART)) {
      throw new ToolError(
        "open_failed",
        "Cannot open: zip has no [Content_Types].xml (not an OPC package).",
        ["Check the path; the message says what the file actually is."],
      );
    }
    return pkg;
  }

  /** Entry names: original zip order, then appended parts in creation order. */
  entryNames(): string[] {
    return [...this.originalOrder, ...this.appended];
  }

  has(name: string): boolean {
    return this.parts.has(name);
  }

  /** Raw bytes of a part. Missing part is a programming error, not a ToolError. */
  part(name: string): Uint8Array {
    const bytes = this.parts.get(name);
    if (!bytes) throw new Error(`unknown part: ${name}`);
    return bytes;
  }

  tryPart(name: string): Uint8Array | undefined {
    return this.parts.get(name);
  }

  /** A part decoded as UTF-8 text (for the §3 scanner). */
  partText(name: string): string {
    return decoder.decode(this.part(name));
  }

  /** Replace or create a part; marks it dirty. New parts append at the end. */
  setPart(name: string, content: Uint8Array | string): void {
    if (!this.parts.has(name)) this.appended.push(name);
    this.parts.set(name, typeof content === "string" ? encoder.encode(content) : content);
    this.dirtySet.add(name);
  }

  isDirty(name: string): boolean {
    return this.dirtySet.has(name);
  }

  get dirty(): ReadonlySet<string> {
    return this.dirtySet;
  }

  // -------------------------------------------------------------------------
  // [Content_Types].xml
  // -------------------------------------------------------------------------

  contentTypes(): ContentTypes {
    const xml = this.partText(CONTENT_TYPES_PART);
    const defaults = new Map<string, string>();
    const overrides = new Map<string, string>();
    for (const tag of scanElements(xml, "Default")) {
      const a = attrs(xml, tag);
      const ext = a["Extension"];
      const ct = a["ContentType"];
      if (ext !== undefined && ct !== undefined) defaults.set(ext.toLowerCase(), ct);
    }
    for (const tag of scanElements(xml, "Override")) {
      const a = attrs(xml, tag);
      const pn = a["PartName"];
      const ct = a["ContentType"];
      if (pn !== undefined && ct !== undefined) overrides.set(pn, ct);
    }
    return { defaults, overrides };
  }

  /** Content type of a part (Override first, then extension Default). */
  contentTypeOf(partName: string): string | undefined {
    const { defaults, overrides } = this.contentTypes();
    const override = overrides.get(partName.startsWith("/") ? partName : `/${partName}`);
    if (override !== undefined) return override;
    const dot = partName.lastIndexOf(".");
    if (dot < 0) return undefined;
    return defaults.get(partName.slice(dot + 1).toLowerCase());
  }

  // -------------------------------------------------------------------------
  // Relationships
  // -------------------------------------------------------------------------

  /** The rels part name for a part (`undefined` → package-level `_rels/.rels`). */
  static relsPartFor(partName?: string): string {
    if (partName === undefined || partName === "") return "_rels/.rels";
    const slash = partName.lastIndexOf("/");
    const dir = slash < 0 ? "" : partName.slice(0, slash + 1);
    const base = slash < 0 ? partName : partName.slice(slash + 1);
    return `${dir}_rels/${base}.rels`;
  }

  /** Parsed relationships of a part (empty if the rels part is absent). */
  rels(partName?: string): Relationship[] {
    const relsName = Package.relsPartFor(partName);
    if (!this.parts.has(relsName)) return [];
    const xml = this.partText(relsName);
    const out: Relationship[] = [];
    for (const tag of scanElements(xml, "Relationship")) {
      const a = attrs(xml, tag);
      out.push({
        id: a["Id"] ?? "",
        type: a["Type"] ?? "",
        target: a["Target"] ?? "",
        targetMode: a["TargetMode"] ?? "Internal",
      });
    }
    return out;
  }

  // -------------------------------------------------------------------------
  // Save (algorithms.md §9 steps 2–4; the §8 validation gate is wired in by
  // docx_save when the validator lands — Package serializes mechanically)
  // -------------------------------------------------------------------------

  /**
   * Serialize the package: source entries in original order with untouched
   * parts' content byte-identical, new parts appended; zeroed DOS timestamps,
   * no extra fields or comments, deflate level 6.
   */
  toBytes(): Uint8Array {
    const zippable: Zippable = {};
    for (const name of this.entryNames()) {
      zippable[name] = this.part(name);
    }
    return zipSync(zippable, { level: 6, mtime: DOS_EPOCH });
  }

  /** Atomic write: serialize to a temp file in the destination dir, then rename. */
  save(dest: string): void {
    const bytes = this.toBytes();
    let tmp: string | undefined;
    try {
      const fs = nodeFs();
      const path = nodePath();
      const resolved = path.resolve(dest);
      tmp = path.join(path.dirname(resolved), `.${path.basename(resolved)}.tmp-${process.pid}`);
      fs.writeFileSync(tmp, bytes);
      fs.renameSync(tmp, resolved);
    } catch (e) {
      if (tmp !== undefined) {
        try {
          nodeFs().rmSync(tmp, { force: true });
        } catch {
          /* best-effort cleanup */
        }
      }
      throw new ToolError(
        "save_failed",
        `I/O failure writing output to ${dest}: ${(e as Error).message}.`,
        ["Check the path and permissions."],
      );
    }
  }
}
