/**
 * Package validator and mechanical repair (algorithms.md §8/§8a).
 *
 * `validateDoc` runs the five MVP checks (a–e) and returns the pinned issue
 * list — exact ordering, messages, and fix hints matter, because the
 * conformance harness deep-compares `docx_validate` results across the Python
 * and TypeScript implementations. `repairDoc` applies the §8a fixes by
 * splicing raw part text (never re-serializing), then re-validates.
 */
import { maxRevisionId } from "./edits.js";
import { Package } from "./opc.js";
import type { DocHandle } from "./session.js";
import { type SpliceEdit, attrs, elementExtent, nextTag, spliceAll } from "./xmlscan.js";

export interface ValidationIssue {
  severity: "error" | "warning";
  part: string;
  message: string;
  fix_hint: string;
}

const CONTENT_TYPES_PART = "[Content_Types].xml";
export const COMMENTS_PART = "word/comments.xml";
export const FOOTNOTES_PART = "word/footnotes.xml";

const REVISION_NAMES = ["w:ins", "w:del"] as const;

/**
 * Relationship types (last path segment) Word consumes without an explicit
 * r:id reference in the document part — exempt from the unreferenced warning.
 */
const IMPLICIT_REL_TYPES = new Set([
  "styles",
  "settings",
  "webSettings",
  "fontTable",
  "numbering",
  "theme",
  "customXml",
  "comments",
  "commentsExtended",
  "footnotes",
  "endnotes",
  "glossaryDocument",
]);

/** §8a content types for repaired `Default` entries; anything else gets octet-stream. */
const DEFAULT_CONTENT_TYPES: Readonly<Record<string, string>> = {
  rels: "application/vnd.openxmlformats-package.relationships+xml",
  xml: "application/xml",
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  gif: "image/gif",
};
const FALLBACK_CONTENT_TYPE = "application/octet-stream";

const R_ID_RE = /\sr:id\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
const R_REF_RE = /\sr:(?:id|embed|link)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
const W_ID_VALUE_RE = /(\sw:id\s*=\s*")([^"]*)(")/;

// ---------------------------------------------------------------------------
// Scan helpers
// ---------------------------------------------------------------------------

/** Package entries subject to check a, in zip order (directories skipped). */
function contentParts(pkg: Package): string[] {
  return pkg.entryNames().filter((n) => !n.endsWith("/") && n !== CONTENT_TYPES_PART);
}

/** The part's extension (lowercased), `""` when the basename has none. */
function extensionOf(partName: string): string {
  const base = partName.slice(partName.lastIndexOf("/") + 1);
  const dot = base.lastIndexOf(".");
  return dot < 0 ? "" : base.slice(dot + 1).toLowerCase();
}

function relsParts(pkg: Package): string[] {
  return pkg.entryNames().filter((n) => n.endsWith(".rels"));
}

/** Inverse of `Package.relsPartFor` (`null` = the package root). */
function sourcePartFor(relsName: string): string | null {
  const slash = relsName.lastIndexOf("/");
  const dir = slash < 0 ? "" : relsName.slice(0, slash);
  const base = slash < 0 ? relsName : relsName.slice(slash + 1);
  const parentSlash = dir.lastIndexOf("/");
  const parent = parentSlash < 0 ? "" : dir.slice(0, parentSlash); // strip trailing "_rels"
  const source = base.endsWith(".rels") ? base.slice(0, -".rels".length) : base;
  if (!source) return null;
  return parent ? `${parent}/${source}` : source;
}

/** Resolve an Internal relationship target to a package part name. */
function resolveRelTarget(sourcePart: string | null, target: string): string {
  const joined = target.startsWith("/")
    ? target.slice(1)
    : (sourcePart ?? "").slice(0, (sourcePart ?? "").lastIndexOf("/") + 1) + target;
  const segs: string[] = [];
  for (const s of joined.split("/")) {
    if (s === "" || s === ".") continue;
    if (s === "..") segs.pop();
    else segs.push(s);
  }
  return segs.join("/");
}

/** Distinct matched attribute values in order of first occurrence. */
function attrValues(text: string, re: RegExp): string[] {
  const out: string[] = [];
  re.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const value = m[1] ?? m[2] ?? "";
    if (!out.includes(value)) out.push(value);
  }
  return out;
}

interface NamedElement {
  name: string;
  /** Full element extent [start, end). */
  start: number;
  end: number;
  /** End of the start tag (`>` exclusive) — the attribute search region. */
  startTagEnd: number;
  id: string;
}

/** Every named element with its `w:id`, in document order. */
function elementIds(xml: string, names: readonly string[]): NamedElement[] {
  const out: NamedElement[] = [];
  let i = 0;
  for (;;) {
    const t = nextTag(xml, i);
    if (!t) return out;
    i = t.end;
    if (t.kind === "end" || !names.includes(t.name)) continue;
    const el = elementExtent(xml, t);
    out.push({
      name: t.name,
      start: el.start,
      end: el.end,
      startTagEnd: t.end,
      id: attrs(xml, t)["w:id"] ?? "",
    });
  }
}

/** Reference ids (distinct, first occurrence) and `[id, w:type]` definitions. */
function storyRefsAndDefs(
  pkg: Package,
  mainText: string,
  refName: string,
  defsPart: string,
  defName: string,
): { refs: string[]; defs: [string, string][] } {
  const refs: string[] = [];
  for (const el of elementIds(mainText, [refName])) {
    if (!refs.includes(el.id)) refs.push(el.id);
  }
  const defs: [string, string][] = [];
  if (pkg.has(defsPart)) {
    const xml = pkg.partText(defsPart);
    let i = 0;
    for (;;) {
      const t = nextTag(xml, i);
      if (!t) break;
      i = t.end;
      if (t.kind === "end" || t.name !== defName) continue;
      const a = attrs(xml, t);
      defs.push([a["w:id"] ?? "", a["w:type"] ?? ""]);
    }
  }
  return { refs, defs };
}

// ---------------------------------------------------------------------------
// Validation (§8/§8a)
// ---------------------------------------------------------------------------

/** All §8 checks, issues in the pinned a → e order. */
export function validateDoc(doc: DocHandle): ValidationIssue[] {
  const pkg = doc.pkg;
  const main = doc.documentPartName;
  const mainText = doc.documentXml();
  const mainRelsName = Package.relsPartFor(main);
  const issues: ValidationIssue[] = [];

  // a — content-type coverage, package entries in zip order.
  for (const name of contentParts(pkg)) {
    if (pkg.contentTypeOf(name) !== undefined) continue;
    const ext = extensionOf(name);
    issues.push({
      severity: "error",
      part: name,
      message:
        `Part ${name} is not covered by [Content_Types].xml ` +
        `(no Override, no Default for extension '${ext}').`,
      fix_hint: "docx_repair adds a content-type Default for the extension.",
    });
  }

  // b — r:id references of the document part resolve in its rels.
  const relIds = new Set(pkg.rels(main).map((r) => r.id));
  for (const rid of attrValues(mainText, R_ID_RE)) {
    if (relIds.has(rid)) continue;
    issues.push({
      severity: "error",
      part: main,
      message: `r:id ${rid} is referenced in ${main} but not defined in ${mainRelsName}.`,
      fix_hint:
        "Add the missing relationship or remove the referencing element; not auto-repairable.",
    });
  }

  // c — every non-External relationship target exists; then unreferenced warnings.
  for (const relsName of relsParts(pkg)) {
    const source = sourcePartFor(relsName);
    for (const rel of pkg.rels(source ?? undefined)) {
      if (rel.targetMode === "External") continue;
      const target = resolveRelTarget(source, rel.target);
      if (pkg.has(target)) continue;
      issues.push({
        severity: "error",
        part: relsName,
        message: `Relationship ${rel.id} targets missing part ${target}.`,
        fix_hint: "docx_repair drops the orphaned relationship.",
      });
    }
  }
  const referenced = new Set(attrValues(mainText, R_REF_RE));
  for (const rel of pkg.rels(main)) {
    const stripped = rel.type.replace(/\/+$/, "");
    const shortType = stripped.slice(stripped.lastIndexOf("/") + 1);
    if (IMPLICIT_REL_TYPES.has(shortType) || referenced.has(rel.id)) continue;
    issues.push({
      severity: "warning",
      part: mainRelsName,
      message: `Relationship ${rel.id} (${shortType}) is never referenced.`,
      fix_hint: "Harmless; remove the unused relationship to tidy the package.",
    });
  }

  // d — w:ins/w:del id uniqueness (counted together), first-occurrence order.
  const counts = new Map<string, number>();
  for (const el of elementIds(mainText, REVISION_NAMES)) {
    counts.set(el.id, (counts.get(el.id) ?? 0) + 1);
  }
  for (const [revId, n] of counts) {
    if (n <= 1) continue;
    issues.push({
      severity: "error",
      part: main,
      message: `Duplicate revision id ${revId} on ${n} w:ins/w:del elements.`,
      fix_hint: "docx_repair renumbers the later duplicates.",
    });
  }

  // e — comment and footnote references resolve both directions.
  const stories: [string, string, string, string][] = [
    ["Comment", "w:commentReference", COMMENTS_PART, "w:comment"],
    ["Footnote", "w:footnoteReference", FOOTNOTES_PART, "w:footnote"],
  ];
  for (const [noun, refName, defsPart, defName] of stories) {
    const { refs, defs } = storyRefsAndDefs(pkg, mainText, refName, defsPart, defName);
    const defIds = new Set(defs.map(([defId]) => defId));
    for (const refId of refs) {
      if (defIds.has(refId)) continue;
      issues.push({
        severity: "error",
        part: defsPart,
        message: `${noun} id=${refId} referenced in body but missing.`,
        fix_hint: "docx_repair removes the orphaned reference.",
      });
    }
    for (const [defId, defType] of defs) {
      if (refs.includes(defId) || defType === "separator" || defType === "continuationSeparator") {
        continue;
      }
      issues.push({
        severity: "warning",
        part: defsPart,
        message: `${noun} id=${defId} defined but never referenced.`,
        fix_hint: "Harmless; delete the unused definition to tidy the package.",
      });
    }
  }

  return issues;
}

/** §8a: valid iff no error-severity issue (warnings never block). */
export function isValid(issues: readonly ValidationIssue[]): boolean {
  return !issues.some((issue) => issue.severity === "error");
}

// ---------------------------------------------------------------------------
// Repair (§8/§8a)
// ---------------------------------------------------------------------------

function dropOrphanedRelationships(pkg: Package, fixed: string[]): void {
  for (const relsName of relsParts(pkg)) {
    const source = sourcePartFor(relsName);
    const xml = pkg.partText(relsName);
    const edits: SpliceEdit[] = [];
    let i = 0;
    for (;;) {
      const t = nextTag(xml, i);
      if (!t) break;
      i = t.end;
      if (t.kind === "end" || t.name !== "Relationship") continue;
      const a = attrs(xml, t);
      if ((a["TargetMode"] ?? "Internal") === "External") continue;
      const target = resolveRelTarget(source, a["Target"] ?? "");
      if (pkg.has(target)) continue;
      const el = elementExtent(xml, t);
      edits.push({ start: el.start, end: el.end, text: "" });
      fixed.push(`removed orphaned relationship ${a["Id"] ?? ""} (${relsName})`);
      i = el.end;
    }
    if (edits.length > 0) pkg.setPart(relsName, spliceAll(xml, edits));
  }
}

function addMissingContentTypeDefaults(pkg: Package, fixed: string[]): void {
  const missing: string[] = [];
  for (const name of contentParts(pkg)) {
    if (pkg.contentTypeOf(name) !== undefined) continue;
    const ext = extensionOf(name);
    if (ext && !missing.includes(ext)) missing.push(ext); // extension-less: not fixable here
  }
  if (missing.length === 0) return;
  const xml = pkg.partText(CONTENT_TYPES_PART);
  const close = xml.lastIndexOf("</Types>");
  if (close < 0) return;
  const inserted = missing
    .map(
      (ext) =>
        `<Default Extension="${ext}" ` +
        `ContentType="${DEFAULT_CONTENT_TYPES[ext] ?? FALLBACK_CONTENT_TYPE}"/>`,
    )
    .join("");
  pkg.setPart(CONTENT_TYPES_PART, spliceAll(xml, [{ start: close, end: close, text: inserted }]));
  for (const ext of missing) fixed.push(`added content-type Default for extension '${ext}'`);
}

function renumberDuplicateRevisionIds(doc: DocHandle, fixed: string[]): void {
  const pkg = doc.pkg;
  const main = doc.documentPartName;
  const xml = pkg.partText(main);
  let nextId = maxRevisionId(xml) + 1; // §8a: later duplicates take max+1 onward
  const seen = new Set<string>();
  const edits: SpliceEdit[] = [];
  for (const el of elementIds(xml, REVISION_NAMES)) {
    const m = W_ID_VALUE_RE.exec(xml.slice(el.start, el.startTagEnd));
    if (!m) continue;
    const value = m[2] as string;
    if (!seen.has(value)) {
      seen.add(value);
      continue;
    }
    const valueStart = el.start + m.index + (m[1] as string).length;
    edits.push({ start: valueStart, end: valueStart + value.length, text: String(nextId) });
    fixed.push(`renumbered duplicate revision id ${value} -> ${nextId}`);
    nextId += 1;
  }
  if (edits.length > 0) {
    pkg.setPart(main, spliceAll(xml, edits));
    doc.invalidate();
  }
}

function removeOrphanedStoryReferences(doc: DocHandle, fixed: string[]): void {
  const pkg = doc.pkg;
  const main = doc.documentPartName;
  const stories: [string, string, string, string, string[]][] = [
    [
      "comment",
      "w:commentReference",
      COMMENTS_PART,
      "w:comment",
      ["w:commentRangeStart", "w:commentRangeEnd"],
    ],
    ["footnote", "w:footnoteReference", FOOTNOTES_PART, "w:footnote", []],
  ];
  for (const [noun, refName, defsPart, defName, rangeNames] of stories) {
    const data = pkg.partText(main);
    const { refs, defs } = storyRefsAndDefs(pkg, data, refName, defsPart, defName);
    const defIds = new Set(defs.map(([defId]) => defId));
    const orphaned = new Set(refs.filter((refId) => !defIds.has(refId)));
    if (orphaned.size === 0) continue;
    const edits: SpliceEdit[] = [];
    for (const el of elementIds(data, [refName, ...rangeNames])) {
      if (!orphaned.has(el.id)) continue;
      edits.push({ start: el.start, end: el.end, text: "" });
      if (el.name === refName) fixed.push(`removed orphaned ${noun} reference id=${el.id}`);
    }
    if (edits.length > 0) {
      pkg.setPart(main, spliceAll(data, edits));
      doc.invalidate();
    }
  }
}

/**
 * Apply the §8a fixes in order; returns `{fixed, remaining}`.
 *
 * `remaining` is the message of every error-severity issue still present
 * after re-validation. The caller reports both; fixes are applied either way
 * (`repair_incomplete` is reserved for a fix that cannot be applied at all).
 */
export function repairDoc(doc: DocHandle): { fixed: string[]; remaining: string[] } {
  const fixed: string[] = [];
  dropOrphanedRelationships(doc.pkg, fixed);
  addMissingContentTypeDefaults(doc.pkg, fixed);
  renumberDuplicateRevisionIds(doc, fixed);
  removeOrphanedStoryReferences(doc, fixed);
  const remaining = validateDoc(doc)
    .filter((issue) => issue.severity === "error")
    .map((issue) => issue.message);
  return { fixed, remaining };
}
