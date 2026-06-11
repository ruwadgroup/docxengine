/**
 * Result-shape conformance: every spec tool routes through `dispatch()` and
 * returns an object whose keys are all declared in the tool's spec
 * `result_schema.properties`. This is the TS mirror of Python's
 * test_result_shapes — it spot-checks each tool's success result against the
 * published contract so the dispatcher surface can never drift from spec/tools.
 *
 * Each tool gets one representative successful call built on the shared
 * fixtures. The check is a subset assertion (result keys ⊆ declared props):
 * a tool may legitimately omit optional fields, but it must never invent a key
 * the spec does not declare.
 */
import { strToU8, zipSync, type Zippable } from "fflate";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { describe, expect, it } from "vitest";

import { Session, TOOL_SPECS, dispatch } from "../src/index.js";
import { DEFAULT_PARTS, type DocxParts, buildDocx, docWithBody } from "./fixtures.js";

function u32(n: number): number[] {
  return [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255];
}

/** A minimal PNG (IHDR only) — enough for media dimension parsing. */
function makePng(w: number, h: number): Uint8Array {
  const ihdr = [
    0,
    0,
    0,
    13,
    0x49,
    0x48,
    0x44,
    0x52,
    ...u32(w),
    ...u32(h),
    8,
    6,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
  ];
  return new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, ...ihdr]);
}

/** Write a .docx with the given parts to a temp file and return the path. */
function writeDocxFile(name: string, parts: DocxParts): string {
  const zippable: Zippable = {};
  for (const [partName, xml] of Object.entries(parts)) zippable[partName] = strToU8(xml);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-shapes-"));
  const file = path.join(dir, name);
  fs.writeFileSync(file, zipSync(zippable, { level: 0 }));
  return file;
}

const RESULT_SCHEMAS = new Map(
  TOOL_SPECS.map((t) => [t.name, t.result_schema as Record<string, unknown> | undefined]),
);

/** Assert that every key in `result` is declared in the tool's result_schema. */
function assertShape(tool: string, result: unknown): void {
  const schema = RESULT_SCHEMAS.get(tool);
  expect(schema, `${tool} must declare a result_schema in spec/tools/`).toBeDefined();
  expect(result, `${tool} must return a plain object`).toBeTypeOf("object");
  expect(Array.isArray(result)).toBe(false);
  const props = (schema?.["properties"] ?? {}) as Record<string, unknown>;
  const declared = new Set(Object.keys(props));
  const actual = Object.keys(result as Record<string, unknown>);
  const undeclared = actual.filter((k) => !declared.has(k));
  expect(undeclared, `${tool} returned keys not in its result_schema`).toEqual([]);
  expect(actual.length, `${tool} returned an empty result object`).toBeGreaterThan(0);
}

/** First paragraph anchor for the open doc (a stable "after" target). */
function firstParaAnchor(session: Session, docId: string): string {
  return session
    .get(docId)
    .anchorIndex()
    .filter((e) => e.kind === "p")[0]!.anchor;
}

function secondParaAnchor(session: Session, docId: string): string {
  return session
    .get(docId)
    .anchorIndex()
    .filter((e) => e.kind === "p")[1]!.anchor;
}

describe("result shapes match spec result_schemas", () => {
  it("every spec tool declares a result_schema", () => {
    const missing = TOOL_SPECS.filter((t) => t.result_schema === undefined).map((t) => t.name);
    expect(missing).toEqual([]);
  });

  it("each tool's success result is a subset of its declared result_schema", () => {
    // One session, opened once; tools that mutate run in sequence against it.
    const session = new Session();
    const bytes = Buffer.from(buildDocx()).toString("base64");
    const open = dispatch(session, "docx_open", { bytes }) as Record<string, unknown>;
    const docId = open["doc_id"] as string;
    assertShape("docx_open", open);

    const seen = new Set<string>(["docx_open"]);
    const check = (tool: string, args: Record<string, unknown>): Record<string, unknown> => {
      const res = dispatch(session, tool, args) as Record<string, unknown>;
      assertShape(tool, res);
      seen.add(tool);
      return res;
    };

    // Read surface.
    check("docx_outline", { doc_id: docId });
    check("docx_read", { doc_id: docId });
    check("docx_search", { doc_id: docId, query: "term" });

    // Edit surface. Anchors are content-hashed, so re-derive a fresh anchor
    // immediately before each mutating call rather than reusing a cached one.
    check("docx_replace", { doc_id: docId, old: "five (5)", new: "three (3)" });
    check("docx_edit_paragraph", {
      doc_id: docId,
      anchor: firstParaAnchor(session, docId),
      text: "New heading text",
    });
    const ins = check("docx_insert", {
      doc_id: docId,
      after: firstParaAnchor(session, docId),
      content: "Inserted paragraph.",
    });
    check("docx_delete", {
      doc_id: docId,
      anchor: (ins["new_anchors"] as string[])[0],
    });
    check("docx_revision", { doc_id: docId, op: "list" });

    // Lifecycle / validation.
    check("docx_validate", { doc_id: docId });
    check("docx_repair", { doc_id: docId });
    const savePath = path.join(
      fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-shapes-save-")),
      "out.docx",
    );
    check("docx_save", { doc_id: docId, path: savePath });

    // Tables.
    check("docx_table", {
      doc_id: docId,
      op: "create",
      after: firstParaAnchor(session, docId),
      rows: 2,
      cols: 2,
    });

    // Styles + formatting.
    check("docx_style", { doc_id: docId, op: "list" });
    check("docx_format", {
      doc_id: docId,
      anchor: firstParaAnchor(session, docId),
      props: { bold: true },
    });

    // Lists.
    check("docx_list", {
      doc_id: docId,
      op: "create",
      after: firstParaAnchor(session, docId),
      kind: "ul",
      items: [{ text: "Bullet" }],
    });

    // Comments.
    check("docx_comment", {
      doc_id: docId,
      op: "add",
      anchor: firstParaAnchor(session, docId),
      text: "A note",
      author: "Reviewer",
    });

    // Sections.
    check("docx_section", { doc_id: docId, op: "list" });

    // Media.
    const imgPath = path.join(
      fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-shapes-img-")),
      "logo.png",
    );
    fs.writeFileSync(imgPath, makePng(64, 48));
    check("docx_media", {
      doc_id: docId,
      op: "insert",
      after: secondParaAnchor(session, docId),
      image: imgPath,
      width_cm: 4,
    });

    // Fields.
    check("docx_field", {
      doc_id: docId,
      op: "insert_toc",
      after: firstParaAnchor(session, docId),
      levels: 3,
    });

    // Convert (md is the parity path; renderer-free).
    check("docx_convert", { doc_id: docId, to: "md" });

    // Render preview (structural fallback; soffice not installed).
    check("docx_render_preview", { doc_id: docId });

    // Create (own doc, no doc_id needed).
    check("docx_create", { content_md: "# Title\n\nA paragraph.\n" });

    // Template fill (needs a template file with a placeholder).
    const tmplPath = writeDocxFile("template.docx", {
      ...DEFAULT_PARTS,
      "word/document.xml": docWithBody(
        "<w:p><w:r><w:t>Hello {{name}}</w:t></w:r></w:p><w:sectPr/>",
      ),
    });
    check("docx_template_fill", { template: tmplPath, data: { name: "World" } });

    // Sanity: every spec tool was exercised exactly once.
    const all = new Set(TOOL_SPECS.map((t) => t.name));
    const unexercised = [...all].filter((name) => !seen.has(name));
    expect(unexercised, "every spec tool must have a result-shape case").toEqual([]);
  });
});
