/**
 * Dispatcher, spec-embedding, and Document facade tests.
 * Mirrors the Python stage-4 cases (python/tests/test_dispatch.py).
 */
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { unzipSync } from "fflate";
import { beforeAll, describe, expect, it } from "vitest";

import {
  Document,
  ERROR_CODES,
  MVP_TOOLS,
  Session,
  TOOL_SPECS,
  ToolError,
  anthropicTools,
  call,
  dispatch,
  openaiTools,
  toolSchemas,
} from "../src/index.js";
import { DEFAULT_PARTS, buildDocx, docWithBody } from "./fixtures.js";

const PARA_TRACKED_DUP_IDS =
  "<w:p>" +
  '<w:r><w:t xml:space="preserve">Payment due in </w:t></w:r>' +
  '<w:del w:id="1" w:author="J.Doe" w:date="2026-01-01T00:00:00Z">' +
  "<w:r><w:delText>30</w:delText></w:r></w:del>" +
  '<w:ins w:id="1" w:author="J.Doe" w:date="2026-01-01T00:00:00Z">' +
  "<w:r><w:t>45</w:t></w:r></w:ins>" +
  '<w:r><w:t xml:space="preserve"> days</w:t></w:r>' +
  "</w:p>";

let docxPath: string;

beforeAll(() => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-dispatch-"));
  docxPath = path.join(dir, "fixture.docx");
  fs.writeFileSync(docxPath, buildDocx());
});

function captureToolError(fn: () => unknown): ToolError {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(ToolError);
    return e as ToolError;
  }
  throw new Error("expected a ToolError");
}

describe("dispatch", () => {
  it("routes MVP tools", () => {
    const session = new Session();
    const result = dispatch(session, "docx_open", { path: docxPath }) as Record<string, unknown>;
    expect(result["doc_id"]).toBe("d1");
    const outline = dispatch(session, "docx_outline", { doc_id: "d1" }) as Record<string, unknown>;
    expect(Array.isArray(outline["outline"])).toBe(true);
  });

  it("missing required arg is invalid_args", () => {
    const err = captureToolError(() => dispatch(new Session(), "docx_read", {}));
    expect(err.code).toBe("invalid_args");
    expect(err.message).toContain("doc_id");
    expect(err.message.startsWith("docx_read:")).toBe(true);
  });

  it("multiple missing args are all named", () => {
    const err = captureToolError(() => dispatch(new Session(), "docx_replace", { doc_id: "d1" }));
    expect(err.code).toBe("invalid_args");
    expect(err.message).toContain("old");
    expect(err.message).toContain("new");
  });

  it("non-object args is invalid_args", () => {
    const err = captureToolError(() => dispatch(new Session(), "docx_read", ["d1"]));
    expect(err.code).toBe("invalid_args");
  });

  it("every spec tool is now implemented (no not_implemented surface)", () => {
    // Phase 2 stage 3 completes the surface: every spec tool routes to a
    // handler, so nothing in spec/tools/ returns not_implemented anymore.
    const undeclined = TOOL_SPECS.map((t) => t.name).filter((name) => !MVP_TOOLS.has(name));
    expect(undeclined).toEqual([]);
  });

  it("unknown (non-spec) tool is not_implemented", () => {
    const err = captureToolError(() => dispatch(new Session(), "docx_frobnicate", {}));
    expect(err.code).toBe("not_implemented");
  });

  it("every spec tool dispatches (validates args, never declines)", () => {
    const session = new Session();
    dispatch(session, "docx_open", { path: docxPath });
    for (const tool of TOOL_SPECS.map((t) => t.name)) {
      // Calling with empty args should never yield not_implemented; it either
      // succeeds or raises a domain error (invalid_args/doc_not_found/etc.).
      try {
        dispatch(session, tool, {});
      } catch (e) {
        expect(e).toBeInstanceOf(ToolError);
        expect((e as ToolError).code).not.toBe("not_implemented");
      }
    }
  });

  it("extra args are ignored", () => {
    const result = dispatch(new Session(), "docx_open", {
      path: docxPath,
      unexpected: true,
    }) as Record<string, unknown>;
    expect(result["doc_id"]).toBe("d1");
  });

  it("tool errors carry spec codes", () => {
    const err = captureToolError(() => dispatch(new Session(), "docx_read", { doc_id: "d404" }));
    expect(err.code).toBe("doc_not_found");
    expect(ERROR_CODES.has(err.code)).toBe(true);
  });

  it("the default session persists across call()s", async () => {
    const opened = await call("docx_open", { path: docxPath });
    const read = await call("docx_read", { doc_id: opened["doc_id"] as string });
    expect(String(read["content"])).toContain("Master Services Agreement");
  });
});

describe("spec surface", () => {
  it("toolSchemas cover all spec tools", () => {
    const schemas = toolSchemas();
    const names = schemas.map((s) => s.name);
    expect(names).toEqual([...names].sort());
    expect(new Set(names)).toEqual(new Set(TOOL_SPECS.map((t) => t.name)));
    for (const tool of MVP_TOOLS) expect(names).toContain(tool);
    for (const schema of schemas) {
      expect(typeof schema.description).toBe("string");
      expect(schema.input_schema["type"]).toBe("object");
    }
  });

  it("openaiTools shape", () => {
    const tools = openaiTools();
    expect(tools.every((t) => t.type === "function")).toBe(true);
    const replace = tools.find((t) => t.function.name === "docx_replace");
    expect(replace?.function.parameters["required"]).toEqual(["doc_id", "old", "new"]);
  });

  it("anthropicTools shape", () => {
    const tools = anthropicTools();
    const save = tools.find((t) => t.name === "docx_save");
    const properties = save?.input_schema["properties"] as Record<string, unknown>;
    expect(Object.keys(properties)).toContain("doc_id");
    expect(save?.input_schema["required"]).toEqual(["doc_id", "path"]);
  });

  it("returned schemas are copies", () => {
    const first = anthropicTools();
    (first[0]?.input_schema["properties"] as Record<string, unknown>)["injected"] = {};
    const fresh = anthropicTools();
    expect(
      Object.keys(fresh[0]?.input_schema["properties"] as Record<string, unknown>),
    ).not.toContain("injected");
  });

  it("error codes include the spec set", () => {
    for (const code of [
      "anchor_stale",
      "doc_not_found",
      "invalid_args",
      "not_implemented",
      "validation_failed",
    ]) {
      expect(ERROR_CODES.has(code)).toBe(true);
    }
  });
});

describe("Document facade", () => {
  it("open, paragraphs, find, replace, save", async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-document-"));
    const doc = await Document.open(docxPath);
    const paragraphs = doc.paragraphs();
    expect(paragraphs[0]?.anchor).toBe("P1#515a");
    expect(paragraphs[0]?.text).toBe("Master Services Agreement");

    const found = doc.find("five (5)");
    expect(found?.anchor).toBe(paragraphs[1]?.anchor);

    expect(doc.dirty).toBe(false);
    const result = (await doc.call("docx_replace", {
      old: "five (5) years",
      new: "three (3) years",
    })) as Record<string, unknown>;
    expect(result["n_replaced"]).toBe(1);
    expect(doc.dirty).toBe(true);

    const out = path.join(dir, "out.docx");
    const saved = await doc.save(out);
    expect(saved.ok).toBe(true);
    expect(doc.dirty).toBe(false);
    const entries = unzipSync(fs.readFileSync(out));
    const documentXml = new TextDecoder().decode(entries["word/document.xml"]);
    expect(documentXml).toContain("three (3) years");
  });

  it("opens from bytes", async () => {
    const doc = await Document.open(buildDocx());
    expect(doc.paragraphs()).toHaveLength(3);
  });

  it("save runs the gate", async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-document-"));
    const doc = await Document.open(
      buildDocx({ ...DEFAULT_PARTS, "word/document.xml": docWithBody(PARA_TRACKED_DUP_IDS) }),
    );
    let err: ToolError | null = null;
    try {
      await doc.save(path.join(dir, "out.docx"));
    } catch (e) {
      expect(e).toBeInstanceOf(ToolError);
      err = e as ToolError;
    }
    expect(err?.code).toBe("validation_failed");
  });
});
