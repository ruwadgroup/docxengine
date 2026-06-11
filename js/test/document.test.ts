/** Native `Document` API: full-coverage parity with the tool surface. */
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { describe, expect, it } from "vitest";

import { Document, DocumentParagraph, Session, ToolError } from "../src/index.js";
import { DEFAULT_PARTS, DOCUMENT_RELS_XML, buildDocx, docWithBody } from "./fixtures.js";

const TOOL_METHODS = [
  "outline",
  "read",
  "search",
  "replace",
  "editParagraph",
  "insert",
  "delete",
  "revision",
  "comment",
  "table",
  "style",
  "format",
  "list",
  "section",
  "media",
  "field",
  "validate",
  "repair",
  "convert",
  "renderPreview",
  "save",
  "toBytes",
] as const;

function templateBytes(): Uint8Array {
  return buildDocx({
    ...DEFAULT_PARTS,
    "word/document.xml": docWithBody("<w:p><w:r><w:t>Hello {{name}}</w:t></w:r></w:p>"),
  });
}

function corruptBytes(): Uint8Array {
  // An orphan relationship: opens cleanly, fails validation (error severity).
  const orphan = DOCUMENT_RELS_XML.replace(
    "</Relationships>",
    '<Relationship Id="rId9" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
      'Target="media/missing.png"/></Relationships>',
  );
  return buildDocx({ ...DEFAULT_PARTS, "word/_rels/document.xml.rels": orphan });
}

describe("Document native API", () => {
  it("exposes a method per tool plus the constructors", async () => {
    const doc = await Document.create({ contentMd: "# Hi" });
    for (const name of TOOL_METHODS) {
      expect(typeof (doc as unknown as Record<string, unknown>)[name]).toBe("function");
    }
    for (const ctor of ["open", "create", "fillTemplate", "attach"] as const) {
      expect(typeof Document[ctor]).toBe("function");
    }
  });

  it("opens from bytes and reports paragraphs with styleId", async () => {
    const doc = await Document.open(buildDocx());
    const paras = doc.paragraphs();
    expect(paras[0].anchor).toBe("P1#515a");
    expect(paras[0].text).toBe("Master Services Agreement");
    expect(paras[0].style).toBe("Heading1"); // raw styleId, matching Python
    expect(paras[1].style).toBeNull();
  });

  it("find returns a paragraph view; search returns match dicts", async () => {
    const doc = await Document.open(buildDocx());
    const para = doc.find("five (5)");
    expect(para).toBeInstanceOf(DocumentParagraph);
    expect(doc.find("nonexistent")).toBeNull();
    const hits = (await doc.search("five (5)")) as { matches: { anchor: string }[] };
    expect(hits.matches[0].anchor).toBe(para?.anchor);
  });

  it("creates and edits, round-tripping through the in-memory handle", async () => {
    const doc = await Document.create({ contentMd: "# Title\n\nThe term is five (5) years." });
    expect(doc.dirty).toBe(false);
    await doc.replace("five (5) years", "three (3) years");
    expect(doc.dirty).toBe(true);
    expect(doc.find("three (3) years")).not.toBeNull();
  });

  it("covers the edit surface (insert, revision, comment, table)", async () => {
    const doc = await Document.open(buildDocx());
    const anchor = doc.paragraphs()[0].anchor;
    const inserted = (await doc.insert("New intro.", { after: anchor })) as {
      new_anchors: string[];
    };
    expect(inserted.new_anchors.length).toBe(1);
    expect((await doc.revision("list")) as Record<string, unknown>).toHaveProperty("revisions");
    await doc.comment("add", { anchor, text: "note", author: "QA" });
    await doc.table("create", {
      after: anchor,
      rows: 2,
      cols: 2,
      data: [
        ["a", "b"],
        ["c", "d"],
      ],
    });
    expect(((await doc.validate()) as { valid: boolean }).valid).toBe(true);
  });

  it("drives the paragraph primitives", async () => {
    const doc = await Document.open(buildDocx());
    const para = doc.find("five (5)");
    expect(para).not.toBeNull();
    await para?.replace("five (5) years", "two (2) years");
    expect(doc.find("two (2) years")).not.toBeNull();
    const fresh = doc.paragraphs()[0];
    const out = (await fresh.insertAfter("After heading.")) as { new_anchors: string[] };
    expect(out.new_anchors.length).toBe(1);
  });

  it("converts to markdown", async () => {
    const doc = await Document.open(buildDocx());
    expect((await doc.convert("md")) as Record<string, unknown>).toHaveProperty("content");
  });

  it("save runs the validation gate and writes atomically", async () => {
    const doc = await Document.open(buildDocx());
    await doc.replace("five (5) years", "three (3) years");
    const out = path.join(os.tmpdir(), `docxengine-doc-${Date.now()}.docx`);
    const saved = await doc.save(out);
    expect(saved.ok).toBe(true);
    expect(doc.dirty).toBe(false);
    expect(fs.existsSync(out)).toBe(true);
    fs.rmSync(out);

    const bad = await Document.open(corruptBytes());
    await expect(bad.save(path.join(os.tmpdir(), "x.docx"))).rejects.toMatchObject({
      code: "validation_failed",
    });
  });

  it("toBytes round-trips and gates invalid packages (no filesystem)", async () => {
    const doc = await Document.open(buildDocx());
    await doc.replace("five (5) years", "three (3) years");
    const bytes = doc.toBytes();
    expect(bytes[0]).toBe(0x50); // "P"
    expect(doc.dirty).toBe(false);
    expect((await Document.open(bytes)).find("three (3) years")).not.toBeNull();

    let threw: ToolError | null = null;
    try {
      (await Document.open(corruptBytes())).toBytes();
    } catch (e) {
      threw = e as ToolError;
    }
    expect(threw?.code).toBe("validation_failed");
  });

  it("fills a template from bytes and from a path", async () => {
    const fromBytes = await Document.fillTemplate(templateBytes(), { name: "World" });
    expect(fromBytes.find("Hello World")).not.toBeNull();

    const tpl = path.join(os.tmpdir(), `docxengine-tpl-${Date.now()}.docx`);
    fs.writeFileSync(tpl, templateBytes());
    const fromPath = await Document.fillTemplate(tpl, { name: "Ada" });
    expect(fromPath.find("Hello Ada")).not.toBeNull();
    fs.rmSync(tpl);
  });

  it("fillTemplate strict raises on unfilled placeholders", async () => {
    await expect(
      Document.fillTemplate(templateBytes(), {}, { strict: true }),
    ).rejects.toMatchObject({ code: "placeholder_unfilled" });
  });

  it("attach shares the session's handle", async () => {
    const session = new Session();
    const a = await Document.open(buildDocx(), { session });
    const b = Document.attach(session, a.id);
    expect(b.id).toBe(a.id);
    expect(b.session).toBe(session);
    await a.replace("five (5) years", "three (3) years");
    expect(b.dirty).toBe(true); // same underlying handle
  });
});
