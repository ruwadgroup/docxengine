import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { strFromU8, strToU8, unzipSync } from "fflate";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { Package, ToolError } from "../src/index.js";
import { DEFAULT_PARTS, DOCUMENT_XML, buildDocx } from "./fixtures.js";

describe("Package.open", () => {
  it("preserves the source zip entry order", () => {
    const pkg = Package.open(buildDocx());
    expect(pkg.entryNames()).toEqual(Object.keys(DEFAULT_PARTS));
  });

  it("exposes raw part bytes verbatim", () => {
    const pkg = Package.open(buildDocx());
    expect(pkg.part("word/document.xml")).toEqual(strToU8(DOCUMENT_XML));
    expect(pkg.partText("word/document.xml")).toBe(DOCUMENT_XML);
  });

  it("rejects non-zip bytes with open_failed", () => {
    let err: unknown;
    try {
      Package.open(strToU8("this is not a zip"));
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(ToolError);
    expect((err as ToolError).code).toBe("open_failed");
    expect((err as ToolError).toJSON()).toMatchObject({ error: "open_failed" });
  });

  it("rejects a zip without [Content_Types].xml", () => {
    const bogus = buildDocx({ "word/document.xml": DOCUMENT_XML });
    expect(() => Package.open(bogus)).toThrowError(/Content_Types/);
  });

  it("rejects an unreadable path with open_failed", () => {
    try {
      Package.open("/nonexistent/nope.docx");
      expect.unreachable();
    } catch (e) {
      expect((e as ToolError).code).toBe("open_failed");
    }
  });
});

describe("untouched-part byte round-trip (§9)", () => {
  it("re-saves every untouched part with byte-identical decompressed content", () => {
    const pkg = Package.open(buildDocx());
    const out = unzipSync(pkg.toBytes());
    expect(Object.keys(out)).toEqual(Object.keys(DEFAULT_PARTS));
    for (const [name, xml] of Object.entries(DEFAULT_PARTS)) {
      expect(out[name], name).toEqual(strToU8(xml));
    }
  });

  it("zeroes entry timestamps to DOS 1980-01-01 00:00:00", () => {
    const bytes = Buffer.from(Package.open(buildDocx()).toBytes());
    // Local file headers: PK\x03\x04 … mod time at +10 (2 bytes), mod date at +12.
    let found = 0;
    for (let i = 0; i + 14 <= bytes.length; i++) {
      if (
        bytes[i] === 0x50 &&
        bytes[i + 1] === 0x4b &&
        bytes[i + 2] === 0x03 &&
        bytes[i + 3] === 0x04
      ) {
        expect(bytes.readUInt16LE(i + 10)).toBe(0); // time 00:00:00
        expect(bytes.readUInt16LE(i + 12)).toBe(0x0021); // date 1980-01-01
        found++;
      }
    }
    expect(found).toBe(Object.keys(DEFAULT_PARTS).length);
  });

  it("is deterministic: two serializations are byte-identical", () => {
    const pkg = Package.open(buildDocx());
    expect(pkg.toBytes()).toEqual(pkg.toBytes());
  });
});

describe("setPart / dirty tracking", () => {
  it("marks modified parts dirty and re-serializes only them", () => {
    const pkg = Package.open(buildDocx());
    expect(pkg.dirty.size).toBe(0);
    const modified = DOCUMENT_XML.replace("Master", "Updated");
    pkg.setPart("word/document.xml", modified);
    expect(pkg.isDirty("word/document.xml")).toBe(true);
    expect([...pkg.dirty]).toEqual(["word/document.xml"]);

    const out = unzipSync(pkg.toBytes());
    expect(strFromU8(out["word/document.xml"] as Uint8Array)).toBe(modified);
    // Untouched siblings stay byte-identical.
    expect(out["word/styles.xml"]).toEqual(strToU8(DEFAULT_PARTS["word/styles.xml"] as string));
    expect(out["_rels/.rels"]).toEqual(strToU8(DEFAULT_PARTS["_rels/.rels"] as string));
  });

  it("appends new parts after the originals, in creation order", () => {
    const pkg = Package.open(buildDocx());
    pkg.setPart("word/comments.xml", "<w:comments/>");
    pkg.setPart("word/footnotes.xml", "<w:footnotes/>");
    expect(pkg.entryNames()).toEqual([
      ...Object.keys(DEFAULT_PARTS),
      "word/comments.xml",
      "word/footnotes.xml",
    ]);
    const out = unzipSync(pkg.toBytes());
    expect(Object.keys(out)).toEqual(pkg.entryNames());
    expect(strFromU8(out["word/comments.xml"] as Uint8Array)).toBe("<w:comments/>");
  });
});

describe("content types and rels parsing", () => {
  it("parses Defaults and Overrides", () => {
    const pkg = Package.open(buildDocx());
    const ct = pkg.contentTypes();
    expect(ct.defaults.get("rels")).toContain("relationships+xml");
    expect(ct.overrides.get("/word/document.xml")).toContain("wordprocessingml.document.main+xml");
    expect(pkg.contentTypeOf("word/document.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    );
    expect(pkg.contentTypeOf("_rels/.rels")).toBe(
      "application/vnd.openxmlformats-package.relationships+xml",
    );
    expect(pkg.contentTypeOf("word/styles.xml")).toBe("application/xml");
    expect(pkg.contentTypeOf("word/unknown.bin")).toBeUndefined();
  });

  it("parses package-level and part-level relationships", () => {
    const pkg = Package.open(buildDocx());
    expect(pkg.rels()).toEqual([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        target: "word/document.xml",
        targetMode: "Internal",
      },
    ]);
    const docRels = pkg.rels("word/document.xml");
    expect(docRels).toHaveLength(1);
    expect(docRels[0]).toMatchObject({ id: "rId1", target: "styles.xml" });
    expect(pkg.rels("word/styles.xml")).toEqual([]);
  });

  it("maps part names to their rels parts", () => {
    expect(Package.relsPartFor()).toBe("_rels/.rels");
    expect(Package.relsPartFor("word/document.xml")).toBe("word/_rels/document.xml.rels");
    expect(Package.relsPartFor("word/glossary/document.xml")).toBe(
      "word/glossary/_rels/document.xml.rels",
    );
  });
});

describe("save(path)", () => {
  let dir: string;
  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-opc-"));
  });
  afterEach(() => {
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it("writes atomically and the result re-opens cleanly", () => {
    const pkg = Package.open(buildDocx());
    const dest = path.join(dir, "out.docx");
    pkg.save(dest);
    expect(fs.readdirSync(dir)).toEqual(["out.docx"]); // no temp file left behind
    const reopened = Package.open(dest);
    expect(reopened.entryNames()).toEqual(pkg.entryNames());
    expect(reopened.partText("word/document.xml")).toBe(DOCUMENT_XML);
  });

  it("raises save_failed on an unwritable destination", () => {
    const pkg = Package.open(buildDocx());
    try {
      pkg.save(path.join(dir, "missing-subdir", "out.docx"));
      expect.unreachable();
    } catch (e) {
      expect((e as ToolError).code).toBe("save_failed");
    }
  });
});
