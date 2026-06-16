/**
 * Adversarial / fuzz tests for the §27 hardening (SECURITY.md threat model).
 *
 * Mirrors the Python suite (`python/tests/test_adversarial.py`): zip bombs (part
 * count, total size, per-part size, compression ratio), hostile XML (DTD/entity
 * declarations → XXE / billion-laughs), pathological XML nesting, and path
 * traversal normalization.
 */
import { strToU8, zipSync, type Zippable } from "fflate";
import { afterEach, describe, expect, it } from "vitest";

import { Package, ToolError } from "../src/index.js";
import { assertXmlDepth } from "../src/xmlscan.js";
import { DEFAULT_PARTS, type DocxParts, buildDocx, docWithBody } from "./fixtures.js";

const LIMIT_ENV = [
  "DOCXENGINE_MAX_PARTS",
  "DOCXENGINE_MAX_TOTAL_BYTES",
  "DOCXENGINE_MAX_PART_BYTES",
  "DOCXENGINE_MAX_COMPRESSION_RATIO",
  "DOCXENGINE_MAX_XML_DEPTH",
];

afterEach(() => {
  for (const k of LIMIT_ENV) delete process.env[k];
});

function docxWith(overrides: DocxParts): Uint8Array {
  return buildDocx({ ...DEFAULT_PARTS, ...overrides });
}

/** Open a package and force every XML part to be read (triggers §27 screening). */
function openAndReadAll(bytes: Uint8Array): Package {
  const pkg = Package.open(bytes);
  for (const name of pkg.entryNames()) {
    if (name.endsWith(".xml") || name.endsWith(".rels")) pkg.partText(name);
  }
  return pkg;
}

describe("decompression bombs", () => {
  it("refuses too many parts", () => {
    process.env["DOCXENGINE_MAX_PARTS"] = "3"; // fixture has 5
    try {
      Package.open(buildDocx());
      expect.unreachable("expected doc_too_large");
    } catch (e) {
      expect(e).toBeInstanceOf(ToolError);
      expect((e as ToolError).code).toBe("doc_too_large");
    }
  });

  it("refuses an oversized total", () => {
    process.env["DOCXENGINE_MAX_TOTAL_BYTES"] = "100";
    expect(() => Package.open(buildDocx())).toThrowError(/cap/);
  });

  it("refuses an oversized single part", () => {
    process.env["DOCXENGINE_MAX_PART_BYTES"] = "50";
    try {
      Package.open(buildDocx());
      expect.unreachable("expected doc_too_large");
    } catch (e) {
      expect((e as ToolError).code).toBe("doc_too_large");
    }
  });

  it("catches a high-ratio zip bomb above the floor", () => {
    const bomb = "A".repeat(256 * 1024);
    try {
      Package.open(docxWith({ "word/bomb.xml": bomb }));
      expect.unreachable("expected doc_too_large");
    } catch (e) {
      expect((e as ToolError).code).toBe("doc_too_large");
      expect((e as ToolError).message.toLowerCase()).toMatch(/ratio|zip bomb/);
    }
  });

  it("does not flag a small highly-compressible part (below the ratio floor)", () => {
    const ok = "A".repeat(4 * 1024);
    const pkg = Package.open(docxWith({ "word/note.xml": ok }));
    expect(pkg.partText("word/note.xml")).toBe(ok);
  });

  it("falls back to the default on an invalid env value", () => {
    process.env["DOCXENGINE_MAX_PARTS"] = "not-a-number";
    expect(() => Package.open(buildDocx())).not.toThrow();
  });
});

describe("hostile XML", () => {
  const billionLaughs =
    '<?xml version="1.0"?>\n' +
    '<!DOCTYPE lolz [<!ENTITY lol "lol">' +
    '<!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">]>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    "<w:body><w:p><w:r><w:t>&lol2;</w:t></w:r></w:p></w:body></w:document>";

  it("rejects a DOCTYPE/ENTITY declaration in the document", () => {
    const pkg = Package.open(docxWith({ "word/document.xml": billionLaughs }));
    try {
      pkg.partText("word/document.xml");
      expect.unreachable("expected malicious_content");
    } catch (e) {
      expect((e as ToolError).code).toBe("malicious_content");
    }
  });

  it("rejects an external-entity (XXE) DOCTYPE", () => {
    const xxe =
      '<?xml version="1.0"?>\n' +
      '<!DOCTYPE r [<!ENTITY ext SYSTEM "file:///etc/passwd">]>' +
      docWithBody("<w:p><w:r><w:t>x</w:t></w:r></w:p>").replace(/^<\?xml[^>]*\?>\n?/, "");
    const pkg = Package.open(docxWith({ "word/document.xml": xxe }));
    expect(() => pkg.partText("word/document.xml")).toThrowError(ToolError);
  });

  it("does not screen non-XML parts (no false positive)", () => {
    // A media part whose bytes contain the literal marker must pass through —
    // only .xml/.rels parts are screened.
    const zippable: Zippable = {};
    for (const [name, xml] of Object.entries(DEFAULT_PARTS)) zippable[name] = strToU8(xml);
    const pngLike = strToU8("PNGDATA<!ENTITY not really xml>");
    zippable["word/media/image1.png"] = pngLike;
    const pkg = Package.open(zipSync(zippable));
    expect(() => pkg.part("word/media/image1.png")).not.toThrow();
    expect(pkg.part("word/media/image1.png")).toEqual(pngLike);
  });
});

describe("XML nesting depth", () => {
  it("refuses pathologically deep nesting", () => {
    process.env["DOCXENGINE_MAX_XML_DEPTH"] = "50";
    const nested = "<w:x>".repeat(200) + "</w:x>".repeat(200);
    try {
      assertXmlDepth("word/document.xml", `<w:body>${nested}</w:body>`);
      expect.unreachable("expected doc_too_large");
    } catch (e) {
      expect((e as ToolError).code).toBe("doc_too_large");
    }
  });

  it("allows normal nesting", () => {
    expect(() =>
      assertXmlDepth("word/document.xml", "<w:p><w:r><w:t>ok</w:t></w:r></w:p>"),
    ).not.toThrow();
  });
});

describe("path traversal", () => {
  // resolveRelTarget is internal to validate.ts; exercise it through validation
  // by pointing a relationship at an escaping target and confirming it cannot
  // resolve to a part outside the package (it simply won't be found).
  it("clamps escaping .. targets so they never resolve outside the package", () => {
    const evilRels =
      '<?xml version="1.0"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom" Target="../../../../etc/passwd"/>' +
      "</Relationships>";
    // Opening + reading must not throw a path error or reach the filesystem;
    // the escaping target is just an in-package name that does not exist.
    const pkg = openAndReadAll(docxWith({ "_rels/.rels": evilRels }));
    expect(pkg.has("/etc/passwd")).toBe(false);
    expect(pkg.has("etc/passwd")).toBe(false);
  });
});

describe("malformed XML", () => {
  it("an unterminated tag is tolerated by the scanner (no hang)", () => {
    // The scanner returns null at a truncated tag rather than looping forever.
    const broken = docWithBody("<w:p><w:r><w:t>unterminated");
    const pkg = Package.open(docxWith({ "word/document.xml": broken }));
    expect(() =>
      assertXmlDepth("word/document.xml", pkg.partText("word/document.xml")),
    ).not.toThrow();
  });
});
