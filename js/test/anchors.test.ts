import { describe, expect, it } from "vitest";

import {
  ToolError,
  anchorHash,
  buildAnchorIndex,
  emitTextElement,
  findElement,
  normalizedText,
  parseAnchor,
  pieceAt,
  resolveAnchor,
  scopeText,
  splice,
  textWithMap,
} from "../src/index.js";
import { DOCUMENT_XML, docWithBody } from "./fixtures.js";

describe("normalizedText (§1)", () => {
  it("follows the worked example: NFC → collapse → strip", () => {
    expect(normalizedText("Master Services  Agreement ")).toBe("Master Services Agreement");
  });

  it("collapses every §1 whitespace run to one ASCII space", () => {
    expect(normalizedText("a\t\nb c d　e f g")).toBe("a b c d e f g");
    expect(normalizedText("  x ")).toBe("x");
  });

  it("does not treat U+FEFF as whitespace (JS \\s is non-conformant)", () => {
    expect(normalizedText("﻿a﻿b﻿")).toBe("﻿a﻿b﻿");
  });

  it("applies NFC so composed and decomposed forms hash identically", () => {
    expect(normalizedText("Café")).toBe("Café");
    expect(anchorHash(normalizedText("Café"))).toBe(anchorHash(normalizedText("Café")));
  });

  it("normalizes an empty paragraph to '' which hashes to e3b0", () => {
    expect(normalizedText("   \t ")).toBe("");
    expect(anchorHash("")).toBe("e3b0");
  });
});

describe("anchor index (§1)", () => {
  it("reproduces the spec worked example: P1#515a", () => {
    const index = buildAnchorIndex(DOCUMENT_XML);
    expect(index[0]?.anchor).toBe("P1#515a");
    expect(index[0]?.normalized).toBe("Master Services Agreement");
  });

  it("orders body blocks, numbers tables separately, skips the trailing sectPr", () => {
    const index = buildAnchorIndex(DOCUMENT_XML);
    expect(index.map((e) => `${e.kind}${e.ordinal}`)).toEqual(["p1", "p2", "tbl1", "p3"]);
    expect(index[2]?.anchor).toBe("T1");
    expect(index[3]?.anchor).toBe(`P3#e3b0`); // <w:p/> — empty paragraph
  });

  it("is stable across rsid fragmentation and run splits", () => {
    const single = docWithBody(
      "<w:p><w:r><w:t>The term is five (5) years from the Effective Date.</w:t></w:r></w:p>",
    );
    const fragmented = docWithBody(
      '<w:p><w:r w:rsidR="00AA0001"><w:t xml:space="preserve">The term is </w:t></w:r>' +
        '<w:r w:rsidR="00BB0002"><w:t xml:space="preserve">five (5) </w:t></w:r>' +
        '<w:r w:rsidR="00CC0003" w:rsidRPr="00DD0004"><w:rPr><w:b/></w:rPr>' +
        "<w:t>years from the Effective Date.</w:t></w:r></w:p>",
    );
    const a = buildAnchorIndex(single)[0]?.anchor;
    const b = buildAnchorIndex(fragmented)[0]?.anchor;
    expect(a).toBe(b);
    // …and the hash matches the same content inside the full fixture document (P2).
    const hash = (a as string).split("#")[1];
    expect(buildAnchorIndex(DOCUMENT_XML)[1]?.anchor).toBe(`P2#${hash}`);
  });

  it("excludes w:delText (the hash sees the document as-if-accepted)", () => {
    const withDel = docWithBody(
      '<w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r>' +
        '<w:del w:id="1" w:author="X" w:date="2026-01-01T00:00:00Z">' +
        "<w:r><w:delText>gone</w:delText></w:r></w:del>" +
        "<w:r><w:t>this</w:t></w:r></w:p>",
    );
    const plain = docWithBody("<w:p><w:r><w:t>Keep this</w:t></w:r></w:p>");
    expect(buildAnchorIndex(withDel)[0]?.anchor).toBe(buildAnchorIndex(plain)[0]?.anchor);
  });

  it("decodes entities before hashing", () => {
    const escaped = docWithBody("<w:p><w:r><w:t>A &amp; B &lt;C&gt;</w:t></w:r></w:p>");
    const index = buildAnchorIndex(escaped);
    expect(index[0]?.normalized).toBe("A & B <C>");
    expect(index[0]?.anchor).toBe(`P1#${anchorHash("A & B <C>")}`);
  });
});

describe("parseAnchor / resolveAnchor (§1 validation)", () => {
  it("parses paragraph and table anchors", () => {
    expect(parseAnchor("P12#a7b2")).toEqual({ kind: "p", ordinal: 12, hash: "a7b2" });
    expect(parseAnchor("T3")).toEqual({ kind: "tbl", ordinal: 3 });
  });

  it.each(["", "P0#abcd", "P1#ABCD", "P1#abc", "P1#abcde", "P1", "T0", "X1", "p1#abcd"])(
    "rejects malformed anchor %j with anchor_invalid",
    (bad) => {
      try {
        parseAnchor(bad);
        expect.unreachable();
      } catch (e) {
        expect(e).toBeInstanceOf(ToolError);
        expect((e as ToolError).code).toBe("anchor_invalid");
      }
    },
  );

  it("resolves a current anchor to its block slice", () => {
    const entry = resolveAnchor(DOCUMENT_XML, "P1#515a");
    expect(DOCUMENT_XML.slice(entry.start, entry.start + 4)).toBe("<w:p");
    expect(DOCUMENT_XML.slice(entry.end - 6, entry.end)).toBe("</w:p>");
    expect(resolveAnchor(DOCUMENT_XML, "T1").kind).toBe("tbl");
  });

  it("raises anchor_stale on hash mismatch", () => {
    try {
      resolveAnchor(DOCUMENT_XML, "P1#beef");
      expect.unreachable();
    } catch (e) {
      expect((e as ToolError).code).toBe("anchor_stale");
      expect((e as ToolError).suggestions).toEqual(["docx_read(window:P1)"]);
    }
  });

  it("raises anchor_not_found for out-of-range ordinals", () => {
    for (const anchor of ["P99#e3b0", "T2"]) {
      try {
        resolveAnchor(DOCUMENT_XML, anchor);
        expect.unreachable();
      } catch (e) {
        expect((e as ToolError).code).toBe("anchor_not_found");
      }
    }
  });
});

describe("xmlscan text + offset map (§4 step 1)", () => {
  it("maps the §4 worked example offsets across fragmented runs", () => {
    const xml = docWithBody(
      '<w:p><w:r><w:t xml:space="preserve">The term is five (5) </w:t></w:r>' +
        "<w:r><w:rPr><w:b/></w:rPr><w:t>years from the Effective Date.</w:t></w:r></w:p>",
    );
    const p = findElement(xml, "w:p");
    expect(p).not.toBeNull();
    const { text, pieces } = textWithMap(xml, p!);
    expect(text).toBe("The term is five (5) years from the Effective Date.");
    const tPieces = pieces.filter((x) => x.kind === "t");
    expect(tPieces).toHaveLength(2);
    // Run 1 covers indices 0–21 (exclusive), run 2 starts at 21.
    expect(tPieces[0]?.textOffset).toBe(0);
    expect(tPieces[1]?.textOffset).toBe(21);
    expect(pieceAt(pieces, 20)).toBe(tPieces[0]);
    expect(pieceAt(pieces, 21)).toBe(tPieces[1]);
    expect(pieceAt(pieces, text.length)).toBeNull();
    // Each piece knows its enclosing run extent.
    for (const piece of tPieces) {
      expect(xml.slice(piece.runStart, piece.runStart + 4)).toBe("<w:r");
      expect(xml.slice(piece.runEnd - 6, piece.runEnd)).toBe("</w:r>");
    }
  });

  it("carries delText pieces without offsets and keeps scopeText del-free", () => {
    const xml = docWithBody(
      '<w:p><w:r><w:t xml:space="preserve">a </w:t></w:r>' +
        "<w:del><w:r><w:delText>x</w:delText></w:r></w:del>" +
        "<w:r><w:t>b</w:t></w:r></w:p>",
    );
    const p = findElement(xml, "w:p")!;
    const { text, pieces } = textWithMap(xml, p);
    expect(text).toBe("a b");
    expect(scopeText(xml, p)).toBe("a b");
    const del = pieces.find((x) => x.kind === "delText");
    expect(del?.text).toBe("x");
    expect(del?.textOffset).toBe(-1);
  });

  it("splices a w:t in place, leaving every other byte untouched (§3)", () => {
    const xml = docWithBody("<w:p><w:r><w:t>old text</w:t></w:r></w:p>");
    const p = findElement(xml, "w:p")!;
    const piece = textWithMap(xml, p).pieces[0]!;
    const out = splice(xml, piece.el.contentStart, piece.el.contentEnd, "new &amp; improved");
    expect(out).toBe(xml.replace(">old text<", ">new &amp; improved<"));
  });

  it("emits xml:space=preserve exactly when text starts/ends with §1 whitespace", () => {
    expect(emitTextElement("w:t", "plain")).toBe("<w:t>plain</w:t>");
    expect(emitTextElement("w:t", " lead")).toBe('<w:t xml:space="preserve"> lead</w:t>');
    expect(emitTextElement("w:t", "trail ")).toBe('<w:t xml:space="preserve">trail </w:t>');
    expect(emitTextElement("w:delText", "a & b < c")).toBe(
      "<w:delText>a &amp; b &lt; c</w:delText>",
    );
    expect(emitTextElement("w:t", "﻿x")).toBe("<w:t>﻿x</w:t>");
  });
});
