/**
 * Phase-2 stage-1: docx_list and numbering (algorithms.md §17). Mirrors the
 * Python lists cases — create (ol/ul, multilevel), numbering.xml creation +
 * wiring, convert to/from list, set_level, restart, projection annotations.
 */
import { describe, expect, it } from "vitest";

import { Session, ToolError, docxList, docxOpen, docxRead } from "../src/index.js";
import { DEFAULT_PARTS, type DocxParts, buildDocx, docWithBody } from "./fixtures.js";

function partsWith(extra: DocxParts): DocxParts {
  return { ...DEFAULT_PARTS, ...extra };
}

function openBody(body: string, extra: DocxParts = {}) {
  const session = new Session();
  const parts = partsWith({ "word/document.xml": docWithBody(body), ...extra });
  const res = docxOpen(session, { bytes: Buffer.from(buildDocx(parts)).toString("base64") });
  return { session, docId: res.doc_id };
}

function pAnchor(session: Session, docId: string, ordinal: number): string {
  return session
    .get(docId)
    .anchorIndex()
    .filter((e) => e.kind === "p")[ordinal - 1]!.anchor;
}

function docXml(session: Session, docId: string): string {
  return session.get(docId).documentXml();
}

function numberingXml(session: Session, docId: string): string {
  return session.get(docId).pkg.partText("word/numbering.xml");
}

const INTRO = "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>";

describe("docx_list create", () => {
  it("creates numbering.xml (override + rel) and an ol with cascading levels", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = pAnchor(session, docId, 1);
    const res = docxList(session, {
      doc_id: docId,
      op: "create",
      after,
      kind: "ol",
      items: [{ text: "First" }, { text: "Second" }],
    });
    expect(res.n_affected).toBe(2);
    expect(res.new_anchors?.length).toBe(2);

    const doc = session.get(docId);
    expect(doc.pkg.has("word/numbering.xml")).toBe(true);
    const num = numberingXml(session, docId);
    expect(num).toContain('<w:abstractNum w:abstractNumId="1">');
    // ol level 0/1/2 cascade decimal/lowerLetter/lowerRoman with %n. lvlText.
    expect(num).toContain(
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>',
    );
    expect(num).toContain(
      '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>',
    );
    expect(num).toContain('<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>');

    // content-type Override + document rel.
    expect(doc.pkg.partText("[Content_Types].xml")).toContain('PartName="/word/numbering.xml"');
    expect(doc.pkg.partText("word/_rels/document.xml.rels")).toContain('Target="numbering.xml"');

    // ListParagraph style ensured; each item carries numPr + pStyle.
    expect(doc.pkg.partText("word/styles.xml")).toContain('w:styleId="ListParagraph"');
    expect(docXml(session, docId)).toContain(
      '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>',
    );
  });

  it("ul uses bullet glyphs", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = pAnchor(session, docId, 1);
    docxList(session, {
      doc_id: docId,
      op: "create",
      after,
      kind: "ul",
      items: [{ text: "Bullet" }],
    });
    const num = numberingXml(session, docId);
    expect(num).toContain('<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>');
    expect(num).toContain('<w:lvlText w:val="◦"/>');
    expect(num).toContain('<w:lvlText w:val="▪"/>');
  });

  it("multi-level items set their ilvl", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = pAnchor(session, docId, 1);
    docxList(session, {
      doc_id: docId,
      op: "create",
      after,
      kind: "ol",
      items: [{ text: "Top" }, { text: "Nested", level: 1 }],
    });
    const xml = docXml(session, docId);
    expect(xml).toContain('<w:ilvl w:val="0"/>');
    expect(xml).toContain('<w:ilvl w:val="1"/>');
  });

  it("projects list items with the §2 List annotation", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = pAnchor(session, docId, 1);
    docxList(session, {
      doc_id: docId,
      op: "create",
      after,
      kind: "ol",
      items: [{ text: "Alpha" }],
    });
    const read = docxRead(session, { doc_id: docId });
    expect(read.content).toContain("List:ol L1] Alpha");
  });

  it("a second list allocates the next abstractNum/num ids", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = pAnchor(session, docId, 1);
    docxList(session, { doc_id: docId, op: "create", after, kind: "ol", items: [{ text: "A" }] });
    const after2 = pAnchor(session, docId, 1);
    docxList(session, {
      doc_id: docId,
      op: "create",
      after: after2,
      kind: "ul",
      items: [{ text: "B" }],
    });
    const num = numberingXml(session, docId);
    expect(num).toContain('<w:abstractNum w:abstractNumId="1">');
    expect(num).toContain('<w:abstractNum w:abstractNumId="2">');
    expect(num).toContain('<w:num w:numId="1">');
    expect(num).toContain('<w:num w:numId="2">');
  });
});

describe("docx_list convert", () => {
  it("converts plain paragraphs to an ol list", () => {
    const body =
      "<w:p><w:r><w:t>One</w:t></w:r></w:p>" + "<w:p><w:r><w:t>Two</w:t></w:r></w:p><w:sectPr/>";
    const { session, docId } = openBody(body);
    const res = docxList(session, { doc_id: docId, op: "convert", range: "P1..P2", to: "ol" });
    expect(res.n_affected).toBe(2);
    const xml = docXml(session, docId);
    expect((xml.match(/<w:numPr>/g) ?? []).length).toBe(2);
    expect(xml).toContain('<w:pStyle w:val="ListParagraph"/>');
  });

  it("converts list items back to plain paragraphs (removes numPr)", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = pAnchor(session, docId, 1);
    docxList(session, { doc_id: docId, op: "create", after, kind: "ol", items: [{ text: "X" }] });
    const itemAnchor = pAnchor(session, docId, 2);
    docxList(session, { doc_id: docId, op: "convert", anchor: itemAnchor, to: "paragraphs" });
    expect(docXml(session, docId)).not.toContain("<w:numPr>");
  });

  it("convert preserves a pre-existing pPr child", () => {
    const body =
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered</w:t></w:r></w:p><w:sectPr/>';
    const { session, docId } = openBody(body);
    docxList(session, {
      doc_id: docId,
      op: "convert",
      anchor: pAnchor(session, docId, 1),
      to: "ul",
    });
    const xml = docXml(session, docId);
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('<w:pStyle w:val="ListParagraph"/>');
    expect(xml).toContain("<w:numPr>");
  });
});

describe("docx_list set_level", () => {
  it("rewrites the item's ilvl", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = pAnchor(session, docId, 1);
    docxList(session, { doc_id: docId, op: "create", after, kind: "ol", items: [{ text: "X" }] });
    const itemAnchor = pAnchor(session, docId, 2);
    docxList(session, { doc_id: docId, op: "set_level", anchor: itemAnchor, level: 2 });
    expect(docXml(session, docId)).toContain('<w:ilvl w:val="2"/>');
  });
});

describe("docx_list restart", () => {
  it("allocates a new num referencing the same abstractNum with a startOverride", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = pAnchor(session, docId, 1);
    docxList(session, { doc_id: docId, op: "create", after, kind: "ol", items: [{ text: "X" }] });
    const itemAnchor = pAnchor(session, docId, 2);
    docxList(session, { doc_id: docId, op: "restart", anchor: itemAnchor, at: 5 });
    const num = numberingXml(session, docId);
    expect(num).toContain(
      '<w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>',
    );
    // The paragraph now points at numId 2.
    expect(docXml(session, docId)).toContain('<w:numId w:val="2"/>');
  });

  it("restart on a non-list paragraph is anchor_invalid", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sectPr/>");
    const err = (() => {
      try {
        docxList(session, {
          doc_id: docId,
          op: "restart",
          anchor: pAnchor(session, docId, 1),
          at: 2,
        });
      } catch (e) {
        return e as ToolError;
      }
      throw new Error("expected throw");
    })();
    expect(err.code).toBe("anchor_invalid");
  });
});
