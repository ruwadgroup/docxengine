/**
 * Phase-2 stage-3: docx_create from markdown (algorithms.md §22). Mirrors the
 * Python create cases — block mapping (headings, quotes, rules, lists, tables),
 * inline runs (bold/italic/code), the deterministic skeleton parts, validation
 * gate, and the n_paragraphs count.
 */
import { describe, expect, it } from "vitest";

import {
  Session,
  docxConvert,
  docxCreate,
  docxRead,
  docxValidate,
  parseInline,
} from "../src/index.js";

function create(md: string) {
  const session = new Session();
  const res = docxCreate(session, { content_md: md });
  return { session, res };
}

function docXml(session: Session, docId: string): string {
  return session.get(docId).documentXml();
}

describe("docx_create block mapping", () => {
  it("maps ATX headings to Heading{n} styles", () => {
    const { session, res } = create("# Title\n\n## Sub\n\n### Deep\n");
    expect(res.n_paragraphs).toBe(3);
    const xml = docXml(session, res.doc_id);
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>');
    expect(xml).toContain('<w:pStyle w:val="Heading2"/>');
    expect(xml).toContain('<w:pStyle w:val="Heading3"/>');
    const proj = docxRead(session, { doc_id: res.doc_id }).content;
    expect(proj).toContain("[P1#");
    expect(proj).toContain("H1] Title");
  });

  it("maps blockquotes to the Quote style", () => {
    const { session, res } = create("> A quoted line\n");
    const xml = docXml(session, res.doc_id);
    expect(xml).toContain('<w:pStyle w:val="Quote"/>');
    expect(xml).toContain("A quoted line");
  });

  it("maps --- / *** to a bordered empty paragraph", () => {
    const { session, res } = create("Above\n\n---\n\nBelow\n");
    const xml = docXml(session, res.doc_id);
    expect(xml).toContain("<w:pBdr>");
    expect(xml).toContain('<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>');
    expect(res.n_paragraphs).toBe(3);
  });

  it("maps - / * / 1. list items via numbering.xml", () => {
    const { session, res } = create("- One\n- Two\n\n1. First\n2. Second\n");
    const numbering = session.get(res.doc_id).pkg.partText("word/numbering.xml");
    expect(numbering).toContain("<w:abstractNum");
    expect(numbering).toContain('<w:numFmt w:val="bullet"/>');
    expect(numbering).toContain('<w:numFmt w:val="decimal"/>');
    const proj = docxRead(session, { doc_id: res.doc_id }).content;
    expect(proj).toContain("List:ul L1] One");
    expect(proj).toContain("List:ol L1] First");
  });

  it("maps GitHub task items to a glyph-prefixed ListParagraph (no bullet)", () => {
    const { session, res } = create("- [ ] todo\n- [x] done\n* [X] also done\n");
    const xml = docXml(session, res.doc_id);
    expect(xml).toContain(
      '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr><w:r><w:t>☐ todo</w:t></w:r></w:p>',
    );
    expect(xml).toContain("<w:t>☒ done</w:t>");
    expect(xml).toContain("<w:t>☒ also done</w:t>");
    // The checkbox replaces the bullet — task items allocate no numbering.
    expect(session.get(res.doc_id).pkg.has("word/numbering.xml")).toBe(false);
  });

  it("maps a GitHub pipe table with a separator to a header table (§14)", () => {
    const { session, res } = create("| Term | Value |\n| --- | --- |\n| Fee | $100 |\n");
    const xml = docXml(session, res.doc_id);
    expect(xml).toContain("<w:tbl>");
    expect(xml).toContain('<w:tblStyle w:val="TableGrid"/>');
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>');
    expect(xml).toContain("<w:b/>");
    // Tables are excluded from n_paragraphs.
    expect(res.n_paragraphs).toBe(0);
    const md = docxConvert(session, { doc_id: res.doc_id, to: "md" }).content;
    expect(md).toContain("| Term | Value |");
    expect(md).toContain("| Fee | $100 |");
  });

  it("emits a plain paragraph for ordinary text", () => {
    const { session, res } = create("Just a sentence.\n");
    const xml = docXml(session, res.doc_id);
    expect(xml).toContain("Just a sentence.");
    expect(xml).not.toContain("w:pStyle");
  });
});

describe("docx_create inline runs", () => {
  it("splits **bold**, *italic*, and `code` into runs (§22 worked example)", () => {
    const { session, res } = create("See **clause** `4a`.\n");
    const xml = docXml(session, res.doc_id);
    expect(xml).toContain("<w:r><w:rPr><w:b/></w:rPr><w:t>clause</w:t></w:r>");
    expect(xml).toContain(
      '<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>4a</w:t></w:r>',
    );
  });

  it("parseInline tokenizes mixed inline markdown", () => {
    const runs = parseInline("a **b** c *d* `e`");
    expect(runs.map((r) => r.text)).toEqual(["a ", "b", " c ", "d", " ", "e"]);
    expect(runs[1]?.bold).toBe(true);
    expect(runs[3]?.italic).toBe(true);
    expect(runs[5]?.code).toBe(true);
  });

  it("leaves an unmatched marker literal", () => {
    const runs = parseInline("a * b");
    expect(runs).toEqual([{ text: "a * b" }]);
  });
});

describe("docx_create skeleton + validation", () => {
  it("ships the deterministic skeleton parts in §22 order", () => {
    const { session, res } = create("# Hi\n");
    const names = session.get(res.doc_id).pkg.entryNames();
    expect(names).toContain("word/document.xml");
    expect(names).toContain("word/styles.xml");
    expect(names).toContain("[Content_Types].xml");
    expect(names).toContain("_rels/.rels");
    expect(names).toContain("word/_rels/document.xml.rels");
    expect(names).toContain("docProps/core.xml");
    // document precedes styles precedes content-types.
    expect(names.indexOf("word/document.xml")).toBeLessThan(names.indexOf("word/styles.xml"));
    expect(names.indexOf("word/styles.xml")).toBeLessThan(names.indexOf("[Content_Types].xml"));
  });

  it("ships the base style set", () => {
    const { session, res } = create("text\n");
    const styles = session.get(res.doc_id).pkg.partText("word/styles.xml");
    for (const id of ["Normal", "Heading1", "Heading6", "ListParagraph", "TableGrid", "Quote"]) {
      expect(styles).toContain(`w:styleId="${id}"`);
    }
  });

  it("honors DOCXENGINE_FIXED_DATE in core.xml", () => {
    const prev = process.env["DOCXENGINE_FIXED_DATE"];
    process.env["DOCXENGINE_FIXED_DATE"] = "2026-06-10T00:00:00Z";
    try {
      const { session, res } = create("hi\n");
      const core = session.get(res.doc_id).pkg.partText("docProps/core.xml");
      expect(core).toContain(
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-06-10T00:00:00Z</dcterms:created>',
      );
      expect(core).toContain("2026-06-10T00:00:00Z</dcterms:modified>");
    } finally {
      if (prev === undefined) delete process.env["DOCXENGINE_FIXED_DATE"];
      else process.env["DOCXENGINE_FIXED_DATE"] = prev;
    }
  });

  it("the created document passes the §8 validation gate", () => {
    const { session, res } = create("# Doc\n\n- a\n- b\n\n| h |\n| --- |\n| v |\n");
    const verdict = docxValidate(session, { doc_id: res.doc_id });
    expect(verdict.valid).toBe(true);
    expect(verdict.issues.filter((i) => i.severity === "error")).toEqual([]);
  });

  it("omits numbering.xml when there are no list items", () => {
    const { session, res } = create("# No lists here\n");
    expect(session.get(res.doc_id).pkg.has("word/numbering.xml")).toBe(false);
    const ct = session.get(res.doc_id).pkg.partText("[Content_Types].xml");
    expect(ct).not.toContain("numbering+xml");
  });
});
