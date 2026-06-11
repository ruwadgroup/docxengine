/**
 * Phase-2 stage-1: docx_table (algorithms.md §14). Mirrors the Python
 * tables cases — create with the §14 worked example, set_cells, insert/delete
 * row/col, merge, projection round-trip.
 */
import { describe, expect, it } from "vitest";

import {
  Session,
  ToolError,
  bodyProjBlocks,
  colLettersToIndex,
  docxOpen,
  docxRead,
  docxTable,
  parseA1,
  projectionContext,
  tableProjection,
} from "../src/index.js";
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

const INTRO = "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>";
/** Anchor of a single-paragraph "Intro" body P1. */
function introAnchor(session: Session, docId: string): string {
  const doc = session.get(docId);
  return doc.anchorIndex().filter((e) => e.kind === "p")[0]!.anchor;
}

function docXml(session: Session, docId: string): string {
  return session.get(docId).documentXml();
}

describe("A1 addressing", () => {
  it("maps column letters to 0-based indices", () => {
    expect(colLettersToIndex("A")).toBe(0);
    expect(colLettersToIndex("Z")).toBe(25);
    expect(colLettersToIndex("AA")).toBe(26);
    expect(colLettersToIndex("AB")).toBe(27);
  });

  it("parses A1 refs", () => {
    expect(parseA1("A1")).toEqual({ r: 0, c: 0 });
    expect(parseA1("B2")).toEqual({ r: 1, c: 1 });
    expect(parseA1("C1")).toEqual({ r: 0, c: 2 });
  });
});

describe("docx_table create", () => {
  it("emits the §14 worked-example XML for a 2×2 header table", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    const res = docxTable(session, {
      doc_id: docId,
      op: "create",
      after,
      rows: 2,
      cols: 2,
      data: [
        ["Term", "Value"],
        ["Fee", "$100 "],
      ],
      header: true,
    });
    expect(res.new_anchor).toBe(`T1@after:${after}`);
    const xml = docXml(session, docId);
    expect(xml).toContain(
      '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>',
    );
    expect(xml).toContain('<w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>');
    // Header cell: shading + bold run.
    expect(xml).toContain(
      '<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Term</w:t></w:r></w:p></w:tc>',
    );
    // Body cell: $100 with trailing space → xml:space=preserve.
    expect(xml).toContain(
      '<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">$100 </w:t></w:r></w:p></w:tc>',
    );
    // TableGrid style ensured.
    expect(session.get(docId).pkg.partText("word/styles.xml")).toContain('w:styleId="TableGrid"');
  });

  it("distributes width with the last column absorbing the remainder", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, { doc_id: docId, op: "create", after, rows: 1, cols: 3 });
    const xml = docXml(session, docId);
    // 9026 / 3 = 3008 r2 → 3008,3008,3010.
    expect(xml).toContain(
      '<w:tblGrid><w:gridCol w:w="3008"/><w:gridCol w:w="3008"/><w:gridCol w:w="3010"/></w:tblGrid>',
    );
  });

  it("an empty cell becomes <w:p/>", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, { doc_id: docId, op: "create", after, rows: 1, cols: 1, data: [[""]] });
    expect(docXml(session, docId)).toContain(
      '<w:tcPr><w:tcW w:w="9026" w:type="dxa"/></w:tcPr><w:p/>',
    );
  });

  it("a plain (no header, no style) table omits tblStyle", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, { doc_id: docId, op: "create", after, rows: 1, cols: 1 });
    const xml = docXml(session, docId);
    expect(xml).toContain('<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>');
    expect(xml).not.toContain("tblStyle");
  });

  it("data overflow is anchor_invalid", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    expect(() =>
      docxTable(session, {
        doc_id: docId,
        op: "create",
        after,
        rows: 1,
        cols: 1,
        data: [["a", "b"]],
      }),
    ).toThrowError(ToolError);
  });

  it("projects the new table as a GitHub markdown table (§2)", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, {
      doc_id: docId,
      op: "create",
      after,
      rows: 2,
      cols: 2,
      data: [
        ["Term", "Value"],
        ["Fee", "$100"],
      ],
      header: true,
    });
    const read = docxRead(session, { doc_id: docId });
    expect(read.content).toContain("| Term | Value |");
    expect(read.content).toContain("| --- | --- |");
    expect(read.content).toContain("| Fee | $100 |");
    expect(read.content).toContain(`[T1 2×2 @after:${after}]`);
  });
});

describe("docx_table set_cells", () => {
  it("rewrites a cell by {r,c} and by A1 ref, preserving tcW", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, {
      doc_id: docId,
      op: "create",
      after,
      rows: 2,
      cols: 2,
      data: [
        ["a", "b"],
        ["c", "d"],
      ],
    });
    docxTable(session, {
      doc_id: docId,
      op: "set_cells",
      anchor: "T1",
      cells: [
        { r: 1, c: 1, text: "D!" },
        { ref: "A1", text: "A!" },
      ],
    });
    const xml = docXml(session, docId);
    expect(xml).toContain(
      '<w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A!</w:t></w:r></w:p>',
    );
    expect(xml).toContain("<w:t>D!</w:t>");
    expect(xml).not.toContain("<w:t>d</w:t>");
  });

  it("out-of-range cell is anchor_invalid", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, { doc_id: docId, op: "create", after, rows: 1, cols: 1 });
    expect(() =>
      docxTable(session, {
        doc_id: docId,
        op: "set_cells",
        anchor: "T1",
        cells: [{ r: 5, c: 0, text: "x" }],
      }),
    ).toThrowError(ToolError);
  });
});

describe("docx_table rows and columns", () => {
  function make() {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, {
      doc_id: docId,
      op: "create",
      after,
      rows: 2,
      cols: 2,
      data: [
        ["a", "b"],
        ["c", "d"],
      ],
    });
    return { session, docId };
  }

  it("insert_row clones the cell structure with blank text", () => {
    const { session, docId } = make();
    docxTable(session, { doc_id: docId, op: "insert_row", anchor: "T1", at: 1 });
    const read = docxRead(session, { doc_id: docId });
    // New blank row between row a|b and row c|d.
    expect(read.content).toContain("| a | b |\n| --- | --- |\n|  |  |\n| c | d |");
  });

  it("insert_row at rows appends", () => {
    const { session, docId } = make();
    docxTable(session, { doc_id: docId, op: "insert_row", anchor: "T1", at: 2 });
    const read = docxRead(session, { doc_id: docId });
    expect(read.content).toContain("| c | d |\n|  |  |");
  });

  it("insert_col adds a gridCol and a blank cell per row", () => {
    const { session, docId } = make();
    docxTable(session, { doc_id: docId, op: "insert_col", anchor: "T1", at: 1 });
    const xml = docXml(session, docId);
    const gridCols = (xml.match(/<w:gridCol /g) ?? []).length;
    expect(gridCols).toBe(3);
    const read = docxRead(session, { doc_id: docId });
    expect(read.content).toContain("| a |  | b |");
  });

  it("delete_row removes the row", () => {
    const { session, docId } = make();
    docxTable(session, { doc_id: docId, op: "delete_row", anchor: "T1", at: 0 });
    const read = docxRead(session, { doc_id: docId });
    expect(read.content).toContain("| c | d |");
    expect(read.content).not.toContain("| a | b |");
  });

  it("delete_col removes the gridCol and that cell in every row", () => {
    const { session, docId } = make();
    docxTable(session, { doc_id: docId, op: "delete_col", anchor: "T1", at: 0 });
    const xml = docXml(session, docId);
    expect((xml.match(/<w:gridCol /g) ?? []).length).toBe(1);
    const read = docxRead(session, { doc_id: docId });
    expect(read.content).toContain("| b |");
    expect(read.content).not.toContain("| a |");
  });
});

describe("docx_table merge", () => {
  it("horizontal merge sets gridSpan and removes covered cells", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, {
      doc_id: docId,
      op: "create",
      after,
      rows: 1,
      cols: 3,
      data: [["x", "y", "z"]],
    });
    docxTable(session, { doc_id: docId, op: "merge", anchor: "T1", range: "A1:C1" });
    const xml = docXml(session, docId);
    expect(xml).toContain('<w:gridSpan w:val="3"/>');
    // The covered y, z cells are gone.
    expect(xml).not.toContain("<w:t>y</w:t>");
    expect(xml).not.toContain("<w:t>z</w:t>");
  });

  it("vertical merge sets vMerge restart then continue", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, {
      doc_id: docId,
      op: "create",
      after,
      rows: 2,
      cols: 1,
      data: [["top"], ["bottom"]],
    });
    docxTable(session, { doc_id: docId, op: "merge", anchor: "T1", range: "A1:A2" });
    const xml = docXml(session, docId);
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain("<w:vMerge/>");
  });

  it("rectangular merge applies gridSpan per row and vMerge down the left column", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, {
      doc_id: docId,
      op: "create",
      after,
      rows: 3,
      cols: 3,
      data: [
        ["a", "b", "c"],
        ["d", "e", "f"],
        ["g", "h", "i"],
      ],
    });
    docxTable(session, { doc_id: docId, op: "merge", anchor: "T1", range: "A1:B2" });
    const xml = docXml(session, docId);
    // §14: marks are the FIRST tcPr children; vMerge (written second) precedes
    // gridSpan, both before tcW. Top-left keeps text "a".
    expect(xml).toContain(
      '<w:tcPr><w:vMerge w:val="restart"/><w:gridSpan w:val="2"/><w:tcW w:w="3008" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p>',
    );
    // Continuation: bare vMerge + gridSpan, empty paragraph.
    expect(xml).toContain(
      '<w:tcPr><w:vMerge/><w:gridSpan w:val="2"/><w:tcW w:w="3008" w:type="dxa"/></w:tcPr><w:p/>',
    );
    // Covered cells b, e gone; third row untouched.
    expect(xml).not.toContain("<w:t>b</w:t>");
    expect(xml).not.toContain("<w:t>e</w:t>");
    expect(xml).toContain("<w:t>g</w:t>");
  });
});

describe("docx_table style", () => {
  it("adds the TableGrid style to an existing table", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, { doc_id: docId, op: "create", after, rows: 1, cols: 1 });
    docxTable(session, { doc_id: docId, op: "style", anchor: "T1", style: "Table Grid" });
    const xml = docXml(session, docId);
    expect(xml).toContain('<w:tblStyle w:val="TableGrid"/>');
    expect(session.get(docId).pkg.partText("word/styles.xml")).toContain('w:styleId="TableGrid"');
  });
});

describe("docx_table projection helper parity", () => {
  it("tableProjection renders the created table identically to docx_read", () => {
    const { session, docId } = openBody(INTRO + "<w:sectPr/>");
    const after = introAnchor(session, docId);
    docxTable(session, {
      doc_id: docId,
      op: "create",
      after,
      rows: 1,
      cols: 2,
      data: [["k", "v"]],
    });
    const doc = session.get(docId);
    const blocks = bodyProjBlocks(doc);
    const tblBlock = blocks.find((b) => b.kind === "tbl")!;
    void projectionContext(doc);
    const line = tableProjection(tblBlock, after);
    expect(line).toContain(`[T1 1×2 @after:${after}]`);
    expect(line).toContain("| k | v |");
  });
});
