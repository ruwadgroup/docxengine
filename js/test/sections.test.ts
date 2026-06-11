/**
 * Phase-2 stage-2: docx_section (algorithms.md §15). Mirrors the Python
 * sections cases — list, set_geometry (page size, orientation, margins,
 * columns), set_header/set_footer (part + content-type + rel + reference),
 * insert_break (sectPr cloned into a pPr).
 */
import { describe, expect, it } from "vitest";

import {
  Session,
  ToolError,
  docxOpen,
  docxSection,
  docxValidate,
  headerFooterBody,
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

function pAnchor(session: Session, docId: string, ordinal: number): string {
  return session
    .get(docId)
    .anchorIndex()
    .filter((e) => e.kind === "p")[ordinal - 1]!.anchor;
}

function part(session: Session, docId: string, name: string): string {
  return session.get(docId).pkg.partText(name);
}

const BODY_SECT =
  "<w:p><w:r><w:t>Body</w:t></w:r></w:p>" +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

describe("docx_section list", () => {
  it("reports page size, orientation, columns, header/footer presence", () => {
    const { session, docId } = openBody(BODY_SECT);
    const res = docxSection(session, { doc_id: docId, op: "list" });
    expect(res.sections).toEqual([
      {
        id: "S1",
        break_type: "nextPage",
        page_size: "Letter",
        orientation: "portrait",
        columns: 1,
        has_header: false,
        has_footer: false,
      },
    ]);
  });
});

describe("docx_section set_geometry", () => {
  it("applies an A4 landscape page with margins and columns", () => {
    const { session, docId } = openBody(BODY_SECT);
    const res = docxSection(session, {
      doc_id: docId,
      op: "set_geometry",
      section: "S1",
      page_size: "A4",
      orientation: "landscape",
      margins: { top: 2, bottom: 2 },
      columns: 2,
    });
    expect(res.section).toBe("S1");
    const xml = part(session, docId, "word/document.xml");
    // A4 portrait is 11906×16838; landscape swaps to 16838×11906.
    expect(xml).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
    // top/bottom 2cm = 1134 twips; right/left kept at 1440.
    expect(xml).toContain('w:top="1134"');
    expect(xml).toContain('w:bottom="1134"');
    expect(xml).toContain('w:right="1440"');
    expect(xml).toContain('<w:cols w:num="2" w:space="708"/>');

    const list = docxSection(session, { doc_id: docId, op: "list" });
    expect(list.sections![0]).toMatchObject({
      page_size: "A4",
      orientation: "landscape",
      columns: 2,
    });
  });

  it("portrait removes w:orient", () => {
    const { session, docId } = openBody(BODY_SECT);
    docxSection(session, {
      doc_id: docId,
      op: "set_geometry",
      page_size: "A4",
      orientation: "landscape",
    });
    docxSection(session, { doc_id: docId, op: "set_geometry", orientation: "portrait" });
    const xml = part(session, docId, "word/document.xml");
    expect(xml).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(xml).not.toContain('w:orient="landscape"');
  });

  it("an unknown section id is anchor_not_found", () => {
    const { session, docId } = openBody(BODY_SECT);
    expect(() =>
      docxSection(session, { doc_id: docId, op: "set_geometry", section: "S5", page_size: "A4" }),
    ).toThrowError(ToolError);
  });
});

describe("docx_section set_header / set_footer", () => {
  it("creates a header part, content-type, rel, and a reference (before pgSz)", () => {
    const { session, docId } = openBody(BODY_SECT);
    const res = docxSection(session, {
      doc_id: docId,
      op: "set_header",
      content: "Confidential",
      variant: "default",
    });
    expect(res.section).toBe("S1");
    // header1.xml part with the markdown→paragraph content.
    expect(part(session, docId, "word/header1.xml")).toContain("<w:hdr xmlns:w=");
    expect(part(session, docId, "word/header1.xml")).toContain("<w:t>Confidential</w:t>");
    // content-type Override + document rel.
    expect(part(session, docId, "[Content_Types].xml")).toContain('PartName="/word/header1.xml"');
    expect(part(session, docId, "word/_rels/document.xml.rels")).toContain("/relationships/header");
    // The reference precedes pgSz in the sectPr.
    const xml = part(session, docId, "word/document.xml");
    const refIdx = xml.indexOf("<w:headerReference");
    const pgSzIdx = xml.indexOf("<w:pgSz");
    expect(refIdx).toBeGreaterThan(-1);
    expect(refIdx).toBeLessThan(pgSzIdx);
    expect(xml).toContain('<w:headerReference w:type="default" r:id=');

    expect(docxValidate(session, { doc_id: docId })).toEqual({ valid: true, issues: [] });
  });

  it("a footer with variant first maps to w:type='first'", () => {
    const { session, docId } = openBody(BODY_SECT);
    docxSection(session, { doc_id: docId, op: "set_footer", content: "Page", variant: "first" });
    expect(part(session, docId, "word/document.xml")).toContain(
      '<w:footerReference w:type="first"',
    );
    const list = docxSection(session, { doc_id: docId, op: "list" });
    expect(list.sections![0]!.has_footer).toBe(true);
  });
});

describe("headerFooterBody", () => {
  it("maps non-blank lines to plain paragraphs", () => {
    expect(headerFooterBody("One\nTwo")).toBe(
      "<w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p>",
    );
    expect(headerFooterBody("")).toBe("<w:p/>");
  });
});

describe("docx_section insert_break", () => {
  it("clones the body sectPr into the paragraph's pPr with a break type", () => {
    const { session, docId } = openBody(
      "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>" +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    );
    const after = pAnchor(session, docId, 1);
    const res = docxSection(session, {
      doc_id: docId,
      op: "insert_break",
      after,
      break_type: "nextPage",
    });
    expect(res.new_anchor).toBe(after); // text unchanged → same anchor
    const xml = part(session, docId, "word/document.xml");
    // The first paragraph gains a pPr/sectPr with the type and cloned geometry.
    expect(xml).toContain(
      '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>',
    );
    // Now two sections exist.
    const list = docxSection(session, { doc_id: docId, op: "list" });
    expect(list.sections).toHaveLength(2);
    expect(list.sections![0]!.id).toBe("S1");
    expect(list.sections![1]!.id).toBe("S2");
  });

  it("a malformed after anchor is anchor_stale/invalid", () => {
    const { session, docId } = openBody(BODY_SECT);
    expect(() =>
      docxSection(session, { doc_id: docId, op: "insert_break", after: "P1#0000" }),
    ).toThrowError(ToolError);
  });
});
