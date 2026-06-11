/**
 * Phase-2 stage-2: docx_field (algorithms.md §20). Mirrors the Python fields
 * cases — insert_toc run-triple, insert_page_number (footer ensured), update
 * (settings.xml updateFields). Computed values never appear in results.
 */
import { describe, expect, it } from "vitest";

import { Session, ToolError, docxField, docxOpen, docxValidate } from "../src/index.js";
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

const TITLE = "<w:p><w:r><w:t>Title</w:t></w:r></w:p><w:sectPr/>";

describe("docx_field insert_toc", () => {
  it("inserts the TOC field run-triple after the anchor (never fldSimple)", () => {
    const { session, docId } = openBody(TITLE);
    const after = pAnchor(session, docId, 1);
    const res = docxField(session, { doc_id: docId, op: "insert_toc", after, levels: 3 });
    expect(res.new_anchor).toBeDefined();
    const xml = part(session, docId, "word/document.xml");
    expect(xml).toContain(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>Right-click to update field.</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(xml).not.toContain("fldSimple");
    expect(docxValidate(session, { doc_id: docId })).toEqual({ valid: true, issues: [] });
  });

  it("honors a custom levels value", () => {
    const { session, docId } = openBody(TITLE);
    const after = pAnchor(session, docId, 1);
    docxField(session, { doc_id: docId, op: "insert_toc", after, levels: 5 });
    expect(part(session, docId, "word/document.xml")).toContain('TOC \\o "1-5"');
  });

  it("a malformed after anchor is anchor_stale/invalid", () => {
    const { session, docId } = openBody(TITLE);
    expect(() =>
      docxField(session, { doc_id: docId, op: "insert_toc", after: "P1#0000" }),
    ).toThrowError(ToolError);
  });
});

describe("docx_field insert_page_number", () => {
  it("ensures a footer and appends a PAGE field run-triple", () => {
    const { session, docId } = openBody(TITLE);
    const res = docxField(session, { doc_id: docId, op: "insert_page_number", scope: "footer" });
    expect(res.note).toContain("footer");
    // A footer part with the PAGE field exists.
    expect(session.get(docId).pkg.has("word/footer1.xml")).toBe(true);
    const footer = part(session, docId, "word/footer1.xml");
    expect(footer).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>');
    expect(footer).toContain('<w:fldChar w:fldCharType="begin"/>');
    // A footerReference was spliced into the body sectPr.
    expect(part(session, docId, "word/document.xml")).toContain(
      '<w:footerReference w:type="default"',
    );
    expect(part(session, docId, "[Content_Types].xml")).toContain('PartName="/word/footer1.xml"');
    expect(docxValidate(session, { doc_id: docId })).toEqual({ valid: true, issues: [] });
  });

  it("reuses an existing default footer part", () => {
    const { session, docId } = openBody(TITLE);
    docxField(session, { doc_id: docId, op: "insert_page_number", scope: "footer" });
    docxField(session, { doc_id: docId, op: "insert_page_number", scope: "footer" });
    // Still only one footer part; the second PAGE field appended to it.
    expect(session.get(docId).pkg.has("word/footer1.xml")).toBe(true);
    expect(session.get(docId).pkg.has("word/footer2.xml")).toBe(false);
    const footer = part(session, docId, "word/footer1.xml");
    const count = footer.split("PAGE").length - 1;
    expect(count).toBe(2);
  });
});

describe("docx_field update", () => {
  it("sets updateFields true in settings.xml (created on demand)", () => {
    const { session, docId } = openBody(TITLE);
    const res = docxField(session, { doc_id: docId, op: "update" });
    expect(res.updated).toBe(1);
    expect(session.get(docId).pkg.has("word/settings.xml")).toBe(true);
    const settings = part(session, docId, "word/settings.xml");
    expect(settings).toContain("<w:settings");
    expect(settings).toContain('<w:updateFields w:val="true"/>');
    expect(part(session, docId, "[Content_Types].xml")).toContain('PartName="/word/settings.xml"');
    expect(part(session, docId, "word/_rels/document.xml.rels")).toContain(
      "/relationships/settings",
    );
  });

  it("is idempotent when settings already flags updates", () => {
    const settingsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:updateFields w:val="true"/></w:settings>';
    const { session, docId } = openBody(TITLE, { "word/settings.xml": settingsXml });
    docxField(session, { doc_id: docId, op: "update" });
    const settings = part(session, docId, "word/settings.xml");
    // No duplicate updateFields element (idempotent).
    expect(settings.match(/<w:updateFields/g)!.length).toBe(1);
  });
});
