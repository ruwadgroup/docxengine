/**
 * Phase-2 stage-3: docx_convert to md/html (algorithms.md §23). Mirrors the
 * Python convert cases — heading/list/table rendering, bold/italic
 * reconstruction, accepted-view revisions, inline comment notes, alignment and
 * color inline styles, HTML escaping, and the pdf/png adapter gate.
 */
import { strToU8, zipSync, type Zippable } from "fflate";
import { describe, expect, it } from "vitest";

import { Session, docxConvert, docxCreate } from "../src/index.js";
import { DEFAULT_PARTS, type DocxParts, docWithBody } from "./fixtures.js";

function openBody(body: string, extra: DocxParts = {}) {
  const session = new Session();
  const parts: DocxParts = { ...DEFAULT_PARTS, "word/document.xml": docWithBody(body), ...extra };
  const zippable: Zippable = {};
  for (const [name, xml] of Object.entries(parts)) zippable[name] = strToU8(xml);
  const bytes = zipSync(zippable, { level: 0 });
  const doc = session.open(bytes);
  return { session, docId: doc.id };
}

const SECT = "<w:sectPr/>";

const STYLES_WITH_HEADINGS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>' +
  "</w:styles>";

describe("docx_convert to md", () => {
  it("renders headings and paragraphs", () => {
    const { session, docId } = openBody(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
        "<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>" +
        SECT,
      { "word/styles.xml": STYLES_WITH_HEADINGS },
    );
    const md = docxConvert(session, { doc_id: docId, to: "md" }).content as string;
    expect(md).toBe("# Title\n\nBody text.");
  });

  it("reconstructs **bold** and *italic* from run rPr", () => {
    const { session, docId } = openBody(
      "<w:p><w:r><w:t>plain </w:t></w:r>" +
        "<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>" +
        "<w:r><w:t> and </w:t></w:r>" +
        "<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r></w:p>" +
        SECT,
    );
    const md = docxConvert(session, { doc_id: docId, to: "md" }).content as string;
    expect(md).toBe("plain **bold** and *italic*");
  });

  it("renders a tight list of consecutive items", () => {
    const created = new Session();
    const res = docxCreate(created, { content_md: "- a\n- b\n- c\n" });
    const md = docxConvert(created, { doc_id: res.doc_id, to: "md" }).content as string;
    expect(md).toBe("- a\n- b\n- c");
  });

  it("renders a GitHub table", () => {
    const { session, docId } = openBody(
      "<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>" +
        "<w:tr><w:tc><w:p><w:r><w:t>Term</w:t></w:r></w:p></w:tc>" +
        "<w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>" +
        "<w:tr><w:tc><w:p><w:r><w:t>Fee</w:t></w:r></w:p></w:tc>" +
        "<w:tc><w:p><w:r><w:t>$100</w:t></w:r></w:p></w:tc></w:tr>" +
        "</w:tbl>" +
        SECT,
    );
    const md = docxConvert(session, { doc_id: docId, to: "md" }).content as string;
    expect(md).toContain("| Term | Value |");
    expect(md).toContain("| --- | --- |");
    expect(md).toContain("| Fee | $100 |");
  });

  it("shows insertions and omits deletions in the accepted view, with markers", () => {
    const { session, docId } = openBody(
      "<w:p><w:r><w:t>Pay in </w:t></w:r>" +
        '<w:del w:id="1" w:author="X" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30</w:delText></w:r></w:del>' +
        '<w:ins w:id="2" w:author="X" w:date="2026-01-01T00:00:00Z"><w:r><w:t>45</w:t></w:r></w:ins>' +
        "<w:r><w:t> days</w:t></w:r></w:p>" +
        SECT,
    );
    const md = docxConvert(session, { doc_id: docId, to: "md" }).content as string;
    expect(md).toContain("[ins]45");
    expect(md).toContain("[del]");
    expect(md).not.toContain("30");
  });

  it("annotates a comment inline at the range end", () => {
    const comments =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:comment w:id="0" w:author="Jane" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>confirm scope</w:t></w:r></w:p></w:comment>' +
      "</w:comments>";
    const { session, docId } = openBody(
      '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>This is mutual.</w:t></w:r>' +
        '<w:commentRangeEnd w:id="0"/>' +
        '<w:r><w:commentReference w:id="0"/></w:r></w:p>' +
        SECT,
      { "word/comments.xml": comments },
    );
    const md = docxConvert(session, { doc_id: docId, to: "md" }).content as string;
    expect(md).toContain("This is mutual. <!-- comment:Jane: confirm scope -->");
  });
});

describe("docx_convert to html", () => {
  it("emits semantic tags and HTML-escapes text", () => {
    const { session, docId } = openBody(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>A &amp; B</w:t></w:r></w:p>' +
        "<w:p><w:r><w:t>x &lt; y</w:t></w:r></w:p>" +
        SECT,
      { "word/styles.xml": STYLES_WITH_HEADINGS },
    );
    const html = docxConvert(session, { doc_id: docId, to: "html" }).content as string;
    expect(html).toContain("<h1>A &amp; B</h1>");
    expect(html).toContain("<p>x &lt; y</p>");
  });

  it("emits inline styles for alignment and color", () => {
    const { session, docId } = openBody(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
        '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>Red center</w:t></w:r></w:p>' +
        SECT,
    );
    const html = docxConvert(session, { doc_id: docId, to: "html" }).content as string;
    expect(html).toContain('style="text-align:center;color:#FF0000"');
  });

  it("wraps consecutive list items in one <ul>", () => {
    const created = new Session();
    const res = docxCreate(created, { content_md: "- a\n- b\n" });
    const html = docxConvert(created, { doc_id: res.doc_id, to: "html" }).content as string;
    expect(html).toContain("<ul>\n<li>a</li>\n<li>b</li>\n</ul>");
  });

  it("reconstructs <strong>/<em> from run rPr", () => {
    const { session, docId } = openBody(
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>B</w:t></w:r>" +
        "<w:r><w:rPr><w:i/></w:rPr><w:t>I</w:t></w:r></w:p>" +
        SECT,
    );
    const html = docxConvert(session, { doc_id: docId, to: "html" }).content as string;
    expect(html).toContain("<strong>B</strong><em>I</em>");
  });
});

describe("docx_convert pdf/png gate", () => {
  it("pdf without a renderer is render_unavailable (soffice not installed)", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>x</w:t></w:r></w:p>" + SECT);
    let code: string | null = null;
    try {
      docxConvert(session, { doc_id: docId, to: "pdf", path: "/tmp/out.pdf" });
    } catch (e) {
      code = (e as { code?: string }).code ?? null;
    }
    // On a machine with LibreOffice this would succeed; in CI it is unavailable.
    expect(["render_unavailable", "render_failed", "save_failed"]).toContain(code ?? "");
  });

  it("an unknown target is unsupported_format", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>x</w:t></w:r></w:p>" + SECT);
    let code: string | null = null;
    try {
      docxConvert(session, { doc_id: docId, to: "rtf" as unknown as "md" });
    } catch (e) {
      code = (e as { code?: string }).code ?? null;
    }
    expect(code).toBe("unsupported_format");
  });
});
