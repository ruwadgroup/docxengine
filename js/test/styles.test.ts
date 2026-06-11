/**
 * Phase-2 stage-1: docx_style and docx_format (algorithms.md §16). Mirrors the
 * Python styles cases — list with basedOn + in_use, define (id from name,
 * §16 child order, collisions), apply, style_selector merge, direct format.
 */
import { describe, expect, it } from "vitest";

import {
  Session,
  ToolError,
  canonicalizeProps,
  docxFormat,
  docxOpen,
  docxStyle,
  paraPropsInner,
  runPropsInner,
} from "../src/index.js";
import { DEFAULT_PARTS, type DocxParts, buildDocx, docWithBody } from "./fixtures.js";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Heading1"/><w:rPr><w:sz w:val="28"/></w:rPr></w:style>
</w:styles>`;

function partsWith(extra: DocxParts): DocxParts {
  return { ...DEFAULT_PARTS, ...extra };
}

function openBody(body: string, extra: DocxParts = {}) {
  const session = new Session();
  const parts = partsWith({
    "word/document.xml": docWithBody(body),
    "word/styles.xml": STYLES,
    ...extra,
  });
  const res = docxOpen(session, { bytes: Buffer.from(buildDocx(parts)).toString("base64") });
  return { session, docId: res.doc_id };
}

function pAnchor(session: Session, docId: string, ordinal: number): string {
  return session
    .get(docId)
    .anchorIndex()
    .filter((e) => e.kind === "p")[ordinal - 1]!.anchor;
}

function stylesXml(session: Session, docId: string): string {
  return session.get(docId).pkg.partText("word/styles.xml");
}

function docXml(session: Session, docId: string): string {
  return session.get(docId).documentXml();
}

describe("canonicalizeProps + emission (§16)", () => {
  it("maps shorthand to canonical and emits in §16 order", () => {
    const p = canonicalizeProps({ size: 11, bold: true, spacing_after: 6, alignment: "justify" });
    expect(p).toEqual({ size_pt: 11, bold: true, spacing_after_pt: 6, alignment: "both" });
    // rPr order: b, sz; pPr order: jc, spacing.
    expect(runPropsInner(p)).toBe('<w:b/><w:sz w:val="22"/>');
    expect(paraPropsInner(p)).toBe('<w:jc w:val="both"/><w:spacing w:after="120"/>');
  });

  it("color uppercases and strips the leading #", () => {
    expect(runPropsInner(canonicalizeProps({ color: "#1f4e79" }))).toBe(
      '<w:color w:val="1F4E79"/>',
    );
  });

  it("boolean false emits the toggle-off form", () => {
    expect(runPropsInner(canonicalizeProps({ bold: false }))).toBe('<w:b w:val="0"/>');
  });

  it("the line-spacing multiplier `spacing` is ignored", () => {
    expect(paraPropsInner(canonicalizeProps({ spacing: 1.5 }))).toBe("");
  });
});

describe("docx_style list", () => {
  it("reports cascade and in_use counts", () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>C</w:t></w:r></w:p>' +
      "<w:sectPr/>";
    const { session, docId } = openBody(body);
    const res = docxStyle(session, { doc_id: docId, op: "list" });
    expect(res.styles).toEqual([
      { id: "Normal", name: "Normal", type: "paragraph", in_use: 0 },
      { id: "Heading1", name: "heading 1", type: "paragraph", based_on: "Normal", in_use: 1 },
      { id: "Heading2", name: "heading 2", type: "paragraph", based_on: "Heading1", in_use: 2 },
    ]);
  });
});

describe("docx_style define", () => {
  it("derives the id from the name and emits §16 child order", () => {
    const { session, docId } = openBody("<w:p/><w:sectPr/>");
    const res = docxStyle(session, {
      doc_id: docId,
      op: "define",
      name: "Clause",
      based_on: "Normal",
      props: { size_pt: 11, bold: true, spacing_after_pt: 6, alignment: "justify" },
    });
    expect(res.style_id).toBe("Clause");
    expect(stylesXml(session, docId)).toContain(
      '<w:style w:type="paragraph" w:styleId="Clause"><w:name w:val="Clause"/><w:basedOn w:val="Normal"/>' +
        '<w:pPr><w:jc w:val="both"/><w:spacing w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>',
    );
  });

  it("strips whitespace from the name to form the id", () => {
    const { session, docId } = openBody("<w:p/><w:sectPr/>");
    const res = docxStyle(session, { doc_id: docId, op: "define", name: "Body Text" });
    expect(res.style_id).toBe("BodyText");
  });

  it("collisions take the suffix 2, 3, …", () => {
    const { session, docId } = openBody("<w:p/><w:sectPr/>");
    // STYLES already defines `Normal`, so the first collision is Normal2…
    const first = docxStyle(session, { doc_id: docId, op: "define", name: "Normal" });
    expect(first.style_id).toBe("Normal2");
    const second = docxStyle(session, { doc_id: docId, op: "define", name: "Normal" });
    expect(second.style_id).toBe("Normal3");
  });

  it("creates styles.xml when the part is absent", () => {
    const session = new Session();
    const parts = { ...DEFAULT_PARTS };
    delete (parts as DocxParts)["word/styles.xml"];
    parts["word/document.xml"] = docWithBody("<w:p/><w:sectPr/>");
    const res0 = docxOpen(session, { bytes: Buffer.from(buildDocx(parts)).toString("base64") });
    const docId = res0.doc_id;
    docxStyle(session, { doc_id: docId, op: "define", name: "Clause" });
    expect(session.get(docId).pkg.has("word/styles.xml")).toBe(true);
    expect(stylesXml(session, docId)).toContain('w:styleId="Clause"');
  });
});

describe("docx_style apply", () => {
  it("splices pStyle as the first pPr child and returns a fresh anchor", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr/>");
    const a1 = pAnchor(session, docId, 1);
    const res = docxStyle(session, { doc_id: docId, op: "apply", anchor: a1, style: "Heading2" });
    expect(docXml(session, docId)).toContain(
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>',
    );
    expect(res.new_anchor).toBe(a1); // text unchanged → same hash
  });

  it("resolves a style by name", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>Hi</w:t></w:r></w:p><w:sectPr/>");
    const a1 = pAnchor(session, docId, 1);
    docxStyle(session, { doc_id: docId, op: "apply", anchor: a1, style: "heading 2" });
    expect(docXml(session, docId)).toContain('<w:pStyle w:val="Heading2"/>');
  });

  it("an unknown style is style_unknown", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>Hi</w:t></w:r></w:p><w:sectPr/>");
    const a1 = pAnchor(session, docId, 1);
    const err = (() => {
      try {
        docxStyle(session, { doc_id: docId, op: "apply", anchor: a1, style: "Nope" });
      } catch (e) {
        return e as ToolError;
      }
      throw new Error("expected throw");
    })();
    expect(err.code).toBe("style_unknown");
  });
});

describe("docx_format style_selector", () => {
  it("merges props into the style's rPr (creating it in §16 order)", () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:sectPr/>';
    const { session, docId } = openBody(body);
    const res = docxFormat(session, {
      doc_id: docId,
      style_selector: { style: "heading 2" },
      props: { color: "#1F4E79" },
    });
    // Heading2 already has <w:sz w:val="28"/>; color sorts before sz in §16.
    expect(stylesXml(session, docId)).toContain(
      '<w:rPr><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr>',
    );
    // style_selector is one document-wide edit: no paragraphs touched (§16).
    expect(res.affected).toBe(0);
    expect(res.anchors).toEqual([]);
  });

  it("creates a pPr when the style has none", () => {
    const { session, docId } = openBody("<w:p/><w:sectPr/>");
    docxFormat(session, {
      doc_id: docId,
      style_selector: { style: "Heading1" },
      props: { alignment: "center" },
    });
    expect(stylesXml(session, docId)).toContain(
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style>',
    );
  });

  it("an unknown style is style_unknown", () => {
    const { session, docId } = openBody("<w:p/><w:sectPr/>");
    const err = (() => {
      try {
        docxFormat(session, {
          doc_id: docId,
          style_selector: { style: "Nope" },
          props: { bold: true },
        });
      } catch (e) {
        return e as ToolError;
      }
      throw new Error("expected throw");
    })();
    expect(err.code).toBe("style_unknown");
  });
});

describe("docx_format direct", () => {
  it("applies rPr to every run and pPr to the paragraph", () => {
    const body = "<w:p><w:r><w:t>x</w:t></w:r><w:r><w:t>y</w:t></w:r></w:p><w:sectPr/>";
    const { session, docId } = openBody(body);
    const a1 = pAnchor(session, docId, 1);
    const res = docxFormat(session, {
      doc_id: docId,
      anchor: a1,
      props: { bold: true, alignment: "right" },
    });
    const xml = docXml(session, docId);
    expect(xml).toContain(
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>y</w:t></w:r></w:p>',
    );
    expect(res.affected).toBe(1);
    expect(res.anchors).toEqual([a1]);
  });

  it("merges into an existing rPr without disturbing other props", () => {
    const body = "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>z</w:t></w:r></w:p><w:sectPr/>";
    const { session, docId } = openBody(body);
    const a1 = pAnchor(session, docId, 1);
    docxFormat(session, { doc_id: docId, anchor: a1, props: { bold: true } });
    // §16 order: b before i.
    expect(docXml(session, docId)).toContain("<w:rPr><w:b/><w:i/></w:rPr>");
  });

  it("applies across a range and returns ascending anchors", () => {
    const body =
      "<w:p><w:r><w:t>a</w:t></w:r></w:p>" +
      "<w:p><w:r><w:t>b</w:t></w:r></w:p>" +
      "<w:p><w:r><w:t>c</w:t></w:r></w:p><w:sectPr/>";
    const { session, docId } = openBody(body);
    const res = docxFormat(session, { doc_id: docId, range: "P1..P3", props: { italic: true } });
    expect(res.affected).toBe(3);
    expect(res.anchors?.length).toBe(3);
    expect((docXml(session, docId).match(/<w:i\/>/g) ?? []).length).toBe(3);
  });

  it("missing anchor/range/style_selector is anchor_invalid", () => {
    const { session, docId } = openBody("<w:p/><w:sectPr/>");
    const err = (() => {
      try {
        docxFormat(session, { doc_id: docId, props: { bold: true } });
      } catch (e) {
        return e as ToolError;
      }
      throw new Error("expected throw");
    })();
    expect(err.code).toBe("anchor_invalid");
  });
});
