/**
 * Stage-3 edit surface: docx_replace / docx_edit_paragraph / docx_insert /
 * docx_delete. Mirrors the Python stage-3 cases (python/tests/test_edit.py):
 * §4 splice-coalescing, §5 tracked emission (DOCXENGINE_FIXED_DATE pinned),
 * §6 LCS word diff, minimal markdown insert, range deletes, and
 * anchor-validated-first error paths.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  Session,
  ToolError,
  anchorHash,
  docxDelete,
  docxEditParagraph,
  docxInsert,
  docxOpen,
  docxRead,
  docxReplace,
} from "../src/index.js";
import { DEFAULT_PARTS, type DocxParts, buildDocx, docWithBody } from "./fixtures.js";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const FIXED_DATE = "2026-06-10T00:00:00Z";

const P2_TEXT = "The term is five (5) years from the Effective Date.";
const P2A = `P2#${anchorHash(P2_TEXT)}`;
const P2_NEW_TEXT = "The term is three (3) years from the Effective Date.";

// The §5/§6 split-run paragraph: plain run + bold run.
const BOLD_SPLIT =
  '<w:p><w:r><w:t xml:space="preserve">The term is five (5) </w:t></w:r>' +
  "<w:r><w:rPr><w:b/></w:rPr><w:t>years from the Effective Date.</w:t></w:r></w:p>";

beforeEach(() => {
  process.env["DOCXENGINE_FIXED_DATE"] = FIXED_DATE;
});

afterEach(() => {
  delete process.env["DOCXENGINE_FIXED_DATE"];
  delete process.env["DOCXENGINE_AUTHOR"];
});

function partsWith(extra: DocxParts): DocxParts {
  return { ...DEFAULT_PARTS, ...extra };
}

function openParts(parts: DocxParts = DEFAULT_PARTS) {
  const session = new Session();
  const res = docxOpen(session, { bytes: Buffer.from(buildDocx(parts)).toString("base64") });
  return { session, docId: res.doc_id, res };
}

function openBody(body: string, extra: DocxParts = {}) {
  return openParts(partsWith({ "word/document.xml": docWithBody(body), ...extra }));
}

function documentXml(session: Session, docId: string): string {
  return session.get(docId).documentXml();
}

function expectCode(fn: () => unknown, code: string): ToolError {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(ToolError);
    expect((e as ToolError).code).toBe(code);
    return e as ToolError;
  }
  return expect.unreachable() as never;
}

// ---------------------------------------------------------------------------
// docx_replace — §4 untracked
// ---------------------------------------------------------------------------

describe("docx_replace (untracked)", () => {
  it("pins the conformance anchors of the §4 worked example", () => {
    expect(P2A).toBe("P2#d337");
    expect(anchorHash(P2_NEW_TEXT)).toBe("eeb0");
  });

  it("coalesces split runs and splices only the touched w:t (§4 worked example)", () => {
    const { session, docId } = openParts();
    const res = docxReplace(session, {
      doc_id: docId,
      anchor: P2A,
      old: "five (5) years",
      new: "three (3) years",
    });
    expect(res).toEqual({ n_replaced: 1, new_anchor: "P2#eeb0" });
    // Run 1 (first overlap) takes prefix + replacement; run 2 keeps its suffix,
    // gains xml:space="preserve", keeps its <w:b/> and rsid attributes.
    expect(documentXml(session, docId)).toContain(
      '<w:p><w:r w:rsidR="00AB12CD" w:rsidRPr="00AB12CD"><w:t>The term is three (3) years</w:t></w:r>' +
        '<w:r w:rsidR="00FF00AA"><w:rPr><w:b/></w:rPr>' +
        '<w:t xml:space="preserve"> from the Effective Date.</w:t></w:r></w:p>',
    );
    expect(docxRead(session, { doc_id: docId, anchor: "P2#eeb0" }).content).toBe(
      `[P2#eeb0] ${P2_NEW_TEXT}`,
    );
  });

  it("removes runs whose w:t is left empty (§4 rule 4)", () => {
    const { session, docId } = openParts();
    const res = docxReplace(session, {
      doc_id: docId,
      anchor: P2A,
      old: "years from the Effective Date.",
      new: "",
    });
    const newAnchor = `P2#${anchorHash("The term is five (5)")}`;
    expect(res).toEqual({ n_replaced: 1, new_anchor: newAnchor });
    const xml = documentXml(session, docId);
    expect(xml).not.toContain("00FF00AA");
    // The untouched first run survives byte-for-byte.
    expect(xml).toContain(
      '<w:r w:rsidR="00AB12CD" w:rsidRPr="00AB12CD"><w:t xml:space="preserve">The term is five (5) </w:t></w:r>',
    );
  });

  it("validates the anchor hash FIRST: stale beats found text", () => {
    const { session, docId } = openParts();
    expectCode(
      () => docxReplace(session, { doc_id: docId, anchor: "P2#0000", old: "five (5)", new: "x" }),
      "anchor_stale",
    );
    expectCode(
      () => docxReplace(session, { doc_id: docId, anchor: "P99#abcd", old: "five (5)", new: "x" }),
      "anchor_not_found",
    );
    expectCode(
      () => docxReplace(session, { doc_id: docId, anchor: "junk", old: "five (5)", new: "x" }),
      "anchor_invalid",
    );
    expectCode(
      () => docxReplace(session, { doc_id: docId, anchor: "T1", old: "Fee", new: "x" }),
      "anchor_invalid",
    );
    expectCode(
      () => docxReplace(session, { doc_id: docId, anchor: "P0#abcd", old: "five", new: "x" }),
      "anchor_invalid",
    );
  });

  it("raises not_found for absent or empty old text", () => {
    const { session, docId } = openParts();
    expectCode(
      () => docxReplace(session, { doc_id: docId, anchor: P2A, old: "zebra", new: "x" }),
      "not_found",
    );
    expectCode(() => docxReplace(session, { doc_id: docId, old: "zebra", new: "x" }), "not_found");
    expectCode(() => docxReplace(session, { doc_id: docId, old: "", new: "x" }), "not_found");
  });

  it("raises ambiguous_target when multiple matches lack all: true", () => {
    const { session, docId } = openBody(
      "<w:p><w:r><w:t>aaa bbb aaa</w:t></w:r></w:p><w:p><w:r><w:t>aaa</w:t></w:r></w:p>",
    );
    const err = expectCode(
      () => docxReplace(session, { doc_id: docId, old: "aaa", new: "z" }),
      "ambiguous_target",
    );
    expect(err.message).toBe("aaa matches 3 times without all: true.");
    const a1 = `P1#${anchorHash("aaa bbb aaa")}`;
    expectCode(
      () => docxReplace(session, { doc_id: docId, anchor: a1, old: "aaa", new: "z" }),
      "ambiguous_target",
    );
  });

  it("returns new_anchor for a whole-document single match", () => {
    const { session, docId } = openParts();
    const res = docxReplace(session, { doc_id: docId, old: "Master", new: "Prime" });
    expect(res).toEqual({
      n_replaced: 1,
      new_anchor: `P1#${anchorHash("Prime Services Agreement")}`,
    });
  });

  it("all: true replaces every occurrence and is idempotent", () => {
    const { session, docId } = openBody(
      "<w:p><w:r><w:t>Acme Corp ltd</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>x Acme Corp y Acme Corp</w:t></w:r></w:p>",
    );
    const res = docxReplace(session, {
      doc_id: docId,
      old: "Acme Corp",
      new: "GlobalTech",
      all: true,
    });
    expect(res).toEqual({
      n_replaced: 3,
      anchors: [
        `P1#${anchorHash("GlobalTech ltd")}`,
        `P2#${anchorHash("x GlobalTech y GlobalTech")}`,
      ],
    });
    expect(docxRead(session, { doc_id: docId }).content.split("\n")).toEqual([
      `[P1#${anchorHash("GlobalTech ltd")}] GlobalTech ltd`,
      `[P2#${anchorHash("x GlobalTech y GlobalTech")}] x GlobalTech y GlobalTech`,
    ]);
    // Idempotent: nothing left to match, not an error.
    expect(
      docxReplace(session, { doc_id: docId, old: "Acme Corp", new: "GlobalTech", all: true }),
    ).toEqual({ n_replaced: 0, anchors: [] });
  });

  it("terminates when the replacement contains the old text", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>aa bb aa</w:t></w:r></w:p>");
    const res = docxReplace(session, { doc_id: docId, old: "aa", new: "aaa", all: true });
    expect(res.n_replaced).toBe(2);
    expect(docxRead(session, { doc_id: docId }).content).toContain("aaa bb aaa");
  });

  it("all: true with an anchor scopes to that paragraph only", () => {
    const { session, docId } = openBody(
      "<w:p><w:r><w:t>aaa bbb</w:t></w:r></w:p><w:p><w:r><w:t>aaa</w:t></w:r></w:p>",
    );
    const res = docxReplace(session, {
      doc_id: docId,
      anchor: `P1#${anchorHash("aaa bbb")}`,
      old: "aaa",
      new: "ccc",
      all: true,
    });
    expect(res).toEqual({ n_replaced: 1, anchors: [`P1#${anchorHash("ccc bbb")}`] });
    expect(docxRead(session, { doc_id: docId, range: "P2..P2" }).content).toBe(
      `[P2#${anchorHash("aaa")}] aaa`,
    );
  });
});

// ---------------------------------------------------------------------------
// docx_replace — §5 tracked
// ---------------------------------------------------------------------------

describe("docx_replace (tracked)", () => {
  const SPEC_BODY =
    "<w:p><w:r><w:t>Intro</w:t></w:r>" +
    '<w:ins w:id="6" w:author="X" w:date="2020-01-01T00:00:00Z">' +
    '<w:r><w:t xml:space="preserve"> note</w:t></w:r></w:ins></w:p>' +
    BOLD_SPLIT;

  it("emits the §5 worked example byte-for-byte (ids continue from max 6)", () => {
    const { session, docId } = openBody(SPEC_BODY);
    const res = docxReplace(session, {
      doc_id: docId,
      anchor: "P2#d337",
      old: "five (5) years",
      new: "three (3) years",
      track_changes: true,
      author: "Claude",
    });
    expect(res).toEqual({ n_replaced: 1, new_anchor: "P2#eeb0" });
    expect(documentXml(session, docId)).toContain(
      '<w:p><w:r><w:t xml:space="preserve">The term is </w:t></w:r>' +
        '<w:del w:id="7" w:author="Claude" w:date="2026-06-10T00:00:00Z">' +
        '<w:r><w:delText xml:space="preserve">five (5) </w:delText></w:r>' +
        "<w:r><w:rPr><w:b/></w:rPr><w:delText>years</w:delText></w:r></w:del>" +
        '<w:ins w:id="8" w:author="Claude" w:date="2026-06-10T00:00:00Z">' +
        "<w:r><w:t>three (3) years</w:t></w:r></w:ins>" +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> from the Effective Date.</w:t></w:r></w:p>',
    );
    // Projection stays as-if-accepted with markers at the span ends.
    expect(docxRead(session, { doc_id: docId, anchor: "P2#eeb0" }).content).toBe(
      "[P2#eeb0] The term is [del by Claude] three (3) years [ins by Claude] " +
        "from the Effective Date.",
    );
  });

  it("defaults the author from DOCXENGINE_AUTHOR, then 'DocxEngine'", () => {
    process.env["DOCXENGINE_AUTHOR"] = "Robo";
    const first = openParts();
    docxReplace(first.session, {
      doc_id: first.docId,
      anchor: P2A,
      old: "five",
      new: "six",
      track_changes: true,
    });
    expect(documentXml(first.session, first.docId)).toContain('w:author="Robo"');
    delete process.env["DOCXENGINE_AUTHOR"];
    const second = openParts();
    docxReplace(second.session, {
      doc_id: second.docId,
      anchor: P2A,
      old: "five",
      new: "six",
      track_changes: true,
    });
    expect(documentXml(second.session, second.docId)).toContain('w:author="DocxEngine"');
  });

  it("allocates wrapper ids in document order across an all: true operation", () => {
    const { session, docId } = openBody(
      "<w:p><w:r><w:t>alpha x</w:t></w:r></w:p><w:p><w:r><w:t>alpha y</w:t></w:r></w:p>",
    );
    const res = docxReplace(session, {
      doc_id: docId,
      old: "alpha",
      new: "beta",
      all: true,
      track_changes: true,
      author: "Claude",
    });
    expect(res.n_replaced).toBe(2);
    expect(res.anchors).toEqual([`P1#${anchorHash("beta x")}`, `P2#${anchorHash("beta y")}`]);
    const xml = documentXml(session, docId);
    const order = [1, 2, 3, 4].map((n) => xml.indexOf(`w:id="${n}"`));
    expect(order.every((p) => p >= 0)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
    expect(xml.indexOf('w:id="2"')).toBeLessThan(xml.indexOf("</w:p>"));
  });

  it("tracked replacement with empty new emits only a w:del", () => {
    const { session, docId } = openParts();
    docxReplace(session, {
      doc_id: docId,
      anchor: P2A,
      old: "five (5) ",
      new: "",
      track_changes: true,
      author: "Claude",
    });
    const xml = documentXml(session, docId);
    expect(xml).toContain('<w:del w:id="1"');
    expect(xml).toContain('<w:delText xml:space="preserve">five (5) </w:delText>');
    expect(xml).not.toContain("<w:ins");
    expect(
      docxRead(session, {
        doc_id: docId,
        anchor: `P2#${anchorHash("The term is years from the Effective Date.")}`,
      }).content,
    ).toContain("[del by Claude]");
  });

  it("is deterministic under DOCXENGINE_FIXED_DATE", () => {
    const runOnce = (): string => {
      const { session, docId } = openParts();
      docxReplace(session, {
        doc_id: docId,
        anchor: P2A,
        old: "five (5) years",
        new: "three (3) years",
        track_changes: true,
        author: "Claude",
      });
      return documentXml(session, docId);
    };
    expect(runOnce()).toBe(runOnce());
  });
});

// ---------------------------------------------------------------------------
// docx_edit_paragraph — §6
// ---------------------------------------------------------------------------

describe("docx_edit_paragraph", () => {
  it("untracked rewrite keeps w:pPr and the first run's rPr (§6a)", () => {
    const { session, docId } = openBody(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
        "<w:r><w:rPr><w:b/></w:rPr><w:t>Old title</w:t></w:r>" +
        '<w:r><w:t xml:space="preserve"> here</w:t></w:r></w:p>',
    );
    const res = docxEditParagraph(session, {
      doc_id: docId,
      anchor: `P1#${anchorHash("Old title here")}`,
      text: "New title",
    });
    expect(res.new_anchor).toBe(`P1#${anchorHash("New title")}`);
    expect(documentXml(session, docId)).toContain(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
        "<w:r><w:rPr><w:b/></w:rPr><w:t>New title</w:t></w:r></w:p>",
    );
    expect(docxRead(session, { doc_id: docId, anchor: res.new_anchor }).content).toBe(
      `[${res.new_anchor} H1] New title`,
    );
  });

  it("tracked rewrite emits the §6 minimal diff — never delete-all/insert-all", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>three year term</w:t></w:r></w:p>");
    const res = docxEditParagraph(session, {
      doc_id: docId,
      anchor: `P1#${anchorHash("three year term")}`,
      text: "five year initial term",
      track_changes: true,
      author: "Claude",
    });
    // §6a: n = max(#deleted, #inserted) = max(1, 2).
    expect(res).toEqual({
      new_anchor: `P1#${anchorHash("five year initial term")}`,
      diff: "~2 words changed",
    });
    expect(documentXml(session, docId)).toContain(
      "<w:p>" +
        '<w:del w:id="1" w:author="Claude" w:date="2026-06-10T00:00:00Z">' +
        '<w:r><w:delText xml:space="preserve">three </w:delText></w:r></w:del>' +
        '<w:ins w:id="2" w:author="Claude" w:date="2026-06-10T00:00:00Z">' +
        '<w:r><w:t xml:space="preserve">five </w:t></w:r></w:ins>' +
        '<w:r><w:t xml:space="preserve">year </w:t></w:r>' +
        '<w:ins w:id="3" w:author="Claude" w:date="2026-06-10T00:00:00Z">' +
        '<w:r><w:t xml:space="preserve">initial </w:t></w:r></w:ins>' +
        "<w:r><w:t>term</w:t></w:r></w:p>",
    );
  });

  it("keep and replace spans preserve per-run formatting (§6a)", () => {
    const { session, docId } = openBody(BOLD_SPLIT);
    docxEditParagraph(session, {
      doc_id: docId,
      anchor: `P1#${anchorHash(P2_TEXT)}`,
      text: "The term is five (5) years from the Start Date.",
      track_changes: true,
      author: "Claude",
    });
    const xml = documentXml(session, docId);
    // Kept spans re-emit per-run portions with each run's own rPr.
    expect(xml).toContain('<w:r><w:t xml:space="preserve">The term is five (5) </w:t></w:r>');
    expect(xml).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">years from the </w:t></w:r>',
    );
    // The deleted span keeps the bold run's rPr; the replacement ins inherits it (§5).
    expect(xml).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:delText xml:space="preserve">Effective </w:delText>' +
        "</w:r></w:del>",
    );
    expect(xml).toContain(
      `<w:ins w:id="2" w:author="Claude" w:date="${FIXED_DATE}">` +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Start </w:t></w:r></w:ins>',
    );
    expect(xml).toContain("<w:r><w:rPr><w:b/></w:rPr><w:t>Date.</w:t></w:r>");
  });

  it("pure insertion takes the rPr of the run at the insertion offset (§6a)", () => {
    const { session, docId } = openBody(BOLD_SPLIT);
    docxEditParagraph(session, {
      doc_id: docId,
      anchor: `P1#${anchorHash(P2_TEXT)}`,
      text: "The term is five (5) whole years from the Effective Date.",
      track_changes: true,
      author: "Claude",
    });
    const xml = documentXml(session, docId);
    // Insertion offset 21 is exactly where the bold run starts → bold rPr.
    expect(xml).toContain(
      `<w:ins w:id="1" w:author="Claude" w:date="${FIXED_DATE}">` +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">whole </w:t></w:r></w:ins>',
    );
    expect(xml).not.toContain("<w:del ");
  });

  it("tracked rewrite of an empty paragraph becomes a single insertion", () => {
    const { session, docId } = openParts();
    const res = docxEditParagraph(session, {
      doc_id: docId,
      anchor: "P3#e3b0",
      text: "Added text",
      track_changes: true,
    });
    expect(res.new_anchor).toBe(`P3#${anchorHash("Added text")}`);
    expect(documentXml(session, docId)).toContain(
      '<w:p><w:ins w:id="1" w:author="DocxEngine" w:date="2026-06-10T00:00:00Z">' +
        "<w:r><w:t>Added text</w:t></w:r></w:ins></w:p>",
    );
  });

  it("reports a zero-word diff for an unchanged rewrite", () => {
    const { session, docId } = openParts();
    const res = docxEditParagraph(session, {
      doc_id: docId,
      anchor: P2A,
      text: P2_TEXT,
      track_changes: true,
    });
    expect(res).toEqual({ new_anchor: P2A, diff: "~0 words changed" });
    expect(documentXml(session, docId)).not.toContain("<w:ins");
  });

  it("uses the singular noun for a one-word diff", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>alpha beta</w:t></w:r></w:p>");
    const res = docxEditParagraph(session, {
      doc_id: docId,
      anchor: `P1#${anchorHash("alpha beta")}`,
      text: "alpha gamma",
    });
    expect(res.diff).toBe("~1 word changed");
  });

  it("validates the anchor before rewriting", () => {
    const { session, docId } = openParts();
    expectCode(
      () => docxEditParagraph(session, { doc_id: docId, anchor: "P2#0000", text: "x" }),
      "anchor_stale",
    );
    expectCode(
      () => docxEditParagraph(session, { doc_id: docId, anchor: "P9#abcd", text: "x" }),
      "anchor_not_found",
    );
    expectCode(
      () => docxEditParagraph(session, { doc_id: docId, anchor: "T1", text: "x" }),
      "anchor_invalid",
    );
  });
});

// ---------------------------------------------------------------------------
// docx_insert
// ---------------------------------------------------------------------------

describe("docx_insert", () => {
  const STYLED_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:styles ${W_NS}>` +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>' +
    "</w:styles>";

  it("inserts plain text after an anchor and returns fresh anchors", () => {
    const { session, docId } = openParts();
    const res = docxInsert(session, { doc_id: docId, after: "P1#515a", content: "Hello world" });
    const anchor = `P2#${anchorHash("Hello world")}`;
    expect(res).toEqual({ new_anchors: [anchor] });
    expect(docxRead(session, { doc_id: docId, range: "P2..P2" }).content).toBe(
      `[${anchor}] Hello world`,
    );
    // The old P2 shifted to P3.
    expect(docxRead(session, { doc_id: docId, range: "P3..P3" }).content).toBe(
      `[P3#${anchorHash(P2_TEXT)}] ${P2_TEXT}`,
    );
  });

  it("inserts before an anchor", () => {
    const { session, docId } = openParts();
    const res = docxInsert(session, { doc_id: docId, before: "P1#515a", content: "Preamble" });
    expect(res).toEqual({ new_anchors: [`P1#${anchorHash("Preamble")}`] });
    expect(docxRead(session, { doc_id: docId, range: "P2..P2" }).content).toBe(
      "[P2#515a H1] Master Services Agreement",
    );
  });

  it("parses minimal markdown: headings, bullets, CRLF, blank lines", () => {
    const { session, docId } = openParts();
    const content = "# Title\r\n\n#### Deep\n- item one\n* item two\n   \nplain";
    const res = docxInsert(session, { doc_id: docId, after: "P1#515a", content });
    expect(res.new_anchors).toEqual([
      `P2#${anchorHash("Title")}`,
      `P3#${anchorHash("Deep")}`,
      `P4#${anchorHash("item one")}`,
      `P5#${anchorHash("item two")}`,
      `P6#${anchorHash("plain")}`,
    ]);
    const xml = documentXml(session, docId);
    expect(xml).toContain(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>',
    );
    expect(xml).toContain('<w:pStyle w:val="Heading4"/>');
    expect(xml).toContain(
      '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>' +
        "<w:r><w:t>item one</w:t></w:r></w:p>",
    );
    expect(xml).toContain("<w:p><w:r><w:t>plain</w:t></w:r></w:p>");
    const projected = docxRead(session, { doc_id: docId }).content;
    expect(projected).toContain(`[P2#${anchorHash("Title")} H1] Title`);
    expect(projected).toContain(`[P3#${anchorHash("Deep")} H4] Deep`);
  });

  it("resolves a style override against styles.xml ('Heading 2' → Heading2)", () => {
    const { session, docId } = openParts(partsWith({ "word/styles.xml": STYLED_XML }));
    docxInsert(session, {
      doc_id: docId,
      after: "P1#515a",
      content: "## sub\n- item",
      style: "Heading 2",
    });
    // The override applies to every inserted paragraph (markdown styles lose).
    expect(documentXml(session, docId)).toContain(
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>sub</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>item</w:t></w:r></w:p>',
    );
    expect(documentXml(session, docId)).not.toContain("ListParagraph");
    expectCode(
      () =>
        docxInsert(session, {
          doc_id: docId,
          after: "P1#515a",
          content: "x",
          style: "No Such Style",
        }),
      "style_unknown",
    );
  });

  it("wraps tracked insertions in one w:ins per paragraph with §5 metadata", () => {
    process.env["DOCXENGINE_AUTHOR"] = "Robo";
    const { session, docId } = openParts();
    const res = docxInsert(session, {
      doc_id: docId,
      after: "P1#515a",
      content: "Inserted clause\nSecond clause",
      track_changes: true,
    });
    expect(res.new_anchors).toEqual([
      `P2#${anchorHash("Inserted clause")}`,
      `P3#${anchorHash("Second clause")}`,
    ]);
    expect(documentXml(session, docId)).toContain(
      '<w:p><w:ins w:id="1" w:author="Robo" w:date="2026-06-10T00:00:00Z">' +
        "<w:r><w:t>Inserted clause</w:t></w:r></w:ins></w:p>" +
        '<w:p><w:ins w:id="2" w:author="Robo" w:date="2026-06-10T00:00:00Z">' +
        "<w:r><w:t>Second clause</w:t></w:r></w:ins></w:p>",
    );
    expect(docxRead(session, { doc_id: docId, anchor: res.new_anchors[0] as string }).content).toBe(
      `[P2#${anchorHash("Inserted clause")}] Inserted clause [ins by Robo]`,
    );
  });

  it("treats whitespace-only content as a no-op", () => {
    const { session, docId } = openParts();
    const before = documentXml(session, docId);
    expect(docxInsert(session, { doc_id: docId, after: P2A, content: "  \n \n" })).toEqual({
      new_anchors: [],
    });
    expect(documentXml(session, docId)).toBe(before);
  });

  it("requires exactly one of after/before and validates the anchor first", () => {
    const { session, docId } = openParts();
    expectCode(() => docxInsert(session, { doc_id: docId, content: "x" }), "anchor_invalid");
    expectCode(
      () =>
        docxInsert(session, { doc_id: docId, after: "P1#515a", before: "P2#d337", content: "x" }),
      "anchor_invalid",
    );
    expectCode(
      () => docxInsert(session, { doc_id: docId, after: "P1#dead", content: "x" }),
      "anchor_stale",
    );
    expectCode(
      () => docxInsert(session, { doc_id: docId, after: "P77#abcd", content: "x" }),
      "anchor_not_found",
    );
  });
});

// ---------------------------------------------------------------------------
// docx_delete
// ---------------------------------------------------------------------------

describe("docx_delete", () => {
  it("deletes a single paragraph by anchor and shifts ordinals", () => {
    const { session, docId } = openParts();
    expect(docxDelete(session, { doc_id: docId, anchor: P2A })).toEqual({ ok: true, deleted: 1 });
    const lines = docxRead(session, { doc_id: docId }).content.split("\n");
    expect(lines[0]).toBe("[P1#515a H1] Master Services Agreement");
    expect(lines[lines.length - 1]).toBe("[P2#e3b0]");
    expect(lines.join("\n")).not.toContain("five (5)");
  });

  it("deletes a contiguous paragraph range", () => {
    const { session, docId } = openParts();
    expect(docxDelete(session, { doc_id: docId, range: "P1..P2" })).toEqual({
      ok: true,
      deleted: 2,
    });
    expect(docxRead(session, { doc_id: docId, range: "P1..P1" }).content).toBe("[P1#e3b0]");
  });

  it("validates range endpoints: hash when present, ordinal always (§6a)", () => {
    const { session, docId } = openParts();
    expectCode(() => docxDelete(session, { doc_id: docId, range: "P2#0000..P3" }), "anchor_stale");
    expectCode(
      () => docxDelete(session, { doc_id: docId, range: `${P2A}..P99` }),
      "anchor_not_found",
    );
    expectCode(() => docxDelete(session, { doc_id: docId, range: "bogus" }), "anchor_invalid");
    expectCode(() => docxDelete(session, { doc_id: docId, range: "P3..P1" }), "anchor_invalid");
    expectCode(
      () => docxDelete(session, { doc_id: docId, anchor: P2A, range: "P1..P2" }),
      "anchor_invalid",
    );
    expectCode(() => docxDelete(session, { doc_id: docId }), "anchor_invalid");
    // Valid endpoint hashes pass.
    expect(docxDelete(session, { doc_id: docId, range: `P1#515a..${P2A}` })).toEqual({
      ok: true,
      deleted: 2,
    });
  });

  it("tracked deletion wraps each paragraph's content after w:pPr in one w:del", () => {
    const { session, docId } = openBody(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Gone soon</w:t></w:r></w:p>' +
        "<w:p/>" +
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Second</w:t></w:r><w:r><w:t xml:space="preserve"> clause.</w:t></w:r></w:p>',
    );
    const res = docxDelete(session, {
      doc_id: docId,
      range: "P1..P3",
      track_changes: true,
      author: "Claude",
    });
    expect(res).toEqual({ ok: true, deleted: 3 });
    const xml = documentXml(session, docId);
    // w:pPr survives outside the wrapper; runs keep their own bytes with
    // w:t renamed to w:delText.
    expect(xml).toContain(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
        '<w:del w:id="1" w:author="Claude" w:date="2026-06-10T00:00:00Z">' +
        "<w:r><w:delText>Gone soon</w:delText></w:r></w:del></w:p>",
    );
    // The empty paragraph is counted but unchanged; ids stay sequential.
    expect(xml).toContain("<w:p/>");
    expect(xml).toContain(
      '<w:p><w:del w:id="2" w:author="Claude" w:date="2026-06-10T00:00:00Z">' +
        "<w:r><w:rPr><w:b/></w:rPr><w:delText>Second</w:delText></w:r>" +
        '<w:r><w:delText xml:space="preserve"> clause.</w:delText></w:r></w:del></w:p>',
    );
    // As-if-accepted projection: empty paragraphs with deletion markers.
    expect(docxRead(session, { doc_id: docId }).content.split("\n")).toEqual([
      "[P1#e3b0 H1] [del by Claude]",
      "[P2#e3b0]",
      "[P3#e3b0] [del by Claude]",
    ]);
  });

  it("tracked deletion keeps the paragraph count (the mark survives in MVP)", () => {
    const { session, docId } = openParts();
    docxDelete(session, { doc_id: docId, anchor: P2A, track_changes: true, author: "Claude" });
    // As-if-accepted text is now empty, but the paragraph mark survives…
    expect(docxRead(session, { doc_id: docId, range: "P2..P2" }).content).toBe(
      "[P2#e3b0] [del by Claude]",
    );
    // …so the following paragraph keeps its ordinal.
    expect(docxRead(session, { doc_id: docId, range: "P3..P3" }).content).toBe("[P3#e3b0]");
  });

  it("counts empty paragraphs in a tracked range without emitting wrappers", () => {
    const { session, docId } = openParts();
    const res = docxDelete(session, { doc_id: docId, range: "P3..P3", track_changes: true });
    expect(res).toEqual({ ok: true, deleted: 1 });
    expect(documentXml(session, docId)).not.toContain("<w:del");
  });
});
