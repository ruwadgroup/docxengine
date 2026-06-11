/**
 * Stage-2 read surface: docx_open / docx_outline / docx_read / docx_search.
 * Mirrors the Python stage-2 cases: headings (basedOn walk), split runs,
 * ins/del/comment markers, fragmented search, pagination.
 */
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { describe, expect, it } from "vitest";

import {
  READ_CHAR_CAP,
  Session,
  ToolError,
  anchorHash,
  docxOpen,
  docxOutline,
  docxRead,
  docxSearch,
} from "../src/index.js";
import { DEFAULT_PARTS, type DocxParts, buildDocx, docWithBody } from "./fixtures.js";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const P2_TEXT = "The term is five (5) years from the Effective Date.";
const P2A = `P2#${anchorHash(P2_TEXT)}`;

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
// Reusable parts
// ---------------------------------------------------------------------------

const CHAIN_STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<w:styles ${W_NS}>` +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="SectionHead"><w:basedOn w:val="Heading2"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="SubHead"><w:basedOn w:val="SectionHead"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Looper"><w:basedOn w:val="Looper"/></w:style>' +
  "</w:styles>";

const NUMBERING_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<w:numbering ${W_NS}>` +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/></w:lvl>' +
  "</w:abstractNum>" +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  "</w:numbering>";

const COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<w:comments ${W_NS}>` +
  '<w:comment w:id="7" w:author="J.Doe" w:initials="JD">' +
  "<w:p><w:r><w:t>tighten this</w:t></w:r></w:p>" +
  "</w:comment></w:comments>";

const FOOTNOTES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<w:footnotes ${W_NS}>` +
  '<w:footnote w:id="1"><w:p><w:r><w:t>Footnote text</w:t></w:r></w:p></w:footnote>' +
  "</w:footnotes>";

const INS_BODY =
  '<w:p><w:r><w:t xml:space="preserve">First </w:t></w:r>' +
  '<w:ins w:id="1" w:author="Jane" w:date="2026-01-01T00:00:00Z">' +
  "<w:r><w:t>obligation</w:t></w:r></w:ins></w:p>";

const DEL_BODY =
  '<w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r>' +
  '<w:del w:id="2" w:author="Bob" w:date="2026-01-01T00:00:00Z">' +
  '<w:r><w:delText xml:space="preserve">gone </w:delText></w:r></w:del>' +
  "<w:r><w:t>this</w:t></w:r></w:p>";

const COMMENT_BODY =
  "<w:p><w:r><w:t>Hello</w:t></w:r>" + '<w:r><w:commentReference w:id="7"/></w:r></w:p>';

const HEADER2_XML = `<w:hdr ${W_NS}><w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>`;
const HEADER10_XML = `<w:hdr ${W_NS}><w:p><w:r><w:t>Second header</w:t></w:r></w:p></w:hdr>`;

// ---------------------------------------------------------------------------
// docx_open
// ---------------------------------------------------------------------------

describe("docx_open", () => {
  it("hands out sequential doc ids d1, d2 within one session", () => {
    const { session, docId } = openParts();
    expect(docId).toBe("d1");
    const second = docxOpen(session, {
      bytes: Buffer.from(buildDocx()).toString("base64"),
    });
    expect(second.doc_id).toBe("d2");
    expect(session.get("d1").id).toBe("d1");
  });

  it("summarizes the document per the §2a pinned format", () => {
    const { res } = openParts();
    expect(res).toEqual({
      doc_id: "d1",
      summary: "Master Services Agreement — 3 paragraphs, 1 section, 1 table",
      n_paragraphs: 3,
      has_tracked_changes: false,
      has_comments: false,
    });
  });

  it("flags tracked changes and comments", () => {
    expect(openBody(INS_BODY).res.has_tracked_changes).toBe(true);
    expect(openBody(DEL_BODY).res.has_tracked_changes).toBe(true);
    const withComment = openBody(COMMENT_BODY, { "word/comments.xml": COMMENTS_XML }).res;
    expect(withComment.has_comments).toBe(true);
    expect(withComment.has_tracked_changes).toBe(false);
  });

  it("opens from a filesystem path", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-read-"));
    try {
      const file = path.join(dir, "fixture.docx");
      fs.writeFileSync(file, buildDocx());
      const res = docxOpen(new Session(), { path: file });
      expect(res.doc_id).toBe("d1");
      expect(res.n_paragraphs).toBe(3);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it("rejects a call with neither path nor bytes, and unreadable paths", () => {
    expectCode(() => docxOpen(new Session(), {}), "open_failed");
    expectCode(() => docxOpen(new Session(), { path: "/nonexistent/x.docx" }), "open_failed");
  });

  it("raises doc_not_found for unknown doc ids", () => {
    const { session } = openParts();
    expectCode(() => docxOutline(session, { doc_id: "d99" }), "doc_not_found");
  });
});

// ---------------------------------------------------------------------------
// docx_outline
// ---------------------------------------------------------------------------

describe("docx_outline", () => {
  it("lists headings with anchors and tables with dims + @after anchor", () => {
    const { session, docId } = openParts();
    expect(docxOutline(session, { doc_id: docId })).toEqual({
      outline: [{ anchor: "P1#515a", level: 1, text: "Master Services Agreement" }],
      tables: [{ anchor: "T1", dims: "2×2", after: P2A }],
    });
  });

  it("resolves heading levels through the styles.xml basedOn chain", () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="SubHead"/></w:pPr><w:r><w:t>Chained heading</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Looper"/></w:pPr><w:r><w:t>Cycle</w:t></w:r></w:p>' +
      "<w:p><w:r><w:t>Body text</w:t></w:r></w:p>";
    const { session, docId } = openBody(body, { "word/styles.xml": CHAIN_STYLES_XML });
    const { outline } = docxOutline(session, { doc_id: docId });
    expect(outline).toEqual([
      { anchor: `P1#${anchorHash("Chained heading")}`, level: 2, text: "Chained heading" },
    ]);
  });

  it("omits 'after' for a table that precedes every paragraph", () => {
    const body =
      "<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid>" +
      "<w:tr><w:tc><w:p><w:r><w:t>only</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
      "<w:p><w:r><w:t>after the table</w:t></w:r></w:p>";
    const { session, docId } = openBody(body);
    const { tables } = docxOutline(session, { doc_id: docId });
    expect(tables).toEqual([{ anchor: "T1", dims: "1×1" }]);
  });
});

// ---------------------------------------------------------------------------
// docx_read
// ---------------------------------------------------------------------------

describe("docx_read", () => {
  const TABLE_LINES = [
    `[T1 2×2 @after:${P2A}]`,
    "| Term | Value |",
    "| --- | --- |",
    "| Fee | $100 |",
  ];

  it("reads a single anchored paragraph (heading annotation, coalesced split runs)", () => {
    const { session, docId } = openParts();
    expect(docxRead(session, { doc_id: docId, anchor: "P1#515a" })).toEqual({
      content: "[P1#515a H1] Master Services Agreement",
    });
  });

  it("resolves by ordinal even when the anchor hash is stale (refresh path)", () => {
    const { session, docId } = openParts();
    expect(docxRead(session, { doc_id: docId, anchor: "P1#0000" }).content).toBe(
      "[P1#515a H1] Master Services Agreement",
    );
  });

  it("expands a window of body blocks around the anchor", () => {
    const { session, docId } = openParts();
    const res = docxRead(session, { doc_id: docId, anchor: P2A, window: 1 });
    expect(res.content).toBe(
      ["[P1#515a H1] Master Services Agreement", `[${P2A}] ${P2_TEXT}`, ...TABLE_LINES].join("\n"),
    );
  });

  it("reads the whole body when neither anchor nor range is given", () => {
    const { session, docId } = openParts();
    const res = docxRead(session, { doc_id: docId });
    expect(res.content.split("\n")).toEqual([
      "[P1#515a H1] Master Services Agreement",
      `[${P2A}] ${P2_TEXT}`,
      ...TABLE_LINES,
      "[P3#e3b0]",
    ]);
    expect(res.continuation).toBeUndefined();
  });

  it("reads paragraph ranges, including tables that lie between the endpoints", () => {
    const { session, docId } = openParts();
    expect(docxRead(session, { doc_id: docId, range: "P1..P2" }).content).toBe(
      ["[P1#515a H1] Master Services Agreement", `[${P2A}] ${P2_TEXT}`].join("\n"),
    );
    expect(docxRead(session, { doc_id: docId, range: "P2..P3" }).content).toBe(
      [`[${P2A}] ${P2_TEXT}`, ...TABLE_LINES, "[P3#e3b0]"].join("\n"),
    );
    // Hash suffixes on range endpoints are accepted and ignored.
    expect(docxRead(session, { doc_id: docId, range: `${P2A}..P3#e3b0` }).content).toBe(
      docxRead(session, { doc_id: docId, range: "P2..P3" }).content,
    );
  });

  it("clamps the window at the document edges", () => {
    const { session, docId } = openParts();
    expect(docxRead(session, { doc_id: docId, anchor: "P1#515a", window: 5 }).content).toBe(
      docxRead(session, { doc_id: docId }).content,
    );
  });

  it("counts blocks for windows, so a table fills a window slot", () => {
    const { session, docId } = openParts();
    expect(docxRead(session, { doc_id: docId, anchor: "P3#e3b0", window: 1 }).content).toBe(
      [...TABLE_LINES, "[P3#e3b0]"].join("\n"),
    );
  });

  it("lets anchor win when both anchor and range are given", () => {
    const { session, docId } = openParts();
    expect(docxRead(session, { doc_id: docId, anchor: P2A, range: "P1..P1" }).content).toBe(
      `[${P2A}] ${P2_TEXT}`,
    );
  });

  it("reads a table anchor with @after pointing at the nearest preceding paragraph", () => {
    const { session, docId } = openParts();
    expect(docxRead(session, { doc_id: docId, anchor: "T1" }).content).toBe(TABLE_LINES.join("\n"));
  });

  it("renders an empty paragraph as the bracket alone, no trailing space", () => {
    const { session, docId } = openParts();
    expect(docxRead(session, { doc_id: docId, range: "P3..P3" }).content).toBe("[P3#e3b0]");
  });

  it("appends [ins by …] at the end of the inserted span; hash stays as-if-accepted", () => {
    const { session, docId } = openBody(INS_BODY);
    const anchor = `P1#${anchorHash("First obligation")}`;
    expect(docxRead(session, { doc_id: docId, anchor }).content).toBe(
      `[${anchor}] First obligation [ins by Jane]`,
    );
  });

  it("marks deletions at the deletion point without showing w:delText", () => {
    const { session, docId } = openBody(DEL_BODY);
    const anchor = `P1#${anchorHash("Keep this")}`;
    expect(docxRead(session, { doc_id: docId, anchor }).content).toBe(
      `[${anchor}] Keep [del by Bob] this`,
    );
  });

  it("projects a fully deleted paragraph as the marker alone", () => {
    const body =
      '<w:p><w:del w:id="9" w:author="Bob" w:date="2026-01-01T00:00:00Z">' +
      "<w:r><w:delText>gone entirely</w:delText></w:r></w:del></w:p><w:p/>";
    const { session, docId } = openBody(body);
    expect(docxRead(session, { doc_id: docId }).content).toBe("[P1#e3b0] [del by Bob]\n[P2#e3b0]");
  });

  it("appends comment markers with the author resolved from comments.xml", () => {
    const { session, docId } = openBody(COMMENT_BODY, { "word/comments.xml": COMMENTS_XML });
    const anchor = `P1#${anchorHash("Hello")}`;
    expect(docxRead(session, { doc_id: docId, anchor }).content).toBe(
      `[${anchor}] Hello [comment:C7 by J.Doe]`,
    );
  });

  it("annotates lists via numbering.xml (ol vs ul, level, heading order)", () => {
    const body =
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      "<w:r><w:t>First obligation</w:t></w:r></w:p>" +
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>' +
      "<w:r><w:t>Bullet point</w:t></w:r></w:p>" +
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:numPr><w:ilvl w:val="1"/>' +
      '<w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Numbered heading</w:t></w:r></w:p>';
    const { session, docId } = openBody(body, { "word/numbering.xml": NUMBERING_XML });
    expect(docxRead(session, { doc_id: docId }).content.split("\n")).toEqual([
      `[P1#${anchorHash("First obligation")} List:ol L1] First obligation`,
      `[P2#${anchorHash("Bullet point")} List:ul L1] Bullet point`,
      `[P3#${anchorHash("Numbered heading")} H2 List:ol L2] Numbered heading`,
    ]);
  });

  it("reads non-body stories; a missing story part reads as empty", () => {
    const { session, docId } = openParts(partsWith({ "word/footnotes.xml": FOOTNOTES_XML }));
    expect(docxRead(session, { doc_id: docId, scope: "footnotes" }).content).toBe(
      `[P1#${anchorHash("Footnote text")}] Footnote text`,
    );
    expect(docxRead(session, { doc_id: docId, scope: "headers" }).content).toBe("");
  });

  it("reads the comments story (each w:comment paragraph gets a story anchor)", () => {
    const { session, docId } = openParts(partsWith({ "word/comments.xml": COMMENTS_XML }));
    expect(docxRead(session, { doc_id: docId, scope: "comments" }).content).toBe(
      `[P1#${anchorHash("tighten this")}] tighten this`,
    );
  });

  it("concatenates header parts numerically (header2 before header10), ordinals across parts", () => {
    const { session, docId } = openParts(
      partsWith({ "word/header10.xml": HEADER10_XML, "word/header2.xml": HEADER2_XML }),
    );
    expect(docxRead(session, { doc_id: docId, scope: "headers" }).content).toBe(
      [
        `[P1#${anchorHash("First header")}] First header`,
        `[P2#${anchorHash("Second header")}] Second header`,
      ].join("\n"),
    );
  });

  it("rejects bad anchors, ranges, and scopes with the spec error codes", () => {
    const { session, docId } = openParts();
    for (const bad of ["junk", "X9", "P0#1234", "P1#XYZW", "p1#abcd"]) {
      expectCode(() => docxRead(session, { doc_id: docId, anchor: bad }), "anchor_invalid");
    }
    expectCode(() => docxRead(session, { doc_id: docId, anchor: "P99#abcd" }), "anchor_not_found");
    expectCode(() => docxRead(session, { doc_id: docId, anchor: "T2" }), "anchor_not_found");
    expectCode(() => docxRead(session, { doc_id: docId, range: "P5..P2" }), "anchor_invalid");
    expectCode(() => docxRead(session, { doc_id: docId, range: "bogus" }), "anchor_invalid");
    expectCode(() => docxRead(session, { doc_id: docId, range: "P1..P99" }), "anchor_not_found");
    expectCode(() => docxRead(session, { doc_id: docId, scope: "margins" }), "anchor_invalid");
  });

  it("paginates long ranges with a continuation range token", () => {
    const n = 30;
    const body = Array.from(
      { length: n },
      (_, i) => `<w:p><w:r><w:t>p${i + 1} ${"x".repeat(990)}</w:t></w:r></w:p>`,
    ).join("");
    const { session, docId } = openBody(body);
    const first = docxRead(session, { doc_id: docId });
    expect(first.continuation).toMatch(/^P[0-9]+\.\.P30$/);
    expect(first.content.length).toBeLessThanOrEqual(READ_CHAR_CAP);

    const lines: string[] = [];
    let pages = 0;
    let res = first;
    for (;;) {
      lines.push(...res.content.split("\n"));
      pages++;
      if (res.continuation === undefined) break;
      res = docxRead(session, { doc_id: docId, range: res.continuation });
    }
    expect(pages).toBeGreaterThan(1);
    expect(lines).toHaveLength(n);
    expect(lines[0]).toMatch(/^\[P1#[0-9a-f]{4}\] p1 /);
    expect(lines[n - 1]).toMatch(/^\[P30#[0-9a-f]{4}\] p30 /);
  });

  it("always returns the first block even when it alone exceeds the budget", () => {
    const body =
      `<w:p><w:r><w:t>${"y".repeat(READ_CHAR_CAP + 100)}</w:t></w:r></w:p>` +
      "<w:p><w:r><w:t>tail</w:t></w:r></w:p>";
    const { session, docId } = openBody(body);
    const res = docxRead(session, { doc_id: docId });
    expect(res.content.length).toBeGreaterThan(READ_CHAR_CAP);
    expect(res.content.split("\n")).toHaveLength(1);
    expect(res.continuation).toBe("P2..P2");
  });
});

// ---------------------------------------------------------------------------
// docx_search
// ---------------------------------------------------------------------------

describe("docx_search", () => {
  it("matches text fragmented across rsid-split runs (coalesced text)", () => {
    const { session, docId } = openParts();
    const res = docxSearch(session, { doc_id: docId, query: "five (5) years from" });
    expect(res).toEqual({
      matches: [
        {
          anchor: P2A,
          snippet: P2_TEXT,
          context: "Master Services Agreement",
        },
      ],
      n_matches: 1,
    });
  });

  it("returns one entry per occurrence in document order", () => {
    const { session, docId } = openBody("<w:p><w:r><w:t>alpha beta alpha</w:t></w:r></w:p>");
    const res = docxSearch(session, { doc_id: docId, query: "alpha" });
    expect(res.n_matches).toBe(2);
    expect(res.matches.map((m) => m.anchor)).toEqual([
      `P1#${anchorHash("alpha beta alpha")}`,
      `P1#${anchorHash("alpha beta alpha")}`,
    ]);
  });

  it("truncates snippets with ellipses on the cut sides", () => {
    const text = `${"A".repeat(60)} needle ${"B".repeat(60)}`;
    const { session, docId } = openBody(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`);
    const res = docxSearch(session, { doc_id: docId, query: "needle" });
    const snippet = res.matches[0]?.snippet as string;
    expect(snippet.startsWith("…")).toBe(true);
    expect(snippet.endsWith("…")).toBe(true);
    expect(snippet).toContain(" needle ");
  });

  it("sees the document as-if-accepted: delText is never matched, ins text is", () => {
    const del = openBody(DEL_BODY);
    expect(docxSearch(del.session, { doc_id: del.docId, query: "gone" }).n_matches).toBe(0);
    const ins = openBody(INS_BODY);
    expect(docxSearch(ins.session, { doc_id: ins.docId, query: "obligation" }).n_matches).toBe(1);
  });

  it("matches text spanning a tracked-change boundary; markers are not searchable", () => {
    const { session, docId } = openBody(INS_BODY);
    expect(docxSearch(session, { doc_id: docId, query: "First obligation" }).n_matches).toBe(1);
    expect(docxSearch(session, { doc_id: docId, query: "[ins by Jane]" }).n_matches).toBe(0);
  });

  it("searches the comments story", () => {
    const { session, docId } = openParts(partsWith({ "word/comments.xml": COMMENTS_XML }));
    const res = docxSearch(session, { doc_id: docId, query: "tighten", scope: "comments" });
    expect(res.n_matches).toBe(1);
    expect(res.matches[0]?.anchor).toBe(`P1#${anchorHash("tighten this")}`);
  });

  it("supports regex queries and rejects invalid patterns", () => {
    const { session, docId } = openParts();
    const res = docxSearch(session, { doc_id: docId, query: "f[io]ve \\(5\\)", regex: true });
    expect(res.n_matches).toBe(1);
    expect(res.matches[0]?.anchor).toBe(P2A);
    expectCode(() => docxSearch(session, { doc_id: docId, query: "(", regex: true }), "not_found");
    expectCode(() => docxSearch(session, { doc_id: docId, query: "" }), "not_found");
  });

  it("restricts matching to a paragraph-range scope", () => {
    const { session, docId } = openParts();
    expect(docxSearch(session, { doc_id: docId, query: "Master", scope: "P2..P3" }).n_matches).toBe(
      0,
    );
    expect(docxSearch(session, { doc_id: docId, query: "Master" }).n_matches).toBe(1);
    expectCode(
      () => docxSearch(session, { doc_id: docId, query: "Master", scope: "weird" }),
      "not_found",
    );
  });

  it("omits context when no heading precedes the match", () => {
    const body =
      "<w:p><w:r><w:t>Intro words</w:t></w:r></w:p>" +
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
      "<w:p><w:r><w:t>Body under title</w:t></w:r></w:p>";
    const { session, docId } = openBody(body);
    const before = docxSearch(session, { doc_id: docId, query: "Intro" });
    expect(before.matches[0]?.context).toBeUndefined();
    const after = docxSearch(session, { doc_id: docId, query: "under" });
    expect(after.matches[0]?.context).toBe("Title");
  });

  it("returns an empty result (not an error) when nothing matches", () => {
    const { session, docId } = openParts();
    expect(docxSearch(session, { doc_id: docId, query: "zebra" })).toEqual({
      matches: [],
      n_matches: 0,
    });
  });

  it("searches non-body stories", () => {
    const { session, docId } = openParts(partsWith({ "word/footnotes.xml": FOOTNOTES_XML }));
    const res = docxSearch(session, { doc_id: docId, query: "Footnote", scope: "footnotes" });
    expect(res.n_matches).toBe(1);
    expect(res.matches[0]?.anchor).toBe(`P1#${anchorHash("Footnote text")}`);
  });
});
