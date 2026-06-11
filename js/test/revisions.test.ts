/**
 * Stage-3 revision surface: docx_revision list/accept/reject/accept_all/
 * reject_all per spec/algorithms.md §7/§6a. Mirrors the Python stage-3 cases
 * (python/tests/test_revisions.py) and the redlines conformance corpus
 * (Alice owns R1–R2, Bob owns R3–R4).
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  Session,
  ToolError,
  anchorHash,
  docxOpen,
  docxRead,
  docxReplace,
  docxRevision,
} from "../src/index.js";
import { DEFAULT_PARTS, type DocxParts, buildDocx, docWithBody } from "./fixtures.js";

const FIXED_DATE = "2026-06-10T00:00:00Z";
const DATE_ALICE = "2026-01-15T09:30:00Z";
const DATE_BOB = "2026-02-20T16:45:00Z";

// Mirrors conformance/corpus/redlines: Alice revises P2, Bob revises P3.
const REDLINES_BODY =
  "<w:p><w:r><w:t>Revision History</w:t></w:r></w:p>" +
  '<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r>' +
  `<w:del w:id="1" w:author="Alice" w:date="${DATE_ALICE}">` +
  "<w:r><w:delText>ten percent</w:delText></w:r></w:del>" +
  `<w:ins w:id="2" w:author="Alice" w:date="${DATE_ALICE}">` +
  "<w:r><w:t>twelve percent</w:t></w:r></w:ins>" +
  '<w:r><w:t xml:space="preserve"> of net revenue.</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">Notices must be sent </w:t></w:r>' +
  `<w:ins w:id="3" w:author="Bob" w:date="${DATE_BOB}">` +
  '<w:r><w:t xml:space="preserve">by certified mail </w:t></w:r></w:ins>' +
  "<w:r><w:t>to the address below</w:t></w:r>" +
  `<w:del w:id="4" w:author="Bob" w:date="${DATE_BOB}">` +
  '<w:r><w:delText xml:space="preserve"> within five days</w:delText></w:r></w:del>' +
  "<w:r><w:t>.</w:t></w:r></w:p>";

const P2_ASIF = "The fee is twelve percent of net revenue.";
const P2_REJECTED = "The fee is ten percent of net revenue.";
const P3_ASIF = "Notices must be sent by certified mail to the address below.";
const P3_REJECTED = "Notices must be sent to the address below within five days.";

const P2A = `P2#${anchorHash(P2_ASIF)}`;
const P3A = `P3#${anchorHash(P3_ASIF)}`;

beforeEach(() => {
  process.env["DOCXENGINE_FIXED_DATE"] = FIXED_DATE;
});

afterEach(() => {
  delete process.env["DOCXENGINE_FIXED_DATE"];
  delete process.env["DOCXENGINE_AUTHOR"];
});

function openBody(body: string) {
  const session = new Session();
  const parts: DocxParts = { ...DEFAULT_PARTS, "word/document.xml": docWithBody(body) };
  const res = docxOpen(session, { bytes: Buffer.from(buildDocx(parts)).toString("base64") });
  return { session, docId: res.doc_id };
}

function openRedlines() {
  return openBody(REDLINES_BODY);
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
// op: list
// ---------------------------------------------------------------------------

describe("docx_revision list", () => {
  it("lists every revision in document order with R-ids, anchors, and raw text", () => {
    const { session, docId } = openRedlines();
    expect(docxRevision(session, { doc_id: docId, op: "list" })).toEqual({
      revisions: [
        {
          id: "R1",
          type: "del",
          author: "Alice",
          date: DATE_ALICE,
          anchor: P2A,
          text: "ten percent",
        },
        {
          id: "R2",
          type: "ins",
          author: "Alice",
          date: DATE_ALICE,
          anchor: P2A,
          text: "twelve percent",
        },
        {
          id: "R3",
          type: "ins",
          author: "Bob",
          date: DATE_BOB,
          anchor: P3A,
          text: "by certified mail ",
        },
        {
          id: "R4",
          type: "del",
          author: "Bob",
          date: DATE_BOB,
          anchor: P3A,
          text: " within five days",
        },
      ],
    });
  });

  it("filters by author, date prefix, and after/before bounds", () => {
    const { session, docId } = openRedlines();
    const ids = (filter: Record<string, string>) =>
      (docxRevision(session, { doc_id: docId, op: "list", filter }).revisions ?? []).map(
        (r) => r.id,
      );
    expect(ids({ author: "Bob" })).toEqual(["R3", "R4"]);
    expect(ids({ date: "2026-01-15" })).toEqual(["R1", "R2"]);
    expect(ids({ after: "2026-02-01" })).toEqual(["R3", "R4"]);
    expect(ids({ before: "2026-02-01" })).toEqual(["R1", "R2"]);
    expect(ids({ author: "Alice", date: "2026-02" })).toEqual([]);
  });

  it("raises doc_not_found for unknown documents and not_found for unknown ops", () => {
    const { session, docId } = openRedlines();
    expectCode(() => docxRevision(session, { doc_id: "d99", op: "list" }), "doc_not_found");
    expectCode(
      () => docxRevision(session, { doc_id: docId, op: "merge" as unknown as "list" }),
      "not_found",
    );
  });
});

// ---------------------------------------------------------------------------
// accept / reject
// ---------------------------------------------------------------------------

describe("docx_revision accept/reject", () => {
  it("accepts by author filter without touching other authors (§7)", () => {
    const { session, docId } = openRedlines();
    const res = docxRevision(session, {
      doc_id: docId,
      op: "accept",
      filter: { author: "Alice" },
    });
    expect(res).toEqual({
      accepted: 2,
      remaining_by_author: { Bob: 2 },
      anchors: [P2A], // the as-if-accepted hash is unchanged by accepting
    });
    const xml = documentXml(session, docId);
    expect(xml).not.toContain("ten percent");
    // Alice's paragraph resolved and re-merged into a single clean run.
    expect(xml).toContain(
      "<w:p><w:r><w:t>The fee is twelve percent of net revenue.</w:t></w:r></w:p>",
    );
    // Bob's wrappers are byte-for-byte untouched.
    expect(xml).toContain(`<w:ins w:id="3" w:author="Bob" w:date="${DATE_BOB}">`);
    expect(xml).toContain(`<w:del w:id="4" w:author="Bob" w:date="${DATE_BOB}">`);
  });

  it("accept_all resolves everything and is idempotent", () => {
    const { session, docId } = openRedlines();
    const res = docxRevision(session, { doc_id: docId, op: "accept_all" });
    expect(res).toEqual({
      accepted: 4,
      remaining_by_author: {},
      anchors: [P2A, P3A],
    });
    expect(docxRead(session, { doc_id: docId, range: "P2..P3" }).content.split("\n")).toEqual([
      `[${P2A}] ${P2_ASIF}`,
      `[${P3A}] ${P3_ASIF}`,
    ]);
    const xml = documentXml(session, docId);
    expect(xml).not.toContain("<w:ins");
    expect(xml).not.toContain("<w:del");
    expect(docxRevision(session, { doc_id: docId, op: "accept_all" })).toEqual({
      accepted: 0,
      remaining_by_author: {},
      anchors: [],
    });
  });

  it("accept_all ignores id and filter (§6a)", () => {
    const { session, docId } = openRedlines();
    const res = docxRevision(session, {
      doc_id: docId,
      op: "accept_all",
      filter: { author: "Alice" },
    });
    expect(res.accepted).toBe(4);
  });

  it("reject_all restores the original text (delText → w:t) and re-merges runs", () => {
    const { session, docId } = openRedlines();
    const res = docxRevision(session, { doc_id: docId, op: "reject_all" });
    expect(res).toEqual({
      rejected: 4,
      remaining_by_author: {},
      anchors: [`P2#${anchorHash(P2_REJECTED)}`, `P3#${anchorHash(P3_REJECTED)}`],
    });
    const xml = documentXml(session, docId);
    expect(xml).toContain(
      "<w:p><w:r><w:t>The fee is ten percent of net revenue.</w:t></w:r></w:p>",
    );
    expect(xml).toContain(
      "<w:p><w:r><w:t>Notices must be sent to the address below within five days.</w:t></w:r></w:p>",
    );
    expect(xml).not.toContain("w:delText");
    const content = docxRead(session, { doc_id: docId }).content;
    expect(content).not.toContain("twelve percent");
    expect(content).not.toContain("certified mail");
  });

  it("accepts a single revision by id, leaving the rest in place", () => {
    const { session, docId } = openRedlines();
    const res = docxRevision(session, { doc_id: docId, op: "accept", id: "R3" });
    expect(res).toEqual({
      accepted: 1,
      remaining_by_author: { Alice: 2, Bob: 1 },
      anchors: [P3A],
    });
    const xml = documentXml(session, docId);
    expect(xml).not.toContain('w:id="3"');
    expect(xml).toContain('w:id="4"');
  });

  it("an id selecting nothing resolves nothing (§7 idempotency, not an error)", () => {
    const { session, docId } = openRedlines();
    for (const id of ["R99", "X9"]) {
      expect(docxRevision(session, { doc_id: docId, op: "accept", id })).toEqual({
        accepted: 0,
        remaining_by_author: { Alice: 2, Bob: 2 },
        anchors: [],
      });
    }
  });

  it("reject of a single del restores its delText as text", () => {
    const { session, docId } = openRedlines();
    const res = docxRevision(session, { doc_id: docId, op: "reject", id: "R4" });
    expect(res.rejected).toBe(1);
    expect(docxRead(session, { doc_id: docId }).content).toContain(
      "to the address below within five days.",
    );
  });

  it("accepts by date filters: prefix, after (≥), before (<)", () => {
    const prefix = openRedlines();
    const byPrefix = docxRevision(prefix.session, {
      doc_id: prefix.docId,
      op: "accept",
      filter: { date: "2026-01-15" },
    });
    expect(byPrefix.accepted).toBe(2);
    expect(byPrefix.remaining_by_author).toEqual({ Bob: 2 });

    const after = openRedlines();
    const byAfter = docxRevision(after.session, {
      doc_id: after.docId,
      op: "accept",
      filter: { after: "2026-02-01" },
    });
    expect(byAfter.accepted).toBe(2);
    expect(byAfter.remaining_by_author).toEqual({ Alice: 2 });

    const before = openRedlines();
    const byBefore = docxRevision(before.session, {
      doc_id: before.docId,
      op: "accept",
      filter: { before: "2026-02-01" },
    });
    expect(byBefore.accepted).toBe(2);
    expect(byBefore.remaining_by_author).toEqual({ Bob: 2 });
  });

  it("removes empty paragraph-mark wrappers on accept", () => {
    const body =
      "<w:p><w:pPr><w:rPr>" +
      `<w:ins w:id="5" w:author="Alice" w:date="${DATE_ALICE}"/>` +
      "</w:rPr></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>";
    const { session, docId } = openBody(body);
    const listed = docxRevision(session, { doc_id: docId, op: "list" }).revisions ?? [];
    expect(listed).toEqual([
      {
        id: "R5",
        type: "ins",
        author: "Alice",
        date: DATE_ALICE,
        anchor: `P1#${anchorHash("Hi")}`,
        text: "",
      },
    ]);
    const res = docxRevision(session, { doc_id: docId, op: "accept_all" });
    expect(res.accepted).toBe(1);
    expect(documentXml(session, docId)).not.toContain("<w:ins");
  });
});

// ---------------------------------------------------------------------------
// §7 post-pass: run merging
// ---------------------------------------------------------------------------

describe("docx_revision run-merge post-pass", () => {
  it("does not merge runs whose rPr differ", () => {
    const body =
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>" +
      `<w:ins w:id="1" w:author="A" w:date="${DATE_ALICE}">` +
      '<w:r><w:t xml:space="preserve"> plain</w:t></w:r></w:ins></w:p>';
    const { session, docId } = openBody(body);
    docxRevision(session, { doc_id: docId, op: "accept_all" });
    expect(documentXml(session, docId)).toContain(
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>" +
        '<w:r><w:t xml:space="preserve"> plain</w:t></w:r></w:p>',
    );
  });

  it("is rsid-blind and keeps the first run's start tag verbatim", () => {
    const body =
      '<w:p><w:r w:rsidR="00AA0001"><w:t>x</w:t></w:r>' +
      `<w:ins w:id="1" w:author="Alice" w:date="${DATE_ALICE}">` +
      "<w:r><w:t>z</w:t></w:r></w:ins>" +
      '<w:r w:rsidR="00BB0002"><w:t>y</w:t></w:r></w:p>';
    const { session, docId } = openBody(body);
    const res = docxRevision(session, { doc_id: docId, op: "accept_all" });
    expect(res.accepted).toBe(1);
    expect(documentXml(session, docId)).toContain(
      '<w:p><w:r w:rsidR="00AA0001"><w:t>xzy</w:t></w:r></w:p>',
    );
  });

  it("merges identically-formatted runs (rPr compared, first rPr kept)", () => {
    const body =
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>One</w:t></w:r>" +
      `<w:ins w:id="1" w:author="A" w:date="${DATE_ALICE}">` +
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> two</w:t></w:r></w:ins></w:p>';
    const { session, docId } = openBody(body);
    docxRevision(session, { doc_id: docId, op: "accept_all" });
    expect(documentXml(session, docId)).toContain(
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>One two</w:t></w:r></w:p>",
    );
  });

  it("only touches the affected paragraphs", () => {
    // P1 has two mergeable runs but no revisions; resolving P2 must not touch P1.
    const body =
      "<w:p><w:r><w:t>split</w:t></w:r><w:r><w:t> runs</w:t></w:r></w:p>" +
      `<w:p><w:ins w:id="1" w:author="Alice" w:date="${DATE_ALICE}">` +
      "<w:r><w:t>added</w:t></w:r></w:ins></w:p>";
    const { session, docId } = openBody(body);
    docxRevision(session, { doc_id: docId, op: "accept_all" });
    const xml = documentXml(session, docId);
    expect(xml).toContain("<w:p><w:r><w:t>split</w:t></w:r><w:r><w:t> runs</w:t></w:r></w:p>");
    expect(xml).toContain("<w:p><w:r><w:t>added</w:t></w:r></w:p>");
  });
});

// ---------------------------------------------------------------------------
// Round trip with the tracked edit surface
// ---------------------------------------------------------------------------

describe("tracked edit → revision round trip", () => {
  const P2_TEXT = "The term is five (5) years from the Effective Date.";

  it("tracked replace then accept by author lands on the clean result", () => {
    const { session, docId } = openRedlines();
    docxReplace(session, {
      doc_id: docId,
      anchor: `P1#${anchorHash("Revision History")}`,
      old: "Revision History",
      new: "Change Log",
      track_changes: true,
      author: "Claude",
    });
    const listed = docxRevision(session, {
      doc_id: docId,
      op: "list",
      filter: { author: "Claude" },
    });
    expect((listed.revisions ?? []).map((r) => r.id)).toEqual(["R5", "R6"]);
    const res = docxRevision(session, {
      doc_id: docId,
      op: "accept",
      filter: { author: "Claude" },
    });
    expect(res).toEqual({
      accepted: 2,
      remaining_by_author: { Alice: 2, Bob: 2 },
      anchors: [`P1#${anchorHash("Change Log")}`],
    });
    const content = docxRead(session, { doc_id: docId }).content;
    expect(content.startsWith(`[P1#${anchorHash("Change Log")}] Change Log`)).toBe(true);
    expect(content).not.toContain("[del by Claude]");
  });

  it("a tracked replace nested in an existing w:ins resolves with its container", () => {
    // Claude's del/ins pair lands INSIDE Alice's w:ins R2; resolving R2
    // unwraps it, leaving Claude's pair pending (§6a nested-candidate rule).
    const { session, docId } = openRedlines();
    docxReplace(session, {
      doc_id: docId,
      anchor: P2A,
      old: "twelve percent",
      new: "fifteen percent",
      track_changes: true,
      author: "Claude",
    });
    const first = docxRevision(session, { doc_id: docId, op: "accept_all" });
    expect(first.accepted).toBe(4);
    expect(first.remaining_by_author).toEqual({ Claude: 2 });
    const second = docxRevision(session, { doc_id: docId, op: "accept_all" });
    expect(second.accepted).toBe(2);
    expect(second.remaining_by_author).toEqual({});
    const after = `P2#${anchorHash("The fee is fifteen percent of net revenue.")}`;
    expect(docxRead(session, { doc_id: docId, range: "P2..P2" }).content).toBe(
      `[${after}] The fee is fifteen percent of net revenue.`,
    );
  });

  it("reject_all restores the pre-edit text", () => {
    const { session, docId } = openBody(`<w:p><w:r><w:t>${P2_TEXT}</w:t></w:r></w:p>`);
    const anchor = `P1#${anchorHash(P2_TEXT)}`;
    docxReplace(session, {
      doc_id: docId,
      anchor,
      old: "five (5) years",
      new: "three (3) years",
      track_changes: true,
      author: "Claude",
    });
    const res = docxRevision(session, { doc_id: docId, op: "reject_all" });
    expect(res.rejected).toBe(2);
    expect(res.anchors).toEqual([anchor]);
    expect(docxRead(session, { doc_id: docId, range: "P1..P1" }).content).toBe(
      `[${anchor}] ${P2_TEXT}`,
    );
    expect(documentXml(session, docId)).toContain(`<w:p><w:r><w:t>${P2_TEXT}</w:t></w:r></w:p>`);
  });
});
