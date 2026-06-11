/**
 * Phase-2 stage-2: docx_comment (algorithms.md §18). Mirrors the Python
 * comments cases — five-place add wiring, w15 reply/resolve threading,
 * delete (with replies), list, and initials.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  Session,
  ToolError,
  commentInitials,
  docxComment,
  docxOpen,
  docxValidate,
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

const TWO = "<w:p><w:r><w:t>Hello world</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>";

beforeEach(() => {
  process.env["DOCXENGINE_FIXED_DATE"] = "2026-06-11T00:00:00Z";
});
afterEach(() => {
  delete process.env["DOCXENGINE_FIXED_DATE"];
  delete process.env["DOCXENGINE_AUTHOR"];
});

describe("commentInitials", () => {
  it("uppercases the first letter of each whitespace word", () => {
    expect(commentInitials("Jane Q. Doe")).toBe("JQD");
    expect(commentInitials("jane")).toBe("J");
    expect(commentInitials("")).toBe("");
    expect(commentInitials("  spaced  out ")).toBe("SO");
  });
});

describe("docx_comment add", () => {
  it("wires all five places: range markers, reference run, comment part, style", () => {
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    const anchor = pAnchor(session, docId, 1);
    const res = docxComment(session, {
      doc_id: docId,
      op: "add",
      anchor,
      text: "Should this be mutual?",
      author: "Jane Q. Doe",
    });
    expect(res).toEqual({ comment_id: "C0", anchor });

    const docXml = part(session, docId, "word/document.xml");
    // (1) range start before runs, (2) range end + (3) reference run after.
    expect(docXml).toContain(
      '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Hello world</w:t></w:r>' +
        '<w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' +
        '<w:commentReference w:id="0"/></w:r></w:p>',
    );
    // (4) the comment part entry with initials + a w14:paraId.
    const comments = part(session, docId, "word/comments.xml");
    expect(comments).toContain(
      '<w:comment w:id="0" w:author="Jane Q. Doe" w:date="2026-06-11T00:00:00Z" w:initials="JQD">',
    );
    expect(comments).toContain("w14:paraId=");
    expect(comments).toContain("<w:t>Should this be mutual?</w:t>");
    // (5) ensured CommentReference style.
    expect(part(session, docId, "word/styles.xml")).toContain(
      '<w:style w:type="character" w:styleId="CommentReference">',
    );
    // Content-type + document rel registered.
    expect(part(session, docId, "[Content_Types].xml")).toContain('PartName="/word/comments.xml"');
    expect(part(session, docId, "word/_rels/document.xml.rels")).toContain(
      "/relationships/comments",
    );
    // The result document validates clean.
    expect(docxValidate(session, { doc_id: docId })).toEqual({ valid: true, issues: [] });
  });

  it("places the range start after a leading w:pPr", () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:sectPr/>';
    const { session, docId } = openBody(body);
    const anchor = pAnchor(session, docId, 1);
    docxComment(session, { doc_id: docId, op: "add", anchor, text: "x", author: "A" });
    const docXml = part(session, docId, "word/document.xml");
    expect(docXml).toContain(
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:commentRangeStart w:id="0"/>',
    );
  });

  it("allocates ids as max existing + 1", () => {
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    const a1 = pAnchor(session, docId, 1);
    const a2 = pAnchor(session, docId, 2);
    const first = docxComment(session, {
      doc_id: docId,
      op: "add",
      anchor: a1,
      text: "one",
      author: "A",
    });
    expect(first.comment_id).toBe("C0");
    const second = docxComment(session, {
      doc_id: docId,
      op: "add",
      anchor: a2,
      text: "two",
      author: "B",
    });
    expect(second.comment_id).toBe("C1");
  });

  it("uses DOCXENGINE_AUTHOR when author is omitted", () => {
    process.env["DOCXENGINE_AUTHOR"] = "EnvUser";
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    const anchor = pAnchor(session, docId, 1);
    docxComment(session, { doc_id: docId, op: "add", anchor, text: "x" });
    expect(part(session, docId, "word/comments.xml")).toContain('w:author="EnvUser"');
  });
});

describe("docx_comment reply / resolve", () => {
  it("appends a w15 commentEx with the thread root as parent, and resolves it", () => {
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    const anchor = pAnchor(session, docId, 1);
    const root = docxComment(session, {
      doc_id: docId,
      op: "add",
      anchor,
      text: "Q?",
      author: "Jane",
    });
    const reply = docxComment(session, {
      doc_id: docId,
      op: "reply",
      comment_id: root.comment_id,
      text: "Yes.",
      author: "J.Doe",
    });
    expect(reply.comment_id).toBe("C1");
    const ext = part(session, docId, "word/commentsExtended.xml");
    // The reply's commentEx references the root's paraId as parent, done=0.
    expect(ext).toContain("w15:paraIdParent=");
    expect(ext).toContain('w15:done="0"');
    // commentsExtended part registered.
    expect(part(session, docId, "[Content_Types].xml")).toContain("commentsExtended");

    docxComment(session, { doc_id: docId, op: "resolve", comment_id: root.comment_id });
    const ext2 = part(session, docId, "word/commentsExtended.xml");
    expect(ext2).toContain('w15:done="1"');
  });

  it("reply on an unknown id is not_found", () => {
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    expect(() =>
      docxComment(session, { doc_id: docId, op: "reply", comment_id: "C9", text: "x" }),
    ).toThrowError(ToolError);
  });
});

describe("docx_comment list", () => {
  it("returns one entry per thread root with anchor, replies, resolved", () => {
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    const anchor = pAnchor(session, docId, 1);
    const root = docxComment(session, {
      doc_id: docId,
      op: "add",
      anchor,
      text: "Q?",
      author: "Jane",
    });
    docxComment(session, {
      doc_id: docId,
      op: "reply",
      comment_id: root.comment_id,
      text: "A.",
      author: "Bob",
    });
    docxComment(session, { doc_id: docId, op: "resolve", comment_id: root.comment_id });

    const list = docxComment(session, { doc_id: docId, op: "list" });
    expect(list.comments).toHaveLength(1);
    const entry = list.comments![0]!;
    expect(entry.id).toBe("C0");
    expect(entry.anchor).toBe(anchor);
    expect(entry.author).toBe("Jane");
    expect(entry.text).toBe("Q?");
    expect(entry.resolved).toBe(true);
    expect(entry.replies).toEqual([{ author: "Bob", date: "2026-06-11T00:00:00Z", text: "A." }]);
  });

  it("empty doc lists no comments", () => {
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    expect(docxComment(session, { doc_id: docId, op: "list" })).toEqual({ comments: [] });
  });
});

describe("docx_comment delete", () => {
  it("removes all five places for the id and its replies", () => {
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    const anchor = pAnchor(session, docId, 1);
    const root = docxComment(session, {
      doc_id: docId,
      op: "add",
      anchor,
      text: "Q?",
      author: "Jane",
    });
    docxComment(session, {
      doc_id: docId,
      op: "reply",
      comment_id: root.comment_id,
      text: "A.",
      author: "Bob",
    });

    docxComment(session, { doc_id: docId, op: "delete", comment_id: root.comment_id });
    const docXml = part(session, docId, "word/document.xml");
    expect(docXml).not.toContain("commentRangeStart");
    expect(docXml).not.toContain("commentReference");
    const comments = part(session, docId, "word/comments.xml");
    expect(comments).not.toContain("<w:comment ");
    // The list is now empty.
    expect(docxComment(session, { doc_id: docId, op: "list" })).toEqual({ comments: [] });
  });

  it("delete on an unknown id is not_found", () => {
    const { session, docId } = openBody(TWO + "<w:sectPr/>");
    docxComment(session, {
      doc_id: docId,
      op: "add",
      anchor: pAnchor(session, docId, 1),
      text: "x",
      author: "A",
    });
    expect(() =>
      docxComment(session, { doc_id: docId, op: "delete", comment_id: "C9" }),
    ).toThrowError(ToolError);
  });
});
