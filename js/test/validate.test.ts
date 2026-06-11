/**
 * Validator, repair, and save-gate tests (algorithms.md §8/§8a/§9).
 * Mirrors the Python stage-4 cases (python/tests/test_validate.py).
 */
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  type DocHandle,
  Session,
  ToolError,
  docxRepair,
  docxSave,
  docxValidate,
  repairDoc,
  validateDoc,
} from "../src/index.js";
import {
  DEFAULT_PARTS,
  DOCUMENT_RELS_XML,
  type DocxParts,
  buildDocx,
  docWithBody,
} from "./fixtures.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const ORPHAN_REL =
  '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/' +
  'relationships/image" Target="media/image1.png"/>';

// Tracked-changes paragraph (same as the Python conftest PARA_TRACKED):
// w:del id=1 + w:ins id=2 — the duplicate-id fixtures collide them on id=1.
const PARA_TRACKED =
  "<w:p>" +
  '<w:r><w:t xml:space="preserve">Payment due in </w:t></w:r>' +
  '<w:del w:id="1" w:author="J.Doe" w:date="2026-01-01T00:00:00Z">' +
  "<w:r><w:delText>30</w:delText></w:r></w:del>" +
  '<w:ins w:id="2" w:author="J.Doe" w:date="2026-01-01T00:00:00Z">' +
  "<w:r><w:t>45</w:t></w:r></w:ins>" +
  '<w:r><w:t xml:space="preserve"> days</w:t></w:r>' +
  "</w:p>";

const TRACKED_DOCUMENT_XML = docWithBody(PARA_TRACKED);

function withOrphanRel(parts: DocxParts): DocxParts {
  return {
    ...parts,
    "word/_rels/document.xml.rels": DOCUMENT_RELS_XML.replace(
      "</Relationships>",
      `${ORPHAN_REL}</Relationships>`,
    ),
  };
}

function withDuplicateRevIds(parts: DocxParts): DocxParts {
  return {
    ...parts,
    "word/document.xml": TRACKED_DOCUMENT_XML.replace('w:ins w:id="2"', 'w:ins w:id="1"'),
  };
}

/** The task fixture: an orphaned relationship plus duplicate revision ids. */
function corruptDocx(): Uint8Array {
  return buildDocx(withDuplicateRevIds(withOrphanRel(DEFAULT_PARTS)));
}

function openDoc(parts: DocxParts = DEFAULT_PARTS): DocHandle {
  return new Session().open(buildDocx(parts));
}

function warningsOf(doc: DocHandle): string[] {
  return validateDoc(doc)
    .filter((i) => i.severity === "warning")
    .map((i) => i.message);
}

function tmpDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-validate-"));
}

function captureToolError(fn: () => unknown): ToolError {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(ToolError);
    return e as ToolError;
  }
  throw new Error("expected a ToolError");
}

describe("validate", () => {
  it("clean fixture is valid with no issues", () => {
    expect(validateDoc(openDoc())).toEqual([]);
  });

  it("orphaned relationship is a check-c error", () => {
    const doc = openDoc(withOrphanRel(DEFAULT_PARTS));
    const issues = validateDoc(doc);
    const errors = issues.filter((i) => i.severity === "error");
    expect(errors).toHaveLength(1);
    expect(errors[0]?.part).toBe("word/_rels/document.xml.rels");
    expect(errors[0]?.message).toBe(
      "Relationship rId9 targets missing part word/media/image1.png.",
    );
    expect(errors[0]?.fix_hint).toContain("docx_repair");
    // The image rel is also never referenced: pinned §8a warning, never blocking.
    expect(warningsOf(doc)).toEqual(["Relationship rId9 (image) is never referenced."]);
  });

  it("dangling r:id is a check-b error", () => {
    const doc = openDoc({
      ...DEFAULT_PARTS,
      "word/document.xml": docWithBody(
        `<w:p><w:hyperlink xmlns:r="${R_NS}" r:id="rId8" w:history="1">` +
          "<w:r><w:t>full report</w:t></w:r></w:hyperlink></w:p>",
      ),
    });
    const issues = validateDoc(doc);
    expect(issues.map((i) => i.severity)).toEqual(["error"]);
    expect(issues[0]?.part).toBe("word/document.xml");
    expect(issues[0]?.message).toBe(
      "r:id rId8 is referenced in word/document.xml " +
        "but not defined in word/_rels/document.xml.rels.",
    );
    expect(issues[0]?.fix_hint).toContain("not auto-repairable");
  });

  it("duplicate revision ids are a check-d error", () => {
    const doc = openDoc(withDuplicateRevIds(DEFAULT_PARTS));
    const issues = validateDoc(doc);
    expect(issues.map((i) => i.severity)).toEqual(["error"]);
    expect(issues[0]?.part).toBe("word/document.xml");
    expect(issues[0]?.message).toBe("Duplicate revision id 1 on 2 w:ins/w:del elements.");
  });

  it("uncovered part is a check-a error", () => {
    const doc = openDoc({ ...DEFAULT_PARTS, "word/media/image1.png": "not really a png" });
    const issues = validateDoc(doc);
    expect(issues.map((i) => i.severity)).toEqual(["error"]);
    expect(issues[0]?.part).toBe("word/media/image1.png");
    expect(issues[0]?.message).toBe(
      "Part word/media/image1.png is not covered by [Content_Types].xml " +
        "(no Override, no Default for extension 'png').",
    );
  });

  it("comment reference without definition is an error", () => {
    const doc = openDoc({
      ...DEFAULT_PARTS,
      "word/document.xml": docWithBody(
        '<w:p><w:r><w:t>Noted.</w:t></w:r><w:r><w:commentReference w:id="3"/></w:r></w:p>',
      ),
    });
    const issues = validateDoc(doc);
    expect(issues.map((i) => i.severity)).toEqual(["error"]);
    expect(issues[0]?.part).toBe("word/comments.xml");
    expect(issues[0]?.message).toBe("Comment id=3 referenced in body but missing.");
  });

  it("unreferenced definitions warn and separators are exempt", () => {
    const doc = openDoc({
      ...DEFAULT_PARTS,
      "word/document.xml": docWithBody("<w:p><w:r><w:t>Plain.</w:t></w:r></w:p>"),
      "word/comments.xml":
        `<w:comments xmlns:w="${W_NS}">` +
        '<w:comment w:id="1" w:author="J.Doe"><w:p/></w:comment></w:comments>',
      "word/footnotes.xml":
        `<w:footnotes xmlns:w="${W_NS}">` +
        '<w:footnote w:type="separator" w:id="0"><w:p/></w:footnote>' +
        '<w:footnote w:type="continuationSeparator" w:id="1"><w:p/></w:footnote>' +
        "</w:footnotes>",
    });
    const issues = validateDoc(doc);
    expect(issues.map((i) => i.severity)).toEqual(["warning"]);
    expect(issues[0]?.message).toBe("Comment id=1 defined but never referenced.");
  });

  it("issue order is pinned a then c then d", () => {
    const parts = withDuplicateRevIds(withOrphanRel(DEFAULT_PARTS));
    parts["docProps/thumbnail.jpeg"] = "binary"; // uncovered, unrelated to the orphan rel
    const issues = validateDoc(openDoc(parts));
    const kinds = issues
      .filter((i) => i.severity === "error")
      .map((i) => i.message.split(" ", 1)[0]);
    expect(kinds).toEqual(["Part", "Relationship", "Duplicate"]);
  });

  it("docx_validate tool shape", () => {
    const session = new Session();
    const doc = session.open(corruptDocx());
    const result = docxValidate(session, { doc_id: doc.id });
    expect(result.valid).toBe(false);
    for (const issue of result.issues) {
      expect(Object.keys(issue).sort()).toEqual(["fix_hint", "message", "part", "severity"]);
    }
    expect(result.issues.filter((i) => i.severity === "error")).toHaveLength(2);
  });
});

describe("repair", () => {
  it("repairs orphaned rel and duplicate ids", () => {
    const doc = new Session().open(corruptDocx());
    const { fixed, remaining } = repairDoc(doc);
    expect(fixed).toEqual([
      "removed orphaned relationship rId9 (word/_rels/document.xml.rels)",
      "renumbered duplicate revision id 1 -> 2",
    ]);
    expect(remaining).toEqual([]);
    expect(validateDoc(doc)).toEqual([]); // §8a: validate must then be clean
  });

  it("renumber mirrors corpus semantics", () => {
    // corrupt-dup-ids corpus: del id=5 + ins id=5 + ins id=6 -> second 5 becomes 7.
    const doc = openDoc({
      ...DEFAULT_PARTS,
      "word/document.xml": docWithBody(
        "<w:p>" +
          '<w:del w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z">' +
          "<w:r><w:delText>March 1</w:delText></w:r></w:del>" +
          '<w:ins w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z">' +
          "<w:r><w:t>April 1</w:t></w:r></w:ins>" +
          '<w:ins w:id="6" w:author="A" w:date="2026-01-01T00:00:00Z">' +
          "<w:r><w:t>promptly </w:t></w:r></w:ins>" +
          "</w:p>",
      ),
    });
    const { fixed, remaining } = repairDoc(doc);
    expect(fixed).toEqual(["renumbered duplicate revision id 5 -> 7"]);
    expect(remaining).toEqual([]);
    const data = doc.pkg.partText("word/document.xml");
    expect(data).toContain('<w:del w:id="5"'); // first in document order keeps its id
    expect(data).toContain('<w:ins w:id="7"');
  });

  it("dangling r:id lands in remaining", () => {
    const parts = withOrphanRel(DEFAULT_PARTS);
    parts["word/document.xml"] = docWithBody(
      `<w:p><w:hyperlink xmlns:r="${R_NS}" r:id="rId8" w:history="1">` +
        "<w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>",
    );
    const { fixed, remaining } = repairDoc(openDoc(parts));
    expect(fixed).toEqual(["removed orphaned relationship rId9 (word/_rels/document.xml.rels)"]);
    expect(remaining).toEqual([
      "r:id rId8 is referenced in word/document.xml " +
        "but not defined in word/_rels/document.xml.rels.",
    ]);
  });

  it("adds missing content-type Default", () => {
    const doc = openDoc({ ...DEFAULT_PARTS, "word/media/image1.png": "binary" });
    const { fixed, remaining } = repairDoc(doc);
    expect(fixed).toEqual(["added content-type Default for extension 'png'"]);
    expect(remaining).toEqual([]);
    expect(doc.pkg.partText("[Content_Types].xml")).toContain(
      '<Default Extension="png" ContentType="image/png"/>',
    );
  });

  it("removes orphaned comment reference", () => {
    const doc = openDoc({
      ...DEFAULT_PARTS,
      "word/document.xml": docWithBody(
        "<w:p>" +
          '<w:commentRangeStart w:id="3"/>' +
          "<w:r><w:t>Noted.</w:t></w:r>" +
          '<w:commentRangeEnd w:id="3"/>' +
          '<w:r><w:commentReference w:id="3"/></w:r>' +
          "</w:p>",
      ),
    });
    const { fixed, remaining } = repairDoc(doc);
    expect(fixed).toEqual(["removed orphaned comment reference id=3"]);
    expect(remaining).toEqual([]);
    const data = doc.pkg.partText("word/document.xml");
    expect(data).not.toContain("commentReference");
    expect(data).not.toContain("commentRangeStart");
    expect(validateDoc(doc)).toEqual([]);
  });

  it("docx_repair tool marks dirty", () => {
    const session = new Session();
    const doc = session.open(corruptDocx());
    expect(doc.dirty).toBe(false);
    const result = docxRepair(session, { doc_id: doc.id });
    expect(result.remaining).toEqual([]);
    expect(doc.dirty).toBe(true);
  });
});

describe("save gate", () => {
  it("save refuses an invalid package", () => {
    const session = new Session();
    const doc = session.open(corruptDocx());
    const out = path.join(tmpDir(), "out.docx");
    const err = captureToolError(() => docxSave(session, { doc_id: doc.id, path: out }));
    expect(err.code).toBe("validation_failed");
    expect(err.suggestions[0]).toBe("Run docx_repair, then re-validate.");
    expect(fs.existsSync(out)).toBe(false);
  });

  it("save succeeds after repair", () => {
    const session = new Session();
    const doc = session.open(corruptDocx());
    docxRepair(session, { doc_id: doc.id });
    const out = path.join(tmpDir(), "out.docx");
    const result = docxSave(session, { doc_id: doc.id, path: out });
    expect(result.ok).toBe(true);
    expect(result.validated).toBe(true);
    expect(result.bytes).toBe(fs.statSync(out).size);
    expect(doc.dirty).toBe(false);
    expect(validateDoc(new Session().open(out))).toEqual([]);
  });

  it("warnings never block save", () => {
    // An unreferenced image relationship to an existing part: warning only.
    const doc_parts: DocxParts = {
      ...DEFAULT_PARTS,
      "word/_rels/document.xml.rels": DOCUMENT_RELS_XML.replace(
        "</Relationships>",
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/' +
          '2006/relationships/image" Target="styles.xml"/></Relationships>',
      ),
    };
    const session = new Session();
    const doc = session.open(buildDocx(doc_parts));
    expect(warningsOf(doc)).toEqual(["Relationship rId9 (image) is never referenced."]);
    const out = path.join(tmpDir(), "out.docx");
    const result = docxSave(session, { doc_id: doc.id, path: out });
    expect(result.ok).toBe(true);
  });

  it("saved output round-trips", () => {
    const session = new Session();
    const doc = session.open(buildDocx());
    const out = path.join(tmpDir(), "out.docx");
    docxSave(session, { doc_id: doc.id, path: out });
    const entries = unzipSync(fs.readFileSync(out));
    expect(Object.keys(entries)).toEqual(Object.keys(DEFAULT_PARTS));
  });
});
