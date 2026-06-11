/**
 * Phase-2 stage-3: docx_template_fill (algorithms.md §21). Mirrors the Python
 * template cases — split-run placeholders coalesced, loop sections over whole
 * paragraphs, single-row table loops, inverted/conditional sections, comments,
 * unfilled tracking, strict mode, and XML-only escaping.
 */
import { strToU8, zipSync, type Zippable } from "fflate";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";

import { Session, ToolError, docxRead, docxTemplateFill } from "../src/index.js";
import { DEFAULT_PARTS, type DocxParts, docWithBody } from "./fixtures.js";

const tmpFiles: string[] = [];

afterEach(() => {
  for (const f of tmpFiles.splice(0)) {
    try {
      fs.rmSync(f, { force: true });
    } catch {
      /* ignore */
    }
  }
});

/** Write a .docx with the given body to a temp path and return that path. */
function templateFile(body: string, extra: DocxParts = {}): string {
  const parts: DocxParts = { ...DEFAULT_PARTS, "word/document.xml": docWithBody(body), ...extra };
  const zippable: Zippable = {};
  for (const [name, xml] of Object.entries(parts)) zippable[name] = strToU8(xml);
  const bytes = zipSync(zippable, { level: 0 });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-tmpl-"));
  const file = path.join(dir, "template.docx");
  fs.writeFileSync(file, bytes);
  tmpFiles.push(file);
  return file;
}

function run(session: Session, file: string, data: Record<string, unknown>, strict = false) {
  return docxTemplateFill(session, { template: file, data, strict });
}

function projection(session: Session, docId: string): string {
  return docxRead(session, { doc_id: docId }).content;
}

const SECT = "<w:sectPr/>";

describe("docx_template_fill var substitution", () => {
  it("fills a simple placeholder", () => {
    const file = templateFile("<w:p><w:r><w:t>Hello {{name}}!</w:t></w:r></w:p>" + SECT);
    const session = new Session();
    const res = run(session, file, { name: "World" });
    expect(res.filled).toBe(1);
    expect(res.unfilled).toEqual([]);
    expect(projection(session, res.doc_id)).toContain("Hello World!");
  });

  it("coalesces a placeholder split across two runs (§21 worked example)", () => {
    const file = templateFile(
      "<w:p><w:r><w:t>Client: {{Cli</w:t></w:r><w:r><w:t>ent}}</w:t></w:r></w:p>" + SECT,
    );
    const session = new Session();
    const res = run(session, file, { Client: "GlobalTech & Co" });
    expect(res.filled).toBe(1);
    expect(projection(session, res.doc_id)).toContain("Client: GlobalTech & Co");
    // XML-escaping only: the part holds &amp; (not HTML-escaped).
    const xml = session.get(res.doc_id).documentXml();
    expect(xml).toContain("GlobalTech &amp; Co");
  });

  it("leaves missing vars verbatim and lists them in unfilled (dedup, order)", () => {
    const file = templateFile("<w:p><w:r><w:t>{{a}} {{b}} {{a}} {{c}}</w:t></w:r></w:p>" + SECT);
    const session = new Session();
    const res = run(session, file, { b: "B" });
    expect(res.filled).toBe(1);
    expect(res.unfilled).toEqual(["a", "c"]);
    expect(projection(session, res.doc_id)).toContain("{{a}} B {{a}} {{c}}");
  });

  it("stringifies numbers and booleans", () => {
    const file = templateFile("<w:p><w:r><w:t>n={{n}} b={{b}}</w:t></w:r></w:p>" + SECT);
    const session = new Session();
    const res = run(session, file, { n: 42, b: true });
    expect(projection(session, res.doc_id)).toContain("n=42 b=true");
  });

  it("drops {{!comment}} tags", () => {
    const file = templateFile("<w:p><w:r><w:t>A{{! ignore me }}B</w:t></w:r></w:p>" + SECT);
    const session = new Session();
    const res = run(session, file, {});
    expect(res.unfilled).toEqual([]);
    expect(projection(session, res.doc_id)).toContain("AB");
  });

  it("strict mode raises placeholder_unfilled on a missing var", () => {
    const file = templateFile("<w:p><w:r><w:t>{{missing}}</w:t></w:r></w:p>" + SECT);
    const session = new Session();
    let err: ToolError | null = null;
    try {
      run(session, file, {}, true);
    } catch (e) {
      err = e as ToolError;
    }
    expect(err?.code).toBe("placeholder_unfilled");
  });
});

describe("docx_template_fill sections", () => {
  it("expands a loop over whole paragraphs, cloning the inner paragraph", () => {
    const file = templateFile(
      "<w:p><w:r><w:t>Items:</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>{{#items}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>- {{text}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>{{/items}}</w:t></w:r></w:p>" +
        SECT,
    );
    const session = new Session();
    const res = run(session, file, {
      items: [{ text: "Alpha" }, { text: "Beta" }, { text: "Gamma" }],
    });
    expect(res.loops_expanded).toEqual({ items: 3 });
    expect(res.filled).toBe(3);
    const proj = projection(session, res.doc_id);
    expect(proj).toContain("- Alpha");
    expect(proj).toContain("- Beta");
    expect(proj).toContain("- Gamma");
    expect(proj).not.toContain("{{");
  });

  it("renders an empty loop to nothing", () => {
    const file = templateFile(
      "<w:p><w:r><w:t>{{#items}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>- {{text}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>{{/items}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
        SECT,
    );
    const session = new Session();
    const res = run(session, file, { items: [] });
    expect(res.loops_expanded).toEqual({ items: 0 });
    const proj = projection(session, res.doc_id);
    expect(proj).toContain("After");
    expect(proj).not.toContain("text");
  });

  it("renders an inverted section when the value is falsy/empty", () => {
    const file = templateFile(
      "<w:p><w:r><w:t>{{^items}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>No items.</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>{{/items}}</w:t></w:r></w:p>" +
        SECT,
    );
    const session = new Session();
    const res = run(session, file, { items: [] });
    // Inverted sections are not loops.
    expect(res.loops_expanded).toEqual({});
    expect(projection(session, res.doc_id)).toContain("No items.");
  });

  it("hides an inverted section when the value is truthy", () => {
    const file = templateFile(
      "<w:p><w:r><w:t>{{^items}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>No items.</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>{{/items}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>End</w:t></w:r></w:p>" +
        SECT,
    );
    const session = new Session();
    const res = run(session, file, { items: [{ text: "x" }] });
    expect(projection(session, res.doc_id)).not.toContain("No items.");
  });

  it("resolves {{.}} for scalar loop elements", () => {
    const file = templateFile(
      "<w:p><w:r><w:t>{{#tags}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>#{{.}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>{{/tags}}</w:t></w:r></w:p>" +
        SECT,
    );
    const session = new Session();
    const res = run(session, file, { tags: ["one", "two"] });
    const proj = projection(session, res.doc_id);
    expect(proj).toContain("#one");
    expect(proj).toContain("#two");
  });

  it("clones a single table row whose cells hold the section tags", () => {
    const body =
      "<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>" +
      "<w:tr><w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>" +
      "<w:tc><w:p><w:r><w:t>Role</w:t></w:r></w:p></w:tc></w:tr>" +
      "<w:tr><w:tc><w:p><w:r><w:t>{{#people}}{{name}}</w:t></w:r></w:p></w:tc>" +
      "<w:tc><w:p><w:r><w:t>{{role}}{{/people}}</w:t></w:r></w:p></w:tc></w:tr>" +
      "</w:tbl>" +
      SECT;
    const file = templateFile(body);
    const session = new Session();
    const res = run(session, file, {
      people: [
        { name: "Ada", role: "Eng" },
        { name: "Bob", role: "PM" },
      ],
    });
    expect(res.loops_expanded).toEqual({ people: 2 });
    const xml = session.get(res.doc_id).documentXml();
    expect(xml).toContain("Ada");
    expect(xml).toContain("Bob");
    expect(xml).toContain("Eng");
    expect(xml).toContain("PM");
    expect(xml).not.toContain("{{");
  });
});

describe("docx_template_fill errors", () => {
  it("rejects a non-mustache syntax", () => {
    const file = templateFile("<w:p><w:r><w:t>x</w:t></w:r></w:p>" + SECT);
    const session = new Session();
    let err: ToolError | null = null;
    try {
      docxTemplateFill(session, {
        template: file,
        data: {},
        syntax: "handlebars" as unknown as "mustache",
      });
    } catch (e) {
      err = e as ToolError;
    }
    expect(err?.code).toBe("template_syntax");
  });

  it("raises template_syntax on an unclosed section", () => {
    const file = templateFile(
      "<w:p><w:r><w:t>{{#open}}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>inner</w:t></w:r></w:p>" +
        SECT,
    );
    const session = new Session();
    let err: ToolError | null = null;
    try {
      run(session, file, { open: [{}] });
    } catch (e) {
      err = e as ToolError;
    }
    expect(err?.code).toBe("template_syntax");
  });
});
