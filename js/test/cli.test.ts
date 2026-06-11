/**
 * CLI contract tests (algorithms.md §11) — a real subprocess over a stdin
 * pipe, run against the built `js/dist/cli.js` (build before testing).
 * Mirrors the Python stage-4 cases (python/tests/test_cli.py).
 */
import { spawnSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { unzipSync } from "fflate";
import { beforeAll, describe, expect, it } from "vitest";

import { DEFAULT_PARTS, buildDocx, docWithBody } from "./fixtures.js";

const CLI_PATH = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "dist", "cli.js");

const PARA_TRACKED_DUP_IDS =
  "<w:p>" +
  '<w:r><w:t xml:space="preserve">Payment due in </w:t></w:r>' +
  '<w:del w:id="1" w:author="J.Doe" w:date="2026-01-01T00:00:00Z">' +
  "<w:r><w:delText>30</w:delText></w:r></w:del>" +
  '<w:ins w:id="1" w:author="J.Doe" w:date="2026-01-01T00:00:00Z">' +
  "<w:r><w:t>45</w:t></w:r></w:ins>" +
  '<w:r><w:t xml:space="preserve"> days</w:t></w:r>' +
  "</w:p>";

let tmp: string;
let docxPath: string;
let corruptPath: string;

beforeAll(() => {
  expect(fs.existsSync(CLI_PATH), `build first: ${CLI_PATH} is missing`).toBe(true);
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-cli-"));
  docxPath = path.join(tmp, "fixture.docx");
  fs.writeFileSync(docxPath, buildDocx());
  corruptPath = path.join(tmp, "corrupt.docx");
  fs.writeFileSync(
    corruptPath,
    buildDocx({ ...DEFAULT_PARTS, "word/document.xml": docWithBody(PARA_TRACKED_DUP_IDS) }),
  );
});

/** Pipe request lines into `node dist/cli.js`; parsed responses + return code. */
function runCli(lines: string[]): { responses: Record<string, unknown>[]; rc: number } {
  const proc = spawnSync(process.execPath, [CLI_PATH], {
    input: lines.map((line) => `${line}\n`).join(""),
    encoding: "utf8",
    timeout: 60_000,
  });
  const responses = proc.stdout
    .split("\n")
    .filter((line) => line.trim() !== "")
    .map((line) => JSON.parse(line) as Record<string, unknown>);
  return { responses, rc: proc.status ?? -1 };
}

function request(tool: string, args: Record<string, unknown> = {}): string {
  return JSON.stringify({ tool, args });
}

describe("round trip", () => {
  it("open, replace, save", () => {
    const out = path.join(tmp, "round-trip-out.docx");
    const { responses, rc } = runCli([
      request("docx_open", { path: docxPath }),
      request("docx_replace", { doc_id: "d1", old: "five (5) years", new: "three (3) years" }),
      request("docx_save", { doc_id: "d1", path: out }),
    ]);
    expect(rc).toBe(0);
    expect(responses).toHaveLength(3); // exactly one response line per request, in order
    const [opened, replaced, saved] = responses;
    expect(opened?.["doc_id"]).toBe("d1");
    expect(opened?.["n_paragraphs"]).toBe(3);
    expect(replaced?.["n_replaced"]).toBe(1);
    expect(typeof replaced?.["new_anchor"]).toBe("string");
    expect(saved).toEqual({ ok: true, validated: true, bytes: fs.statSync(out).size });
    const entries = unzipSync(fs.readFileSync(out));
    const documentXml = new TextDecoder().decode(entries["word/document.xml"]);
    expect(documentXml).toContain("three (3) years");
  });

  it("doc state persists for the process lifetime", () => {
    const { responses, rc } = runCli([
      request("docx_open", { path: docxPath }),
      request("docx_open", { path: docxPath }),
      request("docx_read", { doc_id: "d2" }),
    ]);
    expect(rc).toBe(0);
    expect(responses.slice(0, 2).map((r) => r["doc_id"])).toEqual(["d1", "d2"]);
    expect(String(responses[2]?.["content"])).toContain("Master Services Agreement");
  });

  it("save refusal round-trips as an error object", () => {
    const out = path.join(tmp, "refusal-out.docx");
    const { responses, rc } = runCli([
      request("docx_open", { path: corruptPath }),
      request("docx_save", { doc_id: "d1", path: out }),
      request("docx_repair", { doc_id: "d1" }),
      request("docx_save", { doc_id: "d1", path: out }),
    ]);
    expect(rc).toBe(0);
    expect(responses[1]?.["error"]).toBe("validation_failed");
    expect(responses[2]?.["fixed"]).toEqual(["renumbered duplicate revision id 1 -> 2"]);
    expect(responses[2]?.["remaining"]).toEqual([]);
    expect(responses[3]).toEqual({ ok: true, validated: true, bytes: fs.statSync(out).size });
  });
});

describe("protocol", () => {
  it("error payload shape", () => {
    const { responses, rc } = runCli([request("docx_read", { doc_id: "d404" })]);
    expect(rc).toBe(0);
    expect(responses).toEqual([
      {
        error: "doc_not_found",
        message: "Unknown or expired doc_id: d404.",
        suggestions: ["Call docx_open again."],
      },
    ]);
  });

  it("unknown (non-spec) tools decline with not_implemented", () => {
    // The full spec surface is implemented; only a tool that is not in
    // spec/tools/ declines with not_implemented.
    const { responses } = runCli([request("docx_frobnicate", { doc_id: "d1" })]);
    expect(responses).toEqual([
      {
        error: "not_implemented",
        message: "Tool docx_frobnicate is not defined in spec/tools/.",
        suggestions: ["See docs/tools/index.md for the tool catalog."],
      },
    ]);
  });

  it("a malformed JSON line does not kill the process", () => {
    const { responses, rc } = runCli([
      "{this is not json",
      request("docx_open", { path: docxPath }),
    ]);
    expect(rc).toBe(0);
    expect(responses[0]?.["error"]).toBe("invalid_args");
    expect(responses[1]?.["doc_id"]).toBe("d1");
  });

  it("non-object requests are invalid_args", () => {
    const { responses } = runCli(['["docx_open"]', '{"args": {}}']);
    expect(responses.map((r) => r["error"])).toEqual(["invalid_args", "invalid_args"]);
  });

  it("missing required args report invalid_args", () => {
    const { responses } = runCli([request("docx_replace", { doc_id: "d1" })]);
    expect(responses[0]?.["error"]).toBe("invalid_args");
    expect(String(responses[0]?.["message"])).toContain("old");
  });

  it("blank lines are skipped and EOF exits zero", () => {
    const { responses, rc } = runCli(["", "   ", request("docx_open", { path: docxPath })]);
    expect(rc).toBe(0);
    expect(responses).toHaveLength(1);
  });
});
