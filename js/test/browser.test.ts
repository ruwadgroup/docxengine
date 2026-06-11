/**
 * Browser-safety guard (docs/sdks/javascript.md §Browser): the modules behind
 * the package entry point — open from bytes, read/search, the whole `call()`
 * graph — must not carry static `node:` imports; only the Node-only CLI may.
 * Filesystem access resolves lazily via src/nodeenv.ts.
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { Session, ToolError, docxOpen, docxRead, docxSearch } from "../src/index.js";
import { buildDocx } from "./fixtures.js";

const srcDir = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "src");

function tsFiles(dir: string): string[] {
  const out: string[] = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...tsFiles(full));
    else if (entry.name.endsWith(".ts")) out.push(full);
  }
  return out;
}

describe("browser-safe module graph", () => {
  it("keeps static `node:` imports out of every module except the CLI", () => {
    const offenders: string[] = [];
    for (const file of tsFiles(srcDir)) {
      if (path.basename(file) === "cli.ts") continue; // Node-only entry point
      const source = fs.readFileSync(file, "utf8");
      // Static import forms only; `typeof import("node:…")` types are erased.
      if (/from\s+["']node:/.test(source) || /^import\s+["']node:/m.test(source)) {
        offenders.push(path.relative(srcDir, file));
      }
    }
    expect(offenders).toEqual([]);
  });

  it("opens from raw bytes and reads/searches without touching the filesystem", () => {
    const session = new Session();
    const docId = session.open(buildDocx()).id;
    const read = docxRead(session, { doc_id: docId });
    expect(read.content).toContain("Master Services Agreement");
    const found = docxSearch(session, { doc_id: docId, query: "Agreement" });
    expect(found.n_matches).toBeGreaterThan(0);
  });

  it("rejects invalid base64 bytes with open_failed", () => {
    try {
      docxOpen(new Session(), { bytes: "this is !!! not base64" });
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(ToolError);
      expect((e as ToolError).code).toBe("open_failed");
    }
  });
});
