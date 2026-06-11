/**
 * Phase-2 stage-3: render adapter (algorithms.md §24). LibreOffice is not
 * installed on the test machine, so the soffice path is exercised with a stub
 * executable on a temp PATH (a tiny shell script that writes canned output);
 * the structural fallback is the locally-tested real path.
 */
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";

import {
  Session,
  detectSoffice,
  docxConvert,
  docxCreate,
  docxRenderPreview,
  structuralPreview,
} from "../src/index.js";

const cleanup: string[] = [];
const savedEnv: Record<string, string | undefined> = {};

function setEnv(key: string, value: string | undefined): void {
  if (!(key in savedEnv)) savedEnv[key] = process.env[key];
  if (value === undefined) delete process.env[key];
  else process.env[key] = value;
}

afterEach(() => {
  for (const [k, v] of Object.entries(savedEnv)) {
    if (v === undefined) delete process.env[k];
    else process.env[k] = v;
  }
  for (const key of Object.keys(savedEnv)) delete savedEnv[key];
  for (const dir of cleanup.splice(0)) {
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      /* ignore */
    }
  }
});

/** A stub `soffice` that writes `input.pdf` into the --outdir and prints a version. */
function stubSoffice(): { dir: string; bin: string } {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-stub-"));
  cleanup.push(dir);
  const bin = path.join(dir, "soffice");
  const script = [
    "#!/bin/sh",
    'if [ "$1" = "--version" ]; then',
    '  echo "LibreOffice 24.8.1.2 abc"',
    "  exit 0",
    "fi",
    "# Parse --outdir and the trailing input file.",
    "OUTDIR=.",
    "while [ $# -gt 0 ]; do",
    '  if [ "$1" = "--outdir" ]; then OUTDIR="$2"; shift; fi',
    "  LAST=$1",
    "  shift",
    "done",
    'BASE=$(basename "$LAST" .docx)',
    'printf "%%PDF-1.4 stub" > "$OUTDIR/$BASE.pdf"',
    "exit 0",
  ].join("\n");
  fs.writeFileSync(bin, script);
  fs.chmodSync(bin, 0o755);
  return { dir, bin };
}

function newDoc(): { session: Session; docId: string } {
  const session = new Session();
  const res = docxCreate(session, { content_md: "# Title\n\nBody text here.\n" });
  return { session, docId: res.doc_id };
}

describe("render adapter detection", () => {
  it("detects nothing when no soffice is on PATH or the defaults", () => {
    setEnv("DOCXENGINE_SOFFICE", undefined);
    setEnv("PATH", "/nonexistent-bin-dir");
    expect(detectSoffice()).toBeNull();
  });

  it("honors DOCXENGINE_SOFFICE pointing at the stub", () => {
    const { bin } = stubSoffice();
    setEnv("DOCXENGINE_SOFFICE", bin);
    expect(detectSoffice()).toBe(bin);
  });

  it("finds a stub soffice on PATH", () => {
    const { dir, bin } = stubSoffice();
    setEnv("DOCXENGINE_SOFFICE", undefined);
    setEnv("PATH", dir);
    expect(detectSoffice()).toBe(bin);
  });
});

describe("structural fallback (no renderer)", () => {
  it("docx_render_preview returns the structural projection without image links", () => {
    setEnv("DOCXENGINE_SOFFICE", undefined);
    setEnv("PATH", "/nonexistent-bin-dir");
    const { session, docId } = newDoc();
    const res = docxRenderPreview(session, { doc_id: docId });
    expect(res.renderer).toBe("structural");
    expect(res.structural).toContain("[P1#");
    expect(res.structural).toContain("Title");
    expect(res.pages.length).toBeGreaterThanOrEqual(1);
    // No renderer ran, so no fake image links — just the page numbers.
    expect(res.pages.every((p) => p.image === undefined)).toBe(true);
    expect(res.note).toContain("DOCXENGINE_SOFFICE");
  });

  it("structuralPreview estimates pages as ceil(chars/1800)", () => {
    setEnv("DOCXENGINE_SOFFICE", undefined);
    setEnv("PATH", "/nonexistent-bin-dir");
    const { session, docId } = newDoc();
    const sp = structuralPreview(session.get(docId));
    expect(sp.pages).toBeNull();
    expect(sp.renderer).toBe("structural");
    expect(sp.estimatedPages).toBe(Math.max(1, Math.ceil(sp.structural.length / 1800)));
  });

  it("docx_convert to pdf without a renderer is render_unavailable", () => {
    setEnv("DOCXENGINE_SOFFICE", undefined);
    setEnv("PATH", "/nonexistent-bin-dir");
    const { session, docId } = newDoc();
    let code: string | null = null;
    try {
      docxConvert(session, { doc_id: docId, to: "pdf", path: "/tmp/out.pdf" });
    } catch (e) {
      code = (e as { code?: string }).code ?? null;
    }
    expect(code).toBe("render_unavailable");
  });
});

describe("LibreOffice path (stub soffice)", () => {
  it("docx_convert to pdf runs the stub and writes the output", () => {
    const { bin } = stubSoffice();
    setEnv("DOCXENGINE_SOFFICE", bin);
    const outDir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-out-"));
    cleanup.push(outDir);
    const out = path.join(outDir, "result.pdf");
    const { session, docId } = newDoc();
    const res = docxConvert(session, { doc_id: docId, to: "pdf", path: out });
    expect(res.path).toBe(out);
    expect(res.renderer).toBe("libreoffice 24.8.1.2");
    expect(fs.existsSync(out)).toBe(true);
    expect(fs.readFileSync(out, "utf-8")).toContain("%PDF");
  });

  it("docx_render_preview names the libreoffice renderer when detected", () => {
    const { bin } = stubSoffice();
    setEnv("DOCXENGINE_SOFFICE", bin);
    const { session, docId } = newDoc();
    const res = docxRenderPreview(session, { doc_id: docId, pages: [1, 2] });
    expect(res.renderer).toBe("libreoffice 24.8.1.2");
    expect(res.pages.map((p) => p.page)).toEqual([1, 2]);
    expect(res.pages[0]?.image).toBe(`docx://${docId}/preview/page-1.png`);
  });
});
