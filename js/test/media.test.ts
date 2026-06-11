/**
 * Phase-2 stage-2: docx_media (algorithms.md §19). Mirrors the Python media
 * cases — EMU sizing + aspect scaling, PNG/JPEG header dimension parsing, the
 * inline drawing run, the part/rel/content-type registration, extract, replace.
 */
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import {
  Session,
  ToolError,
  cmToEmu,
  docxMedia,
  docxOpen,
  docxValidate,
  drawingRun,
  jpegDimensions,
  pngDimensions,
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

function u32(n: number): number[] {
  return [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255];
}

/** A minimal PNG (IHDR only) with the given width/height. */
function makePng(w: number, h: number): Uint8Array {
  const ihdr = [
    0,
    0,
    0,
    13,
    0x49,
    0x48,
    0x44,
    0x52,
    ...u32(w),
    ...u32(h),
    8,
    6,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
  ];
  return new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, ...ihdr]);
}

/** A minimal JPEG with a SOF0 marker carrying the given width/height. */
function makeJpeg(w: number, h: number): Uint8Array {
  return new Uint8Array([
    0xff,
    0xd8,
    0xff,
    0xc0,
    0,
    17,
    8,
    (h >> 8) & 255,
    h & 255,
    (w >> 8) & 255,
    w & 255,
    3,
    1,
    0x22,
    0,
    2,
    0x11,
    1,
    3,
    0x11,
    1,
    0xff,
    0xd9,
  ]);
}

const INTRO = "<w:p><w:r><w:t>Logo here</w:t></w:r></w:p><w:sectPr/>";

describe("pixel-dimension parsing", () => {
  it("reads PNG width/height from the IHDR", () => {
    expect(pngDimensions(makePng(640, 480))).toEqual({ width: 640, height: 480 });
    expect(pngDimensions(new Uint8Array([1, 2, 3]))).toBeNull();
  });

  it("reads JPEG width/height from the first SOFn", () => {
    expect(jpegDimensions(makeJpeg(800, 600))).toEqual({ width: 800, height: 600 });
    expect(jpegDimensions(new Uint8Array([1, 2, 3]))).toBeNull();
  });

  it("cmToEmu rounds cm × 360000", () => {
    expect(cmToEmu(4)).toBe(1440000);
    expect(cmToEmu(3)).toBe(1080000);
  });
});

describe("drawingRun", () => {
  it("emits the §19 inline drawing run", () => {
    const r = drawingRun("rId8", 1440000, 1080000, "image1");
    expect(r).toContain('<wp:extent cx="1440000" cy="1080000"/>');
    expect(r).toContain('<a:blip r:embed="rId8"/>');
    expect(r).toContain('<a:ext cx="1440000" cy="1080000"/>');
  });
});

describe("docx_media insert", () => {
  it("writes the part, rel, content-type, and an aspect-scaled drawing", () => {
    const dir = mkdtempSync(join(tmpdir(), "docxmedia-"));
    const imgPath = join(dir, "logo.png");
    writeFileSync(imgPath, makePng(4, 2)); // 2:1 aspect
    const { session, docId } = openBody(INTRO);
    const after = pAnchor(session, docId, 1);
    const res = docxMedia(session, {
      doc_id: docId,
      op: "insert",
      after,
      image: imgPath,
      width_cm: 4,
    });
    expect(res.media_id).toBe("M1");
    // 4cm = 1440000 EMU; height scaled by 2/4 → 720000.
    const xml = part(session, docId, "word/document.xml");
    expect(xml).toContain('<wp:extent cx="1440000" cy="720000"/>');
    expect(xml).toContain("<a:blip r:embed=");
    // The image part, content-type Default, and image rel exist.
    expect(session.get(docId).pkg.has("word/media/image1.png")).toBe(true);
    expect(part(session, docId, "[Content_Types].xml")).toContain(
      'Extension="png" ContentType="image/png"',
    );
    expect(part(session, docId, "word/_rels/document.xml.rels")).toContain("/relationships/image");
    expect(docxValidate(session, { doc_id: docId })).toEqual({ valid: true, issues: [] });
  });

  it("uses both dimensions verbatim when given", () => {
    const dir = mkdtempSync(join(tmpdir(), "docxmedia-"));
    const imgPath = join(dir, "logo.png");
    writeFileSync(imgPath, makePng(4, 2));
    const { session, docId } = openBody(INTRO);
    const after = pAnchor(session, docId, 1);
    docxMedia(session, {
      doc_id: docId,
      op: "insert",
      after,
      image: imgPath,
      width_cm: 4,
      height_cm: 3,
    });
    expect(part(session, docId, "word/document.xml")).toContain(
      '<wp:extent cx="1440000" cy="1080000"/>',
    );
  });

  it("allocates image{k} as max existing + 1", () => {
    const dir = mkdtempSync(join(tmpdir(), "docxmedia-"));
    const imgPath = join(dir, "logo.png");
    writeFileSync(imgPath, makePng(10, 10));
    const { session, docId } = openBody(INTRO);
    const after = pAnchor(session, docId, 1);
    docxMedia(session, { doc_id: docId, op: "insert", after, image: imgPath, width_cm: 2 });
    docxMedia(session, { doc_id: docId, op: "insert", after, image: imgPath, width_cm: 2 });
    expect(session.get(docId).pkg.has("word/media/image1.png")).toBe(true);
    expect(session.get(docId).pkg.has("word/media/image2.png")).toBe(true);
  });
});

describe("docx_media extract / replace", () => {
  it("extracts the M{id} part bytes to a path", () => {
    const dir = mkdtempSync(join(tmpdir(), "docxmedia-"));
    const imgPath = join(dir, "logo.png");
    const pngBytes = makePng(8, 8);
    writeFileSync(imgPath, pngBytes);
    const { session, docId } = openBody(INTRO);
    const after = pAnchor(session, docId, 1);
    docxMedia(session, { doc_id: docId, op: "insert", after, image: imgPath, width_cm: 2 });

    const outPath = join(dir, "out.png");
    const res = docxMedia(session, { doc_id: docId, op: "extract", media_id: "M1", path: outPath });
    expect(res.path).toBe(outPath);
    expect(new Uint8Array(readFileSync(outPath))).toEqual(pngBytes);
  });

  it("replaces the M{id} part bytes keeping the rel", () => {
    const dir = mkdtempSync(join(tmpdir(), "docxmedia-"));
    const imgPath = join(dir, "logo.png");
    writeFileSync(imgPath, makePng(8, 8));
    const { session, docId } = openBody(INTRO);
    const after = pAnchor(session, docId, 1);
    docxMedia(session, { doc_id: docId, op: "insert", after, image: imgPath, width_cm: 2 });
    const relsBefore = part(session, docId, "word/_rels/document.xml.rels");

    const newPath = join(dir, "logo2.png");
    writeFileSync(newPath, makePng(16, 16));
    docxMedia(session, { doc_id: docId, op: "replace", media_id: "M1", image: newPath });
    // The rel set is unchanged (same rId, same target).
    expect(part(session, docId, "word/_rels/document.xml.rels")).toBe(relsBefore);
    expect(new Uint8Array(session.get(docId).pkg.part("word/media/image1.png"))).toEqual(
      makePng(16, 16),
    );
  });

  it("an unknown media id is not_found", () => {
    const { session, docId } = openBody(INTRO);
    expect(() =>
      docxMedia(session, { doc_id: docId, op: "extract", media_id: "M9", path: "/tmp/x.png" }),
    ).toThrowError(ToolError);
  });
});
