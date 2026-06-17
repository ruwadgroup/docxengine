/**
 * Resource and content-safety limits for parsing untrusted packages.
 *
 * DocxEngine opens documents from untrusted sources, so the OPC layer bounds a
 * hostile package's cost *before* it is paid (zip bombs) and refuses hostile XML
 * (DTD/entity declarations). Every cap is read from the environment on each open
 * so deployments can tune them; the defaults are generous for real documents and
 * tight against abuse. The Python engine enforces the identical checks (parity).
 *
 * See SECURITY.md, ROADMAP.md Phase 3, and spec/algorithms.md §27.
 */
import { ToolError } from "./errors.js";

// Defaults — generous for genuine documents, bounded against abuse.
const DEFAULT_MAX_PARTS = 10_000;
const DEFAULT_MAX_TOTAL_BYTES = 512 * 1024 * 1024; // 512 MiB uncompressed, whole package
const DEFAULT_MAX_PART_BYTES = 128 * 1024 * 1024; // 128 MiB uncompressed, single part
const DEFAULT_MAX_COMPRESSION_RATIO = 200; // uncompressed / compressed, per part
const DEFAULT_MAX_XML_DEPTH = 1_000; // element nesting depth

/**
 * Below this uncompressed size a part is never flagged by the ratio check:
 * small, highly compressible parts are not a decompression bomb.
 */
export const RATIO_FLOOR_BYTES = 64 * 1024;

const XML_SUFFIXES = [".xml", ".rels"];

function intEnv(name: string, def: number): number {
  const raw = typeof process !== "undefined" && process.env ? process.env[name] : undefined;
  if (raw === undefined || raw.trim() === "") return def;
  const v = Number(raw);
  if (!Number.isInteger(v) || v <= 0) return def;
  return v;
}

export const maxParts = (): number => intEnv("DOCXENGINE_MAX_PARTS", DEFAULT_MAX_PARTS);
export const maxTotalBytes = (): number =>
  intEnv("DOCXENGINE_MAX_TOTAL_BYTES", DEFAULT_MAX_TOTAL_BYTES);
export const maxPartBytes = (): number =>
  intEnv("DOCXENGINE_MAX_PART_BYTES", DEFAULT_MAX_PART_BYTES);
export const maxCompressionRatio = (): number =>
  intEnv("DOCXENGINE_MAX_COMPRESSION_RATIO", DEFAULT_MAX_COMPRESSION_RATIO);
export const maxXmlDepth = (): number => intEnv("DOCXENGINE_MAX_XML_DEPTH", DEFAULT_MAX_XML_DEPTH);

/** True for parts parsed as XML (where DTD/entity declarations are a threat). */
export function isXmlPart(name: string): boolean {
  const lowered = name.toLowerCase();
  return XML_SUFFIXES.some((s) => lowered.endsWith(s));
}

/**
 * Reject a DTD/entity declaration in an XML part (XXE / billion-laughs). XML
 * keywords are case-sensitive; a conformant DOCTYPE/ENTITY is uppercase, so a
 * substring scan is exact and cheap.
 */
export function forbidDoctype(name: string, text: string): void {
  if (text.includes("<!DOCTYPE") || text.includes("<!ENTITY")) {
    throw new ToolError(
      "malicious_content",
      `Refusing ${name}: contains a DTD/entity declaration (DOCTYPE/ENTITY).`,
      ["Conformant Word documents never declare a DTD; treat this file as untrusted."],
    );
  }
}
