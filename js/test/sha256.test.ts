/**
 * The hand-rolled SHA-256 (src/sha256.ts) must match `node:crypto` exactly —
 * the §1 anchor hash depends on it, and cross-language anchor parity with the
 * Python implementation depends on the anchor hash.
 */
import { createHash } from "node:crypto";

import { describe, expect, it } from "vitest";

import { anchorHash } from "../src/index.js";
import { sha256Hex } from "../src/sha256.js";

function reference(data: Uint8Array): string {
  return createHash("sha256").update(data).digest("hex");
}

const utf8 = new TextEncoder();

describe("sha256Hex", () => {
  it("matches the FIPS 180-4 known-answer vectors", () => {
    expect(sha256Hex(utf8.encode(""))).toBe(
      "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
    );
    expect(sha256Hex(utf8.encode("abc"))).toBe(
      "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
    );
    expect(sha256Hex(utf8.encode("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"))).toBe(
      "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1",
    );
  });

  it("matches node:crypto across message-schedule block boundaries", () => {
    // 55/56/64 straddle the one-vs-two-block padding edge; larger sizes
    // exercise multi-block runs.
    for (const len of [0, 1, 3, 31, 54, 55, 56, 57, 63, 64, 65, 119, 120, 127, 128, 1000, 4096]) {
      const data = new Uint8Array(len);
      for (let i = 0; i < len; i++) data[i] = (i * 31 + 7) & 0xff; // deterministic fill
      expect(sha256Hex(data), `len=${len}`).toBe(reference(data));
    }
  });

  it("matches node:crypto on multi-byte UTF-8 text", () => {
    for (const text of [
      "Master Services Agreement",
      "Café — naïve façade",
      "中文测试 / 日本語 / 한국어",
      "emoji \u{1F600}\u{1F680} and \u{10FFFF}",
      "mixed Ω≈ç√∫˜µ ASCII tail",
    ]) {
      const bytes = utf8.encode(text);
      expect(sha256Hex(bytes), text).toBe(reference(bytes));
    }
  });

  it("backs anchorHash per the §1 worked examples", () => {
    expect(anchorHash("")).toBe("e3b0");
    expect(anchorHash("Master Services Agreement")).toBe("515a");
  });
});
