/**
 * Lightweight XML scanner per spec/algorithms.md §3/§4.
 *
 * Parts are held as UTF-8 text; this scanner tokenizes just enough to locate
 * element boundaries as offsets in the decoded string. Edits are text splices —
 * there is no DOM build-then-serialize step, so attribute order, namespace
 * prefixes, inter-tag whitespace and rsid attributes in untouched regions
 * survive verbatim.
 */

// ---------------------------------------------------------------------------
// Whitespace (algorithms.md §1 step 3) — exactly the Unicode White_Space=Yes
// set. Do NOT use `\s` (it adds U+FEFF, which is non-conformant).
// ---------------------------------------------------------------------------

export const WS_CLASS =
  "\\t\\n\\u000B\\f\\r \\u0085\\u00A0\\u1680\\u2000-\\u200A\\u2028\\u2029\\u202F\\u205F\\u3000";

/** Matches one maximal run of §1 whitespace. */
export const WS_RUN_RE = new RegExp(`[${WS_CLASS}]+`, "g");

/** True iff `ch` (first code point) is in the §1 whitespace set. */
export function isWhitespaceChar(ch: string): boolean {
  const c = ch.codePointAt(0);
  if (c === undefined) return false;
  return (
    (c >= 0x09 && c <= 0x0d) ||
    c === 0x20 ||
    c === 0x85 ||
    c === 0xa0 ||
    c === 0x1680 ||
    (c >= 0x2000 && c <= 0x200a) ||
    c === 0x2028 ||
    c === 0x2029 ||
    c === 0x202f ||
    c === 0x205f ||
    c === 0x3000
  );
}

// ---------------------------------------------------------------------------
// Escaping / entities (algorithms.md §3 emission rules)
// ---------------------------------------------------------------------------

/** Text content escapes exactly `&` `<` `>`; non-ASCII is written literally. */
export function escapeText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Attribute values additionally escape `"`. */
export function escapeAttr(s: string): string {
  return escapeText(s).replace(/"/g, "&quot;");
}

const ENTITY_RE = /&(?:lt|gt|amp|quot|apos|#x[0-9a-fA-F]+|#[0-9]+);/g;

/** Decode the five XML entities plus numeric character references. */
export function decodeEntities(s: string): string {
  if (!s.includes("&")) return s;
  return s.replace(ENTITY_RE, (m) => {
    switch (m) {
      case "&lt;":
        return "<";
      case "&gt;":
        return ">";
      case "&amp;":
        return "&";
      case "&quot;":
        return '"';
      case "&apos;":
        return "'";
    }
    return m.startsWith("&#x")
      ? String.fromCodePoint(parseInt(m.slice(3, -1), 16))
      : String.fromCodePoint(parseInt(m.slice(2, -1), 10));
  });
}

// ---------------------------------------------------------------------------
// Tag scanning
// ---------------------------------------------------------------------------

export interface Tag {
  kind: "start" | "end" | "empty";
  /** Qualified name as written (e.g. `w:p`). */
  name: string;
  /** Offset of `<`. */
  start: number;
  /** One past `>`. */
  end: number;
  /** One past the name (attributes begin here for start/empty tags). */
  nameEnd: number;
}

function isNameDelim(c: number): boolean {
  // whitespace, '/', '>'
  return c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d || c === 0x2f || c === 0x3e;
}

/**
 * Find the next element tag in `xml[from:to)`, skipping text, comments,
 * processing instructions, CDATA sections and DOCTYPE declarations.
 */
export function nextTag(xml: string, from: number, to: number = xml.length): Tag | null {
  let i = from;
  for (;;) {
    const lt = xml.indexOf("<", i);
    if (lt < 0 || lt >= to) return null;
    if (xml.startsWith("<!--", lt)) {
      const e = xml.indexOf("-->", lt + 4);
      if (e < 0) return null;
      i = e + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", lt)) {
      const e = xml.indexOf("]]>", lt + 9);
      if (e < 0) return null;
      i = e + 3;
      continue;
    }
    if (xml.startsWith("<?", lt)) {
      const e = xml.indexOf("?>", lt + 2);
      if (e < 0) return null;
      i = e + 2;
      continue;
    }
    if (xml.startsWith("<!", lt)) {
      const e = xml.indexOf(">", lt + 2);
      if (e < 0) return null;
      i = e + 1;
      continue;
    }
    if (xml.startsWith("</", lt)) {
      const gt = xml.indexOf(">", lt + 2);
      if (gt < 0 || gt >= to) return null;
      return {
        kind: "end",
        name: xml.slice(lt + 2, gt).trim(),
        start: lt,
        end: gt + 1,
        nameEnd: gt,
      };
    }
    // Start or empty-element tag; scan attributes respecting quoted values.
    let j = lt + 1;
    while (j < to && !isNameDelim(xml.charCodeAt(j))) j++;
    const name = xml.slice(lt + 1, j);
    const nameEnd = j;
    while (j < to) {
      const c = xml.charCodeAt(j);
      if (c === 0x22 || c === 0x27) {
        const q = xml.indexOf(xml[j] as string, j + 1);
        if (q < 0) return null;
        j = q + 1;
        continue;
      }
      if (c === 0x3e) {
        const empty = xml.charCodeAt(j - 1) === 0x2f; // '/>'
        return { kind: empty ? "empty" : "start", name, start: lt, end: j + 1, nameEnd };
      }
      j++;
    }
    return null;
  }
}

const ATTR_RE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** Parse the attributes of a start/empty tag into a name → decoded-value map. */
export function attrs(xml: string, tag: Tag): Record<string, string> {
  const out: Record<string, string> = {};
  if (tag.kind === "end") return out;
  const segEnd = tag.kind === "empty" ? tag.end - 2 : tag.end - 1;
  const seg = xml.slice(tag.nameEnd, segEnd);
  ATTR_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = ATTR_RE.exec(seg)) !== null) {
    out[m[1] as string] = decodeEntities(m[2] ?? m[3] ?? "");
  }
  return out;
}

/** Read one attribute value (decoded) from a start/empty tag. */
export function getAttr(xml: string, tag: Tag, name: string): string | undefined {
  return attrs(xml, tag)[name];
}

// ---------------------------------------------------------------------------
// Element extents
// ---------------------------------------------------------------------------

export interface ElementSlice {
  name: string;
  /** Offset of the start tag's `<`. */
  start: number;
  /** One past the end tag's `>` (or the self-closing tag's `>`). */
  end: number;
  /** One past the start tag's `>`. */
  startTagEnd: number;
  /** Content range [contentStart, contentEnd); empty for self-closed. */
  contentStart: number;
  contentEnd: number;
  selfClosed: boolean;
  /** One past the name in the start tag (attribute region begins here). */
  nameEnd: number;
}

/** Resolve the full extent of the element opened by `tag`. */
export function elementExtent(xml: string, tag: Tag, to: number = xml.length): ElementSlice {
  if (tag.kind === "empty") {
    return {
      name: tag.name,
      start: tag.start,
      end: tag.end,
      startTagEnd: tag.end,
      contentStart: tag.end,
      contentEnd: tag.end,
      selfClosed: true,
      nameEnd: tag.nameEnd,
    };
  }
  if (tag.kind === "end") throw new Error(`elementExtent on end tag </${tag.name}>`);
  let depth = 0;
  let i = tag.end;
  for (;;) {
    const t = nextTag(xml, i, to);
    if (!t) throw new Error(`unclosed element <${tag.name}>`);
    if (t.name === tag.name) {
      if (t.kind === "start") depth++;
      else if (t.kind === "end") {
        if (depth === 0) {
          return {
            name: tag.name,
            start: tag.start,
            end: t.end,
            startTagEnd: tag.end,
            contentStart: tag.end,
            contentEnd: t.start,
            selfClosed: false,
            nameEnd: tag.nameEnd,
          };
        }
        depth--;
      }
    }
    i = t.end;
  }
}

/** First descendant element with the given qualified name in `[from, to)`. */
export function findElement(
  xml: string,
  name: string,
  from = 0,
  to: number = xml.length,
): ElementSlice | null {
  let i = from;
  for (;;) {
    const t = nextTag(xml, i, to);
    if (!t) return null;
    if (t.name === name && t.kind !== "end") return elementExtent(xml, t, to);
    i = t.end;
  }
}

/** Direct child elements of a content range, in document order. */
export function childElements(
  xml: string,
  contentStart: number,
  contentEnd: number,
): ElementSlice[] {
  const out: ElementSlice[] = [];
  let i = contentStart;
  for (;;) {
    const t = nextTag(xml, i, contentEnd);
    if (!t) return out;
    if (t.kind === "end") {
      i = t.end;
      continue;
    }
    const el = elementExtent(xml, t, contentEnd);
    out.push(el);
    i = el.end;
  }
}

// ---------------------------------------------------------------------------
// Body-level blocks (algorithms.md §1)
// ---------------------------------------------------------------------------

/** Locate `w:body` in word/document.xml text. */
export function findBody(xml: string): ElementSlice {
  const body = findElement(xml, "w:body");
  if (!body) throw new Error("document has no w:body element");
  return body;
}

/**
 * Body-level blocks: the direct children of `w:body` in document order
 * (`w:p`, `w:tbl`, the trailing `w:sectPr`, …).
 */
export function bodyBlocks(xml: string): ElementSlice[] {
  const body = findBody(xml);
  return childElements(xml, body.contentStart, body.contentEnd);
}

// ---------------------------------------------------------------------------
// Run / w:t iteration with offset map (algorithms.md §4 step 1)
// ---------------------------------------------------------------------------

export interface TextPiece {
  kind: "t" | "delText";
  /** Decoded character data of the `w:t` / `w:delText`. */
  text: string;
  /** Extent of the `w:t` / `w:delText` element itself. */
  el: ElementSlice;
  /** Extent of the enclosing `w:r` (−1 if none — malformed, tolerated). */
  runStart: number;
  runEnd: number;
  /** Start index in the concatenated searchable text; −1 for `delText`. */
  textOffset: number;
}

/**
 * Every `w:t`/`w:delText` descendant of a scope (typically one paragraph), in
 * document order, with the enclosing run extents and offsets into the
 * concatenated searchable text. `w:delText` is carried but contributes
 * nothing to the searchable string (`textOffset` −1) — the document is seen
 * as-if-accepted. `w:tab`/`w:br` contribute nothing in MVP.
 */
export function textPieces(
  xml: string,
  scope: { contentStart: number; contentEnd: number },
): TextPiece[] {
  const pieces: TextPiece[] = [];
  const runStack: { start: number; pieceIdx: number[] }[] = [];
  let i = scope.contentStart;
  let offset = 0;
  for (;;) {
    const t = nextTag(xml, i, scope.contentEnd);
    if (!t) break;
    if (t.name === "w:r" && t.kind !== "empty") {
      if (t.kind === "start") {
        runStack.push({ start: t.start, pieceIdx: [] });
      } else {
        const r = runStack.pop();
        if (r) {
          for (const idx of r.pieceIdx) {
            const p = pieces[idx] as TextPiece;
            p.runStart = r.start;
            p.runEnd = t.end;
          }
        }
      }
      i = t.end;
      continue;
    }
    if ((t.name === "w:t" || t.name === "w:delText") && t.kind !== "end") {
      const el = elementExtent(xml, t, scope.contentEnd);
      const text = el.selfClosed ? "" : decodeEntities(xml.slice(el.contentStart, el.contentEnd));
      const kind: TextPiece["kind"] = t.name === "w:t" ? "t" : "delText";
      const piece: TextPiece = {
        kind,
        text,
        el,
        runStart: -1,
        runEnd: -1,
        textOffset: kind === "t" ? offset : -1,
      };
      if (kind === "t") offset += text.length;
      const top = runStack[runStack.length - 1];
      if (top) top.pieceIdx.push(pieces.length);
      pieces.push(piece);
      i = el.end;
      continue;
    }
    i = t.end;
  }
  return pieces;
}

/**
 * Concatenated `w:t` character data of a scope, in document order,
 * `w:delText` excluded (algorithms.md §1 step 1). Raw — not normalized.
 */
export function scopeText(
  xml: string,
  scope: { contentStart: number; contentEnd: number },
): string {
  let s = "";
  for (const p of textPieces(xml, scope)) if (p.kind === "t") s += p.text;
  return s;
}

/** Searchable text plus the offset map, in one pass. */
export function textWithMap(
  xml: string,
  scope: { contentStart: number; contentEnd: number },
): { text: string; pieces: TextPiece[] } {
  const pieces = textPieces(xml, scope);
  let text = "";
  for (const p of pieces) if (p.kind === "t") text += p.text;
  return { text, pieces };
}

/** The `t`-piece covering searchable-text index `index`, or null. */
export function pieceAt(pieces: TextPiece[], index: number): TextPiece | null {
  for (const p of pieces) {
    if (p.kind !== "t") continue;
    if (index >= p.textOffset && index < p.textOffset + p.text.length) return p;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Splice helpers (algorithms.md §3)
// ---------------------------------------------------------------------------

/** Replace `s[start:end)` with `replacement`; every other byte untouched. */
export function splice(s: string, start: number, end: number, replacement: string): string {
  return s.slice(0, start) + replacement + s.slice(end);
}

export interface SpliceEdit {
  start: number;
  end: number;
  text: string;
}

/** Apply non-overlapping edits in one pass (any input order). */
export function spliceAll(s: string, edits: readonly SpliceEdit[]): string {
  const sorted = [...edits].sort((a, b) => a.start - b.start || a.end - b.end);
  let out = "";
  let pos = 0;
  for (const e of sorted) {
    if (e.start < pos) throw new Error("overlapping splice edits");
    out += s.slice(pos, e.start) + e.text;
    pos = e.end;
  }
  return out + s.slice(pos);
}

/** True iff the §3 rule requires `xml:space="preserve"` on an emitted text element. */
export function needsSpacePreserve(text: string): boolean {
  if (text.length === 0) return false;
  return isWhitespaceChar(text[0] as string) || isWhitespaceChar(text[text.length - 1] as string);
}

/**
 * Emit a `w:t`/`w:delText` element per §3: `xml:space="preserve"` iff the text
 * starts or ends with §1 whitespace; content escaped (`&` `<` `>` only).
 */
export function emitTextElement(name: "w:t" | "w:delText", text: string): string {
  const attr = needsSpacePreserve(text) ? ' xml:space="preserve"' : "";
  return `<${name}${attr}>${escapeText(text)}</${name}>`;
}
