# Algorithms

The language-agnostic algorithm specification. The Python (`python/`) and TypeScript (`js/`) implementations both follow this document exactly; the conformance harness ([conformance/README.md](../conformance/README.md)) asserts their outputs are **byte-equivalent after normalization** (§10). Where this spec pins a choice, the choice is normative — "reasonable alternative" implementations are conformance failures.

Behavioral background lives in [docs/core/](../docs/core/); this file pins the bytes.

## 1. Anchors

Grammar: `P{ordinal}#{hash}` for paragraphs, `T{ordinal}` for tables.

- **ordinal** — 1-based position among body-level `w:p` elements: the `w:p` direct children of `w:body`, in document order. Paragraphs nested inside tables (or any other container) do not get body ordinals in MVP. Tables anchor as `T{ordinal}` over body-level `w:tbl` elements (table internals are Phase 2). The trailing `w:sectPr` is not a paragraph.
- **hash** — the first **4 lowercase hex chars** of the SHA-256 of the UTF-8 encoding of the paragraph's _normalized text_.

**Normalized text** (also used by projection §2 and table cells):

1. Concatenate the character data of every `w:t` descendant of the paragraph, in document order. `w:delText` is **excluded** (the hash sees the document as-if-accepted). `w:tab`, `w:br`, and all other elements contribute nothing.
2. Apply Unicode NFC (`unicodedata.normalize("NFC", s)` / `s.normalize("NFC")`).
3. Collapse every maximal run of whitespace to one ASCII space (U+0020). "Whitespace" is exactly the Unicode `White_Space=Yes` set: U+0009–U+000D, U+0020, U+0085, U+00A0, U+1680, U+2000–U+200A, U+2028, U+2029, U+202F, U+205F, U+3000. Do **not** use language defaults (JS `\s` adds U+FEFF; that is non-conformant).
4. Strip leading and trailing spaces.

An empty paragraph normalizes to `""` and hashes to `e3b0` (SHA-256 of the empty byte string begins `e3b0c442…`).

**Validation before every edit**: recompute the hash at the given ordinal. Mismatch → `anchor_stale`; ordinal out of range → `anchor_not_found`; unparseable anchor string → `anchor_invalid`. Batches validate all anchors upfront, atomically (`batch_aborted`).

### Worked example

```xml
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:rPr><w:b/></w:rPr><w:t>Master</w:t></w:r>
  <w:r><w:t xml:space="preserve"> Services</w:t></w:r>
  <w:proofErr w:type="spellStart"/>
  <w:r><w:t xml:space="preserve">  Agreement </w:t></w:r>
</w:p>
```

Concatenated `w:t` text: `"Master Services  Agreement "` → NFC (unchanged) → collapse: `"Master Services Agreement "` → strip: `"Master Services Agreement"`. SHA-256 = `515a77ed…` → hash `515a`. As the 1st body paragraph, the anchor is **`P1#515a`**.

## 2. Projection line format

One line per body-level block, joined with `\n` (no trailing newline):

```
[{anchor}{annotations}] {text}
```

- `annotations` is `""` or `" "` + space-separated tokens in this fixed order:
  1. heading: `H1`…`H9` if the paragraph's effective style is `Heading1`…`Heading9` (the `pStyle` styleId itself, or reached via the style's `basedOn` chain);
  2. list: `List:ol L{n}` or `List:ul L{n}` where `n` = `w:ilvl` + 1 and `ul` iff the resolved numbering level's `numFmt` is `bullet`, else `ol`.
- Exactly **one space** between `]` and the text. If the normalized text is empty, the line is the bracket alone with no trailing space.
- Markers are appended **inside the text**, at the end of the span they describe: `[ins by {author}]`, `[del by {author}]`, `[comment:C{id} by {author}]`. Marker insertion happens on the concatenated raw text (§1 step 1) **before** normalization: walking the paragraph in document order, at each `w:ins`/`w:del` wrapper end and at each `w:commentReference` insert the marker with one space on each side (§1 steps 2–4 then absorb doubled/edge spaces). Comment ids render as `C{w:id}`; authors come from `word/comments.xml` (`unknown` when unresolvable). Text is shown as-if-accepted; `w:delText` content is not shown in the default projection, so a `[del by …]` marker marks the deletion point.

```
[P1#515a H1] Master Services Agreement
[P4#d4e5] "Confidential Information" means all information disclosed… [comment:C1 by J.Doe]
[P12#e7f8 List:ol L1] First obligation [ins by Jane]
```

**Tables** project as a header line followed by a GitHub-style markdown table of normalized cell text:

```
[T1 3×2 @after:P5#9a01]
| Term | Value |
| --- | --- |
| Fee | $100 |
```

- `{rows}×{cols}` uses U+00D7; rows = count of `w:tr`; cols = count of `w:gridCol` in `w:tblGrid` (fallback: max `w:tc` count per row).
- `@after:{prev_anchor}` names the nearest preceding body-level paragraph's full anchor; a table with no preceding paragraph uses `@start`.
- Row 1 is rendered as the markdown header; the separator row is `| --- |`-style with one `---` per column. Cell text = the cell's paragraphs' normalized texts joined with a single space; `|` in cell text is escaped as `\|`.

### 2a. Read-surface pins (`docx_open` / `docx_outline` / `docx_read` / `docx_search`)

- **doc ids** — `d1`, `d2`, … in open order, for the process lifetime (§11).
- **List annotation** resolves the paragraph's direct `w:numPr` only (style-supplied numbering is Phase 2). `w:numId` absent or `0` → no annotation; a missing `word/numbering.xml` or unresolvable chain → `ol`; missing `w:ilvl` → level 0.
- **`docx_open`** — `n_paragraphs` = body-level `w:p` count; `has_tracked_changes` = any `w:ins`/`w:del` element in the document part; `has_comments` = any `w:commentReference`. `summary` = `"{title} — {p} paragraphs, {s} sections, {t} tables"` (em dash U+2014; nouns singular when the count is exactly 1): title = the first body paragraph with an effective heading level and non-empty normalized text, else the first body paragraph with non-empty normalized text, else `Untitled`; sections = count of `w:sectPr` elements in the document part; tables = body-level `w:tbl` count.
- **`docx_outline`** — `outline` lists body paragraphs whose effective style resolves to `Heading1`…`Heading9` (`level` 1–9, `text` = normalized text, markers excluded); `tables` lists body tables with `dims` = `{rows}×{cols}` (§2) and `after` = the nearest preceding body paragraph's full anchor, key omitted when no paragraph precedes.
- **`docx_read`** — `anchor` resolves by ordinal only (the hash is **not** validated: a read is how a stale anchor is refreshed). `window` counts body-level blocks (paragraphs and tables) on each side of the anchor block. `range` grammar: `P{a}..P{b}`, each endpoint optionally suffixed `#hhhh` (ignored); `a > b` → `anchor_invalid`; endpoints are paragraph ordinals and the selection is every body block between them inclusive; an endpoint ordinal that does not exist → `anchor_not_found`. Neither anchor nor range → the whole story. `anchor` wins when both are given. Non-body scopes project every `w:p` of the story part(s) — `headers`/`footers` concatenate `word/header{n}.xml` / `word/footer{n}.xml` ascending by `n` — with `P{ordinal}#{hash}` anchors computed over the story's own paragraph sequence (valid only within that scope); a missing story part reads as empty content.
- **Pagination** — blocks accumulate into `content` (lines joined with `\n`); when appending the next **paragraph** (never a table — tables ride with what precedes them) would push `content` past **24 000 characters** and at least one block is already emitted, the response stops there and `continuation` = `P{next}..P{last}` (ordinals of the first unemitted paragraph and the last paragraph of the selection), which the caller passes back as `range`.
- **`docx_search`** — matching runs over the §4 concatenated string of each paragraph in scope (body-level paragraphs; table-cell text is Phase 2). Literal matching is case-sensitive, non-overlapping, left-to-right; `regex: true` compiles with the host language's engine (documented MVP divergence — the conformance corpus avoids regex cases) and zero-length matches are skipped. One result entry per occurrence in document order. `snippet` = §1 normalization of `raw[max(0, start−40) .. min(len, end+40))` with `…` (U+2026) prepended/appended on each side that was truncated. `context` = normalized text of the nearest paragraph **at or before** the match whose effective style is a heading, key omitted when there is none (the heading scan ignores the scope filter). Zero matches → `{"matches": [], "n_matches": 0}`, not an error; an empty query or invalid regex → `not_found`.

## 3. XML editing strategy: splice, don't re-serialize

Every package part is held as **raw bytes**. Only parts actually modified by an operation are ever re-encoded; untouched parts pass through byte-for-byte (§9).

Modifications use **text splicing**: a lightweight scanner tokenizes the part's XML (UTF-8 text) just enough to locate element boundaries — start tag, end tag, attributes, character data — as byte offsets in the original buffer. An edit replaces exactly the spliced byte range(s) and leaves every other byte of the part untouched where possible. There is no DOM build-then-serialize step for editing; attribute order, namespace prefixes, whitespace between sibling elements, and rsid attributes in untouched regions survive verbatim.

Emission rules for new XML produced by a splice:

- Text content escapes exactly `&` → `&amp;`, `<` → `&lt;`, `>` → `&gt;`; attribute values additionally `"` → `&quot;`. Non-ASCII characters are written literally (parts are UTF-8).
- An emitted `w:t`/`w:delText` carries `xml:space="preserve"` iff its text starts or ends with a whitespace character (§1 set); otherwise the attribute is omitted.

## 4. Run coalescing for search/replace

To match text that Word has fragmented across runs:

1. Build the paragraph's concatenated `w:t` text (document order, `w:delText` excluded) together with an **offset map**: each text index → `(run, w:t element, char offset within that w:t)`. `w:tab`/`w:br` contribute nothing to the searchable string in MVP (a match may therefore span one; documented limitation).
2. Match `old` (literal) against the concatenated string.
3. Apply the replacement to the **first overlapping `w:t`**: its new content is its original prefix before the match start, plus the full replacement text. Every subsequent overlapping `w:t` keeps only its suffix after the match end (usually empty).
4. Runs whose `w:t` is left empty are removed entirely (the whole `w:r`).
5. The formatting (`rPr`) of the first matched run wins for the replacement text. `rsid*` attributes never block coalescing or matching — they are formatting noise.

### Worked example

```xml
<w:p><w:r><w:t xml:space="preserve">The term is five (5) </w:t></w:r
><w:r><w:rPr><w:b/></w:rPr><w:t>years from the Effective Date.</w:t></w:r></w:p>
```

Concatenated: `"The term is five (5) years from the Effective Date."` (run 1 covers indices 0–20, run 2 covers 21–51). Replace `old="five (5) years"` → `new="three (3) years"`: match spans [12, 26).

- Run 1's `w:t` (first overlap) → prefix `"The term is "` + replacement → `The term is three (3) years`.
- Run 2's `w:t` → suffix from index 26 → `" from the Effective Date."` (gains `xml:space="preserve"`; keeps its `<w:b/>`).

Only those two `w:t` contents (plus the one added attribute) change in the part's bytes.

## 5. Tracked-changes emission

With `track_changes: true`, the same match instead splits at the boundaries:

- **Deletion**: the matched portion of each overlapping run becomes a run inside one `<w:del>`, with that run's original `rPr` and `w:t` → `w:delText`.
- **Insertion**: the replacement text becomes one run inside `<w:ins>`, formatted with the first matched run's `rPr` (§4 rule 5).
- Wrapper attributes, in exactly this order: `w:id`, `w:author`, `w:date`.
- **id allocation**: scan the part for all existing `w:ins` and `w:del` `w:id` values (counting both); next id = max + 1; each new wrapper in the operation takes the next id in emission (document) order. No existing revisions → start at 1.
- **date**: the value of env `DOCXENGINE_FIXED_DATE` if set (verbatim); else current UTC, ISO-8601 with seconds, `Z` suffix (`2026-06-10T14:03:07Z`). Conformance runs always set `DOCXENGINE_FIXED_DATE`.
- **author**: the `author` argument; default env `DOCXENGINE_AUTHOR`, else `"DocxEngine"`.

### Worked example

The §4 replace with `track_changes: true, author: "Claude"`, `DOCXENGINE_FIXED_DATE=2026-06-10T00:00:00Z`, max existing revision id 6:

```xml
<w:p><w:r><w:t xml:space="preserve">The term is </w:t></w:r
><w:del w:id="7" w:author="Claude" w:date="2026-06-10T00:00:00Z"
><w:r><w:delText xml:space="preserve">five (5) </w:delText></w:r
><w:r><w:rPr><w:b/></w:rPr><w:delText>years</w:delText></w:r></w:del
><w:ins w:id="8" w:author="Claude" w:date="2026-06-10T00:00:00Z"
><w:r><w:t>three (3) years</w:t></w:r></w:ins
><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> from the Effective Date.</w:t></w:r></w:p>
```

Note the per-run formatting preserved inside `w:del` (the bold `years`), and the insertion taking the first matched run's (empty) `rPr`.

## 6. Word-level diff (`docx_edit_paragraph`)

A full-paragraph rewrite under tracking is auto-diffed:

1. **Tokenize** old and new text into alternating word/whitespace tokens (maximal runs of §1 whitespace are whitespace tokens). Form **units** of word + its following whitespace; leading whitespace attaches to the first unit.
2. **LCS** over the unit sequences, comparing full unit strings. Deterministic backtrack: with `L[i][j]` = LCS length of `old[i:]` vs `new[j:]`, walk forward from `(0,0)`: if `old[i] == new[j]` and `L[i][j] == L[i+1][j+1] + 1` → match (first-found, i.e. tiebreak toward earlier old tokens); else if `L[i+1][j] >= L[i][j+1]` → delete `old[i]`; else → insert `new[j]`.
3. **Emit** maximal runs of consecutive deletes/inserts as one `w:del`/`w:ins` span each; at the same position the `w:del` precedes the `w:ins`. Ids/author/date per §5.

`old: "three year term"` → `new: "five year initial term"` yields del(`three `)+ins(`five `), keep(`year `), ins(`initial `), keep(`term`) — never delete-all/insert-all.

### 6a. Edit-surface pins (MVP)

Behavior the tool schemas leave open, pinned for cross-implementation parity:

- **§6 insert-only spans** — a pure insertion's `w:ins` run takes the `rPr` of the run containing the insertion offset (the run being split, or the run starting exactly there); an end-of-paragraph insertion takes the last run's `rPr`; an empty paragraph yields no `rPr`. Delete spans keep each overlapped run's own `rPr` (§5).
- **§6 keep/replace spans** — a kept span re-emits each overlapped run's kept portion as one run with that run's own `rPr` (consecutive `w:t` pieces of the same run concatenate). A `w:ins` that directly follows a `w:del` at the same position (a replacement) takes the **first deleted run's** `rPr` (§5), not the insertion-offset rule. Wrapper ids are allocated in emission order (§5).
- **`docx_edit_paragraph` untracked** — the paragraph's content children are replaced by [`w:pPr` verbatim, if present] + one run carrying the first existing run's `rPr` verbatim and the new text (no run when the new text is empty).
- **`docx_edit_paragraph` result** — `new_anchor` = the paragraph's fresh anchor; `diff` = `"~{n} words changed"` where `n` = max(#deleted units, #inserted units) and the noun is `word` when `n` is exactly 1.
- **`docx_insert` minimal markdown** — `content` splits on `\n` (one trailing `\r` stripped per line); lines that are empty/whitespace-only emit nothing. `#{1..9}` + space → paragraph with `pStyle` `Heading{n}`; leading `- ` or `* ` → paragraph with `pStyle` `ListParagraph` (`numPr` wiring is Phase 2); anything else → plain paragraph (inline markdown is not interpreted). A `style` argument overrides every inserted paragraph's `pStyle`: the styleId is the argument verbatim if defined in `word/styles.xml`, else the argument with whitespace removed if defined, else `style_unknown`. With tracking, each inserted paragraph's run is wrapped in its own `w:ins` (§5 ids in document order). Exactly one of `after`/`before` must be given (`anchor_invalid`).
- **`docx_delete`** — exactly one of `anchor`/`range` (`anchor_invalid`). Range endpoints use the §2a grammar, but on this edit tool an endpoint carrying `#hhhh` is validated like an anchor (mismatch → `anchor_stale`). The range deletes paragraphs `a..b` only; body tables between them are untouched (table ops are Phase 2). Tracked deletion wraps each non-empty paragraph's full run content in one `w:del` per paragraph (the paragraph mark survives in MVP); empty paragraphs are counted but emit no wrapper.
- **`docx_replace`** — zero matches with `all: true` → `{n_replaced: 0, anchors: []}` (idempotent), without → `not_found`; >1 match without `all: true` → `ambiguous_target`. `all: true` returns `anchors` = the fresh anchors of affected paragraphs ascending; otherwise `new_anchor`. Matches are non-overlapping, left-to-right, found per paragraph before any splice.
- **`docx_revision`** — revision ids render as `R{w:id}`; `list` entries carry the containing body block's anchor (omitted when the wrapper sits outside any body block) and the wrapper's own text (`w:t` for ins, `w:delText` for del, raw). Filters: `author`/`date` per §7, plus `after` (`w:date >= after`) and `before` (`w:date < before`) as lexicographic ISO comparisons. `accept`/`reject` without `id`/`filter` resolve everything; an `id` selecting nothing resolves nothing (§7 idempotency — `{accepted: 0}`, not an error); `accept_all`/`reject_all` ignore `id`/`filter`. A candidate nested inside another candidate resolves with its container. The §7 merge post-pass runs over the affected paragraphs, and `anchors` returns their fresh anchors in document order with `remaining_by_author` counting what is left.

## 7. Revision accept/reject

- **Accept `w:ins`** — unwrap: keep child content, drop the wrapper element.
- **Accept `w:del`** — remove the element and its content.
- **Reject** — the exact inverse (reject ins = remove; reject del = unwrap with `w:delText` → `w:t`).
- **Filters**: `author` = exact string match; `date` = ISO **date-prefix** match (`"2026-06-10"` matches any `w:date` starting with it).
- Idempotent: nothing matching → `{accepted: 0}`, not an error.
- **Post-pass**: after resolution, adjacent sibling runs whose `rPr` are identical _after dropping `rsid*` attributes_ merge into one run (their `w:t` contents concatenate; `xml:space` per §3). Affected anchors are recomputed and returned fresh.

## 8. Validator (MVP checks)

| #   | Check                                                                                                       | Severity |
| --- | ----------------------------------------------------------------------------------------------------------- | -------- |
| a   | `[Content_Types].xml` covers every part (an `Override` for the part name, or a `Default` for its extension) | error    |
| b   | Every `r:id` referenced in `word/document.xml` resolves in `word/_rels/document.xml.rels`                   | error    |
| c   | Every relationship target (non-External) exists in the package                                              | error    |
| d   | `w:ins`/`w:del` `w:id` values are unique (counted together)                                                 | error    |
| e   | Comment and footnote references resolve **both directions** (every ref has a definition, and vice versa)    | error\*  |

\* a definition with no reference is a **warning**; a reference with no definition is an **error**. Relationships that nothing references are warnings.

Severity `error` blocks `docx_save` (`validation_failed`); `warning` never blocks.

### 8a. Pinned issue surface

`docx_validate` results are parity-compared deep-equal, so the issue list is pinned. `valid` is `true` iff no issue has severity `error`. Issues are emitted in check order **a → e**, scan order within a check as below; every issue carries `severity`, `part`, `message`, `fix_hint`.

- **a** — package entries in zip order, skipping directory entries and `[Content_Types].xml` itself. Uncovered part → error, `part` = the part name, message `Part {name} is not covered by [Content_Types].xml (no Override, no Default for extension '{ext}').` (`{ext}` = lowercased extension of the entry's basename, `""` when it has none), fix_hint `docx_repair adds a content-type Default for the extension.`
- **b** — `r:id` attribute values of the main document part, distinct, in order of first occurrence. Id absent from the part's rels → error, `part` = the document part, message `r:id {rid} is referenced in {part} but not defined in {rels part}.`, fix_hint `Add the missing relationship or remove the referencing element; not auto-repairable.`
- **c** — every `.rels` part in zip-entry order, relationships in document order; a non-External target that does not resolve to a package part → error, `part` = the rels part, message `Relationship {rId} targets missing part {target}.`, fix_hint `docx_repair drops the orphaned relationship.` Then the unreferenced-relationship **warnings**: a relationship of the main document part (any TargetMode) whose type's last path segment is **not** in the implicit set {`styles`, `settings`, `webSettings`, `fontTable`, `numbering`, `theme`, `customXml`, `comments`, `commentsExtended`, `footnotes`, `endnotes`, `glossaryDocument`} and whose Id never appears as an `r:id`/`r:embed`/`r:link` attribute value in the document part → warning, `part` = the rels part, message `Relationship {rId} ({shortType}) is never referenced.`, fix_hint `Harmless; remove the unused relationship to tidy the package.`
- **d** — `w:ins`/`w:del` of the document part in document order; each `w:id` value carried by more than one element → one error per value, ordered by first occurrence, `part` = the document part, message `Duplicate revision id {id} on {n} w:ins/w:del elements.`, fix_hint `docx_repair renumbers the later duplicates.`
- **e** — comments then footnotes. Refs = `w:commentReference`/`w:footnoteReference` `w:id` values in the document part; defs = `w:comment`/`w:footnote` `w:id` values in `word/comments.xml`/`word/footnotes.xml` (a missing part defines nothing). Per story: errors first (refs with no def, first-occurrence order), then warnings (defs with no ref, document order; footnote defs whose `w:type` is `separator`/`continuationSeparator` are exempt). `part` = the comments/footnotes part name. Messages `Comment id={id} referenced in body but missing.` / `Comment id={id} defined but never referenced.` (`Footnote` likewise); fix_hints `docx_repair removes the orphaned reference.` / `Harmless; delete the unused definition to tidy the package.`

**`docx_repair`** applies the mechanical fixes below in this order, reporting one `fixed[]` string per fix; it then re-validates and returns `remaining[]` = the `message` of every error-severity issue still present. The result `{fixed, remaining}` is returned even when `remaining` is non-empty (`repair_incomplete` is reserved for a fix that cannot be applied at all).

1. drop orphaned relationships (check c) — `removed orphaned relationship {rId} ({rels part})`;
2. add missing content-type `Default` entries (check a; one per distinct extension in first-occurrence order, spliced in before `</Types>`) — `added content-type Default for extension '{ext}'`; the content type comes from the pinned map {`rels`: `application/vnd.openxmlformats-package.relationships+xml`, `xml`: `application/xml`, `png`: `image/png`, `jpeg`/`jpg`: `image/jpeg`, `gif`: `image/gif`}, else `application/octet-stream`; an extension-less part is not fixable;
3. renumber duplicate revision ids (check d; the first element in document order keeps its id; later duplicates take max+1 onward in document order) — `renumbered duplicate revision id {old} -> {new}`;
4. remove orphaned comment/footnote references (check e errors): the `w:commentReference` element plus any `w:commentRangeStart`/`w:commentRangeEnd` carrying the same id, or the `w:footnoteReference` element — `removed orphaned comment reference id={id}` / `removed orphaned footnote reference id={id}` (one string per removed reference element).

## 9. Save

1. Run full validation (§8); refuse with `validation_failed` on any error-severity issue (the error suggests `docx_repair`).
2. Stream the **source zip's entries in their original order**: untouched parts are copied byte-for-byte (decompressed content identical; entries are re-stored to normalize metadata); modified parts are re-serialized from their spliced buffers; **new parts are appended** after the originals in creation order.
3. Entry metadata is normalized: DOS timestamp 1980-01-01 00:00:00, no extra fields, no entry comments, deflate compression level 6.
4. **Atomic write**: serialize to a temp file in the destination directory, then rename over the target.

Deflate output may differ between zlib and fflate at the same level — byte-equivalence across implementations is asserted on **decompressed** part contents (§10), not raw archive bytes.

## 10. Normalization for cross-implementation comparison

The conformance harness compares two output packages as follows:

1. Unzip both; the ordered lists of entry names must be equal.
2. For each entry, compare decompressed bytes. Equal → pass.
3. **Canonical-XML fallback** (only for `.xml`/`.rels` entries): parse both, sort each element's attributes by attribute name (byte order of the qualified name as written), drop whitespace-only text nodes between elements, then compare the canonical serializations. Equal → pass; else the case fails with a part-level diff.

The fallback exists so that attribute-order or inter-tag-whitespace differences in _modified_ parts don't fail conformance; semantic differences always do.

## 11. CLI contract (conformance driver)

Both implementations ship a line-oriented JSON CLI used by the harness:

- Read one JSON object per line from stdin: `{"tool": "docx_replace", "args": {…}}`.
- Write exactly one JSON object per line to stdout, in request order: the tool's result object, or an error object `{"error": code, "message": …, "suggestions": […]}`.
- `doc_id`s persist for the process lifetime; EOF on stdin → exit 0. stderr is free-form logging.
- Tools defined in [spec/tools/](tools/) but not yet implemented return `not_implemented`.
- A request line that is not a JSON object with a string `tool`, or whose `args` lack a schema-required key, yields `invalid_args`.

Entry points: Python `python -m docxengine.cli`; TypeScript `node js/dist/cli.js`.

## 12. Dependencies

- **Python**: standard library only.
- **TypeScript**: `fflate` (zip) only.

Anything else (renderers, converters) is an optional adapter and never required by the algorithms above.

# Phase 2 algorithms

Phase 1 (§1–§12) is in force and proven (15/15 parity). This section pins Phase 2 with the same authority: where a byte is named, it is normative. All emission obeys §3 (splice, don't re-serialize; the §3 escape rules; `xml:space="preserve"` iff leading/trailing whitespace). All new ids follow the Phase-1 idiom **max existing + 1**. All dates use `DOCXENGINE_FIXED_DATE` when set (§5). New parts are appended in creation order (§9) and every `op` ends by passing the §8 validator. Shared property names (`size_pt`, `spacing_before_pt`, `spacing_after_pt`) are the canonical spec names; the tool JSON shorthands (`size`, `spacing_after`, `spacing`) map onto them verbatim (`size`→`size_pt`, points; `spacing_after`→`spacing_after_pt`; `spacing`→line multiplier, out of the closed prop set below and ignored by the style/format writers).

## 13. Anchor sequences (Phase 2 invariant)

Body-level `w:p` ordinals (`P{n}`) and body-level `w:tbl` ordinals (`T{n}`) are **independent** sequences, each numbered in document order over its own element type. Inserting, deleting, or re-anchoring a table never shifts any `P{n}`; inserting a paragraph never shifts any `T{n}`. Paragraphs nested in table cells, headers, footers, or comments are **not** in the body `P` sequence (consistent with §1). `T{ordinal}` counts body-level `w:tbl` in document order; a freshly created table's `new_anchor` is `T{k}@after:{prev}` where `prev` is the §2 `@after:` token of the paragraph it follows (`@start` if none). Media ids `M{ordinal}` count `<a:blip>`/drawing references in document order; comment ids render `C{w:id}`, revisions `R{w:id}`, sections `S{n}` over `w:sectPr` in document order (the body trailing `sectPr` is the last `S`).

## 14. Tables (`docx_table`)

**Model.** A table is `w:tbl` → `w:tblPr` → `w:tblGrid` (`w:gridCol w:w` per column) → `w:tr*` → `w:tc*`, each `w:tc` → `w:tcPr?` → `w:p+`. Addressing: `{r,c}` 0-based, or A1 where the column letter(s) are base-26 (`A`=0 … `Z`=25, `AA`=26) and the 1-based row number maps to `r = row−1` (so `A1` = `{r:0,c:0}`, `B2` = `{r:1,c:1}`). When both `{r,c}` and `ref` are given, `{r,c}` wins.

**create.** `rows`×`cols`, optional `data` (row-major; short rows pad with empty cells, overflow `anchor_invalid`). Grid columns get equal integer widths summing to the §15 default content width in twips (`9026`): `w = floor(9026/cols)`, the **last** column absorbs the remainder. Each `w:tc` is `<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr><w:p>…</w:p></w:tc>`. When `header:true` (or a `style` is requested), emit `<w:tblStyle w:val="TableGrid"/>` first in `w:tblPr` and ensure the `TableGrid` style exists in `styles.xml` (§16 ensure-style); the header row's cells additionally carry `<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>` in `w:tcPr` and bold runs (`<w:b/>`). Cell text becomes one run; empty text → an empty `<w:p/>`.

```xml
<w:tbl>
  <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>
  <w:tr><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Term</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Fee</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">$100 </w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
```

**set_cells.** For each addressed cell, replace its paragraphs' content with one `<w:p>` carrying [the cell's first `w:p`'s `w:pPr` verbatim if present] + one run with the new text (preserving `w:tcPr`); a covered (`vMerge`/`gridSpan`) cell address is `anchor_invalid`. **insert_row** at `at`: clone the structure (cell props, not text) of row `min(at,last)`, blank text, insert before index `at` (`at == rows` appends). **insert_col** at `at`: add one `w:gridCol` (width = the split of the neighbor's width) and one blank `w:tc` per row at column `at`; widths re-floor across the grid. **delete_row/delete_col** remove that index and (for col) its `w:gridCol`; a `vMerge`-origin row deletion promotes the next continuation to origin. **merge** over an A1 `range`: horizontal span sets `<w:gridSpan w:val="{n}"/>` on the left cell and **removes** the covered cells in that row; vertical span sets `<w:vMerge w:val="restart"/>` on the top cell and `<w:vMerge/>` (continue) on each cell below, which keep an empty `<w:p/>`; a rectangular range applies gridSpan to each spanned row then vMerge down the merged left column. A `w:gridSpan`/`w:vMerge` mark is spliced as the **first child** of the cell's `w:tcPr` (creating an empty `w:tcPr` as the cell's first child when absent); when a rectangular merge writes both onto the left column, the `w:vMerge` (written second) therefore precedes the `w:gridSpan`. Projection (§2) renders the table as GitHub markdown unchanged; merged-away cells contribute empty string.

## 15. Sections (`docx_section`)

`w:sectPr` carries `<w:pgSz w:w="" w:h="" w:orient=""/>` and `<w:pgMar w:top w:right w:bottom w:left w:header w:footer w:gutter/>` (twips; 1 cm = 567 twips, 1 in = 1440). Presets (portrait twips): **A4** `w=11906 h=16838`, **Letter** `w=12240 h=15840`, A3 `16838×23811`, A5 `8391×11906`, Legal `12240×20160`, Tabloid `15840×24480`. Default margins `1440` all sides, `header=708 footer=708 gutter=0`; content width = `w − left − right` (A4 default = `9026`, §14). **orientation=landscape** swaps `w`↔`h` and sets `w:orient="landscape"`; portrait removes `w:orient` (default). `set_geometry` splices the named section's existing `pgSz`/`pgMar` in place (creating them inside `w:sectPr` when absent), keeping unspecified attributes. `columns>1` sets `<w:cols w:num="{n}" w:space="708"/>`. `insert_break` after a paragraph clones the body `sectPr` into a `<w:pPr><w:sectPr>…<w:type w:val="{break_type}"/></w:sectPr></w:pPr>` on that paragraph.

**Headers/footers.** `set_header`/`set_footer` for `variant` (`default`→`w:type="default"`, `first`→`first`, `even`→`even`) creates `word/header{N}.xml` / `word/footer{N}.xml` (`N` = max existing + 1), adds the part rel to `word/_rels/document.xml.rels`, adds a content-type `Override` (`…wordprocessingml.header+xml` / `.footer+xml`), and splices `<w:headerReference w:type="{variant}" r:id="{rId}"/>` (footers `<w:footerReference …/>`) into the target `w:sectPr` (references precede `pgSz`). Content is the §22 markdown→paragraph mapping (plain paragraphs only — no lists/tables in headers MVP).

```xml
<!-- word/header2.xml -->
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:hdr>
<!-- in document.xml sectPr -->
<w:headerReference w:type="default" r:id="rId7"/>
```

## 16. Styles (`docx_style`) & formatting (`docx_format`)

**list** walks `word/styles.xml`: each `w:style` → `{id: @w:styleId, name: w:name/@w:val, type: @w:type, based_on: w:basedOn/@w:val (key omitted when absent), in_use: count of body paragraphs whose effective style (the §2 `pStyle`+`basedOn` resolution) is this id}`, document order. **define** appends `<w:style w:type="paragraph" w:styleId="{id}">` before `</w:styles>` where `id` = `name` with whitespace removed (collision → suffix `2`,`3`…); children in **fixed order**: `w:name`, `w:basedOn?`, `w:pPr?`, `w:rPr?`. **apply** splices `<w:pStyle w:val="{id}"/>` as the first `w:pPr` child of the anchor paragraph (resolving a `name` to its `styleId`; unknown → `style_unknown`). **Style-reference resolution** (used by `apply`, `docx_format`'s `style_selector`, and table/list `style`): the argument matches a styleId verbatim, else the argument with §1 whitespace removed if that is a styleId (`"Heading 2"` → `Heading2`), else a style whose `w:name/@w:val` equals the argument; otherwise `style_unknown`. When `apply` finds an existing `w:pStyle`, it is replaced in place; otherwise the new `w:pStyle` is inserted as the first `w:pPr` child.

**Closed prop set & emission order.** Inside `w:rPr`, emit present props in exactly: `w:b` (bold), `w:i` (italic), `w:u w:val="single"` (underline), `w:color w:val="RRGGBB"` (hex, no `#`, uppercased), `w:sz w:val="{2×pt}"` (size_pt → half-points). Inside `w:pPr` (after any `w:pStyle`): `w:jc w:val="{left|center|right|both}"` (alignment; `justify`→`both`), `w:spacing w:before="{20×pt}" w:after="{20×pt}"` (twentieths; only the given attrs). A boolean `false` emits the toggle off form `<w:b w:val="0"/>`.

```xml
<!-- docx_style define name="Clause" based_on="Normal" props={size_pt:11,bold:true,spacing_after_pt:6,alignment:"justify"} -->
<w:style w:type="paragraph" w:styleId="Clause">
  <w:name w:val="Clause"/><w:basedOn w:val="Normal"/>
  <w:pPr><w:jc w:val="both"/><w:spacing w:after="120"/></w:pPr>
  <w:rPr><w:b/><w:sz w:val="22"/></w:rPr>
</w:style>
```

**docx_format.** With `style_selector.style`: resolve to a `styleId`, then **merge** the props into that style's `w:rPr`/`w:pPr` (creating them in the §16 child order, replacing same-named children) — one edit, document-wide, idempotent. With `anchor`/`range`: apply as **direct** formatting — `w:rPr` props splice into every run's `w:rPr` (created as the first `w:r` child), `w:pPr` props into each paragraph's `w:pPr`. `affected` = paragraphs touched; `anchors` = their fresh anchors ascending. `track_changes` on direct formatting wraps each changed run's rPr delta in `<w:rPr><w:ins …/></w:rPr>` is **not** done in MVP; tracked formatting records a `w:pPrChange`/`w:rPrChange` with the prior props and §5 id/author/date.

## 17. Lists (`docx_list`) & numbering

If `word/numbering.xml` is absent, create it (root `<w:numbering>`), add the content-type `Override` (`…wordprocessingml.numbering+xml`), and add the document rel (type `…/numbering`). **ol** abstractNum has 9 levels cascading `decimal, lowerLetter, lowerRoman, decimal, …` with `lvlText` `%1.`,`%2.`,`%3.` and indent `<w:ind w:left="{720×(ilvl+1)}" w:hanging="360"/>`. **ul** cascades bullet glyphs `•`(Symbol-less, char `•`), `◦`, `▪` with `numFmt="bullet"`, `lvlText` the glyph, same indents. `create` allocates `abstractNumId` = max+1, a `w:num` with `w:numId` = max+1 pointing at it, and sets each item's `<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{numId}"/></w:numPr>` as the first `w:pPr` child; the paragraph also gets `pStyle ListParagraph` (ensured). `convert to:ol|ul` reuses/creates one abstractNum for the run and sets numPr; `to:paragraphs` removes `w:numPr`. `set_level` rewrites `w:ilvl`. **restart** allocates a **new** `w:num` (numId max+1) referencing the _same_ abstractNum with `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="{at}"/></w:lvlOverride>` and repoints the target paragraph's `numId`.

```xml
<!-- numbering.xml: ol abstractNum levels 0–1 shown -->
<w:abstractNum w:abstractNumId="3">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num>
<!-- restart at 5: new num, same abstractNum -->
<w:num w:numId="5"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>
```

## 18. Comments (`docx_comment`)

**add** wires all five places, id = max `w:comment/@w:id` + 1 (start 0): (1) `<w:commentRangeStart w:id="{id}"/>` before the anchor paragraph's runs and (2) `<w:commentRangeEnd w:id="{id}"/>` + (3) `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{id}"/></w:r>` after them; (4) create `word/comments.xml` if absent (content-type Override `…wordprocessingml.comments+xml` + document rel type `…/comments`) and append the `w:comment`; (5) ensure the `CommentReference` style. `initials` = the uppercased first letter of each whitespace-separated author word (`"Jane Q. Doe"`→`JQD`, empty author→`""`).

```xml
<!-- in document.xml around the target paragraph -->
<w:commentRangeStart w:id="3"/><w:r><w:t>…paragraph runs…</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="3"/></w:r>
<!-- word/comments.xml -->
<w:comment w:id="3" w:author="Jane Q. Doe" w:date="2026-06-11T00:00:00Z" w:initials="JQD"><w:p><w:r><w:t>Should this be mutual?</w:t></w:r></w:p></w:comment>
```

**reply** appends a new `w:comment` (next id) and, in `word/commentsExtended.xml` (w15; create with content-type + rel on demand), a `<w15:commentEx w15:paraId="{child}" w15:paraIdParent="{parent}" w15:done="0"/>` (each `w:p` in comments carries a `w14:paraId`; reply parent = the thread root's paraId). **resolve** sets `w15:done="1"` on the thread root's `commentEx` (creating it if absent). **delete** removes all five places for the id (and its replies). **list** returns one entry per thread root: `{id, anchor (the body anchor of its range start), author, date, text, resolved (w15:done=1), replies:[{author,date,text}]}`. id allocation = max+1 across comments.

The `w14:paraId` of a created comment's `w:p` is **derived deterministically** for cross-implementation byte parity (Word's random ids are not reproducible): the first **8 uppercase hex chars** of the SHA-256 of the UTF-8 encoding of `paraId:{w:id}:{text}` (the comment's `w:id` and its body text). Both implementations MUST use this identical derivation.

## 19. Media (`docx_media`)

**insert** writes `word/media/image{k}.{ext}` (`k` = max existing + 1; `ext` lowercased from the source path), a document rel (type `…/image`, Target `media/image{k}.{ext}`), a content-type `Default` for the extension (per the §8 repair map; `application/octet-stream` fallback), and splices an inline drawing run after/before the anchor paragraph. EMU = `round(cm × 360000)`. When only one of `width_cm`/`height_cm` is given, parse the source's pixel dimensions and scale by aspect: **PNG** = bytes 16–24 big-endian `width`,`height` after the `IHDR` signature (`\x89PNG\r\n\x1a\n` then `IHDR`); **JPEG** = scan segments from `FFD8`, skip each `FFxx` length-prefixed marker, read 16-bit big-endian `height`,`width` from the first `SOF0/1/2` (`FFC0`/`C1`/`C2`); EMU then `other = round(given × otherPx/givenPx)`.

```xml
<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
  <wp:extent cx="1440000" cy="1080000"/><wp:docPr id="1" name="image1"/>
  <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="image1"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="rId8"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1440000" cy="1080000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
  </a:graphicData></a:graphic>
</wp:inline></w:drawing></w:r>
```

**extract** copies the `M{id}` part's bytes to `path` (`path_denied` outside allowed roots). **replace** overwrites the part's bytes keeping the rel and rId (content-type Default added if the new extension differs). `M{ordinal}` = document order of drawing references.

## 20. Fields (`docx_field`)

**insert_toc** emits, after the anchor, a paragraph holding the run-triple field (never `w:fldSimple`): `instrText` = `TOC \o "1-{levels}" \h \z \u`.

```xml
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Right-click to update field.</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
```

**insert_page_number** ensures the section's footer exists (§15 machinery; create when absent) and appends to it a paragraph with the `PAGE` field run-triple (`instrText` `PAGE`, separate, placeholder `1`, end). **update** sets `<w:updateFields w:val="true"/>` in `word/settings.xml` (create the settings part — content-type Override `…wordprocessingml.settings+xml` + document rel type `…/settings` — when absent; `updateFields` is the first child of `w:settings`). Values materialize only at render; results never report computed page/TOC numbers.

## 21. Templates (`docx_template_fill`)

Mustache subset: `{{var}}`, `{{#s}}…{{/s}}` (loop over array / render-once on truthy non-array), `{{^s}}…{{/s}}` (inverted: render when falsy/empty), `{{!c}}` (comment, dropped). Matching runs on **coalesced** paragraph text (§4 offset map) so split-run placeholders resolve. A `{{var}}` is replaced in place; the §4 first-overlap rule writes the value into the first run and trims the rest. Loop/inverted sections that begin and end within paragraphs of one body region expand by **cloning the spanned paragraphs** per array element (substituting `{{.}}` and `{{key}}` from the element); when the section's open and close tags sit in cells of **exactly one table row**, the **row** is cloned instead. Missing vars are left **verbatim** and listed in `unfilled` (dedup, document order); `strict:true` → `placeholder_unfilled`. Values are stringified: booleans `true`/`false`→`"true"`/`"false"` only as `{{var}}` text (sections treat them as truthiness), numbers via shortest round-trip, arrays/objects only drive sections. **XML-escaping only** (§3: `&`,`<`,`>`), never HTML-escaping. `filled` = placeholders resolved; `loops_expanded` = `{section: element_count}`.

```xml
<!-- template: <w:p><w:r><w:t>Client: {{Cli</w:t></w:r><w:r><w:t>ent}}</w:t></w:r></w:p>, data {Client:"GlobalTech & Co"} -->
<w:p><w:r><w:t xml:space="preserve">Client: GlobalTech &amp; Co</w:t></w:r></w:p>
```

## 22. Create from markdown (`docx_create content_md`)

Deterministic skeleton parts (creation order): `word/document.xml`, `word/styles.xml`, `[Content_Types].xml`, `_rels/.rels`, `word/_rels/document.xml.rels`, `docProps/core.xml` (with `<dcterms:created>`/`<dcterms:modified>` = `DOCXENGINE_FIXED_DATE` or its default). `styles.xml` ships `Normal`, `Heading1`…`Heading6`, `ListParagraph`, `TableGrid`, `Quote`, plus styles ensured on demand. Block mapping: `#`×n + space → `pStyle Heading{n}` (1–6); `>` → `pStyle Quote` (indented); `---`/`***` alone → an empty paragraph with `<w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr>`; `- `/`* `/`1. ` → list items via §17; GitHub `| a | b |` table (with `|---|` separator row) → §14 table (`header:true` when a separator row follows row 1); else a plain paragraph. **Inline**: `**x**`/`__x__`→`<w:b/>` run, `*x*`/`_x_`→`<w:i/>`, `` `x` ``→run with `<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>`; the parser splits text into runs at marker boundaries, escaping per §3. `n_paragraphs` = body-level `w:p` count.

```xml
<!-- "## Scope\nSee **clause** `4a`." -->
<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Scope</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>clause</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>4a</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>
```

## 23. Convert (`docx_convert`)

**to md** renders the §2 projection model to GitHub markdown: headings `#`×level + text; ordered items `{n}. `, unordered `- ` indented two spaces per `ilvl`; tables as the §2 GitHub table (header = row 1); `**bold**`/`*italic*` reconstructed from run `w:b`/`w:i` when emitting detailed runs; comments inline as `<!-- comment:{author}: {text} -->` at the range end; revisions in **accepted view** by default (ins shown, del omitted), with markers `[ins]`/`[del]` configurable off. `&`,`<`,`>` are markdown-literal (not escaped). **to html** emits minimal semantic tags (`<h1>`…, `<p>`, `<ul>/<ol>/<li>`, `<table>/<tr>/<td>`, `<strong>/<em>`) with **inline styles only** for alignment (`style="text-align:center"`) and color (`style="color:#RRGGBB"`); text is HTML-escaped (`&`,`<`,`>`,`"`). **to pdf/png** require the §24 render adapter; unavailable → `render_unavailable`.

```text
# Master Services Agreement

| Term | Value |
| --- | --- |
| Fee | $100 |

This is mutual. <!-- comment:Jane: confirm scope -->
```

### 23a. Stage-3 byte pins (convert / create / template)

These resolve choices the prose above left open; `md`/`html` content is parity-compared deep-equal, so both implementations MUST follow them.

- **md block joins** — body blocks join with a blank line (`"\n\n"`), **except** two consecutive list items, which join with a single `"\n"` (a tight list). Inline `[ins]` precedes the run text and a paragraph carrying any tracked deletion appends a trailing `[del]` (both suppressed when markers are off). Inline comment notes (` <!-- comment:{author}: {text} -->`) append after the paragraph text in `w:commentReference` order; `{text}` is the §1 normalized comment body. No trailing newline.
- **html structure** — each non-list paragraph is one `<p>`/`<h{n}>`; consecutive list items of the same `kind` are wrapped in one `<ul>`/`<ol>` and emit `<li>…</li>`; tables emit `<table><tr><td>…`. Alignment maps `both`→`justify`; `style` packs `text-align` then `color` joined by `;`. Block lines join with `"\n"`.
- **create (`content_md`)** — every emitted body paragraph is a bare `<w:p>` (no `w:spacing`); the document carries a trailing `<w:sectPr>` with the §15 A4 default `pgSz`/`pgMar`. The skeleton `styles.xml` ships `Normal` (default), `Heading1`–`Heading6` (`basedOn` Normal, `outlineLvl` n−1, bold run), `ListParagraph`, `TableGrid`, `Quote`. List items allocate `numbering.xml` numIds starting at 1 (one ol abstractNum, one ul abstractNum, in first-use order); `numbering.xml` is created only when a list item exists. `core.xml` carries `dcterms:created`/`dcterms:modified` = `DOCXENGINE_FIXED_DATE` (else `2026-01-01T00:00:00Z`). `n_paragraphs` counts body-level `w:p` only (tables excluded). Inline markdown (`**`/`__`, `*`/`_`, `` ` ``) is parsed in body paragraphs **and** table cells; markers do not nest (an unmatched marker is literal).
- **template `filled`** — `filled` counts every `{{var}}` substitution that resolved to a value, **including** substitutions inside expanded loop clones (so the §21 worked example with two obligations yields `filled` = Client + EffectiveDate + text×2 = 4). `{{!comment}}` removals and unresolved vars do not count. Inverted sections (`{{^s}}`) are conditions, never recorded in `loops_expanded`. Section open/close tags that sit on their own whole paragraphs drop those tag-only paragraphs and clone only the inner paragraphs between them.

## 24. Render adapter (`docx_convert` pdf/png, `docx_render_preview`)

Detection order: env `DOCXENGINE_SOFFICE`; then `soffice` on `PATH`; then platform defaults `/Applications/LibreOffice.app/Contents/MacOS/soffice`, `/usr/bin/soffice`. Invocation: `soffice --headless --convert-to pdf --outdir {DIR} {FILE}` with a **per-call temp profile** dir via `-env:UserInstallation=file://{tmp}` (cleaned up after). On success `renderer` = `"libreoffice {version}"`. When no binary is found, the **structural fallback**: `{pages: null, structural: <§2 projection + estimated page count = ceil(total_chars / 1800)>, renderer: "structural"}` (no error from preview; `docx_convert` to pdf/png with no adapter is `render_unavailable`). PNG via PDF then `pdftoppm` (or macOS `sips`) when available, else the structural fallback.

## 25. MCP Streamable HTTP + resources (Python server only)

`docxengine-mcp --http --port {p}` runs a stdlib `http.server` (threading). **POST /** accepts a JSON-RPC body (same dispatch as stdio), responds `application/json`. `initialize` allocates an `Mcp-Session-Id` (returned as a response header); every subsequent POST must send it (missing/unknown → JSON-RPC error; an expired session → HTTP **410**). Each session owns its own doc store. **GET /health** → `200` (body `{"status":"ok"}`). **resources/list** returns, per open doc, `docx://{doc_id}/outline` and `docx://{doc_id}/projection`; **resources/read** returns `text/markdown` from the projector (`outline` = §2a outline rendering, `projection` = the full §2 projection). stdio mode is unchanged and remains the conformance transport.
