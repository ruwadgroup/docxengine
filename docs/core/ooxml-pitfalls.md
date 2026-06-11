# OOXML pitfalls

Why DOCX editing is hard for principled, structural reasons — not just verbosity. Every subsystem in the core exists to neutralize one of these. References: ECMA-376/ISO 29500, [MS-DOCX], Eric White's OpenXML series.

## The split-run problem

A paragraph (`w:p`) contains runs (`w:r`); runs cannot nest. Bolding one word splits a paragraph into three separate run elements, each carrying its own formatting. Consequence: **naive find-and-replace fails** because the search string is fragmented across `w:r`/`w:t` boundaries:

```xml
<w:r><w:t>Confiden</w:t></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:t>tial Infor</w:t></w:r>
<w:r><w:t>mation</w:t></w:r>
```

The engine's run-coalescing find/replace concatenates text content, builds an index map back to runs, performs the replacement, and re-coalesces — preserving formatting boundaries that are semantically meaningful and discarding the accidental ones.

## rsid fragmentation

Word stamps `w:rsid*` (revision-save IDs) on runs and paragraphs to improve document-merge accuracy. They have **no semantic meaning for editing** and are a major cause of documents being "broken out into so many runs." The engine strips/ignores them during text operations and never lets them surface in projections.

## Style inheritance is a six-layer cascade

Per ISO 29500 §17, effective formatting resolves through: document defaults → table style → numbering → paragraph style → character style → direct formatting — each layer able to add, remove, or override, with `basedOn` chains inside `styles.xml`. Answering "what color is this heading" requires the full walk. The style cascade resolver does it once, deterministically, so projections show _effective_ values and `docx_format` can edit the right layer (a style definition, not 50 direct overrides).

## Numbering is indirection

List numbering isn't stored with the text. It lives in `numbering.xml` as abstract numbering definitions, referenced through a chain of indirection (`w:numPr` → `w:num` → `w:abstractNum`) compiled at render time into "1." or "(a)". The numbering resolver compiles this chain so projections show real markers and `docx_list` can manipulate lists safely.

## Content is split across parts ("stories")

A docx is a ZIP/OPC package: `document.xml`, `styles.xml`, `numbering.xml`, `comments.xml`, `footnotes.xml`, `header*.xml`, `footer*.xml`, media, plus `.rels` relationship parts and `[Content_Types].xml`. A single logical edit — say, adding a comment — touches **five places**: body XML (range markers), the comments part, a relationship, the content-types manifest, and ID allocation. Get any one wrong and Word silently "repairs" the file. This is exactly what the [validation gate](validation.md) guards.

## Fields, TOC, and page numbers require a renderer

Page numbers and TOC entries are field codes computed at render time. There is no way to resolve them from the XML alone. The engine inserts and updates _field codes_ (`docx_field`) and explicitly tells agents the computed values don't exist until Word/LibreOffice renders — preventing hallucinated page numbers.

## Producer diversity

Documents arrive from Word (many versions), LibreOffice, Google Docs exports, python-docx, and templating systems. They differ in which optional parts exist, whether `w14:paraId` is present, namespace prefixes, and `mc:Ignorable` content. The OPC layer preserves what it doesn't understand (round-trip identity) instead of normalizing it away — see the fidelity invariant in [ARCHITECTURE.md](../../ARCHITECTURE.md#invariants).
