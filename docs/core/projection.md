# Projection: the document agents see

Agents never see raw OOXML by default. Reads return a **Markdown-like projection** — the token-economy principle applied to documents: all content has a fixed cost in context, and distracting content harms performance.

## Format

```
[P1#a7b2  H1]            Master Services Agreement
[P2#f3c1]                This Agreement is entered into as of {{EffectiveDate}}...
[P3#b2c4  H2]            1. Definitions
[P4#d4e5]                "Confidential Information" means... [comment:C1 by J.Doe]
[T1  3×4 @after:P5]      | Term | Value | ... |
[P12#e7f8  List:ol L1]   First obligation
```

Each line: `[anchor  annotations]` followed by the resolved text.

### What is surfaced (salient formatting only)

- Heading level (`H1`–`H6`), resolved through the style cascade
- List type and level (`List:ol L1`, `List:ul L2`), with compiled markers
- Bold/italic when semantically relevant (in `detailed` format)
- Comment markers (`[comment:C1 by J.Doe]`) and revision markers (`[ins by Jane]`, `[del by Bob]`)
- Table presence, dimensions, and position (`T1 3×4 @after:P5`)
- Template placeholders verbatim (`{{EffectiveDate}}`)

### What is hidden

rsids, default fonts, `proofErr`, `bookmarkStart/End` noise, namespace plumbing, and any formatting that merely restates a style default. If an agent needs it, `format: "detailed"` returns resolved run-level formatting — on request, never by default.

## Pagination: map first, zoom second

| Call                                     | Cost     | Returns                               |
| ---------------------------------------- | -------- | ------------------------------------- |
| `docx_outline(doc_id)`                   | cheapest | headings tree with anchors            |
| `docx_read(doc_id, {range: "P10..P30"})` | bounded  | projection of a paragraph range       |
| `docx_read(doc_id, {anchor, window: 6})` | bounded  | projection centered on an anchor      |
| `docx_search(doc_id, {query})`           | bounded  | matching anchors + snippets + context |

No response exceeds ~25k tokens; long ranges paginate with continuation hints. `docx_search` is search-focused (return matches), not list-all — agents should never page through a whole document to find something.

## Resolution guarantees

The projector compiles indirection so agents see _effective_ values:

- Numbering chains (`numbering.xml`) → real markers ("1.", "(a)")
- Style cascade (`styles.xml` + direct formatting) → effective heading levels and salient props
- Revision content: by default the projection shows the document _as if accepted_, with markers; `docx_revision {op: list}` gives the full change inventory.

## Reading order and faithfulness

The projection preserves stored reading order across stories: body first; headers/footers, footnotes, and comments are addressable separately (`docx_read {scope: "footnotes"}`). Content faithfulness — no dropped text, no hallucinated text, correct order — is a conformance-tested property, not a best effort (see [conformance](../conformance/corpus.md)).
