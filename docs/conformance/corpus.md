# Conformance corpus

The corpus is what keeps two parallel implementations honest and what proves the fidelity invariant. Layout and harness rules live in [conformance/README.md](../../conformance/README.md); this page explains what goes in it and why.

## What the corpus tests

1. **Round-trip identity** — for every document: open→save must be byte-stable modulo normalization, reopen must show zero semantic diff, and Word must show zero "repair" prompts.
2. **Cross-implementation parity** — the same document + the same tool call must produce byte-equivalent-after-normalization output in Python and TS. Every conformance case is `(input.docx, tool_call.json, expected_output)`.
3. **Content faithfulness** — projections drop no text, hallucinate no text, and preserve reading order.

## Composition targets

| Category                      | Exercises                                                                          |
| ----------------------------- | ---------------------------------------------------------------------------------- |
| Legal contracts with redlines | multi-author `w:ins`/`w:del`, comments on revisions                                |
| Academic papers               | footnotes/endnotes, citations, cross-references                                    |
| Reports                       | TOC fields, nested tables, images, charts                                          |
| Multi-section documents       | headers/footers per section, page geometry changes                                 |
| Multi-language                | RTL (Arabic), CJK, combining marks — normalization edge cases                      |
| Non-Word producers            | LibreOffice, Google Docs exports, python-docx (no `w14:paraId`!)                   |
| Corrupt-on-purpose            | duplicate IDs, orphaned footnotes, broken rels — `validate`/`repair` fixtures      |
| Adversarial                   | zip-slip part names, entity expansion, oversized parts — must be _rejected_ safely |

## Contributing documents

The most valuable contribution to the project right now. Rules:

- **Never contribute confidential content.** Rebuild the structural pattern with synthetic text (lorem ipsum + fake names/dates).
- Keep documents minimal — the smallest file that still exhibits the structure.
- Name the producer: which application and version created it.
- Include the interesting behavior in the case metadata: "Word 2021 regenerates all paraIds when saving this after adding a comment."

## Normalization (what "byte-equivalent" means)

Raw ZIP bytes differ trivially (timestamps, entry order, compression level). The harness compares after normalizing: fixed entry order, zeroed timestamps, canonicalized XML whitespace/attribute order _in parts the engine touched_, untouched parts compared raw. A case fails if anything outside the normalization set differs.
