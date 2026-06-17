# Cross-renderer fidelity

The [conformance corpus](../../conformance/) and the Python test suite prove the engine produces **deterministic, stable** packages that round-trip without Word repair. Fidelity asks a different, harder question: when a real word processor _lays the document out_, does it look right — and does it look the same in Word, LibreOffice, and Google Docs?

This is the Phase 3 "cross-renderer fidelity checks" item. It is **partly manual by necessity**: Word and Google Docs have no headless, scriptable, license-clean rendering path that runs in CI. So fidelity is split into an automated tier that runs everywhere and a manual protocol for the renderers that need a human or a proprietary app.

## What's automated

`conformance/fidelity/run.py` runs two tiers:

1. **Structural fidelity — always on, no renderer required.** For each layout-sensitive corpus document it runs `docx_render_preview` and checks the structural projection (the always-available fallback adapter, see [render-adapter](../core/render-adapter.md)) is consistent with the document model from `docx_outline` — every heading the outline reports must appear in the projection, and the estimated page count must be sane. This catches projector and preview regressions in every CI run, with nothing installed.

2. **Visual rendering — when LibreOffice is detected.** If `soffice` is on the box (auto-detected, or pinned via `DOCXENGINE_SOFFICE`), the harness renders each document to PDF through the engine's render adapter and writes `conformance/fidelity/manifest.json` (renderer label + output size per document). That manifest is the artifact a maintainer reviews and the input to the manual comparison below.

```bash
.venv/bin/python conformance/fidelity/run.py
```

Exit status is non-zero only on a structural inconsistency, or on a render error **when a renderer is present**. With no renderer installed it reports "structural only" and passes.

> Visual output is intentionally **not** part of the byte-stability checks: renderer output is non-deterministic (timestamps, font substitution, layout-engine version), so it is reviewed, not diffed for equality. See [§10 normalization](../../spec/algorithms.md) for why stability is asserted on the _package_, not the _rendering_.

## The manual protocol (Word · LibreOffice · Google Docs)

Run this when a change could affect layout (tables, sections/columns, numbering, fields/TOC, media sizing, page breaks) or before a release.

**Corpus.** Use the layout-sensitive corpus documents — `minimal`, `tables`, `numbered-lists`, `headers-footers`, `media-doc` — plus any document that reproduces a reported fidelity bug (add it to the corpus via `conformance/harness/make_fixtures.py` so it is regenerated deterministically).

**Procedure.** For each document:

1. Open it in **Microsoft Word** (the reference renderer — it defines "correct" for `.docx`), **LibreOffice Writer**, and **Google Docs**.
2. Compare against this checklist:
   - **Reading order & content** — no dropped, duplicated, or reordered text; tracked changes render as insertions/deletions, not raw markup.
   - **Tables** — column widths, merged cells, borders, and cell content match; no collapsed or overflowing cells.
   - **Lists & numbering** — numbering restarts, indentation levels, and bullet/number glyphs match.
   - **Sections & page layout** — page size/orientation, margins, and **headers/footers** appear on the right pages.
   - **Fields** — TOC entries and page numbers _resolve_ (they are field codes until a renderer computes them; see [render-adapter](../core/render-adapter.md)).
   - **Media** — images appear at the right size (EMU sizing) and position.
3. Record the result in the PR (or a release checklist) as `Word: ok | LibreOffice: ok | Google Docs: <notes>`.

**Known, accepted differences** (not bugs):

- **Font substitution.** Word ships Calibri/Cambria; LibreOffice substitutes Carlito/Caladea (metric-compatible). Install the metric-compatible fonts to minimize layout drift. The render adapter names substitutions in its response so the difference is visible, not hidden.
- **Layout-engine drift.** Line-breaking and hyphenation differ slightly between engines; previews are approximate for pixel-exact validation.
- **Field formatting.** Date/number field _formatting_ can differ by renderer locale.

A difference that is **not** on the accepted list — especially dropped content, broken tables, or wrong reading order — is a fidelity bug: capture the document, add it to the corpus, and file it.

## Adding a reference rendering

To turn a spot-check into a regression guard, commit a reviewed reference for the renderer you can automate (LibreOffice):

1. Run the harness on a box with `soffice` installed; review the rendered PDFs against the manual checklist.
2. Once correct, commit `conformance/fidelity/manifest.json` as the baseline.
3. A future run on the same renderer version that diverges materially (e.g. a document that newly renders empty or with a wildly different size) is a signal to re-review.

Because renderer output is non-deterministic, the baseline is a **review aid**, not an equality assertion — the human checklist remains the source of truth for visual correctness.
