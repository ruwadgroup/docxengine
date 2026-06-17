# DocxEngine Roadmap

A sequenced build plan with measurable gates. Phases are scoped so each ends with something testable against the conformance corpus and the agent benchmark — evaluation-driven from day one.

## Status

**Phases 0–2 complete. Phase 3 (Hardening) substantially landed for the v1.0.0-alpha.1 cut.**

- Phase 0 gate passed: every corpus document round-trips byte-stable with zero diff in both implementations.
- Phase 1 surface implemented and green in both languages, including the MCP stdio server.
- Phase 2 surface implemented and green: tables, styles, lists, sections, comments, media, fields, templates, create, convert (md/html in-engine), render adapter (LibreOffice + structural fallback), MCP Streamable HTTP + resources.
- Phase 3 (this cut): hostile-input hardening in both engines — zip-bomb caps (count/size/ratio), `<!DOCTYPE`/`<!ENTITY` rejection (XXE / billion-laughs), XML depth caps, path-traversal clamping — all tunable via `DOCXENGINE_MAX_*` and covered by adversarial suites + two cross-impl conformance cases (algorithms.md §27). Added the large-document perf benchmark (`bench/perf.py`), the cross-renderer fidelity harness + protocol (`conformance/fidelity/`, `docs/conformance/fidelity.md`), and the Rust/WASM core evaluation (`docs/research/rust-wasm-core.md`).
- Totals: 473 Python tests, 355 TS tests, 36/36 cross-implementation conformance cases, 10/10 agent-benchmark tasks over MCP (zero tool errors, zero Word-repair events).
- **Known perf hotspot (tracked):** `docx_replace {all: true}` splices the whole document once per match → superlinear when a match exists in every paragraph (`bench/perf.py --heavy` measures it). Single/anchored replace is linear. Fix: batch all matches into one splice pass — see Phase 3 performance tuning.
- Still open: the benchmark comparison against the python-docx-wrapper and raw-XML baselines (documented as protocol in `bench/`), large-document streaming, and the batch-splice perf fix above.

## Phase 0 — Foundations (weeks 1–3)

The substrate everything else depends on.

- OPC/ZIP package model (parts, relationships, content-types).
- XML DOM load/save with namespace + `mc:Ignorable` preservation.
- **Round-trip identity test**: open→save must be byte-stable modulo normalization.
- Content-hash anchor index (`P{index}#{hash}`), with `w14:paraId` as an optional seed.
- OOXML validator: ID uniqueness, orphaned rels, content-types, dangling footnotes.
- Shared conformance corpus + harness wired across Python and TS.

**Gate**: every corpus document round-trips with zero diff and zero Word-repair prompts in both implementations.

## Phase 1 — MVP (weeks 4–9)

The smallest surface that beats the baselines on real agent tasks.

- `docx_open` / `docx_outline` / `docx_read` / `docx_search`.
- Markdown projector with `concise`/`detailed` formats and pagination.
- `docx_replace` with split-run coalescing; `docx_insert` / `docx_delete` / `docx_edit_paragraph` (auto word-level diff).
- Tracked-change writer (`w:ins`/`w:del`) + `docx_revision` accept/reject with author filter.
- `docx_validate` / `docx_repair` / `docx_save` with the always-on validation gate.
- MCP server (stdio) + Python and JS packages + OpenAI function-calling adapter.

**Exit criterion (gate to Phase 2)**: beats the python-docx-wrapper and raw-XML baselines on the agent benchmark for text-edit and redline tasks, with **lower token use** and **zero Word-repair events**.

## Phase 2 — Full capability (weeks 10–18)

- Tables (`docx_table`): create, set_cells, merge, insert_row/col.
- Styles/themes: edit style definitions (`docx_style`, `docx_format` via style selectors).
- Lists/numbering, sections/page layout, headers/footers.
- Comments: add/reply/resolve/list/delete. Footnotes/endnotes.
- Media, hyperlinks, content controls.
- Fields/TOC/page-number insertion as field codes (`docx_field`).
- `docx_template_fill` (mustache, loops, conditions); `docx_create`; `docx_convert` (md/html in-engine; pdf/png via render adapter); `docx_render_preview`.
- MCP: Streamable-HTTP transport, session state, MCP resources.

**Gate**: full-surface parity between Python and TS on the expanded conformance corpus; benchmark coverage for table/style/comment/template tasks.

## Phase 3 — Hardening (weeks 19–24)

- Fuzzing: malformed/adversarial docx (zip bombs, duplicate IDs, broken rels, hostile XML).
- Large-document streaming.
- Cross-renderer fidelity checks (Word vs LibreOffice vs Google Docs).
- Performance tuning; v2 evaluation of Rust/WASM core unification.

## Testing & benchmarks

- **Round-trip fidelity corpus** — diverse real-world docx: legal contracts with redlines, academic papers with footnotes/citations, reports with TOC/tables/images, multi-section and multi-language docs. Metric: open→save→reopen produces no Word "repair" prompt and no semantic diff (content faithfulness: no dropped/hallucinated text, correct reading order).
- **Agent task benchmark** — single- and multi-edit natural-language tasks ("change all H2 to blue", "accept Jane's changes", "insert table after ¶12", "fill template") with element-level ground truth. Metrics: task success rate, total runtime of tool calls and tasks, number of tool calls, token consumption, tool errors — measured against a python-docx-wrapper MCP and a raw-XML baseline.
- **Fuzzing & repair** — corrupt-on-purpose docs (duplicate IDs, orphaned footnotes, broken rels): `validate`/`repair` must detect and fix or clearly report.
- **Visual diff** — render to PNG and compare against Word's rendering for layout-sensitive tasks.

## Decision thresholds

| Decision                             | Threshold                                                                                  |
| ------------------------------------ | ------------------------------------------------------------------------------------------ |
| Proceed to Phase 2                   | MVP ≥ baseline task success with lower token use and zero repair events on redline tasks   |
| Add Streamable HTTP + external state | More than a handful of concurrent users, or a hosted/gateway deployment requirement        |
| Invest in a faster renderer          | Preview latency becomes the dominant cost in agent loops                                   |
| Rust/WASM core unification (v2)      | Conformance-caught drift costs more than an FFI/WASM build, or edit throughput bottlenecks |

## Non-goals (for now)

- Rendering engine of our own (fields/TOC/page numbers resolve only in Word/LibreOffice).
- .doc (binary), .odt, .pptx, .xlsx — out of scope for v1; the architecture generalizes later.
- Real-time collaborative editing.
