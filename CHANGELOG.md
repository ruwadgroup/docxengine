# Changelog

All notable changes to DocxEngine will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2026-06-17

Stabilized at **1.0.0**.

### Removed

- **The JavaScript/TypeScript implementation (`@docxengine/core`) and the cross-language conformance/parity harness have been removed.** DocxEngine is now a **Python-only** package shipping the `docxengine-mcp` MCP server (`uvx docxengine-mcp`). The shared `spec/` JSON contract remains the source of truth for the tool schemas, and `conformance/corpus/` (Python tests) plus `conformance/fidelity/` (renderer fidelity) are retained.
- **Repositioned around the MCP server.** The "SDK / library for integrators" framing is dropped — the SDK docs (`docs/sdks/`) and the JS examples (`*/run.mjs`) are removed. The Python package is presented as how you install and run the server; its in-process API stays but is no longer documented as a product surface.
- Removed the Rust/WASM core-unification research doc (the dual-engine future it evaluated no longer applies), and the root Node tooling (pnpm workspace, husky, commitlint, lint-staged, prettier) that only existed to serve the JS package.
- Release is now **PyPI-only** (the npm publish path is gone).

## [1.0.0-alpha.1] - 2026-06-16

First public **alpha**. Phase 3 hardening: the engine now defends itself against hostile input, and the security claims in `SECURITY.md` are backed by code in both engines.

### Security

- **Decompression-bomb defenses** (`doc_too_large`): both engines now bound an untrusted package's cost _before_ paying it — caps on part count, total and per-part uncompressed bytes, and per-part compression ratio (ratio floor 64 KiB). Python checks the zip central directory and decompresses each part through a bounded reader; TypeScript enforces the same caps in an `fflate` pre-decompression `filter`. All caps are tunable via `DOCXENGINE_MAX_PARTS` / `_MAX_TOTAL_BYTES` / `_MAX_PART_BYTES` / `_MAX_COMPRESSION_RATIO` / `_MAX_XML_DEPTH` (spec/algorithms.md §27).
- **Hostile-XML rejection** (`malicious_content`, new error code): XML parts carrying a `<!DOCTYPE`/`<!ENTITY` declaration are refused on first read, closing XXE, external-DTD fetches, and billion-laughs entity expansion — in Python this guards the `ElementTree` parse sites that would otherwise expand internal entities.
- **XML nesting-depth cap**: pathologically nested XML is refused (`doc_too_large`) past `DOCXENGINE_MAX_XML_DEPTH` (default 1000).
- **Path-traversal hardening**: Python `resolve_rel_target` now clamps `..` at the package root (a hostile relationship target can no longer resolve to a name with leading `..`), matching the TypeScript engine.

### Added

- **Adversarial test suites** in both engines (`python/tests/test_adversarial.py`, `js/test/adversarial.test.ts`) plus two cross-implementation conformance cases (`hostile-doctype`, `hostile-zip-bomb`) — 36 conformance cases now pass on py, js, and parity.
- **Large-document performance benchmark** (`bench/perf.py`, `make perf`): wall time and peak Python-heap memory for open/outline/search/read/replace/validate/serialize across document sizes.
- **Cross-renderer fidelity harness** (`conformance/fidelity/run.py`, `make fidelity`) and protocol (`docs/conformance/fidelity.md`): automated structural-fidelity checks everywhere, LibreOffice visual rendering when available, and the documented Word/LibreOffice/Google-Docs manual comparison protocol.
- **Rust/WASM core-unification evaluation** (`docs/research/rust-wasm-core.md`): the ROADMAP Phase 3 v2 decision doc — recommendation is to defer and keep the dual-engine + conformance approach until a measurable trigger fires.

### Changed

- Conformance harness compares error responses by their `error` **code** (the machine contract), masking human-prose `message`/`suggestions`, and accepts an expected error raised at the `docx_open` step (hostile documents are refused at open).

## [0.1.0] - 2026-06-11

### Changed

- **MCP tool surface is now path-based** (was `doc_id`/`docx_save`). Agents pass a file `path` to every tool and edits persist automatically; the 24-tool spec becomes 23 over MCP (`docx_save` dropped). The SDK `call()` surface and the CLI are unchanged.
- _(pre-1.0, breaking)_ Python `Document.find()` now returns a `Paragraph | None` view (match-dict search moved to `Document.search()`).
- _(pre-1.0, breaking)_ JS `Document` instances now own a **private session** instead of the module-level `defaultSession`; pass `{ session }` to share one, or `Document.attach(session, id)` to wrap a handle used with raw `call()`.

### Added

- **Phase 2 surface** in both engines (all 24 tools now live; nothing returns `not_implemented`): tables (`docx_table` create/set_cells/merge/insert/delete row/col), styles (`docx_style`, full `docx_format` with style-definition edits), lists/numbering (`docx_list` with `numbering.xml` creation), sections/headers/footers (`docx_section`), threaded comments with resolve state (`docx_comment`, full five-place OOXML wiring + `commentsExtended`), media (`docx_media` with EMU sizing), fields/TOC (`docx_field`), mustache templates with loops (`docx_template_fill`), Markdown→docx creation (`docx_create`), in-engine md/html conversion (`docx_convert`), and the pluggable render adapter with structural fallback (`docx_render_preview`).
- **File-first MCP server**: the MCP face is now a path-based projection of the `doc_id` contract (`_mcp_facade`, algorithms.md §26). Every tool takes a file `path`; each call opens the file, runs the tool, and validates + atomically saves it back when the edit changed it — no `doc_id`, no separate `docx_save` step (`docx_create`/`docx_template_fill` write immediately). `DOCXENGINE_ROOT` confines paths to a sandbox (`path_denied`). The wire contract, the CLI, and the SDKs are unchanged and keep `doc_id`.
- **MCP Streamable HTTP transport** (`docxengine-mcp --http`): `Mcp-Session-Id` lifecycle with HTTP 410 on expired sessions. Sessions are protocol-only (document state is the filesystem); resources render `docx://{path}/outline` and `/projection` on demand.
- **`export_bytes` / `exportBytes`**: a storage-agnostic bytes-out helper for embedding contexts (browser, serverless) — runs the §8 validation gate, then returns the `.docx` bytes with no filesystem. Not a wire tool.
- **Full-coverage native `Document` API** in both engines: one typed method per tool (`outline`/`read`/`search`/`insert`/`delete`/`edit_paragraph`/`revision`/`comment`/`table`/`style`/`format`/`list`/`section`/`media`/`field`/`validate`/`repair`/`convert`/`render_preview`), `Document.create`/`fill_template`/`attach` constructors, `to_bytes()`, and anchor-scoped `Paragraph` primitives (`replace`/`edit`/`insert_after`/`insert_before`/`delete`).
- **Agent task benchmark** (`bench/`, `make bench`): 10 natural-language tasks with element-level ground-truth checks, driven through the real MCP server with metrics (calls, errors, tokens, Word-repair rate) — 10/10 passing with zero tool errors and zero repair events.
- Conformance suite expanded to 31 cross-implementation cases (tables, styles, lists, sections, comments, media, fields, templates, create, convert parity).
- **Phase 0+1 MVP engines** in parallel Python (`docxengine`, 211 tests) and TypeScript (`@docxengine/core`, 191 tests) implementations: OPC package model with byte-stable round-trip, content-hash anchor index, Markdown projector (`docx_open`/`docx_outline`/`docx_read`/`docx_search`), surgical edits with split-run coalescing (`docx_replace`/`docx_edit_paragraph`/`docx_insert`/`docx_delete`), tracked-change writer + `docx_revision` accept/reject with author/date filters, OOXML validator/repair with the always-on save gate (`docx_validate`/`docx_repair`/`docx_save`), `call()` dispatcher, and OpenAI/Anthropic adapters.
- **MCP stdio server** (`docxengine-mcp`): dependency-free JSON-RPC 2.0, serves all 24 tool schemas, dispatches to the shared core.
- **Tool contract** in `spec/`: 24 JSON Schemas + machine-readable error catalog + `spec/algorithms.md`, the language-agnostic algorithm spec both implementations follow.
- **Conformance suite**: synthetic corpus generator (6 fixture documents), 15 cross-implementation cases, and a harness proving Python↔TS parity (`make conformance` — all 15 pass on py, js, and parity).
- **Examples**: redline-review and bulk-rebrand (runnable, both SDKs), annotated agent-loop transcript, template-to-pdf (Phase 2 preview).
- Design-first repository scaffold: architecture, roadmap, full documentation lanes, contribution tooling (commitlint, husky, lint-staged, prettier), CI/release/CodeQL workflows, and community templates.

[unreleased]: https://github.com/ruwadgroup/docxengine/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/ruwadgroup/docxengine/compare/v1.0.0-alpha.1...v1.0.0
[1.0.0-alpha.1]: https://github.com/ruwadgroup/docxengine/compare/v0.1.0...v1.0.0-alpha.1
[0.1.0]: https://github.com/ruwadgroup/docxengine/releases/tag/v0.1.0
