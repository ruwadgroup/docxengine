# Changelog

All notable changes to DocxEngine will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **Phase 2 surface** in both engines (all 24 tools now live; nothing returns `not_implemented`): tables (`docx_table` create/set_cells/merge/insert/delete row/col), styles (`docx_style`, full `docx_format` with style-definition edits), lists/numbering (`docx_list` with `numbering.xml` creation), sections/headers/footers (`docx_section`), threaded comments with resolve state (`docx_comment`, full five-place OOXML wiring + `commentsExtended`), media (`docx_media` with EMU sizing), fields/TOC (`docx_field`), mustache templates with loops (`docx_template_fill`), Markdown→docx creation (`docx_create`), in-engine md/html conversion (`docx_convert`), and the pluggable render adapter with structural fallback (`docx_render_preview`).
- **MCP Streamable HTTP transport** (`docxengine-mcp --http`): `Mcp-Session-Id` lifecycle with per-session document stores, HTTP 410 on expired sessions, and MCP resources (`docx://{doc_id}/outline`, `/projection`) on both transports.
- **Agent task benchmark** (`bench/`, `make bench`): 10 natural-language tasks with element-level ground-truth checks, driven through the real MCP server with metrics (calls, errors, tokens, Word-repair rate) — 10/10 passing with zero tool errors and zero repair events.
- Conformance suite expanded to 31 cross-implementation cases (tables, styles, lists, sections, comments, media, fields, templates, create, convert parity).
- **Phase 0+1 MVP engines** in parallel Python (`docxengine`, 211 tests) and TypeScript (`@docxengine/core`, 191 tests) implementations: OPC package model with byte-stable round-trip, content-hash anchor index, Markdown projector (`docx_open`/`docx_outline`/`docx_read`/`docx_search`), surgical edits with split-run coalescing (`docx_replace`/`docx_edit_paragraph`/`docx_insert`/`docx_delete`), tracked-change writer + `docx_revision` accept/reject with author/date filters, OOXML validator/repair with the always-on save gate (`docx_validate`/`docx_repair`/`docx_save`), `call()` dispatcher, and OpenAI/Anthropic adapters.
- **MCP stdio server** (`docxengine-mcp`): dependency-free JSON-RPC 2.0, serves all 24 tool schemas, dispatches to the shared core.
- **Tool contract** in `spec/`: 24 JSON Schemas + machine-readable error catalog + `spec/algorithms.md`, the language-agnostic algorithm spec both implementations follow.
- **Conformance suite**: synthetic corpus generator (6 fixture documents), 15 cross-implementation cases, and a harness proving Python↔TS parity (`make conformance` — all 15 pass on py, js, and parity).
- **Examples**: redline-review and bulk-rebrand (runnable, both SDKs), annotated agent-loop transcript, template-to-pdf (Phase 2 preview).
- Design-first repository scaffold: architecture, roadmap, full documentation lanes, contribution tooling (commitlint, husky, lint-staged, prettier), CI/release/CodeQL workflows, and community templates.

[unreleased]: https://github.com/ruwadgroup/docxengine/commits/main
