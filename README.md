<div align="center">

# DocxEngine

**Surgical, fidelity-preserving DOCX editing for AI agents — and for you.**

One deterministic core that edits OOXML directly (unzip → patch XML → rezip), exposed through three thin faces: an **MCP server**, a **Python package** (`docxengine`), and a **JS/TS package** (`@docxengine/core`). Agents see a token-efficient, Markdown-like projection with content-hash-anchored paragraph IDs — never raw XML.

[![License: Apache-2.0](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![CI](https://github.com/ruwadgroup/docxengine/actions/workflows/ci.yml/badge.svg)](https://github.com/ruwadgroup/docxengine/actions/workflows/ci.yml)
[![Python ≥3.12](https://img.shields.io/badge/python-%E2%89%A53.12-blue)](python/)
[![Node ≥22](https://img.shields.io/badge/node-%E2%89%A522-brightgreen)](js/)
[![MCP](https://img.shields.io/badge/MCP-stdio%20%2B%20streamable--http-8A2BE2)](docs/mcp/server.md)
[![Conventional Commits](https://img.shields.io/badge/Conventional%20Commits-1.0.0-yellow.svg)](https://conventionalcommits.org)
[![Status: pre-alpha](https://img.shields.io/badge/status-pre--alpha-orange)](ROADMAP.md)

[Quickstart](docs/start/quickstart.md) · [Concepts](docs/start/concepts.md) · [Tool reference](docs/tools/index.md) · [Architecture](ARCHITECTURE.md) · [MCP server](docs/mcp/server.md) · [Docs](docs/README.md) · [Roadmap](ROADMAP.md)

</div>

---

> **Status: pre-alpha — Phases 0–2 implemented.** All 24 tools work today in both Python and TypeScript: read/search/edit/redlines, tables, styles, lists, sections, comments, media, fields, templates, create, convert, and the render adapter — with 31/31 cross-implementation conformance cases passing and the MCP server speaking both stdio and Streamable HTTP. Packages are not yet published to PyPI/npm — install from source for now. Remaining before 0.1: baseline benchmark comparisons and Phase 3 hardening ([ROADMAP.md](ROADMAP.md)).

## Table of contents

- [Overview](#overview)
- [Features](#features)
- [Why DocxEngine](#why-docxengine)
- [What DocxEngine is not](#what-docxengine-is-not)
- [Architecture](#architecture)
- [The agent view](#the-agent-view)
- [Getting started](#getting-started)
- [Documentation](#documentation)
- [Repository layout](#repository-layout)
- [Roadmap & status](#roadmap--status)
- [Contributing](#contributing)
- [Community & support](#community--support)
- [License](#license)

## Overview

Every mainstream DOCX library has a disqualifying gap for agent use: python-docx has no tracked-changes support (open since 2016), docx-js is generation-focused, docxtemplater is template-bound, Pandoc round-trips are lossy, and LibreOffice headless is heavyweight. The only approach that preserves **tracked changes, comments, and footnotes** is editing the OOXML directly — the same strategy Anthropic's docx skill and the strongest MCP servers converged on.

DocxEngine packages that strategy as a reusable engine:

- A **deterministic core** (no LLM inside) that models the OPC/ZIP package, patches the XML DOM, coalesces split runs, writes real `w:ins`/`w:del` redlines, and validates every edit against OOXML before saving — so Word never silently "repairs" your file.
- An **agent-computer interface** of ~16 high-leverage, namespaced tools (`docx_search`, `docx_replace`, `docx_revision`, …) with structured, corrective errors and idempotent semantics.
- **Stable addressing** via content-hash anchors (`P12#a7b2`) — because `w14:paraId` is not spec-guaranteed stable across Word save cycles and is absent from docs written by non-Word tools.
- A **verification loop**: render-to-PDF/PNG previews (via a pluggable LibreOffice adapter) so agents can self-check their edits.

## Features

- **Fidelity-preserving surgical edits** — replace, insert, delete, and rewrite paragraphs in arbitrary existing documents without disturbing tracked changes, comments, footnotes, styles, or media.
- **Real redlines** — first-class tracked-change writing (`track_changes: true, author: "..."`), plus accept/reject filtered by author or date.
- **Token-efficient reading** — outline first, then paginated, Markdown-like projections with only salient formatting; raw OOXML is never shown by default.
- **Hash-anchored addressing** — every paragraph gets a `P{index}#{hash}` anchor validated before each edit; edits return fresh anchors so agents never re-list mid-batch.
- **Always-on validation gate** — ID uniqueness, orphaned relationships, dangling footnotes, and content-type errors are caught before save, with auto-repair where safe.
- **Comments, tables, styles, sections, lists, media, fields, templates** — the full capability surface is implemented: threaded comments with resolve state, style-definition edits, mustache template merge with loops, Markdown↔docx conversion, and field-code insertion.
- **Triple distribution** — MCP server (stdio + Streamable HTTP), `pip install docxengine`, `npm install @docxengine/core`; the published JSON Schemas plug into any framework, with a thin OpenAI function-calling adapter included.
- **One conformance-tested contract** — the Python and TypeScript implementations are kept honest by a shared JSON tool contract and a cross-implementation conformance corpus.

## Why DocxEngine

Agents are a new class of end-user, and tools must be designed for them rather than wrapped from existing APIs (SWE-agent, NeurIPS 2024). Raw OOXML is distracting context; agents can't "see" the rendered page; and naive find-and-replace fails because Word fragments text across run boundaries. DocxEngine applies the resulting design principles end to end:

| Principle                        | How DocxEngine applies it                                                          |
| -------------------------------- | ---------------------------------------------------------------------------------- |
| Simple, few, high-leverage tools | ~16 namespaced tools across 5 groups, not a 1:1 API wrapper                        |
| Guarded actions                  | every edit is hash-validated and OOXML-validated before it lands                   |
| Token economy                    | outline → windowed reads, `concise`/`detailed` formats, ~25k-token response cap    |
| Feedback loops                   | structured corrective errors + render-based visual self-check                      |
| Determinism                      | the core contains no LLM; the same call on the same document yields the same bytes |

## What DocxEngine is not

- **Not a renderer.** Fields, TOC entries, and page numbers only materialize when Word or LibreOffice renders; the engine inserts and updates _field codes_ and tells agents so explicitly.
- **Not a template DSL.** `docx_template_fill` covers mustache-style merge with loops and conditions, but DocxEngine's center of gravity is _arbitrary surgical edits of existing documents_.
- **Not a python-docx/docx-js wrapper.** Those libraries drop the document features this project exists to preserve; they appear at most in narrow create paths.
- **Not Word automation.** No COM, no Office.js host, no GUI — server-side and offline by design.

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│  Integration faces (thin)                                      │
│  1. MCP server (stdio + streamable-HTTP)                       │
│  2. Python package  (docxengine)   — JSON-in/JSON-out + native │
│  3. JS/TS package   (@docxengine)  — JSON-in/JSON-out + native │
│     + OpenAI function-calling adapter (thin)                   │
├──────────────────────────────────────────────────────────────┤
│  Core engine (deterministic, no LLM)                           │
│   • OPC/ZIP package model      • Style cascade resolver        │
│   • XML DOM patcher            • Numbering resolver            │
│   • Run-coalescing find/replace• Tracked-change writer         │
│   • Content-hash anchor index  • Comment/footnote part manager │
│   • Markdown projector (read)  • OOXML validator + repairer    │
│   • Render adapter (LibreOffice/Word) for verification         │
└──────────────────────────────────────────────────────────────┘
```

v1 ships **parallel Python and TypeScript implementations** against a shared JSON contract ([`spec/`](spec/)) and a shared conformance corpus — pure-`pip` and pure-`npm` installs with zero native toolchain. A Rust/WASM core unification is a v2 evaluation. The full reasoning, including the addressing design and tool surface, is in [ARCHITECTURE.md](ARCHITECTURE.md).

## The agent view

Agents never see raw OOXML. Reads return a Markdown-like projection annotated with stable anchors and only the formatting that matters:

```
[P1#a7b2  H1]            Master Services Agreement
[P2#f3c1]                This Agreement is entered into as of {{EffectiveDate}}...
[P3#b2c4  H2]            1. Definitions
[P4#d4e5]                "Confidential Information" means... [comment:C1 by J.Doe]
[T1  3×4 @after:P5]      | Term | Value | ... |
[P12#e7f8  List:ol L1]   First obligation
```

A typical edit flow:

```json
→ docx_revision {"doc_id":"d1","op":"accept","filter":{"author":"Jane Doe"}}
← {"accepted":12,"remaining_by_author":{"Bob":3},"note":"Resolved <w:ins>/<w:del> for Jane Doe; Bob's 3 revisions untouched."}
```

See [Concepts](docs/start/concepts.md) for anchors, projection, and the validation gate, and the [tool reference](docs/tools/index.md) for all tools.

## Getting started

> Pre-alpha: not yet published to PyPI/npm — install from source. All 24 tools work today; see the [Quickstart](docs/start/quickstart.md) and [examples/](examples/).

```bash
git clone https://github.com/ruwadgroup/docxengine.git && cd docxengine

# Python (+ the MCP server entry point)
pip install -e python

# JS/TS
pnpm install && pnpm --dir js build

# MCP (Claude Desktop / any MCP client) — stdio
docxengine-mcp

# Claude Code
claude mcp add docx -- docxengine-mcp
```

Over MCP the engine is **file-first**: tools take a file `path` and every edit is validated and saved back automatically — no handles to track, no save step. The Python/JS packages keep an in-memory `doc_id`/bytes handle (the right fit for embedding, including browser JS); see the [SDK docs](docs/sdks/python.md).

## Documentation

| Lane                                      | What you'll find                                                            |
| ----------------------------------------- | --------------------------------------------------------------------------- |
| [Start](docs/start/quickstart.md)         | Installation, quickstart flows, core concepts                               |
| [Core](docs/core/ooxml-pitfalls.md)       | OOXML pitfalls, anchors, projection, tracked changes, validation, rendering |
| [Tools](docs/tools/index.md)              | The full agent-computer interface, group by group, plus error design        |
| [MCP](docs/mcp/server.md)                 | Transports, resources, session state, scaling                               |
| [SDKs](docs/sdks/python.md)               | Python & JS packages, framework adapters                                    |
| [Conformance](docs/conformance/corpus.md) | Round-trip fidelity corpus, agent task benchmark                            |
| [Research](docs/research/prior-art.md)    | Prior art, key findings, competitive landscape                              |
| [Reference](docs/reference/glossary.md)   | Glossary, tool schemas, error codes                                         |

Start at [docs/README.md](docs/README.md).

## Repository layout

```
docxengine/
├── spec/            # Language-agnostic JSON tool contract (the source of truth)
├── python/          # docxengine — Python implementation (pip)
├── js/              # @docxengine/core — TypeScript implementation (npm)
├── conformance/     # Shared corpus + cross-implementation harness
├── examples/        # End-to-end agent flows
├── docs/            # Design docs, tool reference, guides
└── .github/         # CI, release, security scanning, templates
```

## Roadmap & status

**Phases 0–2 complete; current phase: 3 — Hardening.** All 24 tools are implemented and conformance-tested in both languages: 455 Python tests, 342 TS tests, 31/31 cross-implementation parity cases, and a 10-task agent benchmark passing end-to-end over the file-first MCP server with zero tool errors and zero Word-repair events. Remaining: benchmark comparisons against the python-docx and raw-XML baselines, fuzzing, large-document streaming, and cross-renderer fidelity. Full plan with decision thresholds: [ROADMAP.md](ROADMAP.md).

## Contributing

Contributions are welcome — especially conformance corpus documents, OOXML edge-case reports, and benchmark tasks. Read [CONTRIBUTING.md](CONTRIBUTING.md) for the ground rules (the invariants), development setup, and commit conventions ([Conventional Commits](https://www.conventionalcommits.org) with enforced scopes).

## Community & support

- **Bugs & features** — [GitHub issues](https://github.com/ruwadgroup/docxengine/issues) (structured templates)
- **Questions & ideas** — [GitHub Discussions](https://github.com/ruwadgroup/docxengine/discussions)
- **Security reports** — privately, per [SECURITY.md](SECURITY.md)
- **Governance** — [GOVERNANCE.md](GOVERNANCE.md)

## License

[Apache-2.0](LICENSE). DocxEngine optionally shells out to external renderers/converters under their own licenses — see [LICENSING.md](LICENSING.md).
