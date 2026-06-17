# DocxEngine Architecture

This document is the long-form design reference: the layering, the document representation agents see, the tool surface (the agent-computer interface), the MCP server design, and the packaging strategy. The research and prior art behind these decisions is summarized in [docs/research/prior-art.md](docs/research/prior-art.md).

## Table of contents

- [Design constraints](#design-constraints)
- [Layering — one core, two faces](#layering--one-core-two-faces)
- [Stable addressing: content-hash anchors](#stable-addressing-content-hash-anchors)
- [The document representation agents see](#the-document-representation-agents-see)
- [The tool surface (ACI)](#the-tool-surface-aci)
- [Error design](#error-design)
- [Representative flows](#representative-flows)
- [MCP server design](#mcp-server-design)
- [Framework-agnostic packaging](#framework-agnostic-packaging)
- [Validation and the repair gate](#validation-and-the-repair-gate)
- [Rendering and verification](#rendering-and-verification)
- [Invariants](#invariants)

## Design constraints

Four facts shape everything below:

1. **Only direct-XML editing preserves fidelity.** Tracked changes, comments, and footnotes are dropped or mangled by every library wrapper (python-docx, docx-js, docxtemplater, Pandoc round-trips). The engine must unzip the OPC package, patch the XML, and rezip — the approach Anthropic's docx skill and SecurityRonin/docx-mcp validated in production.
2. **OOXML is hostile to naive editing for structural reasons.** Runs split arbitrarily (the split-run problem), `w:rsid*` attributes fragment runs without semantic meaning, style resolution is a six-layer cascade, list numbering lives in a separate part behind indirection, and a single logical edit (e.g., adding a comment) touches body XML, a comments part, relationships, content-types, and IDs. Get any of it wrong and Word silently "repairs" (rewrites) the file. See [docs/core/ooxml-pitfalls.md](docs/core/ooxml-pitfalls.md).
3. **`w14:paraId` cannot be the durable address.** It is a Word-2010+ extension absent from docs written by other tools, generated randomly by Word, and _not spec-guaranteed stable across saves_ (documented regenerations in Open-XML-SDK #925). See [docs/core/anchors.md](docs/core/anchors.md).
4. **Agents need a designed interface, not a wrapped API.** SWE-agent (NeurIPS 2024) showed interface design — simple, consolidated, guarded actions with feedback — outperforms raw access. Anthropic's tool guidance adds token-economy rules: high-leverage namespaced tools, human-readable context, ~25k-token response cap, `concise`/`detailed` formats, just-in-time detail loading.

## Layering — one core, two faces

```
┌──────────────────────────────────────────────────────────────┐
│  Integration faces (thin)                                      │
│  1. MCP server (stdio + streamable-HTTP)                       │
│  2. Python package  (docxengine)   — JSON-in/JSON-out + native │
│     + OpenAI/Anthropic function-calling adapters (thin)        │
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

The faces are deliberately thin: they translate transport/registration formats and nothing else. All behavior lives in the core, which is deterministic — the same tool call against the same document bytes produces the same output bytes, with no model in the loop.

The Python engine is pure-stdlib with zero runtime dependencies (`mcp` is an optional extra), so a plain `pip install docxengine` works with no native or WASM toolchain — the decisive factor for adoptability in locked-down private systems. The engine is kept honest by a **conformance corpus** — the same input docx plus the same tool call must produce byte-equivalent-after-normalization output across runs ([conformance/](conformance/)).

## Stable addressing: content-hash anchors

The hardest design problem. Requirements: an address an agent can hold across multiple tool calls that (a) survives saves by arbitrary producers, (b) detects staleness when content changes, and (c) costs few tokens.

**Anchor = `P{index}#{hash}`** — e.g., `P12#a7b2`:

- `index` is the paragraph's ordinal position: a human/agent-readable ordering hint, _not_ trusted for targeting.
- `hash` is the first 4–6 hex chars of a hash over the paragraph's **normalized text** (optionally seeded by `w14:paraId` when present). The hash is the integrity check.
- Every edit tool **validates the hash before touching the paragraph** and rejects with a structured `anchor_stale` error on mismatch.
- Every edit **returns the new anchor(s)** (re-anchoring), so an agent never needs to re-list paragraphs mid-batch.
- `w14:paraId` is surfaced only as a secondary hint, never as the durable address.

Tables get `T{index}` anchors with positional context (`T4@after:P12#e7f8`). Anchors _will_ go stale when content changes — the re-anchoring return values and `anchor_stale` errors are essential, not optional.

## The document representation agents see

Agents never see raw OOXML by default. The read projection is a Markdown-like view annotated with anchors and only the formatting that matters:

```
[P1#a7b2  H1]            Master Services Agreement
[P2#f3c1]                This Agreement is entered into as of {{EffectiveDate}}...
[P3#b2c4  H2]            1. Definitions
[P4#d4e5]                "Confidential Information" means... [comment:C1 by J.Doe]
[T1  3×4 @after:P5]      | Term | Value | ... |
[P12#e7f8  List:ol L1]   First obligation
```

- **Salient formatting only**: heading level, list type/level, bold/italic when semantically relevant, comment/revision markers. Hidden: rsids, default fonts, `proofErr`, and everything else that is noise to an agent.
- **Just-in-time depth**: `docx_outline` returns the cheap headings map; `docx_read` takes a range or an anchor+window; responses default to `concise`, with `detailed` returning resolved run-level formatting.
- **Numbering and styles are resolved**, not raw: the projector compiles the numbering indirection ("1.", "(a)") and walks the style cascade so the agent sees effective values.

Full format specification: [docs/core/projection.md](docs/core/projection.md).

## The tool surface (ACI)

~16 namespaced, high-leverage tools across 5 groups, all sharing a JSON-in/JSON-out contract published in [`spec/`](spec/). Full per-tool reference: [docs/tools/index.md](docs/tools/index.md).

**Read/Navigate**

- `docx_open(path|bytes) → {doc_id, summary, n_paragraphs, has_tracked_changes, has_comments}`
- `docx_outline(doc_id) → headings tree with anchors`
- `docx_read(doc_id, {anchor?, range?, window?, format?: concise|detailed}) → Markdown projection`
- `docx_search(doc_id, {query, regex?, scope?}) → [{anchor, snippet, context}]`

**Surgical edit** — each returns new anchor(s); each accepts `track_changes?: bool` + `author?`

- `docx_replace(doc_id, {anchor|search, old, new, all?}) → {new_anchor, n_replaced}` — handles split-run coalescing internally
- `docx_edit_paragraph(doc_id, {anchor, text}) → {new_anchor}` — full-paragraph rewrite with auto word-level diff (clean redline when tracking is on)
- `docx_insert(doc_id, {after|before: anchor, content, style?}) → {new_anchor}`
- `docx_delete(doc_id, {anchor|range}) → {ok}`
- `docx_format(doc_id, {anchor|range|style_selector, props}) → {affected}`
- `docx_table(doc_id, {op: create|set_cells|merge|insert_row|…})`

**Structure/assets**

- `docx_style(doc_id, {op: list|define|apply, …})`, `docx_section(…)`, `docx_list(…)`
- `docx_comment(doc_id, {op: add|reply|resolve|list|delete, anchor?, text?, author?})`
- `docx_revision(doc_id, {op: list|accept|reject|accept_all|reject_all, filter?: {author, date}})`
- `docx_media(doc_id, {op: insert|extract|replace, …})`, `docx_field(…)` — TOC/page-number insertion as field codes
- `docx_template_fill(doc_id|template, {data, syntax?: mustache})`

**Create/Convert/Verify**

- `docx_create({content_md|spec}) → {doc_id}`
- `docx_convert(doc_id, {to: md|html|pdf|png})` — md/html in-engine; pdf/png via render adapter
- `docx_validate(doc_id) → {valid, issues: [{severity, part, message, fix_hint}]}`
- `docx_repair(doc_id) → {fixed: […], remaining: […]}`
- `docx_render_preview(doc_id, {pages?}) → image refs`
- `docx_save(doc_id, {path})`

Inputs are **forgiving**: accept either anchor or search text; tolerate missing `author` (default from environment); accept Markdown or plain text for content. Operations are **idempotent where possible**: `replace` with `all: true` is re-runnable; `accept_all` no-ops when no revisions remain.

## Error design

Every error is structured and corrective — it tells the agent what went wrong _and what to do next_:

```json
{
  "error": "anchor_stale",
  "message": "P12#a7b2 no longer matches (content changed). Re-read P12 or search.",
  "suggestions": ["docx_read(window:P12)"]
}
```

The error-code catalog lives in [docs/reference/error-codes.md](docs/reference/error-codes.md). This is the SWE-agent "guarded action" principle: the engine refuses invalid edits with feedback rather than corrupting the document.

## Representative flows

**"Change all H2 headings to blue"** — the engine edits the _style definition_ in `styles.xml`, not 7 direct overrides. High-leverage, fidelity-preserving, idempotent:

```json
→ docx_format {"doc_id":"d1","style_selector":{"style":"Heading 2"},"props":{"color":"#1F4E79"}}
← {"affected":7,"anchors":["P3#b2c4","P9#1a2b","…"],"note":"Applied to style 'Heading 2' definition; 7 paragraphs use it."}
```

**"Accept all tracked changes from reviewer Jane Doe"**:

```json
→ docx_revision {"doc_id":"d1","op":"accept","filter":{"author":"Jane Doe"}}
← {"accepted":12,"remaining_by_author":{"Bob":3},"note":"Resolved <w:ins>/<w:del> for Jane Doe; Bob's 3 revisions untouched."}
```

**"Insert a table after paragraph 12"**:

```json
→ docx_table {"doc_id":"d1","op":"create","after":"P12#e7f8","rows":3,"cols":3,
              "data":[["Item","Qty","Price"],["Widget","10","$5"],["Gadget","4","$9"]],"header":true}
← {"new_anchor":"T4@after:P12#e7f8","note":"3×3 table inserted; header row styled with 'Table Grid'."}
```

**"Fill this template with data"**:

```json
→ docx_template_fill {"template":"msa.docx","data":{"EffectiveDate":"2026-07-01","Client":"GlobalTech",
   "obligations":[{"text":"Deliver Q3 report"},{"text":"Maintain SLA"}]},"syntax":"mustache"}
← {"doc_id":"d2","filled":4,"loops_expanded":{"obligations":2},"unfilled":[],"note":"All placeholders resolved."}
```

## MCP server design

- **Transports**: **stdio** (default for Claude Desktop and local clients — run the binary, no auth surface) and **Streamable HTTP** (the current MCP standard since the 2025-03-26 spec; single endpoint, `Mcp-Session-Id` header for stateful sessions, gateway/load-balancer friendly). Legacy HTTP+SSE is not supported.
- **State**: open documents are held server-side keyed by `doc_id` (in-memory OPC model + anchor index) so multi-step edits don't re-parse. Over Streamable HTTP, `doc_id` lifecycle is bound to the `Mcp-Session-Id`; for horizontal scaling, the doc store is externalized (object store + sticky sessions).
- **Resources**: open documents and generated previews are exposed as MCP resources (`docx://d1/outline`, `docx://d1/preview/page-1.png`) so clients can surface them without a tool call.
- **Large documents**: the whole document is never returned; reads paginate under the ~25k-token response cap; `convert`/`render` outputs stream as resource links, not inline blobs; edits batch with a single validate+save at the end.
- **Concurrency**: per-`doc_id` write lock; reads are lock-free against an immutable snapshot; idempotent edit semantics make retries safe.

Details: [docs/mcp/server.md](docs/mcp/server.md) and [docs/mcp/state-and-scaling.md](docs/mcp/state-and-scaling.md).

## Framework-agnostic packaging

- **Contract**: every tool is a pure function `fn(args_json) → result_json` with a published JSON Schema. The same schema feeds MCP `tools/list`, OpenAI function-calling `tools`, and Anthropic tool definitions. [`spec/`](spec/) is the source of truth.
- **Python** (`docxengine`): native objects (`Document`, `Paragraph`) for power users plus a `call(tool_name, args_dict) → dict` dispatcher. Ships `docxengine.openai_tools()` and `docxengine.anthropic_tools()` (the function-calling schema lists) as thin adapters.
- **Adapters are ≤ a few lines each** and never contain business logic — they only translate schema/registration formats. Any other framework consumes the published schemas directly; a custom orchestrator just calls `call(name, args)`.

## Validation and the repair gate

Validation is a first-class, always-on gate — the single highest-leverage reliability feature:

- Checks: internal ID uniqueness, orphaned relationships, dangling footnotes/comments, content-type completeness, cross-reference integrity.
- **Validate before every save**; refuse or auto-repair edits that would trigger Word's "repair" prompt.
- `docx_repair` fixes what it safely can and clearly reports what remains.

## Rendering and verification

Rendering is an **optional, pluggable adapter** — never a hard core dependency, so the engine stays installable in locked-down environments:

- Default: shell out to LibreOffice headless for PDF/PNG previews when available.
- Graceful degradation: a "structural preview" (resolved Markdown + computed layout estimate) when no renderer is present.
- Caveats are documented and surfaced to agents: LibreOffice and Word do not render identically (font substitution); fields/TOC/page numbers only resolve at render time.
- Threshold to invest in a faster renderer: when preview latency (~3s+/doc for LibreOffice) dominates agent-loop cost, evaluate a persistent LibreOffice server pool or a native layout estimator.

## Invariants

These hold across all phases and all faces; PRs that break one are rejected (see [CONTRIBUTING.md](CONTRIBUTING.md)):

1. **Fidelity**: open→save of an untouched document is byte-stable modulo normalization; tracked changes, comments, footnotes, and media survive every edit path.
2. **No silent repair**: a document that validates clean before save must open in Word without a "repair" prompt.
3. **Hash-guarded edits**: no edit lands on a paragraph whose anchor hash fails validation.
4. **Token economy**: no tool response exceeds ~25k tokens; raw OOXML is never returned by default.
5. **Determinism**: the core contains no LLM and no nondeterministic output for identical inputs; the same tool call against the same bytes produces byte-equivalent-after-normalization output for every conformance case.
6. **Thin faces**: MCP/adapters translate formats only; behavior lives in the core.
