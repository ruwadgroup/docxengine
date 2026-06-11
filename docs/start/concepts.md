# Concepts

Five ideas explain the whole engine.

## 1. Direct-XML editing (fidelity)

A `.docx` is a ZIP of XML parts. DocxEngine opens that package, patches the XML in place, and rezips — it never round-trips through a library object model. This is the only approach that preserves tracked changes, comments, footnotes, and everything else Word users care about. The cost is that OOXML is full of traps ([OOXML pitfalls](../core/ooxml-pitfalls.md)); the engine's job is to absorb those traps so agents never see them.

## 2. Anchors (stable addressing)

Every paragraph is addressed as `P{index}#{hash}` — e.g., `P12#a7b2`. The index is an ordering hint; the **hash over the paragraph's normalized text is the integrity check**. Edits validate the hash first and fail with `anchor_stale` if the content changed underneath. Every edit returns fresh anchors, so a multi-step agent never works from stale state. Why not Word's own `w14:paraId`? Because it isn't guaranteed stable across saves and is missing from non-Word documents — full evidence in [Anchors](../core/anchors.md).

## 3. Projection (token economy)

Agents never read raw OOXML. Reads return a Markdown-like projection: anchors, heading levels, list markers, comment/revision flags — and nothing else. Pagination is built in: `docx_outline` gives the cheap map, `docx_read` zooms into a range or window, and `concise`/`detailed` controls run-level formatting. No response exceeds ~25k tokens. Format spec: [Projection](../core/projection.md).

## 4. The validation gate (no silent repair)

Word silently "repairs" (rewrites) files with broken internals — duplicate IDs, orphaned relationships, dangling footnotes — and that rewrite can destroy content. DocxEngine validates the full package before every save and refuses or auto-repairs anything that would trigger repair. A clean `docx_validate` is the engine's contract that Word will open the file untouched. Details: [Validation](../core/validation.md).

## 5. Verification (seeing the result)

Agents can't see a rendered page, and some things (page numbers, TOC) only exist at render time. `docx_render_preview` renders pages to images via LibreOffice (when installed) so an agent can self-check layout-sensitive edits — the plan→edit→check loop that makes agent editing reliable. Details: [Render adapter](../core/render-adapter.md).

## Glossary shortcuts

- **doc_id** — handle for an open document held server/process-side; edits accumulate against it until `docx_save`.
- **story** — an independent text flow: body, header, footer, footnotes, comments. Each lives in its own XML part.
- **run** — a span of uniformly-formatted text inside a paragraph. Word fragments these arbitrarily; the engine coalesces them.
- **redline** — a tracked change (`w:ins`/`w:del`) attributed to an author with a timestamp.

Full glossary: [reference/glossary.md](../reference/glossary.md).
