# Render adapter

Agents lack the visual understanding to operate GUI applications — and some document truth (page numbers, TOC entries, line breaks) only exists at render time. The render adapter closes the loop: it turns the document into PDF/PNG so an agent can _check its own work_.

## Design: optional and pluggable

Rendering is **never a hard dependency** of the core. The adapter interface has two implementations:

1. **LibreOffice headless** (default when detected) — `soffice --headless --convert-to pdf` for `docx_convert` and page images for `docx_render_preview`.
2. **Structural preview** (always available) — resolved Markdown projection plus a computed layout estimate, used when no renderer is installed. Keeps the core fully functional in locked-down environments.

Detection is automatic; deployments can pin or disable an adapter explicitly.

## Tool surface

```json
→ docx_render_preview {"doc_id":"d1","pages":[1,2]}
← {"pages":[{"page":1,"image":"docx://d1/preview/page-1.png"},
            {"page":2,"image":"docx://d1/preview/page-2.png"}],
   "renderer":"libreoffice 24.8","note":"Fonts substituted: Calibri→Carlito"}
```

Over MCP, previews are returned as **resource links**, not inline blobs — clients fetch what they display.

## Caveats (surfaced to agents, not hidden)

- **LibreOffice ≠ Word.** Font substitution and layout engines differ; previews are approximate for pixel-exact validation. Installing metric-compatible fonts (Carlito for Calibri, Caladea for Cambria) reduces but doesn't eliminate drift. Responses name the renderer and any substitutions so agents can calibrate trust.
- **Latency.** Each LibreOffice conversion boots much of the office runtime (~3s+/doc, much worse for huge documents). Previews are for _verification checkpoints_, not every edit. If preview latency becomes the dominant agent-loop cost, the roadmap threshold triggers evaluation of a persistent LibreOffice server pool or a native layout estimator.
- **Fields resolve only here.** `docx_field` inserts/updates field _codes_; the computed values (page N of M, TOC text) appear only in rendered output. Agents are told this explicitly to prevent hallucinated page numbers.

## Security

The renderer processes untrusted documents and is treated as the highest-risk component: deployments should sandbox it (container, seccomp, no network). See [SECURITY.md](../../SECURITY.md).
