# Render adapter

Agents lack the visual understanding to operate GUI applications — and some document truth (page numbers, TOC entries, line breaks) only exists at render time. The render adapter closes the loop: it turns the document into PDF/PNG so an agent can _check its own work_.

## Design: zero-install by default, pluggable, never a hard dependency

Rendering works out of the box with **no manual LibreOffice install**. The adapter has two implementations:

1. **LibreOffice headless** (default) - `soffice --headless --convert-to pdf` for `docx_convert` and page images for `docx_render_preview`. If no `soffice` is already installed, docxengine downloads an official LibreOffice build on first render, verifies it, caches it per-user, and uses it (see below).
2. **Structural preview** (always available) - resolved Markdown projection plus a computed layout estimate, used when auto-fetch is disabled, unsupported, or offline. Keeps the core fully functional in locked-down environments.

The renderer is still **never a hard dependency** of the core: `pip install docxengine` pulls no binaries, and md/html conversion needs nothing. LibreOffice is fetched lazily, only when a pdf/png render is requested and none is present.

### Automatic provisioning

Resolution order for `soffice`: `DOCXENGINE_SOFFICE`, then `soffice` on `PATH`, then platform defaults, then a previously cached auto-fetched build, then **download on demand**. The download comes over HTTPS from The Document Foundation (`download.documentfoundation.org`) and is **SHA-256-verified against the publisher's own `.sha256` sidecar** before anything is extracted or executed. Nothing is fetched when a local `soffice` already exists or auto-fetch is off. Auto-fetch supports macOS (`.dmg`) and Linux x86-64 (`.deb` tarball); other platforms fall back to detection with an actionable message.

| Env var                         | Purpose                                                             |
| ------------------------------- | ------------------------------------------------------------------ |
| `DOCXENGINE_SOFFICE`            | Explicit `soffice` path (skips detection and auto-fetch).           |
| `DOCXENGINE_AUTO_FETCH_SOFFICE` | Set `0`/`false`/`off` to disable auto-fetch (structural fallback).  |
| `DOCXENGINE_SOFFICE_CACHE`      | Cache dir (default `$XDG_CACHE_HOME/docxengine` or `~/.cache/docxengine`). |
| `DOCXENGINE_SOFFICE_VERSION`    | Pin a LibreOffice version instead of resolving the latest stable.   |
| `DOCXENGINE_SOFFICE_MIRROR`     | Base URL of a mirror of the TDF `stable` tree (air-gapped/custom).  |

Deployments can pin or disable the adapter explicitly via these knobs.

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
