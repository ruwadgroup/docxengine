# docxengine (Python)

The Python implementation of [DocxEngine](../README.md): deterministic OOXML editing with tracked changes, hash-anchored addressing, and a token-efficient agent view. Ships the reference MCP server (`docxengine-mcp`).

```bash
pip install docxengine
```

Two surfaces: the **file-first MCP server** (`docxengine-mcp` — every tool takes a path, edits save automatically; `claude mcp add docx -- docxengine-mcp`) and the **storage-agnostic library** (`call()` + `Document`, in-memory `doc_id`/bytes handles you persist with `save(path)` or `to_bytes()`). See the usage guide.

- Usage: [docs/sdks/python.md](../docs/sdks/python.md)
- Contract: [spec/](../spec/)
- Layout: `src/docxengine/` (package), `tests/` (pytest)

This implementation must stay byte-equivalent-after-normalization with [`@docxengine/core`](../js/) on the [conformance corpus](../conformance/) — a feature isn't done until it passes in both.
