# docxengine (Python)

The Python implementation of [DocxEngine](../README.md): deterministic OOXML editing with tracked changes, hash-anchored addressing, and a token-efficient agent view. Ships the reference MCP server (`docxengine-mcp`).

```bash
pip install docxengine
```

- Usage: [docs/sdks/python.md](../docs/sdks/python.md)
- Contract: [spec/](../spec/)
- Layout: `src/docxengine/` (package), `tests/` (pytest)

This implementation must stay byte-equivalent-after-normalization with [`@docxengine/core`](../js/) on the [conformance corpus](../conformance/) — a feature isn't done until it passes in both.
