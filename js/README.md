# @docxengine/core (JS/TS)

The TypeScript implementation of [DocxEngine](../README.md): deterministic OOXML editing with tracked changes, hash-anchored addressing, and a token-efficient agent view. Runs in Node ≥22; create/read paths run in the browser.

```bash
npm install @docxengine/core
```

Storage-agnostic and browser-safe: open from bytes, edit through `call()` or the typed `Document` handle, and persist with `save(path)` (Node) or `toBytes()` (anywhere). The file-first path surface lives in the Python MCP server, not here.

- Usage: [docs/sdks/javascript.md](../docs/sdks/javascript.md)
- Contract: [spec/](../spec/)
- Layout: `src/` (TypeScript), `test/` (vitest)

This implementation must stay byte-equivalent-after-normalization with [`docxengine` (Python)](../python/) on the [conformance corpus](../conformance/) — a feature isn't done until it passes in both.
