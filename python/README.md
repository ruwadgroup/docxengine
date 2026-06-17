# docxengine

Deterministic OOXML editing with tracked changes, hash-anchored addressing, and a token-efficient agent view — exposed as an **MCP server** for AI agents. Editing the XML directly preserves tracked changes, comments, and footnotes that mainstream libraries drop.

```bash
pip install docxengine        # or: uvx docxengine-mcp  (zero-install run)
```

The **file-first MCP server** (`docxengine-mcp`): every tool takes a file `path`, and each edit is validated and saved back automatically — no handle to track, no save step.

```bash
docxengine-mcp                              # stdio
claude mcp add docx -- uvx docxengine-mcp   # Claude Code
```

- MCP server: [docs/mcp/server.md](../docs/mcp/server.md)
- Tool reference: [docs/tools/index.md](../docs/tools/index.md)
- Contract: [spec/](../spec/)
- Layout: `src/docxengine/` (package + server), `tests/` (pytest)

The public tool contract lives in [spec/](../spec/) and is validated against the [conformance corpus](../conformance/).
