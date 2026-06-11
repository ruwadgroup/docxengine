# Installation

> **Pre-alpha**: the Phase 1 MVP is implemented, but packages are not yet published to PyPI/npm — install from source. The registry commands below activate at first release.

## From source (today)

```bash
git clone https://github.com/ruwadgroup/docxengine.git && cd docxengine
pip install -e python          # Python SDK + docxengine-mcp
pnpm install && pnpm --dir js build   # JS/TS SDK → js/dist
```

## Python (at first release)

```bash
pip install docxengine
```

Requires Python ≥3.12. No native dependencies — pure pip install.

## JS/TS (at first release)

```bash
npm install @docxengine/core
# or
pnpm add @docxengine/core
```

Requires Node ≥22. Create/read paths also work in the browser.

## MCP server

The MCP server ships with the Python package as a console script:

```bash
pip install docxengine
docxengine-mcp            # stdio transport (default)
docxengine-mcp --http     # Streamable HTTP (Phase 2)
```

### Claude Desktop / Claude Code

```json
{
  "mcpServers": {
    "docx": {
      "command": "docxengine-mcp"
    }
  }
}
```

## Optional: render adapter

Visual previews (`docx_render_preview`, `docx_convert` to pdf/png) use LibreOffice headless when present:

```bash
# macOS
brew install --cask libreoffice
# Debian/Ubuntu
apt-get install libreoffice --no-install-recommends
```

Without LibreOffice the engine degrades gracefully to a structural preview — everything else works. For pixel-closer fidelity with Word, install the metric-compatible Carlito/Caladea fonts.
