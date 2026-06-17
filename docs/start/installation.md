# Installation

## Python

```bash
pip install docxengine
```

Requires Python ≥3.12. No native dependencies — pure pip install.

## From source

```bash
git clone https://github.com/ruwadgroup/docxengine.git && cd docxengine
pip install -e python          # docxengine + the docxengine-mcp server
```

## MCP server

The MCP server ships with the Python package as a console script. Run it without installing anything via `uvx`:

```bash
uvx docxengine-mcp        # stdio transport (default)
```

Or install the package and run the console script:

```bash
pip install docxengine
docxengine-mcp            # stdio transport (default)
docxengine-mcp --http     # Streamable HTTP
```

### Claude Desktop / Claude Code

```json
{
  "mcpServers": {
    "docx": {
      "command": "uvx",
      "args": ["docxengine-mcp"]
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
