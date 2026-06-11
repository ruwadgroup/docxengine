# Python SDK (`docxengine`)

```bash
pip install docxengine
```

Python ≥3.12, pure-Python install, no native toolchain.

## Two surfaces

### 1. The contract surface: `call()`

Matches MCP and the JS package exactly — same names, same JSON in/out, same errors:

```python
from docxengine import call

doc = call("docx_open", {"path": "contract.docx"})
call("docx_replace", {
    "doc_id": doc["doc_id"],
    "old": "Acme Corp", "new": "GlobalTech Inc", "all": True,
    "track_changes": True, "author": "Claude",
})
call("docx_save", {"doc_id": doc["doc_id"], "path": "out.docx"})
```

This is the surface to use in agent backends: deterministic, schema-validated, transport-identical.

### 2. The native surface

For power users who want objects and iteration:

```python
from docxengine import Document

doc = Document.open("contract.docx")
for p in doc.paragraphs():
    if p.style == "Heading 2":
        print(p.anchor, p.text)

p = doc.find("Confidential Information")
p.replace("five (5) years", "three (3) years", track_changes=True, author="Claude")
doc.save("out.docx")
```

Both surfaces share the same core — the native API is sugar over the same operations, not a second implementation.

## Errors

Contract errors raise `docxengine.ToolError` carrying the structured payload:

```python
from docxengine import call, ToolError

try:
    call("docx_replace", {"doc_id": "d1", "anchor": "P12#a7b2", "old": "x", "new": "y"})
except ToolError as e:
    e.code         # "anchor_stale"
    e.suggestions  # ["docx_read(window:P12)"]
```

## Framework adapters

```python
import docxengine

# OpenAI function calling — the schema list, ready to pass as `tools`
tools = docxengine.openai_tools()

# Anthropic tool use — same schemas, Anthropic shape
tools = docxengine.anthropic_tools()
```

Adapters are a few lines each and contain no logic — any other framework can consume the published JSON Schemas in [`spec/`](../../spec/) directly and dispatch to `call()`.

## MCP server

The reference MCP server ships with this package:

```bash
docxengine-mcp   # stdio transport
```
