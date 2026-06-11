# Python SDK (`docxengine`)

```bash
pip install docxengine
```

Python ≥3.12, pure-Python install, no native toolchain.

The library is **storage-agnostic**: documents are in-memory handles you persist explicitly — to a path or to raw bytes. That is deliberate and distinct from the **MCP server**, which is file-first (every MCP tool takes a `path` and saves automatically; see [MCP server](../mcp/server.md)). Embedding software — a web backend, a serverless function — wants bytes in and bytes out, not a fixed filesystem.

Two surfaces, for two jobs.

## 1. Embed in an agent

Give a model the tool schemas, run a loop, dispatch each call against a `Session` you own, then persist the result however your app wants:

```python
import docxengine

tools = docxengine.anthropic_tools()        # or openai_tools() — the spec, as provider schemas

session = docxengine.Session()              # one per request/tenant, not a global
opened = docxengine.call("docx_open", {"bytes": b64_docx}, session=session)  # bytes in, no fs
doc_id = opened["doc_id"]

# ... the model emits tool calls; dispatch each against the same session ...
docxengine.call(
    "docx_replace",
    {"doc_id": doc_id, "old": "Acme Corp", "new": "GlobalTech Inc", "all": True},
    session=session,
)

# persist when the loop ends — your app owns storage
data = docxengine.export_bytes(session, doc_id=doc_id)   # validated .docx bytes, no fs
# upload `data`, stream it in an HTTP response, or write it to disk — your call
```

`call()` matches the MCP and JS surfaces exactly — same names, JSON in/out, same errors. `export_bytes` runs the same validation gate as a save but **returns bytes instead of writing a file**: returning base64 through the model's context would be token waste, and the host already holds the `Session`.

## 2. Manipulate documents in code

`Document` is a typed, full-coverage handle over the same tools — open from a path or bytes, edit, persist:

```python
from docxengine import Document

doc = Document.open("contract.docx")            # or Document.open(raw_bytes)
for p in doc.paragraphs():
    if p.style == "Heading2":                   # the w:pStyle styleId, or None
        print(p.anchor, p.text)

p = doc.find("Confidential Information")         # -> Paragraph | None
if p:
    p.replace("five (5) years", "three (3) years", track_changes=True, author="Claude")

doc.table("create", after=doc.paragraphs()[0].anchor, rows=2, cols=2,
          data=[["Item", "Price"], ["Setup", "$5,000"]])

doc.save("out.docx")        # validation gate + atomic write
data = doc.to_bytes()       # ...or get the bytes, no filesystem
```

Every tool has a method — `outline / read / search / insert / delete / edit_paragraph / revision / comment / table / style / format / list / section / media / field / validate / repair / convert / render_preview`. A `Paragraph` adds the anchor-scoped primitives `replace / edit / insert_after / insert_before / delete`. Start from scratch with `Document.create(content_md=...)`, fill a template with `Document.fill_template(path_or_bytes, data)`, share a session with `session=...`, or wrap a handle from the agent surface with `Document.attach(session, doc_id)`.

> Paragraphs are throwaway views: after any edit, re-fetch via `paragraphs()` / `find()`. Operating through a stale anchor raises `anchor_stale` — the spec's normal recovery signal.

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
