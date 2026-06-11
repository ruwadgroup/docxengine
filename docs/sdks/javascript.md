# JS/TS SDK (`@docxengine/core`)

```bash
npm install @docxengine/core
```

Node ≥22; create/read paths also run in the browser. Fully typed.

## Two surfaces

### 1. The contract surface: `call()`

Identical names, JSON shapes, and errors to the Python package and the MCP server:

```ts
import { call } from "@docxengine/core";

const doc = await call("docx_open", { path: "contract.docx" });
await call("docx_replace", {
  doc_id: doc.doc_id,
  old: "Acme Corp",
  new: "GlobalTech Inc",
  all: true,
  track_changes: true,
  author: "Claude",
});
await call("docx_save", { doc_id: doc.doc_id, path: "out.docx" });
```

### 2. The native surface

```ts
import { Document } from "@docxengine/core";

const doc = await Document.open("contract.docx");
for (const p of doc.paragraphs()) {
  if (p.style === "Heading 2") console.log(p.anchor, p.text);
}
const p = doc.find("Confidential Information");
await p.replace("five (5) years", "three (3) years", { trackChanges: true, author: "Claude" });
await doc.save("out.docx");
```

## Errors

```ts
import { call, ToolError } from "@docxengine/core";

try {
  await call("docx_replace", { doc_id: "d1", anchor: "P12#a7b2", old: "x", new: "y" });
} catch (e) {
  if (e instanceof ToolError) {
    e.code; // "anchor_stale"
    e.suggestions; // ["docx_read(window:P12)"]
  }
}
```

## Framework adapters

```ts
import { openaiTools, anthropicTools } from "@docxengine/core";

const tools = openaiTools(); // ready for OpenAI `tools`
const atools = anthropicTools(); // ready for Anthropic tool use
```

Any other framework consumes the published JSON Schemas in [`spec/`](../../spec/) directly and dispatches to `call()`.

## Browser

The OPC model and projector run on `Uint8Array` inputs in the browser (open/read/search/convert-to-md). Filesystem-touching tools (`docx_open` by path, `docx_save` by path) and the render adapter are Node-only.
