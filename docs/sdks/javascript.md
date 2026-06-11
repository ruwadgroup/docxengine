# JS/TS SDK (`@docxengine/core`)

```bash
npm install @docxengine/core
```

Node ≥22, fully typed. The core is **browser-safe**: open/edit/read/search/convert-to-md and `toBytes()` run on `Uint8Array` with zero filesystem dependency — only path-based `open`/`save` and the render adapter are Node-only.

The library is **storage-agnostic** by design — documents are in-memory handles you persist explicitly (a path, or raw bytes). That is distinct from the **MCP server**, which is file-first (every tool takes a `path` and saves automatically; see [MCP server](../mcp/server.md)). A browser has no filesystem, so the library never assumes one.

Two surfaces, for two jobs.

## 1. Embed in an agent

Give a model the tool schemas, run a loop on a `Session` you own, then persist however your app wants:

```ts
import { Session, call, anthropicTools, exportBytes } from "@docxengine/core";

const tools = anthropicTools(); // or openaiTools() — the spec, as provider schemas

const session = new Session(); // one per request/tab, not a global
const opened = (await call("docx_open", { bytes: b64Docx })) as { doc_id: string }; // bytes in
// ... the model emits tool calls; dispatch each ...
await call("docx_replace", { doc_id: opened.doc_id, old: "Acme", new: "GlobalTech", all: true });

const data = exportBytes(session, opened.doc_id); // validated .docx bytes, no fs
// download in the browser, upload, or return in a response — your call
```

> `call()` uses a module-level session for quick scripts. For multi-tenant servers, dispatch against an explicit `Session` (as the agent loop does) so handles don't accumulate process-wide. `exportBytes` runs the same validation gate as a save but returns bytes — persistence is the host's job.

## 2. Manipulate documents in code

`Document` is a typed, full-coverage handle over the same tools. Each instance owns a **private session**, so many can run side by side (multi-tenant server, browser tab):

```ts
import { Document } from "@docxengine/core";

const doc = await Document.open("contract.docx"); // or Document.open(uint8Array)
for (const p of doc.paragraphs()) {
  if (p.style === "Heading2") console.log(p.anchor, p.text); // the w:pStyle styleId, or null
}

const p = doc.find("Confidential Information"); // -> DocumentParagraph | null
if (p)
  await p.replace("five (5) years", "three (3) years", { trackChanges: true, author: "Claude" });

await doc.save("out.docx"); // validation gate + atomic write (Node)
const bytes = doc.toBytes(); // ...or get the bytes, no filesystem (browser-safe)
```

Every tool has a method — `outline / read / search / insert / delete / editParagraph / revision / comment / table / style / format / list / section / media / field / validate / repair / convert / renderPreview`. A `DocumentParagraph` adds the anchor-scoped primitives `replace / edit / insertAfter / insertBefore / delete`. Start fresh with `Document.create({ contentMd })`, fill a template with `Document.fillTemplate(pathOrBytes, data)`, share a session with `{ session }`, or wrap a handle with `Document.attach(session, docId)`.

> Paragraphs are throwaway views: after any edit, re-fetch via `paragraphs()` / `find()`. A stale anchor raises `anchor_stale`, the normal recovery signal.

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
