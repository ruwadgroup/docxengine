# Framework adapters

One contract, every framework. Each tool is a pure function `fn(args_json) → result_json` with a published JSON Schema in [`spec/`](../../spec/) — the universal substrate that feeds MCP `tools/list`, OpenAI function-calling `tools`, and Anthropic tool definitions without modification.

## Design rule

**Adapters are ≤ a few lines each and contain no business logic.** They translate schema/registration formats — nothing else. If an adapter needs logic, the contract is wrong; fix the contract.

## Shipped adapters

| Framework               | Python                         | JS/TS                 |
| ----------------------- | ------------------------------ | --------------------- |
| OpenAI function calling | `docxengine.openai_tools()`    | `openaiTools()`       |
| Anthropic tool use      | `docxengine.anthropic_tools()` | `anthropicTools()`    |
| MCP                     | `docxengine-mcp` (server)      | via the Python server |

## Rolling your own

Any orchestrator integrates in two steps — register the schemas, dispatch to `call`:

```python
# 1. Give the model the tool definitions
import docxengine, json
schemas = docxengine.tool_schemas()          # the spec/, as dicts

# 2. When the model emits a tool call, dispatch it
result = docxengine.call(tool_name, tool_args)
```

```ts
import { toolSchemas, call } from "@docxengine/core";
const schemas = toolSchemas();
const result = await call(toolName, toolArgs);
```

That's the whole integration. Errors come back structured ([error design](../tools/errors.md)) so the model can self-correct; nothing framework-specific leaks into the engine.

## Persisting the result

The agent surface is storage-agnostic: open from bytes (`docx_open` accepts base64), and when the loop ends, get the bytes back out — your host decides where they go (upload, blob store, HTTP response, disk). This is a **library helper, not a wire tool**: returning megabytes of base64 through the model's context would be token waste, and your code already holds the session.

```python
data = docxengine.export_bytes(session, doc_id=doc_id)   # validated .docx bytes
```

```ts
const data = exportBytes(session, docId); // Uint8Array, browser-safe
```

`export_bytes`/`exportBytes` runs the same §8 validation gate as a save, then serializes — no filesystem. (The MCP server, being file-first, persists every edit to its path automatically and exposes no bytes tool.)
