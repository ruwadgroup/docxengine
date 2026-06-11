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
