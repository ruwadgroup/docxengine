# Tool schemas

Every tool's JSON Schema lives in [`spec/tools/`](../../spec/tools/) — one file per tool, named after the tool. These files are the **source of truth**: MCP `tools/list`, the OpenAI adapter, the Anthropic adapter, and both SDKs' input validation are all generated from them. Nothing is hand-maintained twice.

## Schema conventions

- `name` — the namespaced tool name (`docx_replace`).
- `description` — written for the _agent_: what it does, when to use it, what to call first. Descriptions are part of the interface and get refined from benchmark transcripts.
- `input_schema` — JSON Schema (draft 2020-12 subset all consumers support). Optional params have `default`s; enums are closed.
- Common params share names and semantics across every tool: `doc_id`, `anchor`, `track_changes`, `author`, `response_format`.

## Example: `docx_replace`

```json
{
  "name": "docx_replace",
  "description": "Replace text within a paragraph (or all paragraphs), preserving formatting across split runs. Returns the new anchor. Use docx_search first to locate text.",
  "input_schema": {
    "type": "object",
    "properties": {
      "doc_id": { "type": "string" },
      "anchor": {
        "type": "string",
        "description": "Paragraph anchor like 'P12#e7f8'. Omit to search whole doc."
      },
      "old": { "type": "string" },
      "new": { "type": "string" },
      "all": { "type": "boolean", "default": false },
      "track_changes": { "type": "boolean", "default": false },
      "author": { "type": "string" }
    },
    "required": ["doc_id", "old", "new"]
  }
}
```

## Consuming the schemas

| Consumer  | How                                                                                      |
| --------- | ---------------------------------------------------------------------------------------- |
| MCP       | served verbatim via `tools/list`                                                         |
| OpenAI    | `docxengine.openai_tools()` / `openaiTools()` wrap them in the function-calling envelope |
| Anthropic | `docxengine.anthropic_tools()` / `anthropicTools()`                                      |
| Custom    | read `spec/tools/*.json`, dispatch to `call(name, args)`                                 |

## Versioning

The contract is a stability surface ([GOVERNANCE.md](../../GOVERNANCE.md)): additive changes preferred; breaking changes require a schema version bump, a changelog deprecation note, and same-release updates to both implementations and the conformance harness.
