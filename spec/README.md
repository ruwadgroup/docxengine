# The tool contract

This directory is the **source of truth** for DocxEngine's public interface. Everything else — the MCP `tools/list`, input validation, and the tests — is generated from or validated against these files.

## Layout

```
spec/
├── README.md          # this file
├── errors.json        # the error-code catalog (machine-readable)
└── tools/             # one JSON Schema per tool, named after the tool
    ├── docx_open.json
    ├── docx_replace.json
    └── …
```

## Rules

1. **No behavior lives here** — only shape: names, descriptions, parameter schemas, result schemas, error codes.
2. **Descriptions are written for agents.** They state what the tool does, when to use it, and what to call first. They are part of the interface and get refined from benchmark transcripts.
3. **Common parameters are identical everywhere**: `doc_id`, `anchor`, `track_changes`, `author`, `response_format` mean the same thing in every tool.
4. **Changes follow governance** ([GOVERNANCE.md](../GOVERNANCE.md)): additive preferred; breaking changes bump the contract version and land with the implementation + test updates in the same release.
5. **JSON Schema dialect**: the draft 2020-12 subset that MCP, OpenAI, and Anthropic all accept (no `$dynamicRef`, no remote `$ref`).

## File shape

```json
{
  "name": "docx_replace",
  "description": "…",
  "input_schema": { "type": "object", "properties": { … }, "required": [ … ] },
  "result_schema": { "type": "object", "properties": { … } },
  "errors": ["anchor_stale", "not_found", "ambiguous_target"]
}
```

`result_schema` and `errors` are DocxEngine extensions consumed by the implementation and its tests; the MCP `tools/list` strips them (MCP advertises only `input_schema`).
