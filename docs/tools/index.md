# Tool reference

The agent-computer interface: **~16 high-leverage, namespaced, forgiving tools** across 5 groups, all sharing the JSON-in/JSON-out contract published in [`spec/`](../../spec/). Designed per SWE-agent + Anthropic principles — few tools, consolidated operations, guarded edits, corrective errors, token-bounded responses.

## Conventions (all tools)

- **Forgiving inputs**: accept anchor _or_ search text; `author` defaults from environment; content accepts Markdown or plain text.
- **Idempotent where possible**: re-running `replace {all:true}` or `accept_all` is safe.
- **Re-anchoring**: every mutation returns the fresh anchor(s) it produced.
- **Token-bounded**: no response exceeds ~25k tokens; long outputs paginate or return resource links.
- **`response_format`**: `concise` (default) | `detailed`.
- Errors are structured and corrective — see [Error design](errors.md).

## Read / Navigate — [details](read-navigate.md)

| Tool           | Signature (abridged)                            | Returns                                                              |
| -------------- | ----------------------------------------------- | -------------------------------------------------------------------- |
| `docx_open`    | `(path\|bytes)`                                 | `{doc_id, summary, n_paragraphs, has_tracked_changes, has_comments}` |
| `docx_outline` | `(doc_id)`                                      | headings tree with anchors                                           |
| `docx_read`    | `(doc_id, {anchor?, range?, window?, format?})` | Markdown projection                                                  |
| `docx_search`  | `(doc_id, {query, regex?, scope?})`             | `[{anchor, snippet, context}]`                                       |

## Surgical edit — [details](edit.md)

Each accepts `track_changes?: bool` + `author?: string`; each returns new anchor(s).

| Tool                  | Signature (abridged)                                      | Notes                                              |
| --------------------- | --------------------------------------------------------- | -------------------------------------------------- |
| `docx_replace`        | `(doc_id, {anchor\|search, old, new, all?})`              | split-run coalescing handled internally            |
| `docx_edit_paragraph` | `(doc_id, {anchor, text})`                                | full rewrite, auto word-level diff → clean redline |
| `docx_insert`         | `(doc_id, {after\|before: anchor, content, style?})`      | content as Markdown or text                        |
| `docx_delete`         | `(doc_id, {anchor\|range})`                               |                                                    |
| `docx_format`         | `(doc_id, {anchor\|range\|style_selector, props})`        | style-selector edits hit the style definition      |
| `docx_table`          | `(doc_id, {op: create\|set_cells\|merge\|insert_row\|…})` | one consolidated table tool                        |

## Structure / assets — [details](structure.md)

| Tool                 | Operations                                                                              |
| -------------------- | --------------------------------------------------------------------------------------- |
| `docx_style`         | `list` · `define` · `apply`                                                             |
| `docx_section`       | page size/margins/orientation, headers/footers                                          |
| `docx_list`          | list creation and (re)numbering                                                         |
| `docx_comment`       | `add` · `reply` · `resolve` · `list` · `delete`                                         |
| `docx_revision`      | `list` · `accept` · `reject` · `accept_all` · `reject_all` with `{author, date}` filter |
| `docx_media`         | `insert` · `extract` · `replace`                                                        |
| `docx_field`         | TOC / page-number insertion as field codes                                              |
| `docx_template_fill` | mustache merge: placeholders, loops, conditions                                         |

## Create / Convert / Verify — [details](create-convert-verify.md)

| Tool                  | Signature (abridged)                 | Notes                                                   |
| --------------------- | ------------------------------------ | ------------------------------------------------------- |
| `docx_create`         | `({content_md\|spec})`               | new document from Markdown or a spec                    |
| `docx_convert`        | `(doc_id, {to: md\|html\|pdf\|png})` | md/html in-engine; pdf/png via render adapter           |
| `docx_validate`       | `(doc_id)`                           | `{valid, issues:[{severity, part, message, fix_hint}]}` |
| `docx_repair`         | `(doc_id)`                           | `{fixed:[…], remaining:[…]}`                            |
| `docx_render_preview` | `(doc_id, {pages?})`                 | image resource refs — agent self-check                  |
| `docx_save`           | `(doc_id, {path})`                   | runs the validation gate first                          |

## Schemas

Every tool's JSON Schema lives in [`spec/tools/`](../../spec/tools/) — the same schema feeds MCP `tools/list`, OpenAI function-calling, and Anthropic tool definitions. See [reference/tool-schemas.md](../reference/tool-schemas.md).
