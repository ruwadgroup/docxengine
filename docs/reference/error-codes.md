# Error codes

The complete catalog of structured error codes. Every error response carries `{error, message, suggestions}` — see [error design](../tools/errors.md) for the principles.

## Addressing

| Code               | Meaning                                             | Typical recovery                                       |
| ------------------ | --------------------------------------------------- | ------------------------------------------------------ |
| `anchor_stale`     | Anchor hash no longer matches the paragraph content | `docx_read {anchor, window}` → retry with fresh anchor |
| `anchor_invalid`   | Malformed anchor string                             | check format `P{index}#{hash}`                         |
| `anchor_not_found` | Index out of range / table anchor missing           | `docx_outline` to re-map                               |

## Documents & session

| Code              | Meaning                                    | Typical recovery                                                     |
| ----------------- | ------------------------------------------ | -------------------------------------------------------------------- |
| `doc_not_found`   | Unknown/expired `doc_id`                   | `docx_open` again                                                    |
| `open_failed`     | Not a valid docx / unreadable path         | check path; `file` says what it is                                   |
| `doc_too_large`   | Exceeds configured memory caps             | open with streaming options / split                                  |
| `path_denied`     | Path outside the server's configured roots | use an allowed path                                                  |
| `not_implemented` | Tool not implemented in this build         | this tool lands in a later phase; see [ROADMAP.md](../../ROADMAP.md) |

## Edits

| Code               | Meaning                                                | Typical recovery                                   |
| ------------------ | ------------------------------------------------------ | -------------------------------------------------- |
| `not_found`        | Search text / `old` string not present                 | broaden query; check the projection for exact text |
| `ambiguous_target` | `old` matches multiple times without `all: true`       | add `all: true` or narrow with an anchor           |
| `style_unknown`    | Named style doesn't exist                              | `docx_style {op: "list"}`                          |
| `batch_aborted`    | An anchor in an atomic batch failed upfront validation | re-read the stale anchors; resubmit                |
| `edit_conflict`    | Concurrent write on the same doc_id (HTTP transport)   | retry — edits serialize on the write lock          |

## Validation & save

| Code                | Meaning                                         | Typical recovery                |
| ------------------- | ----------------------------------------------- | ------------------------------- |
| `validation_failed` | Package would trigger Word repair; save refused | `docx_repair`, then re-validate |
| `repair_incomplete` | `docx_repair` couldn't fix everything safely    | report `remaining` to the user  |
| `save_failed`       | I/O failure writing output                      | check path/permissions          |

## Convert & render

| Code                 | Meaning                           | Typical recovery                                                              |
| -------------------- | --------------------------------- | ----------------------------------------------------------------------------- |
| `render_unavailable` | No render adapter installed       | fall back to `docx_convert {to: "md"}`; install LibreOffice for visual checks |
| `render_failed`      | Renderer errored on this document | response includes renderer stderr summary                                     |
| `unsupported_format` | Unknown `to:` target              | use `md`, `html`, `pdf`, or `png`                                             |

## Templates

| Code                   | Meaning                                              | Typical recovery                             |
| ---------------------- | ---------------------------------------------------- | -------------------------------------------- |
| `placeholder_unfilled` | Data missing for required placeholders (strict mode) | response lists them; extend `data`           |
| `template_syntax`      | Malformed mustache syntax in the template            | response points at the offending placeholder |

## Conventions

- Codes are stable machine identifiers; messages may improve over time, codes never change meaning.
- Benign no-ops are **not** errors: `accept_all` with nothing to accept returns `{accepted: 0}`.
- Severity `warning` validation issues never block; only `error` does.
