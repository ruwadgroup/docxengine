# Examples

End-to-end flows demonstrating the engine from each face. Each is self-contained: a synthetic input generator (`make_input.py`), the exact tool calls, and a Python script (`run.py`). All examples are runnable today (PDF output in template-to-pdf additionally needs LibreOffice).

| Example            | Shows                                                                                             |
| ------------------ | ------------------------------------------------------------------------------------------------- |
| `redline-review/`  | Open a contract, accept one reviewer's tracked changes (author filter), validate, save            |
| `bulk-rebrand/`    | Document-wide tracked replace across split runs, with an idempotency proof                        |
| `agent-loop/`      | Annotated agent transcript: outline → search → edit → `anchor_stale` recovery → validate → save   |
| `template-to-pdf/` | Fill an MSA template from JSON data, validate, convert to Markdown (PDF when LibreOffice present) |
