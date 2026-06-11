# Error design

Errors are part of the interface, not exceptions to it. Every error is **structured and corrective**: it names what went wrong, explains it in plain language, and suggests the next call. A good error turns a failed step into a one-step recovery; a bad error spends fifty agent turns.

## The shape

```json
{
  "error": "anchor_stale",
  "message": "P12#a7b2 no longer matches (content changed). Re-read P12 or search.",
  "suggestions": ["docx_read(window:P12)"]
}
```

- `error` — a stable machine code from the [error catalog](../reference/error-codes.md)
- `message` — human/agent-readable, names the exact thing that failed
- `suggestions` — concrete next calls, not generic advice

## Principles

1. **Guarded, not destructive.** Like a linter-gated edit command, the engine refuses an invalid edit with feedback rather than applying a guess. A refused edit costs one turn; a corrupted document costs the task.
2. **Stale state is detected, never silently tolerated.** `anchor_stale` fires whenever a hash mismatch shows the agent's view has drifted — with the cheap recovery path attached.
3. **Forgiveness before failure.** Tools accept anchor _or_ search text, default missing `author` from environment, and coerce Markdown/plain content — an error only fires when intent is genuinely ambiguous or unsafe.
4. **Idempotency makes retries safe.** `accept_all` with nothing left, `replace {all:true}` with no matches → clean no-op results (`{accepted: 0}`), not errors. Network-level retries over MCP can't double-apply.
5. **Validation errors carry `fix_hint`.** Every `docx_validate` issue says which tool call (usually `docx_repair`) addresses it.

## Recovery patterns agents should know

| Error                | Recovery                                                             |
| -------------------- | -------------------------------------------------------------------- |
| `anchor_stale`       | `docx_read {anchor: "P12", window: 3}` → retry with the fresh anchor |
| `not_found` (search) | broaden query, or `docx_outline` to re-map                           |
| `validation_failed`  | `docx_repair`, then re-`validate`; report `remaining` to the user    |
| `render_unavailable` | fall back to `docx_convert {to: "md"}` for structural verification   |
| `doc_not_found`      | `docx_open` again — the session likely expired                       |

The full catalog with all codes: [reference/error-codes.md](../reference/error-codes.md).
