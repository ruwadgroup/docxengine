# Bulk rebrand

Replace a company name document-wide as a tracked change — the idempotent `all: true` flow.

## Run it

```bash
python make_input.py     # builds report.docx (the old name appears in several paragraphs, one split across runs)
python run.py
```

## The flow

See [calls.json](calls.json):

1. `docx_open`
2. `docx_search {query: "Acme Corp"}` — confirm where it appears (including a hit fragmented across runs)
3. `docx_replace {old: "Acme Corp", new: "GlobalTech Inc", all: true, track_changes: true, author: "Rebrand Bot"}`
4. `docx_replace` again — proves idempotency: `{n_replaced: 0}`
5. `docx_save`

## What to verify

- The split-run occurrence is replaced correctly (the engine coalesces runs before matching).
- Every replacement is a real `w:del`/`w:ins` pair attributed to "Rebrand Bot".
- Step 4 is a clean no-op, not an error.
