# Redline review

Open a contract with tracked changes from two reviewers, accept only Jane Doe's, leave Bob's untouched, and save — the canonical MVP redline flow.

## Run it

```bash
python make_input.py                 # builds contract.docx (synthetic, two authors' redlines)
python run.py
```

## The flow (as raw tool calls)

See [calls.json](calls.json) — the same five calls work over MCP and the CLI:

1. `docx_open` → doc handle + `has_tracked_changes: true`
2. `docx_revision {op: "list"}` → inventory by author
3. `docx_revision {op: "accept", filter: {author: "Jane Doe"}}` → Jane's resolved, Bob's intact
4. `docx_validate` → clean
5. `docx_save` → `contract-reviewed.docx`

## What to verify

- The output opens in Word with **no repair prompt**.
- Word's Review pane shows only Bob's remaining revisions.
- Re-running the accept call is a clean no-op (`{accepted: 0}`).
