# Tracked changes (redlines)

First-class tracked-change support is the single clearest reason this engine exists: python-docx's tracked-changes issue has been open since 2016, and reading a revised document through it returns wrong text. DocxEngine writes and resolves real `w:ins`/`w:del` markup.

## Writing redlines

Every edit tool accepts `track_changes?: bool` and `author?: string`:

```json
→ docx_replace {"doc_id":"d1","anchor":"P4#d4e5","old":"five (5) years","new":"three (3) years",
                "track_changes":true,"author":"Claude"}
```

produces genuine OOXML revision markup — deleted text wrapped in `w:del`/`w:delText`, inserted text in `w:ins`, both attributed with author and timestamp — that Word displays in its Review pane exactly like a human reviewer's edits.

`docx_edit_paragraph` goes further: a full-paragraph rewrite is **auto-diffed at word level**, so tracking on produces a clean, minimal redline instead of "delete everything, insert everything."

## Resolving revisions

```json
→ docx_revision {"doc_id":"d1","op":"list"}
← {"revisions":[{"id":"R1","type":"ins","author":"Jane Doe","date":"2026-05-02","anchor":"P7#22ab","text":"…"}, …]}

→ docx_revision {"doc_id":"d1","op":"accept","filter":{"author":"Jane Doe"}}
← {"accepted":12,"remaining_by_author":{"Bob":3},"note":"Resolved <w:ins>/<w:del> for Jane Doe; Bob's 3 revisions untouched."}
```

- `op`: `list` | `accept` | `reject` | `accept_all` | `reject_all`
- `filter`: by `author` and/or `date` range
- Idempotent: `accept_all` no-ops cleanly when no revisions remain.

Accepting an insertion unwraps the `w:ins`; accepting a deletion removes the `w:del` content; rejecting does the inverse. Adjacent runs re-coalesce afterward, and all affected anchors are returned fresh.

## Semantics agents can rely on

- **Projections default to "as accepted"** with markers (`[ins by Jane]`), so reading text is never wrong-by-default the way library wrappers are.
- Revisions interact with comments and footnotes correctly: accepting a deletion that spans a comment range preserves or cleanly removes the comment per OOXML rules — these cross-part interactions are conformance-tested.
- Mixed-author documents resolve per-filter without disturbing other authors' changes.

## What's tested

The conformance corpus includes legal-contract fixtures with multi-author redlines; the gate is that any accept/reject sequence produces a document Word opens with **zero repair prompts** and with the remaining revisions intact. See [conformance](../conformance/corpus.md).
