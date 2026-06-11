# Validation & repair

The validation gate is the single highest-leverage reliability feature in the engine. It exists because Word, on encountering broken package internals, silently "repairs" the file — rewriting it, sometimes destructively, with only a vague prompt to the user. The engine's contract: **a document that validates clean opens in Word with zero repair prompts.**

## What is checked

| Check                      | Catches                                                                         |
| -------------------------- | ------------------------------------------------------------------------------- |
| Internal ID uniqueness     | duplicate comment IDs, footnote IDs, bookmark IDs, numbering IDs                |
| Relationship integrity     | orphaned `.rels` entries, references to missing parts, dangling rIds            |
| Content-types completeness | parts present in the ZIP but absent from `[Content_Types].xml` (and vice versa) |
| Story consistency          | comment ranges without comments, footnote refs without footnotes                |
| Cross-references           | broken internal hyperlinks and bookmark references                              |
| Revision well-formedness   | unclosed `w:ins`/`w:del` scopes, invalid nesting                                |
| Schema sanity              | element ordering and attribute constraints that trigger Word repair             |

## The gate is always on

- `docx_save` runs full validation first and **refuses to write** a failing document (returning the issue list) unless the failure is auto-repairable.
- Each edit operation validates its own output incrementally — an edit that would corrupt the package fails as that edit, with a corrective error, rather than surfacing at save time.

## Tool surface

```json
→ docx_validate {"doc_id":"d1"}
← {"valid":false,"issues":[
    {"severity":"error","part":"word/comments.xml","message":"Comment id=3 referenced in body but missing",
     "fix_hint":"docx_repair removes the orphaned reference"}]}

→ docx_repair {"doc_id":"d1"}
← {"fixed":["removed orphaned comment reference id=3"],"remaining":[]}
```

`docx_repair` fixes what is safe to fix mechanically (orphaned references, missing content-type entries, duplicate ID renumbering) and clearly reports what it won't touch automatically. Severity `warning` issues (e.g., unknown but well-formed extension parts) never block a save.

## Repair-rate as a benchmark metric

The agent benchmark tracks **Word "repair" rate** as a first-class metric alongside task success and token use — the MVP exit criterion requires zero repair events across all redline tasks ([ROADMAP.md](../../ROADMAP.md)). Corrupt-on-purpose fixtures (duplicate IDs, orphaned footnotes, broken rels) live in the conformance corpus to keep `validate`/`repair` honest.
