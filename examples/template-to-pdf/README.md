# Template to PDF (Phase 2)

Fill an MSA template from JSON data, validate, and convert it for delivery.

> **Status**: fully runnable — `docx_template_fill` and `docx_convert` are implemented. The scripts convert to **Markdown** (`to: "md"`, produced in-engine, no external tooling). **PDF** (`to: "pdf"` with a `path`) additionally requires LibreOffice on the host ([render adapter](../../docs/core/render-adapter.md)).

## Run it

```bash
python make_input.py                 # builds msa-template.docx (mustache placeholders + a loop)
python run.py
```

## The flow

```json
→ docx_template_fill {"template": "msa-template.docx",
   "data": {"EffectiveDate": "2026-07-01", "Client": "GlobalTech",
            "obligations": [{"text": "Deliver the Q3 report"}, {"text": "Maintain the SLA"}]},
   "syntax": "mustache"}
← {"doc_id": "d2", "filled": 4, "loops_expanded": {"obligations": 2}, "unfilled": [],
   "note": "All placeholders resolved."}

→ docx_validate {"doc_id": "d2"}
← {"valid": true, "issues": []}

→ docx_convert {"doc_id": "d2", "to": "md"}
← {"content": "# Master Services Agreement\n\nThis Agreement is effective as of 2026-07-01 ..."}
```

## What to verify

- `unfilled` is empty — the success check for template merges.
- The `{{#obligations}}` loop expands once per array element; `{{Client}}` and `{{EffectiveDate}}` resolve everywhere they appear.
- Placeholders split across runs (e.g. `{{Client}}` fragmented by Word's editing history) still resolve — run coalescing applies to template syntax too.
- For PDF, the output reflects computed fields only because a renderer produced it; the .docx itself stores field codes. PDF works when LibreOffice is installed.
