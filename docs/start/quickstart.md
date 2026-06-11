# Quickstart

The canonical loop is: **open → outline → search/read → edit → validate → save**, with anchors carrying state between calls.

## As tool calls (MCP or any framework)

```json
→ docx_open {"path": "contract.docx"}
← {"doc_id": "d1", "summary": "Master Services Agreement, 42 paragraphs, 3 sections",
   "n_paragraphs": 42, "has_tracked_changes": true, "has_comments": true}

→ docx_search {"doc_id": "d1", "query": "Confidential Information"}
← [{"anchor": "P4#d4e5", "snippet": "\"Confidential Information\" means...", "context": "1. Definitions"}]

→ docx_replace {"doc_id": "d1", "anchor": "P4#d4e5", "old": "five (5) years", "new": "three (3) years",
                "track_changes": true, "author": "Claude"}
← {"new_anchor": "P4#91c2", "n_replaced": 1}

→ docx_save {"doc_id": "d1", "path": "contract-redlined.docx"}
← {"ok": true, "validated": true}
```

Note the re-anchoring: the edit returned `P4#91c2` because the paragraph's content (and therefore its hash) changed. Use the new anchor for any follow-up edit — never the old one.

## Python

```python
from docxengine import call

doc = call("docx_open", {"path": "contract.docx"})
hits = call("docx_search", {"doc_id": doc["doc_id"], "query": "Confidential Information"})
call("docx_replace", {
    "doc_id": doc["doc_id"],
    "anchor": hits[0]["anchor"],
    "old": "five (5) years",
    "new": "three (3) years",
    "track_changes": True,
    "author": "Claude",
})
call("docx_save", {"doc_id": doc["doc_id"], "path": "contract-redlined.docx"})
```

Power users get native objects too (`from docxengine import Document`), but `call()` is the contract surface that matches MCP and the JS package exactly.

## JS/TS

```ts
import { call } from "@docxengine/core";

const doc = await call("docx_open", { path: "contract.docx" });
const hits = await call("docx_search", { doc_id: doc.doc_id, query: "Confidential Information" });
await call("docx_replace", {
  doc_id: doc.doc_id,
  anchor: hits[0].anchor,
  old: "five (5) years",
  new: "three (3) years",
  track_changes: true,
  author: "Claude",
});
await call("docx_save", { doc_id: doc.doc_id, path: "contract-redlined.docx" });
```

## Reading big documents cheaply

Don't read the whole document — map first, zoom second:

```json
→ docx_outline {"doc_id": "d1"}
← [{"anchor": "P1#a7b2", "level": 1, "text": "Master Services Agreement"},
   {"anchor": "P3#b2c4", "level": 2, "text": "1. Definitions"}, "…"]

→ docx_read {"doc_id": "d1", "anchor": "P3#b2c4", "window": 6}
← "[P3#b2c4  H2]  1. Definitions\n[P4#d4e5]  \"Confidential Information\" means… [comment:C1 by J.Doe]\n…"
```

## Next

- [Concepts](concepts.md) — anchors, projection, validation gate
- [Tool reference](../tools/index.md) — every tool, group by group
- [Error design](../tools/errors.md) — what `anchor_stale` means and how to recover
