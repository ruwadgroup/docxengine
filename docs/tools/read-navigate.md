# Read & navigate tools

The just-in-time retrieval pattern: keep lightweight identifiers in context, load detail on demand. Map with `outline`, locate with `search`, zoom with `read`.

## `docx_open`

```json
→ docx_open {"path": "contract.docx"}
← {"doc_id": "d1",
   "summary": "Master Services Agreement — 42 paragraphs, 3 sections, 2 tables",
   "n_paragraphs": 42,
   "has_tracked_changes": true,
   "has_comments": true}
```

Accepts a path or raw bytes. The summary is deliberately human-readable context (not technical IDs) so the agent immediately knows what it's holding. Opening never mutates the file; the `doc_id` handle is the unit of all subsequent state (see [MCP state](../mcp/state-and-scaling.md)).

## `docx_outline`

```json
→ docx_outline {"doc_id": "d1"}
← {"outline": [
    {"anchor": "P1#a7b2", "level": 1, "text": "Master Services Agreement"},
    {"anchor": "P3#b2c4", "level": 2, "text": "1. Definitions"},
    {"anchor": "P9#1a2b", "level": 2, "text": "2. Term"}],
   "tables": [{"anchor": "T1", "dims": "3×4", "after": "P5#cc01"}]}
```

The cheap map — headings resolved through the style cascade, with anchors ready for targeted reads. Always start here on an unfamiliar document.

## `docx_read`

```json
→ docx_read {"doc_id": "d1", "anchor": "P3#b2c4", "window": 4}
→ docx_read {"doc_id": "d1", "range": "P10..P24"}
→ docx_read {"doc_id": "d1", "range": "P10..P24", "format": "detailed"}
→ docx_read {"doc_id": "d1", "scope": "footnotes"}
```

Returns the [Markdown projection](../core/projection.md). `concise` (default) shows anchors + salient formatting; `detailed` adds resolved run-level formatting. Long ranges paginate with a continuation token under the ~25k-token cap. `scope` addresses other stories: `body` (default), `footnotes`, `comments`, `headers`, `footers`.

## `docx_search`

```json
→ docx_search {"doc_id": "d1", "query": "Confidential Information"}
← {"matches": [
    {"anchor": "P4#d4e5", "snippet": "\"Confidential Information\" means any information…",
     "context": "1. Definitions"}],
   "n_matches": 1}
```

Search-focused, not list-all: returns matches with surrounding context, never the whole document. `regex: true` enables pattern search; `scope` restricts to a story or range. Matching operates on **coalesced text**, so a query split across runs in the XML still hits.

## The intended loop

```
docx_open → docx_outline → docx_search → docx_read(window) → edit → re-anchor → save
```

An agent should rarely need more than a few hundred tokens of document content in context at any time.
