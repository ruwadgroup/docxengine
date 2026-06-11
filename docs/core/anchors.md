# Anchors: stable addressing

Stable addressing is the crux of the whole design. An agent doing a multi-step edit needs an address that survives between tool calls, detects when it's gone stale, and costs few tokens.

## Why not `w14:paraId`

Word's own paragraph ID looks tempting and is rejected on evidence:

1. **It's an Office-2010+ Word extension** (`w14` namespace, [MS-DOCX] §2.6.2.3). Documents from python-docx, docx-js, LibreOffice, and Google Docs often have none at all.
2. **Word generates values randomly** — unique within the part, values 0 < x < 0x80000000, but with no semantic relationship to content.
3. **It is not spec-guaranteed stable.** Documented case (Open-XML-SDK #925): after adding a comment and saving, _all_ paraId values were replaced by Word with new values — and the behavior is inconsistent (sometimes they're preserved). An address that survives "sometimes" is not an address.

`w14:paraId` is therefore surfaced only as a secondary hint and an optional hash seed — never as the durable address.

## The design: `P{index}#{hash}`

```
P12#a7b2
 │   └── first 4–6 hex chars of a hash over the paragraph's normalized text
 └────── ordinal position (1-based) — an ordering hint, not trusted for targeting
```

- **Normalization** before hashing: concatenate run text (ignoring run boundaries and rsid splits), normalize whitespace, exclude formatting. The same stored content always yields the same hash, regardless of which tool produced the file or how runs are fragmented.
- **Validation before every edit**: the target paragraph's current hash is recomputed and compared. On mismatch the edit is refused with a structured [`anchor_stale`](../reference/error-codes.md) error — never applied to the wrong paragraph.
- **Re-anchoring after every edit**: each edit returns the new anchor(s) for everything it touched, so agents chain edits without re-listing the document.
- **Tables** anchor as `T{index}`, with positional context in responses (`T4@after:P12#e7f8`).

## Staleness is a feature

Anchors _will_ go stale when content changes — that's the integrity check working. The recovery loop is cheap and explicit:

```json
→ docx_replace {"doc_id":"d1","anchor":"P12#a7b2","old":"...","new":"..."}
← {"error":"anchor_stale",
   "message":"P12#a7b2 no longer matches (content changed). Re-read P12 or search.",
   "suggestions":["docx_read(window:P12)"]}
→ docx_read {"doc_id":"d1","anchor":"P12","window":3}
← "[P12#91c2]  …current content…"      // fresh anchor, retry with it
```

## Hash collisions

A 4–6 hex-char hash is a _guard_, not a cryptographic identity — it disambiguates "the paragraph at position 12 with this content" within a document, where collisions at the same index are vanishingly rare. The pair (index, hash) must match together; the index alone is never trusted, and the hash alone is never searched.

## Batch edits

Multi-operation batches validate **all** anchors upfront before applying any operation (atomic semantics): either every anchor is current and the batch applies, or nothing does. This prevents half-applied batches when one anchor went stale mid-plan.
