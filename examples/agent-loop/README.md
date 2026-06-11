# The agent loop (annotated transcript)

A realistic multi-step agent session showing the intended loop — **map → locate → edit → recover → verify → save** — including the `anchor_stale` recovery that distinguishes this engine from line-number-based editing.

> Static transcript: every request/response below matches the contract in [`spec/`](../../spec/), and every tool shown is implemented today.

**Task given to the agent:** _"In contract.docx, shorten the term to three years and flag the liability cap for legal review."_

```json
→ docx_open {"path": "contract.docx"}
← {"doc_id": "d1", "summary": "Master Services Agreement — 38 paragraphs, 2 sections",
   "n_paragraphs": 38, "has_tracked_changes": false, "has_comments": false}

→ docx_outline {"doc_id": "d1"}
← {"outline": [
     {"anchor": "P1#a7b2", "level": 1, "text": "Master Services Agreement"},
     {"anchor": "P9#1a2b", "level": 2, "text": "2. Term"},
     {"anchor": "P17#c3d9", "level": 2, "text": "5. Liability"}]}

→ docx_search {"doc_id": "d1", "query": "five (5) years"}
← {"matches": [{"anchor": "P10#e7f8", "snippet": "…term of five (5) years, renewing…",
                "context": "2. Term"}], "n_matches": 1}

→ docx_replace {"doc_id": "d1", "anchor": "P10#e7f8", "old": "five (5) years",
                "new": "three (3) years", "track_changes": true, "author": "Claude"}
← {"new_anchor": "P10#91c2", "n_replaced": 1}
```

The agent now edits the same paragraph again — but mistakenly reuses the **old** anchor. The hash guard catches it:

```json
→ docx_edit_paragraph {"doc_id": "d1", "anchor": "P10#e7f8", "text": "…", "track_changes": true}
← {"error": "anchor_stale",
   "message": "P10#e7f8 no longer matches (content changed). Re-read P10 or search.",
   "suggestions": ["docx_read(window:P10)"]}

→ docx_read {"doc_id": "d1", "anchor": "P10", "window": 1}
← "[P10#91c2]  The term of this Agreement is three (3) years [ins by Claude], renewing annually."
```

One cheap re-read, and the agent continues with the fresh anchor — no corrupted edit, no full re-list. It finishes the second part of the task and the verification gate:

```json
→ docx_comment {"doc_id": "d1", "op": "add", "anchor": "P18#0b2c",
                "text": "Legal: please confirm the cap survives the term change.", "author": "Claude"}
← {"comment_id": "C1", "anchor": "P18#0b2c"}

→ docx_validate {"doc_id": "d1"}
← {"valid": true, "issues": []}

→ docx_render_preview {"doc_id": "d1", "pages": [1]}   // requires LibreOffice (structural fallback otherwise)
← {"pages": [{"page": 1, "image": "docx://d1/preview/page-1.png"}], "renderer": "libreoffice 24.8"}

→ docx_save {"doc_id": "d1", "path": "contract-redlined.docx"}
← {"ok": true, "validated": true}
```

## Why this transcript matters

- **Token cost**: the agent never held more than ~40 lines of document content in context.
- **Safety**: the stale anchor was _refused_, not guessed at — one extra call instead of a wrong-paragraph edit.
- **Fidelity**: the output is a real redline + comment another human opens in Word's Review pane.
