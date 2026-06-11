# Surgical edit tools

Every edit tool: validates the anchor hash before touching anything, accepts `track_changes?: bool` + `author?: string`, runs incremental validation on its output, and returns fresh anchor(s). No edit can land on the wrong paragraph, and no edit can corrupt the package.

## `docx_replace`

```json
→ docx_replace {"doc_id":"d1","anchor":"P4#d4e5","old":"five (5) years","new":"three (3) years"}
← {"new_anchor":"P4#91c2","n_replaced":1}

→ docx_replace {"doc_id":"d1","old":"Acme Corp","new":"GlobalTech Inc","all":true}
← {"n_replaced":7,"anchors":["P2#aa01","P8#b2ff","…"]}
```

Omit `anchor` to search the whole document. The split-run problem is handled internally — the engine coalesces fragmented runs, replaces, and restores meaningful formatting boundaries; formatting of the surrounding text is preserved. `all: true` is idempotent (re-running matches nothing).

## `docx_edit_paragraph`

```json
→ docx_edit_paragraph {"doc_id":"d1","anchor":"P7#22ab",
    "text":"The term of this Agreement is two (2) years, renewing annually.",
    "track_changes":true,"author":"Claude"}
← {"new_anchor":"P7#e0c4","diff":"~3 words changed"}
```

Full-paragraph rewrite with **automatic word-level diff**: with tracking on, only the changed words become `w:ins`/`w:del`, producing the clean redline a human reviewer would make — not delete-all/insert-all.

## `docx_insert`

```json
→ docx_insert {"doc_id":"d1","after":"P9#1a2b","content":"## 2.1 Renewal\nThis Agreement renews…","style":"Heading 3"}
← {"new_anchors":["P10#f31a","P11#08d2"]}
```

`content` accepts Markdown (multi-paragraph, headings, lists) or plain text; `style` overrides per-paragraph when needed. Position is `after` or `before` an existing anchor.

## `docx_delete`

```json
→ docx_delete {"doc_id":"d1","range":"P14..P16"}
← {"ok":true,"deleted":3}
```

With `track_changes: true`, deletion becomes a `w:del` redline rather than physical removal.

## `docx_format`

```json
→ docx_format {"doc_id":"d1","style_selector":{"style":"Heading 2"},"props":{"color":"#1F4E79"}}
← {"affected":7,"anchors":["P3#b2c4","P9#1a2b","…"],
   "note":"Applied to style 'Heading 2' definition; 7 paragraphs use it."}
```

The high-leverage path: a `style_selector` edits the **style definition** in `styles.xml` — one change, document-wide effect, idempotent, and fidelity-preserving (no scattered direct overrides). Anchor/range targets apply direct formatting instead. `props`: color, bold, italic, size, alignment, spacing, etc.

## `docx_table`

One consolidated tool for all table operations:

```json
→ docx_table {"doc_id":"d1","op":"create","after":"P12#e7f8","rows":3,"cols":3,
              "data":[["Item","Qty","Price"],["Widget","10","$5"],["Gadget","4","$9"]],"header":true}
← {"new_anchor":"T4@after:P12#e7f8","note":"3×3 table inserted; header row styled with 'Table Grid'."}

→ docx_table {"doc_id":"d1","op":"set_cells","anchor":"T4","cells":[{"r":1,"c":2,"text":"$6"}]}
→ docx_table {"doc_id":"d1","op":"insert_row","anchor":"T4","at":2}
→ docx_table {"doc_id":"d1","op":"merge","anchor":"T4","range":"A1:C1"}
```

Ops: `create`, `set_cells`, `insert_row`, `insert_col`, `delete_row`, `delete_col`, `merge`, `style`. Cell addressing accepts `{r, c}` zero-based pairs or A1 notation.

## Batch semantics

Multiple operations submitted together validate **all anchors upfront** and apply atomically — see [Anchors](../core/anchors.md#batch-edits).
