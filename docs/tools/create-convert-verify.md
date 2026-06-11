# Create, convert & verify tools

The lifecycle bookends: making documents, getting them out in other formats, and proving they're valid before they leave the engine.

## `docx_create`

```json
→ docx_create {"content_md":"# Quarterly Report\n\n## Summary\nRevenue grew 12%…"}
← {"doc_id":"d3","n_paragraphs":8}
```

Creates a new document from Markdown (headings, lists, tables, emphasis) or a structured spec (explicit styles/sections). Output passes the same validation gate as everything else.

## `docx_convert`

```json
→ docx_convert {"doc_id":"d1","to":"md"}
← {"content":"# Master Services Agreement\n…","note":"3 comments and 12 tracked changes annotated inline"}

→ docx_convert {"doc_id":"d1","to":"pdf","path":"out/contract.pdf"}
← {"path":"out/contract.pdf","renderer":"libreoffice 24.8"}
```

`md`/`html` are produced **in-engine** by the projector (lossless for content, annotated for revisions/comments). `pdf`/`png` go through the [render adapter](../core/render-adapter.md) and return paths/resource links, never inline blobs.

## `docx_validate` / `docx_repair`

Covered in depth in [Validation](../core/validation.md):

```json
→ docx_validate {"doc_id":"d1"}
← {"valid":true,"issues":[]}

→ docx_repair {"doc_id":"d1"}
← {"fixed":["renumbered duplicate bookmark id=12"],"remaining":[]}
```

Run `validate` any time; `save` runs it for you and refuses to write a broken package.

## `docx_render_preview`

```json
→ docx_render_preview {"doc_id":"d1","pages":[1]}
← {"pages":[{"page":1,"image":"docx://d1/preview/page-1.png"}],"renderer":"libreoffice 24.8"}
```

The agent self-check step: render, look, fix. Use at verification checkpoints (after a batch of layout-affecting edits), not after every edit — renders cost seconds.

## `docx_save`

```json
→ docx_save {"doc_id":"d1","path":"contract-redlined.docx"}
← {"ok":true,"validated":true,"bytes":48211}
```

Runs the full validation gate, rezips the package preserving everything untouched byte-for-byte (modulo normalization), and writes atomically. Saving never closes the `doc_id` — keep editing and save again.
