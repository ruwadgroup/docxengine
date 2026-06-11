# Structure & asset tools

The Phase 2 surface ([ROADMAP.md](../../ROADMAP.md)): styles, sections, lists, comments, revisions, media, fields, and templates. Consolidated op-style tools — one tool per domain, an `op` parameter per action — keeping the tool count low without hiding capability.

## `docx_style`

```json
→ docx_style {"doc_id":"d1","op":"list"}
← {"styles":[{"id":"Heading2","name":"Heading 2","type":"paragraph","based_on":"Heading1","in_use":7}, …]}

→ docx_style {"doc_id":"d1","op":"define","name":"Clause","based_on":"Normal","props":{"size":11,"spacing_after":6}}
→ docx_style {"doc_id":"d1","op":"apply","anchor":"P14#0b2c","style":"Clause"}
```

`list` reports the resolved cascade (`based_on` chains) so agents understand inheritance before changing it.

## `docx_section`

Page geometry and per-section headers/footers: size, margins, orientation, columns. Section breaks are surfaced in projections (`[S2 break:nextPage]`) and addressable.

## `docx_list`

Creates and renumbers lists against `numbering.xml`, hiding the abstract-numbering indirection. Ops: `create` (ol/ul, multi-level), `restart`, `set_level`, `convert` (paragraphs ↔ list items).

## `docx_comment`

```json
→ docx_comment {"doc_id":"d1","op":"add","anchor":"P4#d4e5","text":"Should this be mutual?","author":"Claude"}
← {"comment_id":"C7","anchor":"P4#d4e5"}

→ docx_comment {"doc_id":"d1","op":"reply","comment_id":"C7","text":"Yes — drafting.","author":"J.Doe"}
→ docx_comment {"doc_id":"d1","op":"resolve","comment_id":"C7"}
→ docx_comment {"doc_id":"d1","op":"list"}
```

A single `add` touches five package locations (range markers, comments part, rels, content-types, IDs) — the engine handles all of them and the validation gate proves it. Threaded replies and resolution state use the modern `commentsExtended` parts when present.

## `docx_revision`

Covered in depth in [Tracked changes](../core/tracked-changes.md): `list` / `accept` / `reject` / `accept_all` / `reject_all` with `{author, date}` filters.

## `docx_media`

```json
→ docx_media {"doc_id":"d1","op":"insert","after":"P20#aa19","image":"logo.png","width_cm":4}
→ docx_media {"doc_id":"d1","op":"extract","media_id":"M2","path":"out/chart.png"}
→ docx_media {"doc_id":"d1","op":"replace","media_id":"M2","image":"chart-v2.png"}
```

Handles the media part, relationship, and content-type registration as one operation.

## `docx_field`

```json
→ docx_field {"doc_id":"d1","op":"insert_toc","after":"P1#a7b2","levels":3}
→ docx_field {"doc_id":"d1","op":"insert_page_number","scope":"footer"}
→ docx_field {"doc_id":"d1","op":"update"}
```

Inserts/updates **field codes**. The computed values (page N of M, TOC entries) materialize only when Word or LibreOffice renders — responses say so explicitly, so agents don't hallucinate page numbers.

## `docx_template_fill`

```json
→ docx_template_fill {"template":"msa.docx",
   "data":{"EffectiveDate":"2026-07-01","Client":"GlobalTech",
           "obligations":[{"text":"Deliver Q3 report"},{"text":"Maintain SLA"}]},
   "syntax":"mustache"}
← {"doc_id":"d2","filled":4,"loops_expanded":{"obligations":2},"unfilled":[],
   "note":"All placeholders resolved."}
```

Mustache-style merge with loops and conditions. Placeholders fragmented across split runs are coalesced before matching. `unfilled` lists any placeholder the data didn't cover — an empty list is the success check.
