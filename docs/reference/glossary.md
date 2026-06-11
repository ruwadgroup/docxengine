# Glossary

| Term                         | Meaning                                                                                                                                    |
| ---------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| **Anchor**                   | DocxEngine's stable paragraph address, `P{index}#{hash}` — ordinal hint + content-hash integrity check. See [anchors](../core/anchors.md). |
| **anchor_stale**             | The error returned when an anchor's hash no longer matches the paragraph content; recover by re-reading.                                   |
| **doc_id**                   | Handle for an open document held in engine/server memory; the unit of session state.                                                       |
| **Field (code)**             | An instruction computed at render time — page numbers, TOC entries. The engine edits codes, never invents values.                          |
| **mc:Ignorable**             | Markup-compatibility attribute letting consumers skip extension namespaces; must be preserved on round-trip.                               |
| **Normalization**            | The defined set of trivial differences (ZIP timestamps, entry order, XML whitespace) ignored when comparing outputs.                       |
| **OOXML / WordprocessingML** | The ECMA-376/ISO-29500 XML vocabulary inside .docx files.                                                                                  |
| **OPC**                      | Open Packaging Conventions — the ZIP+parts+relationships+content-types container format of .docx.                                          |
| **Paragraph (`w:p`)**        | The basic block element; the unit that anchors address.                                                                                    |
| **paraId (`w14:paraId`)**    | Word's extension paragraph ID — unique within a part but not stable across saves; a hint, never the address.                               |
| **Part**                     | One file inside the OPC package (`document.xml`, `styles.xml`, `comments.xml`, …).                                                         |
| **Projection**               | The token-efficient Markdown-like view of a document that agents read. See [projection](../core/projection.md).                            |
| **Redline**                  | A tracked change — `w:ins` (insertion) or `w:del` (deletion) attributed to an author + date.                                               |
| **Relationship (`.rels`)**   | OPC's typed link from one part to another (or to an external target); orphans trigger Word repair.                                         |
| **Repair (Word)**            | Word's silent rewrite of a file with broken internals — the failure mode the validation gate exists to prevent.                            |
| **rsid (`w:rsid*`)**         | Revision-save IDs Word stamps for merge accuracy; semantically meaningless for editing; stripped from projections.                         |
| **Run (`w:r`)**              | A span of uniformly formatted text in a paragraph; fragmented arbitrarily by Word (the split-run problem).                                 |
| **Split-run problem**        | Search text fragmented across run boundaries, defeating naive find-and-replace; solved by run coalescing.                                  |
| **Story**                    | An independent text flow: body, header, footer, footnote, endnote, comment. Each lives in its own part.                                    |
| **Style cascade**            | The six-layer formatting resolution: defaults → table → numbering → paragraph style → character style → direct.                            |
| **Validation gate**          | The always-on pre-save check (IDs, rels, content-types, references) guaranteeing zero Word-repair prompts.                                 |
