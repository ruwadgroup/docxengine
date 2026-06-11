# Agent task benchmark

The benchmark answers the only question that matters: **does an agent using DocxEngine complete document tasks more reliably and more cheaply than the alternatives?** It is built before optimization happens, so every design change is measured, not vibed.

## Tasks

Single- and multi-edit natural-language tasks with **element-level ground truth** — the expected document state is checkable mechanically:

| Task family  | Examples                                                                                |
| ------------ | --------------------------------------------------------------------------------------- |
| Text edits   | "Change the term from five years to three years"                                        |
| Style/format | "Change all H2 headings to blue"                                                        |
| Redlines     | "Accept all of Jane's tracked changes, leave Bob's"                                     |
| Structure    | "Insert a 3×3 pricing table after the payment clause"                                   |
| Comments     | "Add a comment on the indemnity clause asking if it should be mutual"                   |
| Templates    | "Fill this MSA template with the attached data"                                         |
| Multi-step   | "Find every mention of the old entity name, replace tracked, add a comment summarizing" |

## Metrics

Per task and aggregated:

- **Task success rate** (element-level ground-truth match)
- **Total runtime** of individual tool calls and tasks
- **Number of tool calls**
- **Token consumption** (the projection's reason to exist)
- **Tool errors** (and recovery rate — did `anchor_stale` lead to recovery or failure?)
- **Word "repair" rate** (must be zero)

## Baselines

1. **python-docx-wrapper MCP** — what most existing DOCX MCP servers are; inherits tracked-changes/comments gaps.
2. **Raw-XML approach** — unzip + grep + manual XML editing (the docx-skill pattern): powerful but token-heavy and brittle.

The MVP exit criterion ([ROADMAP.md](../../ROADMAP.md)): ≥ baseline task success with **lower token use** and **zero repair events** on text-edit and redline tasks.

## Method notes

- Same model, same prompt scaffold, same task set across all three conditions — only the tool surface changes.
- Transcripts are kept; failures get analyzed and feed tool-description and error-message refinements (evaluation-driven iteration).
- Layout-sensitive tasks add a visual-diff check: render to PNG, compare against Word's rendering.
