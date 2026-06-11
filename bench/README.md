# Benchmark

The agent task benchmark: does an agent using DocxEngine complete real document tasks reliably and cheaply? Each task is a natural-language instruction with **element-level ground truth** — the expected document state is checkable mechanically, against the saved `.docx`, with no human in the loop. Background: [docs/conformance/benchmarks.md](../docs/conformance/benchmarks.md).

## Layout

```
bench/
├── tasks/             # task definitions — one .json per task
│   └── <name>.json           # {name, nl_task, fixture, script, checks, phase?}
├── checker.py         # assertion engine — evaluates checks against the saved output
├── run.py             # the runner: fixture → MCP server → drive → metrics → checks
└── results.json       # last run's per-task metrics (written by run.py)
```

## What it measures

Per task and aggregated, written to `results.json` and printed as a table:

- **Task success rate** — element-level ground-truth match (every check passes).
- **Total runtime** — wall time per call and per task.
- **Number of tool calls** — fewer, higher-leverage calls is the design goal.
- **Token consumption** — approximated as `len(json)/4` over every JSON-RPC request and response. The projection's whole reason to exist is keeping this low; it is the metric to beat against the baselines.
- **Tool errors** — calls that returned a structured error (and, with the LLM driver, whether the agent recovered).
- **Word "repair" rate** — must be zero. A non-clean validation before save (the gate that would otherwise let Word silently repair the file) is counted here.

## Task format

```json
{
  "name": "text-edit",
  "nl_task": "Change the contract term from five (5) years to three (3) years.",
  "fixture": "split-runs",
  "script": [
    { "tool": "docx_replace", "args": { "anchor": "P2#d337", "old": "...", "new": "..." } },
    { "tool": "docx_save", "args": { "path": "{out}" } }
  ],
  "checks": [
    { "type": "doc_text_contains", "text": "three (3) years" },
    { "type": "validate_clean" }
  ]
}
```

- **`fixture`** names a conformance corpus document (`conformance/corpus/<name>/input.docx`), regenerated deterministically by reusing `conformance/harness/make_fixtures.py`. The template task uses the synthetic `msa-template` built by `examples/template-to-pdf/make_input.py`.
- **`script`** is the scripted driver's call sequence — the _perfect agent_ reference trajectory. Output paths use a `{out}` placeholder the runner substitutes; the template task additionally uses `{template}`.
- **`checks`** is the element-level ground truth. See [checker.py](checker.py) for the full list: `doc_text_contains`, `doc_text_absent`, `paragraph_text`, `paragraph_count`, `revision_count`, `validate_clean`, `outline_contains`, `style_color`, `comment_count`, `comment_text_contains`.
- **`phase: 2`** marks a task that exercises a Phase 2 tool (tables, styles, comments, templates). Phase 2 tasks are excluded by default and run with `--phase2`.

## Drivers

The benchmark separates _what_ to achieve (the NL task + checks) from _who_ chooses the calls (the driver):

- **`scripted`** (default, the only driver today) — replays each task's `script` verbatim: the perfect-agent reference trajectory. It establishes the ceiling (calls, tokens, runtime an ideal agent would spend) and keeps the harness green and dependency-free, so regressions in the engine show up immediately.
- **`llm`** (documented, **not implemented yet**) — an actual agent loop that sees only `nl_task` and the tool catalog and chooses its own calls. It measures real task success and recovery. It needs a model API key and an agent scaffold; until that lands, `--driver llm` is rejected.

## Baselines (comparison protocol)

The headline result is **DocxEngine vs. the alternatives on the same tasks, same prompt, same model — only the tool surface changes**. Two baselines are defined; both are **documented here as the comparison protocol but require their own MCP servers**, which are not part of this repo:

1. **python-docx-wrapper MCP** — what most existing DOCX MCP servers are. Inherits python-docx's tracked-changes and comments gaps, so redline and comment tasks are expected to fail or corrupt.
2. **Raw-XML approach** — unzip + grep + manual XML editing (the docx-skill pattern). Powerful but token-heavy and brittle.

The MVP exit criterion ([ROADMAP.md](../ROADMAP.md)): ≥ baseline task success with **lower token use** and **zero repair events** on text-edit and redline tasks.

## Honest caveats

- The **scripted driver does not measure agent reasoning** — it replays a known-good trajectory, so its success rate is a property of the engine, not of any agent. Real task success comes from the LLM driver, which is not built yet.
- The **LLM driver needs an API key** and an agent scaffold; neither exists in this repo today.
- The **baselines are a protocol, not running code** — reproducing the head-to-head needs the two baseline servers stood up separately.
- **Token cost is an approximation** (`len(json)/4` over the JSON-RPC traffic), not a model tokenizer count. It is consistent across runs and good for relative comparison, not for billing.
- Determinism depends on `DOCXENGINE_FIXED_DATE=2026-01-01T00:00:00Z`, which the runner sets so tracked-change dates are stable.

## Running

```bash
make bench                       # MVP tasks only — green against the current engine
.venv/bin/python bench/run.py    # same thing
.venv/bin/python bench/run.py --task text-edit     # one task by name
.venv/bin/python bench/run.py --phase2             # include Phase 2 tasks (activate as tools land)
```

Exit status is non-zero if any selected task fails its checks. Phase 2 tasks are well-formed today but their tools return `not_implemented`; they go green once the Phase 2 workflow lands.
