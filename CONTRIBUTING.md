# Contributing to DocxEngine

Thanks for your interest! DocxEngine is design-first and early-stage, which means contributions right now have outsized impact.

## The most valuable contributions (especially now)

1. **Conformance corpus documents** — real-world docx files (sanitized!) that exercise tracked changes, comments, footnotes, complex tables, multi-section layouts, non-Latin scripts, or files produced by non-Word tools (LibreOffice, Google Docs, python-docx). See [conformance/README.md](conformance/README.md).
2. **OOXML edge-case reports** — documented cases where Word "repairs" a file, regenerates `w14:paraId`, splits runs unexpectedly, or where renderers disagree. File them with the bug template.
3. **Benchmark tasks** — natural-language edit tasks with element-level ground truth for the agent benchmark ([docs/conformance/benchmarks.md](docs/conformance/benchmarks.md)).
4. **Design review** — holes in [ARCHITECTURE.md](ARCHITECTURE.md), the [tool contract](spec/), or the [error catalog](docs/reference/error-codes.md). Open a Discussion.
5. **Implementation** — pick a Phase 0/1 item from [ROADMAP.md](ROADMAP.md) and claim it in an issue before starting.

## Ground rules (the invariants)

Every PR must preserve these — they are the project (full list in [ARCHITECTURE.md](ARCHITECTURE.md#invariants)):

- **Fidelity**: tracked changes, comments, footnotes, and media survive every edit path; open→save of an untouched doc is byte-stable modulo normalization.
- **No silent repair**: documents that validate clean must open in Word without a "repair" prompt.
- **Hash-guarded edits**: no edit lands on a paragraph whose anchor hash fails validation.
- **Token economy**: no tool response exceeds ~25k tokens; raw OOXML is never returned by default.
- **Determinism**: no LLM in the core; identical inputs produce identical bytes.
- **Parity**: Python and TS stay byte-equivalent-after-normalization on the conformance corpus — a feature is not done until it passes in both.
- **Thin faces**: MCP server and framework adapters translate formats only; behavior lives in the core.

## Development setup

| Tool        | Version | Used for                                   |
| ----------- | ------- | ------------------------------------------ |
| Python      | ≥3.12   | `docxengine` package, MCP server           |
| Node.js     | ≥22     | `@docxengine/core`, repo tooling           |
| pnpm        | 10.x    | Repo tooling (prettier, commitlint, husky) |
| LibreOffice | any     | Optional — render adapter tests/previews   |

```bash
git clone https://github.com/ruwadgroup/docxengine.git
cd docxengine
make setup   # pnpm install + husky hooks
```

## Workflow

1. Branch from `main`: `git checkout -b feat/short-description`
2. Make the change. Keep the invariants. Add conformance cases and tests.
3. Run the checks. Update docs if behavior or contract changed.
4. Open a PR using the template. Link the issue. Describe testing.

## Checks (must pass before merge)

- `pnpm format:check` — Prettier on Markdown/JSON/YAML
- `make lint` — language linters (ruff for Python, as code lands)
- `make test` — unit tests in both implementations
- `make conformance` — cross-implementation conformance harness
- Commit messages pass commitlint (enforced by the husky `commit-msg` hook)

## Commit style

We use [Conventional Commits](https://www.conventionalcommits.org). Scopes follow the subsystem names: `core`, `opc`, `anchors`, `projector`, `edit`, `revisions`, `comments`, `tables`, `styles`, `validate`, `render`, `mcp`, `py`, `js`, `adapters`, `spec`, `conformance`, `bench`, `docs`, `examples`, `ci`, `deps`, `release`.

Examples:

```
feat(anchors): seed content hash with w14:paraId when present
fix(edit): coalesce runs split by rsid before replace
docs(tools): document anchor_stale recovery flow
test(conformance): add LibreOffice-authored contract fixture
```

## Reporting bugs & ideas

Use the [issue templates](https://github.com/ruwadgroup/docxengine/issues/new/choose) — bug reports ask for the producing application and a minimal sample document, which is usually the whole diagnosis. Never attach confidential documents; reproduce with synthetic content.

## Code of Conduct

This project follows the [Contributor Covenant](CODE_OF_CONDUCT.md). Be excellent to each other.
