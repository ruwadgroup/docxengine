# Conformance

Shared input fixtures and the renderer-fidelity harness. The corpus pins real OOXML structures the engine must round-trip without triggering Word repair; the Python test suite opens these fixtures directly to assert deterministic, stable output. Background: [docs/conformance/corpus.md](../docs/conformance/corpus.md).

## Layout

```
conformance/
├── corpus/            # input fixtures (.docx) — synthetic or sanitized, never confidential
│   └── <name>/
│       ├── input.docx
│       └── meta.json          # producer, features exercised, notes
├── harness/
│   └── make_fixtures.py       # deterministic corpus generator (reused by bench/)
└── fidelity/          # renderer fidelity harness (structural + LibreOffice when present)
```

## How it's used

- **Python tests** (`python/tests/test_edit.py`, `test_read.py`, `test_revisions.py`) open fixtures from `corpus/` and assert exact results — guarded by `skipif(not CORPUS.is_dir())`.
- **`bench/`** regenerates the corpus deterministically via `harness/make_fixtures.py` (never reimplemented).
- **Fidelity** (`make fidelity`) runs `fidelity/run.py` — structural checks everywhere, LibreOffice visual rendering when available.

## Rules for corpus contributions

1. **Never confidential content** — rebuild structures with synthetic text.
2. Minimal documents — the smallest file exhibiting the structure.
3. `meta.json` must name the producer (application + version) and the features exercised.
4. Corrupt-on-purpose fixtures go in `corpus/corrupt-*` and are expected to fail validation in a _defined_ way.

Regenerate the corpus: `python conformance/harness/make_fixtures.py`. Run fidelity: `make fidelity`.
