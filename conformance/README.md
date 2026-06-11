# Conformance

The cross-implementation conformance suite: the same input document + the same tool call must produce **byte-equivalent-after-normalization** output in Python and TypeScript, and every document must round-trip without triggering Word repair. Background: [docs/conformance/corpus.md](../docs/conformance/corpus.md).

## Layout

```
conformance/
├── corpus/            # input fixtures (.docx) — synthetic or sanitized, never confidential
│   └── <name>/
│       ├── input.docx
│       └── meta.json          # producer, features exercised, notes
├── cases/             # test cases referencing corpus docs
│   └── <name>.json            # {doc, tool, args, expect: {result|error, output_normalized_sha}}
└── harness/           # the runner: executes cases against both implementations and diffs
```

## Case format

```json
{
  "doc": "redline-contract",
  "tool": "docx_replace",
  "args": {
    "anchor": "P4#d4e5",
    "old": "five (5) years",
    "new": "three (3) years",
    "track_changes": true,
    "author": "Test"
  },
  "expect": {
    "result": { "n_replaced": 1 },
    "invariants": ["roundtrip", "no_word_repair", "revisions_preserved"]
  }
}
```

The harness runs each case through `docxengine` (Python) and `@docxengine/core` (TS), normalizes both outputs (fixed ZIP entry order, zeroed timestamps, canonicalized XML in touched parts), and fails on any byte difference between implementations or any violated invariant.

## Rules for corpus contributions

1. **Never confidential content** — rebuild structures with synthetic text.
2. Minimal documents — the smallest file exhibiting the structure.
3. `meta.json` must name the producer (application + version) and the features exercised.
4. Corrupt-on-purpose fixtures go in `corpus/corrupt-*` and are expected to fail validation in a _defined_ way.

Run locally: `make conformance`.
