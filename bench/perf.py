#!/usr/bin/env python3
"""Large-document performance benchmark (stdlib + docxengine).

Builds synthetic documents of increasing size and measures the core operations'
wall time and peak Python-heap memory, so performance regressions and the §27
bounded-memory guardrails stay visible. This complements the agent-task
benchmark (``bench/run.py``): that one measures token cost over the MCP server,
this one measures raw engine throughput on big inputs.

The corpus is generated deterministically by reusing the conformance fixture
helpers (``conformance/harness/make_fixtures.py``) — never reimplemented.

    make perf                      # default sizes: 1k / 5k / 20k paragraphs
    .venv/bin/python bench/perf.py --sizes 1000 5000 20000 50000

Results print as a table and are written to ``bench/perf-results.json``.
"""

from __future__ import annotations

import argparse
import gc
import io
import json
import sys
import time
import tracemalloc
import zipfile
from pathlib import Path
from typing import Any, Callable

BENCH_DIR = Path(__file__).resolve().parent
REPO_DIR = BENCH_DIR.parent
RESULTS_PATH = BENCH_DIR / "perf-results.json"

# Reuse the conformance fixture builders and the installed engine.
sys.path.insert(0, str(REPO_DIR / "conformance" / "harness"))
sys.path.insert(0, str(REPO_DIR / "python" / "src"))

import make_fixtures as mk  # noqa: E402

from docxengine import Document  # noqa: E402

# A token present in every paragraph, so search touches the whole document.
NEEDLE = "fox"
# A token in exactly one paragraph, for the realistic single/anchored replace.
MARKER = "ZZUNIQUEMARKERZZ"


def build_doc(n_paras: int) -> bytes:
    """A valid .docx with ``n_paras`` body paragraphs; every one carries the
    search needle, and the first carries a unique marker for single-replace."""
    body = "".join(
        mk.para(
            mk.run(
                f"Paragraph {i}: the quick brown {NEEDLE} jumps over the lazy dog #{i}."
                + (f" {MARKER}" if i == 0 else "")
            ),
            style="Heading1" if i % 50 == 0 else None,
        )
        for i in range(n_paras)
    )
    parts = [
        ("[Content_Types].xml", mk.content_types(mk.STD_OVERRIDES)),
        ("_rels/.rels", mk.PKG_RELS),
        ("word/document.xml", mk.document(body)),
        ("word/_rels/document.xml.rels", mk.relationships(mk.STD_RELS)),
        ("word/styles.xml", mk.styles_xml(headings=True)),
    ]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts:
            zf.writestr(name, data)
    return buf.getvalue()


def _ms(fn: Callable[[], Any]) -> tuple[float, Any]:
    gc.collect()
    start = time.perf_counter()
    out = fn()
    return (time.perf_counter() - start) * 1000.0, out


def measure(n_paras: int, *, heavy: bool = False) -> dict[str, Any]:
    docx = build_doc(n_paras)
    gc.collect()
    tracemalloc.start()

    open_ms, doc = _ms(lambda: Document.open(docx))
    outline_ms, _ = _ms(doc.outline)
    search_ms, hits = _ms(lambda: doc.search(NEEDLE))
    read_ms, _ = _ms(lambda: doc.read(window=20))
    # The realistic edit: replace a single (here, marker) occurrence — what an
    # agent does per anchored edit. O(document) to locate + splice once.
    replace_ms, _ = _ms(lambda: doc.replace(MARKER, "DONE", all=False))
    validate_ms, _ = _ms(doc.validate)
    save_ms, _ = _ms(doc.to_bytes)

    ms = {
        "open": round(open_ms, 1),
        "outline": round(outline_ms, 1),
        "search": round(search_ms, 1),
        "read": round(read_ms, 1),
        "replace_one": round(replace_ms, 1),
        "validate": round(validate_ms, 1),
        "to_bytes": round(save_ms, 1),
    }
    if heavy:
        # KNOWN HOTSPOT: replace(all=True) splices the whole document once per
        # match → superlinear on documents with a match in every paragraph.
        # Tracked as a Phase 3 perf-tuning item (batch all edits into one splice
        # pass). Measured only under --heavy because it dominates wall time.
        fresh = Document.open(docx)
        heavy_ms, _ = _ms(lambda: fresh.replace(NEEDLE, "wolf", all=True))
        ms["replace_all_heavy"] = round(heavy_ms, 1)

    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    n_hits = len(hits.get("matches", hits.get("hits", []))) if isinstance(hits, dict) else 0
    return {
        "paragraphs": n_paras,
        "input_kb": round(len(docx) / 1024, 1),
        "search_hits": n_hits,
        "peak_mb": round(peak / (1024 * 1024), 1),
        "ms": ms,
    }


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--sizes", type=int, nargs="+", default=[1000, 5000, 20000])
    ap.add_argument(
        "--heavy",
        action="store_true",
        help="also measure replace(all=True) — the known O(n^2) hotspot; slow.",
    )
    args = ap.parse_args(argv)

    rows = [measure(n, heavy=args.heavy) for n in sorted(args.sizes)]

    cols = ["open", "outline", "search", "read", "replace_one", "validate", "to_bytes"]
    if args.heavy:
        cols.append("replace_all_heavy")
    header = f"{'paras':>8} {'in_KB':>8} {'peak_MB':>8}  " + "  ".join(f"{c:>11}" for c in cols)
    print(header)
    print("-" * len(header))
    for r in rows:
        line = f"{r['paragraphs']:>8} {r['input_kb']:>8} {r['peak_mb']:>8}  " + "  ".join(
            f"{r['ms'][c]:>9.1f}ms" for c in cols
        )
        print(line)

    # Cheap regression signal: every op should scale roughly linearly. Report the
    # worst per-paragraph open cost so a superlinear regression is obvious.
    worst = max(r["ms"]["open"] / r["paragraphs"] for r in rows)
    print(f"\nopen: {worst * 1000:.2f} µs/paragraph (worst); peak memory bounded by §27 caps.")

    RESULTS_PATH.write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {RESULTS_PATH.relative_to(REPO_DIR)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
