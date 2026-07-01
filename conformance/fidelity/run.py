#!/usr/bin/env python3
"""Cross-renderer fidelity harness (stdlib + docxengine).

Cross-renderer fidelity asks a question the conformance suite deliberately does
NOT: does a document *look* right when a real word processor lays it out? Word,
LibreOffice, and Google Docs have different layout engines, so this is partly a
**manual protocol** (see docs/conformance/fidelity.md) — Word and Google Docs
cannot be driven headless in CI. What this harness automates:

1. **Structural fidelity (always on).** For each layout-sensitive corpus
   document it runs ``docx_render_preview`` and checks the structural projection
   is internally consistent with the document model (``docx_outline``). This
   catches projector/preview regressions everywhere, with no renderer installed.

2. **Visual rendering (when a renderer is present).** If LibreOffice (``soffice``)
   is detected, it additionally renders each document to PDF via the engine's
   render adapter and records a manifest (renderer label + output size) at
   ``conformance/fidelity/manifest.json`` — the artifact a maintainer reviews,
   and the input to the manual Word/Google-Docs comparison.

Exit status: non-zero only on a structural inconsistency, or a renderer error
when a renderer IS installed. With no renderer it reports "structural only" and
passes — visual fidelity is then the documented manual protocol.

    .venv/bin/python conformance/fidelity/run.py
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Any

HARNESS = Path(__file__).resolve().parent
CONFORMANCE_DIR = HARNESS.parent
REPO_DIR = CONFORMANCE_DIR.parent
CORPUS_DIR = CONFORMANCE_DIR / "corpus"
MANIFEST = HARNESS / "manifest.json"

sys.path.insert(0, str(CONFORMANCE_DIR / "harness"))
sys.path.insert(0, str(REPO_DIR / "src"))

# This is a deterministic regression harness: never auto-download LibreOffice.
# Visual rendering runs only when a real soffice is already present, unless the
# maintainer opts in by exporting DOCXENGINE_AUTO_FETCH_SOFFICE=1 themselves.
os.environ.setdefault("DOCXENGINE_AUTO_FETCH_SOFFICE", "0")

import make_fixtures  # noqa: E402

from docxengine import Document  # noqa: E402
from docxengine._render import detect_soffice  # noqa: E402

# Documents whose correctness is layout-sensitive (tables, lists, sections with
# headers/footers, embedded media, multi-heading flow).
LAYOUT_SENSITIVE = ["minimal", "tables", "numbered-lists", "headers-footers", "media-doc"]


def _structural_ok(doc: Document) -> list[str]:
    """Assert the structural preview is consistent with the document model."""
    reasons: list[str] = []
    preview = doc.render_preview()
    if "renderer" not in preview:
        reasons.append("render_preview returned no renderer label")
    if preview.get("renderer") == "structural":
        text = str(preview.get("structural", ""))
        if not text:
            reasons.append("structural fallback returned an empty projection")
        # Every heading the outline reports must appear in the structural text.
        entries = doc.outline().get("outline")
        for entry in entries if isinstance(entries, list) else []:
            heading = entry.get("text") if isinstance(entry, dict) else None
            if isinstance(heading, str) and heading and heading not in text:
                reasons.append(f"structural projection is missing outline heading {heading!r}")
                break
    else:
        # A real renderer must hand back at least one page.
        if not preview.get("pages"):
            reasons.append(f"renderer {preview.get('renderer')!r} produced no pages")
    return reasons


def main() -> int:
    make_fixtures.build_corpus(CORPUS_DIR, quiet=True)
    soffice = detect_soffice()
    renderer = "libreoffice" if soffice else "structural (no soffice)"
    print(f"renderer: {renderer}\n")

    failures: list[tuple[str, str]] = []
    manifest: dict[str, Any] = {"renderer": renderer, "documents": {}}

    header = f"{'document':>16}  {'structural':>11}  {'visual':>16}"
    print(header)
    print("-" * len(header))

    for name in LAYOUT_SENSITIVE:
        src = CORPUS_DIR / name / "input.docx"
        if not src.is_file():
            continue
        doc = Document.open(str(src))

        reasons = _structural_ok(doc)
        structural = "ok" if not reasons else "FAIL"
        for r in reasons:
            failures.append((name, r))

        visual = "skipped (no renderer)"
        if soffice:
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    out = Path(tmp) / f"{name}.pdf"
                    doc.convert("pdf", path=str(out))
                    size = out.stat().st_size if out.is_file() else 0
                if size <= 0:
                    visual = "FAIL (empty)"
                    failures.append((name, "renderer produced an empty PDF"))
                else:
                    visual = f"{size // 1024} KiB pdf"
                    manifest["documents"][name] = {"pdf_bytes": size}
            except Exception as exc:  # render error with a renderer present is a failure
                visual = "FAIL (error)"
                failures.append((name, f"render error: {exc}"))

        print(f"{name:>16}  {structural:>11}  {visual:>16}")

    if soffice:
        MANIFEST.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
        print(f"\nwrote {MANIFEST.relative_to(REPO_DIR)} — review and compare per the protocol.")
    else:
        print(
            "\nVisual cross-renderer fidelity is the documented manual protocol "
            "(docs/conformance/fidelity.md); structural checks ran above."
        )

    if failures:
        print(f"\n{len(failures)} fidelity issue(s):")
        for name, reason in failures:
            print(f"  {name}: {reason}")
        return 1
    print("\nall fidelity checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
