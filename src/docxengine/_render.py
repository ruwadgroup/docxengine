"""Render adapter (``docx_convert`` pdf/png, ``docx_render_preview``) — algorithms.md §24.

Detection order: env ``DOCXENGINE_SOFFICE``; then ``soffice`` on ``PATH``; then the
platform defaults (``/Applications/LibreOffice.app/Contents/MacOS/soffice``,
``/usr/bin/soffice``). When a binary is found, conversion runs
``soffice --headless --convert-to {fmt} --outdir {DIR} {FILE}`` with a per-call temp
profile (``-env:UserInstallation=file://{tmp}``) and ``renderer = "libreoffice {ver}"``.
When none is found, the **structural fallback** returns the §2 projection plus an
estimated page count (``ceil(total_chars / 1800)``) and ``renderer = "structural"``;
preview never errors, but ``docx_convert`` to pdf/png with no adapter is
``render_unavailable``.

The TypeScript twin (``render.ts``) is the cross-language reference; pdf/png byte
parity is **not** required (renderer output is non-deterministic).
"""

from __future__ import annotations

import math
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass

from . import _projector
from ._errors import ToolError
from ._session import OpenDocument, Session

CHARS_PER_PAGE = 1800

#: §24 platform default locations probed after env + PATH.
PLATFORM_DEFAULTS = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
)

_VERSION_RE = re.compile(r"([0-9]+\.[0-9]+(?:\.[0-9]+)*)")


# ---------------------------------------------------------------------------
# soffice detection (§24)
# ---------------------------------------------------------------------------


def _is_executable(path: str) -> bool:
    return os.path.isfile(path)


def detect_soffice() -> str | None:
    """Locate a usable ``soffice`` executable, or ``None`` when none is installed."""
    env = os.environ.get("DOCXENGINE_SOFFICE")
    if env and _is_executable(env):
        return env
    on_path = shutil.which("soffice")
    if on_path is not None:
        return on_path
    for candidate in PLATFORM_DEFAULTS:
        if _is_executable(candidate):
            return candidate
    return None


def _renderer_label(soffice: str) -> str:
    """The renderer label: ``"libreoffice {version}"`` (or ``"libreoffice"``)."""
    try:
        res = subprocess.run(  # noqa: S603 - soffice path is operator-controlled
            [soffice, "--version"],
            capture_output=True,
            text=True,
            timeout=20,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return "libreoffice"
    out = (res.stdout or "").strip()
    m = _VERSION_RE.search(out)
    return f"libreoffice {m.group(1)}" if m else "libreoffice"


# ---------------------------------------------------------------------------
# LibreOffice invocation
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ConversionOutcome:
    """The produced file path (in a temp outdir) plus the resolved renderer label."""

    produced_path: str
    renderer: str


def build_soffice_args(profile_dir: str, fmt: str, work_dir: str, src_docx: str) -> list[str]:
    """The §24 command line, factored out so tests can assert its construction."""
    return [
        "--headless",
        f"-env:UserInstallation=file://{profile_dir}",
        "--convert-to",
        fmt,
        "--outdir",
        work_dir,
        src_docx,
    ]


def _run_soffice(doc: OpenDocument, soffice: str, fmt: str) -> ConversionOutcome:
    """Save the doc, run soffice → ``fmt``, return the produced file path.

    ``render_failed`` on a non-zero exit or no output file.
    """
    work_dir = tempfile.mkdtemp(prefix="docxengine-render-")
    profile_dir = tempfile.mkdtemp(prefix="docxengine-profile-")
    src_docx = os.path.join(work_dir, "input.docx")
    try:
        doc.package.save(src_docx)
        args = build_soffice_args(profile_dir, fmt, work_dir, src_docx)
        try:
            res = subprocess.run(  # noqa: S603 - soffice path is operator-controlled
                [soffice, *args],
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            raise ToolError(
                "render_failed",
                f"soffice could not be invoked: {exc}.",
                ["Check that soffice runs and can write the output dir."],
            ) from exc
        if res.returncode != 0:
            raise ToolError(
                "render_failed",
                f"soffice exited {res.returncode}: {(res.stderr or '')[:200]}.",
                ["Check that the document is valid and soffice can write the output dir."],
            )
        produced = os.path.join(work_dir, f"input.{fmt}")
        if not os.path.exists(produced):
            raise ToolError(
                "render_failed",
                "soffice produced no output file.",
                ["Inspect soffice stderr; the document may have failed to load."],
            )
        return ConversionOutcome(produced, _renderer_label(soffice))
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Structural fallback (§24)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class StructuralPreview:
    """The §2 projection of the body plus an estimated page count (§24)."""

    structural: str
    estimated_pages: int


def structural_preview(doc: OpenDocument) -> StructuralPreview:
    """The §2 projection of the whole body + ``ceil(total_chars / 1800)`` pages."""
    projection = _projector.project_read(doc.package)["content"]
    estimated = max(1, math.ceil(len(projection) / CHARS_PER_PAGE))
    return StructuralPreview(projection, estimated)


# ---------------------------------------------------------------------------
# docx_convert pdf/png target
# ---------------------------------------------------------------------------


def _render_png(doc: OpenDocument, soffice: str) -> ConversionOutcome:
    """PNG = PDF then ``pdftoppm``/``sips``; ``render_failed`` when neither runs."""
    pdf = _run_soffice(doc, soffice, "pdf")
    work_dir = os.path.dirname(pdf.produced_path)
    base = os.path.join(work_dir, "page")
    out = f"{base}.png"

    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm is not None:
        res = subprocess.run(  # noqa: S603 - operator-controlled tool path
            [pdftoppm, "-png", "-singlefile", pdf.produced_path, base],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if res.returncode == 0 and os.path.exists(out):
            return ConversionOutcome(out, pdf.renderer)
    sips = shutil.which("sips")
    if sips is not None:
        res = subprocess.run(  # noqa: S603 - operator-controlled tool path
            [sips, "-s", "format", "png", pdf.produced_path, "--out", out],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if res.returncode == 0 and os.path.exists(out):
            return ConversionOutcome(out, pdf.renderer)
    raise ToolError(
        "render_failed",
        "No PDF→PNG rasterizer (pdftoppm or sips) available.",
        ["Install poppler (pdftoppm) or run on macOS (sips); pdf conversion still works."],
    )


def render_to_file(doc: OpenDocument, fmt: str, dest: str) -> dict[str, object]:
    """Convert to pdf/png via the adapter; no adapter → ``render_unavailable`` (§24)."""
    soffice = detect_soffice()
    if soffice is None:
        raise ToolError(
            "render_unavailable",
            "No render adapter: LibreOffice (soffice) was not detected.",
            [
                "Install LibreOffice or set DOCXENGINE_SOFFICE; md/html convert without it.",
                "Use docx_render_preview for the structural fallback.",
            ],
        )
    outcome = _run_soffice(doc, soffice, "pdf") if fmt == "pdf" else _render_png(doc, soffice)
    try:
        shutil.copyfile(outcome.produced_path, dest)
    except OSError as exc:
        raise ToolError(
            "save_failed",
            f"Could not write {dest}: {exc.strerror or exc}.",
            ["Check the output path and permissions."],
        ) from exc
    return {
        "path": dest,
        "renderer": outcome.renderer,
        "note": f"Rendered {fmt} via {outcome.renderer}.",
    }


# ---------------------------------------------------------------------------
# docx_render_preview (§24)
# ---------------------------------------------------------------------------


def _page_numbers(requested: list[int] | None, total: int) -> list[int]:
    if requested:
        return [max(1, int(n)) for n in requested]
    return list(range(1, total + 1))


def render_preview(
    doc: OpenDocument,
    pages: list[int] | None,
    doc_id: str,
) -> dict[str, object]:
    """Render preview pages (resource links) or the structural fallback (§24).

    Preview never errors when no renderer is installed — it returns the structural
    projection plus estimated page links.
    """
    soffice = detect_soffice()
    if soffice is None:
        fallback = structural_preview(doc)
        plural = "" if fallback.estimated_pages == 1 else "s"
        return {
            "page_count": fallback.estimated_pages,
            "renderer": "structural",
            "structural": fallback.structural,
            "note": (
                "No render adapter (LibreOffice/soffice) detected — install LibreOffice or set "
                "DOCXENGINE_SOFFICE for rendered page images. Showing the structural projection "
                f"(estimated {fallback.estimated_pages} page{plural}); no image links are returned."
            ),
        }
    renderer = _renderer_label(soffice)
    estimate = structural_preview(doc).estimated_pages
    page_set = _page_numbers(pages, estimate)
    return {
        "pages": [
            {"page": page, "image": f"docx://{doc_id}/preview/page-{page}.png"} for page in page_set
        ],
        "renderer": renderer,
        "note": f"Preview links resolve to {renderer}-rendered page images.",
    }


def docx_render_preview(
    session: Session,
    *,
    doc_id: str,
    pages: list[int] | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """Render preview pages (resource links) or the structural fallback (§24)."""
    doc = session.get(doc_id)
    return render_preview(doc, pages, doc_id)
