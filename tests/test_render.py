"""Render adapter tests (algorithms.md §24).

The structural fallback is the locally-tested real path (no LibreOffice on this
machine): docx_render_preview returns `renderer: "structural"` plus the §2 projection
and estimated page links; docx_convert to pdf/png with no adapter raises
render_unavailable. The LibreOffice path is exercised via a stub `soffice` shell
script on a temp PATH that writes a canned PDF — we assert command construction and
the result shape.
"""

from __future__ import annotations

import os
import stat

import pytest

from docxengine import (
    Session,
    ToolError,
    _render,
    docx_convert,
    docx_create,
    docx_render_preview,
)


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)
    monkeypatch.delenv("DOCXENGINE_SOFFICE", raising=False)


def _no_soffice(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:  # noqa: ANN001
    """Force detection failure: empty PATH and no platform default."""
    empty = tmp_path / "empty-path"
    empty.mkdir()
    monkeypatch.setenv("PATH", str(empty))
    monkeypatch.setattr(_render, "PLATFORM_DEFAULTS", ())
    monkeypatch.setenv("DOCXENGINE_AUTO_FETCH_SOFFICE", "0")


def _create(session: Session) -> str:
    return str(docx_create(session, content_md="# Title\n\nBody.")["doc_id"])


# ---------------------------------------------------------------------------
# Structural fallback (no soffice)
# ---------------------------------------------------------------------------


class TestStructuralFallback:
    def test_preview_returns_structural(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        _no_soffice(monkeypatch, tmp_path)
        session = Session()
        doc_id = _create(session)
        result = docx_render_preview(session, doc_id=doc_id)
        assert result["renderer"] == "structural"
        assert "structural" in result
        assert "[P1#" in str(result["structural"])  # the §2 projection

    def test_preview_fallback_uses_page_count(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        _no_soffice(monkeypatch, tmp_path)
        session = Session()
        doc_id = _create(session)
        result = docx_render_preview(session, doc_id=doc_id, pages=[1, 2])
        # No renderer ran, so no per-page array — just a compact page_count estimate.
        assert "pages" not in result
        assert isinstance(result["page_count"], int) and result["page_count"] >= 1
        assert "DOCXENGINE_SOFFICE" in str(result["note"])

    def test_preview_never_errors_without_renderer(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        _no_soffice(monkeypatch, tmp_path)
        session = Session()
        doc_id = _create(session)
        # No exception; structural fallback is returned.
        result = docx_render_preview(session, doc_id=doc_id)
        assert result["page_count"] >= 1

    def test_estimated_page_count_min_one(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        _no_soffice(monkeypatch, tmp_path)
        session = Session()
        doc_id = _create(session)
        preview = _render.structural_preview(session.get(doc_id))
        assert preview.estimated_pages >= 1

    def test_convert_pdf_render_unavailable(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        _no_soffice(monkeypatch, tmp_path)
        session = Session()
        doc_id = _create(session)
        with pytest.raises(ToolError) as exc:
            docx_convert(session, doc_id=doc_id, to="pdf", path=str(tmp_path / "o.pdf"))
        assert exc.value.code == "render_unavailable"


# ---------------------------------------------------------------------------
# Command construction
# ---------------------------------------------------------------------------


class TestCommandConstruction:
    def test_soffice_args_shape(self) -> None:
        args = _render.build_soffice_args("/tmp/prof", "pdf", "/tmp/work", "/tmp/work/input.docx")
        assert args == [
            "--headless",
            "-env:UserInstallation=file:///tmp/prof",
            "--convert-to",
            "pdf",
            "--outdir",
            "/tmp/work",
            "/tmp/work/input.docx",
        ]


# ---------------------------------------------------------------------------
# LibreOffice path via a stub soffice shell script
# ---------------------------------------------------------------------------

_STUB_SOFFICE = """#!/bin/sh
# Stub soffice: --version prints a version; conversion writes a canned PDF.
for arg in "$@"; do
  case "$arg" in
    --version) echo "LibreOffice 24.8.1.2 abcdef"; exit 0;;
  esac
done
# Parse --outdir and the trailing input file.
outdir=""
src=""
prev=""
for arg in "$@"; do
  if [ "$prev" = "--outdir" ]; then outdir="$arg"; fi
  prev="$arg"
  src="$arg"
done
base=$(basename "$src")
name="${base%.*}"
printf '%%PDF-1.4\\nstub\\n%%%%EOF\\n' > "$outdir/$name.pdf"
exit 0
"""


def _install_stub_soffice(tmp_path) -> str:  # noqa: ANN001
    bindir = tmp_path / "bin"
    bindir.mkdir()
    soffice = bindir / "soffice"
    soffice.write_text(_STUB_SOFFICE)
    soffice.chmod(soffice.stat().st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)
    return str(soffice)


class TestLibreOfficePath:
    def test_detect_via_env(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        stub = _install_stub_soffice(tmp_path)
        monkeypatch.setenv("DOCXENGINE_SOFFICE", stub)
        assert _render.detect_soffice() == stub

    def test_detect_via_path(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        stub = _install_stub_soffice(tmp_path)
        monkeypatch.setenv("PATH", os.path.dirname(stub))
        monkeypatch.setattr(_render, "PLATFORM_DEFAULTS", ())
        assert _render.detect_soffice() == stub

    def test_renderer_label_from_version(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        stub = _install_stub_soffice(tmp_path)
        assert _render._renderer_label(stub) == "libreoffice 24.8.1.2"

    def test_convert_pdf_via_stub(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        stub = _install_stub_soffice(tmp_path)
        monkeypatch.setenv("DOCXENGINE_SOFFICE", stub)
        session = Session()
        doc_id = _create(session)
        dest = tmp_path / "out.pdf"
        result = docx_convert(session, doc_id=doc_id, to="pdf", path=str(dest))
        assert result["path"] == str(dest)
        assert result["renderer"] == "libreoffice 24.8.1.2"
        assert dest.read_bytes().startswith(b"%PDF")

    def test_render_to_file_result_shape(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        stub = _install_stub_soffice(tmp_path)
        monkeypatch.setenv("DOCXENGINE_SOFFICE", stub)
        session = Session()
        doc_id = _create(session)
        dest = tmp_path / "out.pdf"
        result = _render.render_to_file(session.get(doc_id), "pdf", str(dest))
        assert set(result) == {"path", "renderer", "note"}
        assert "Rendered pdf via libreoffice" in str(result["note"])

    def test_preview_with_stub_reports_renderer(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        stub = _install_stub_soffice(tmp_path)
        monkeypatch.setenv("DOCXENGINE_SOFFICE", stub)
        session = Session()
        doc_id = _create(session)
        result = docx_render_preview(session, doc_id=doc_id, pages=[1])
        assert result["renderer"] == "libreoffice 24.8.1.2"
        assert result["pages"][0]["image"] == f"docx://{doc_id}/preview/page-1.png"

    def test_render_failed_on_nonzero_exit(self, monkeypatch, tmp_path) -> None:  # noqa: ANN001
        bindir = tmp_path / "bin"
        bindir.mkdir()
        soffice = bindir / "soffice"
        soffice.write_text("#!/bin/sh\nexit 1\n")
        soffice.chmod(soffice.stat().st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)
        monkeypatch.setenv("DOCXENGINE_SOFFICE", str(soffice))
        session = Session()
        doc_id = _create(session)
        with pytest.raises(ToolError) as exc:
            docx_convert(session, doc_id=doc_id, to="pdf", path=str(tmp_path / "o.pdf"))
        assert exc.value.code == "render_failed"
