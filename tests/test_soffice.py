"""Provisioner tests (algorithms.md §24, spec 001).

Fully offline: the network layer (``urllib.request.urlopen``) is monkeypatched
with an in-memory fake, and archives are synthesised in a tmp dir. No real
LibreOffice download happens here — the real end-to-end fetch is exercised
separately (and manually) on a machine with network access.
"""

from __future__ import annotations

import hashlib
import io
import tarfile
import urllib.error
import urllib.request
from pathlib import Path

import pytest

from docxengine import _soffice


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    for var in (
        "DOCXENGINE_AUTO_FETCH_SOFFICE",
        "DOCXENGINE_SOFFICE_CACHE",
        "DOCXENGINE_SOFFICE_VERSION",
        "DOCXENGINE_SOFFICE_MIRROR",
        "XDG_CACHE_HOME",
    ):
        monkeypatch.delenv(var, raising=False)
    _soffice._last_error = None


# ---------------------------------------------------------------------------
# Fake HTTP layer
# ---------------------------------------------------------------------------


class _FakeResponse:
    def __init__(self, data: bytes) -> None:
        self._buf = io.BytesIO(data)

    def read(self, size: int = -1) -> bytes:
        return self._buf.read() if size is None or size < 0 else self._buf.read(size)

    def __enter__(self) -> _FakeResponse:
        return self

    def __exit__(self, *_: object) -> bool:
        return False


def _install_fake_http(monkeypatch: pytest.MonkeyPatch, routes: dict[str, bytes]) -> None:
    def fake_urlopen(req: object, timeout: float = 0) -> _FakeResponse:  # noqa: ARG001
        url = req.full_url if hasattr(req, "full_url") else str(req)
        if url not in routes:
            raise urllib.error.URLError(f"no route: {url}")
        return _FakeResponse(routes[url])

    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)


# ---------------------------------------------------------------------------
# Config knobs
# ---------------------------------------------------------------------------


class TestConfig:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [(None, True), ("1", True), ("yes", True), ("0", False), ("false", False),
         ("OFF", False), ("", False)],
    )
    def test_auto_fetch_enabled(
        self, monkeypatch: pytest.MonkeyPatch, value: str | None, expected: bool
    ) -> None:
        if value is None:
            monkeypatch.delenv("DOCXENGINE_AUTO_FETCH_SOFFICE", raising=False)
        else:
            monkeypatch.setenv("DOCXENGINE_AUTO_FETCH_SOFFICE", value)
        assert _soffice.auto_fetch_enabled() is expected

    def test_cache_root_explicit(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        monkeypatch.setenv("DOCXENGINE_SOFFICE_CACHE", str(tmp_path / "c"))
        assert _soffice.cache_root() == tmp_path / "c"

    def test_cache_root_xdg(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        monkeypatch.setenv("XDG_CACHE_HOME", str(tmp_path))
        assert _soffice.cache_root() == tmp_path / "docxengine"


# ---------------------------------------------------------------------------
# Version + artifact resolution
# ---------------------------------------------------------------------------


class TestResolution:
    def test_version_pinned(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_SOFFICE_VERSION", "25.8.1")
        assert _soffice.resolve_version() == "25.8.1"

    def test_version_from_listing(self, monkeypatch: pytest.MonkeyPatch) -> None:
        listing = '<a href="26.2.4/">26.2.4/</a> <a href="26.2.10/">26.2.10/</a> 7.6.7/'
        _install_fake_http(monkeypatch, {f"{_soffice.DEFAULT_MIRROR}/": listing.encode()})
        assert _soffice.resolve_version() == "26.2.10"  # numeric, not lexical

    def test_artifact_macos_arm(self) -> None:
        art = _soffice.artifact_for("Darwin", "arm64", "26.2.4")
        assert art is not None and art.kind == "dmg"
        assert art.url.endswith("mac/aarch64/LibreOffice_26.2.4_MacOS_aarch64.dmg")

    def test_artifact_macos_intel(self) -> None:
        art = _soffice.artifact_for("Darwin", "x86_64", "26.2.4")
        assert art is not None and "MacOS_x86-64.dmg" in art.url

    def test_artifact_linux(self) -> None:
        art = _soffice.artifact_for("Linux", "x86_64", "26.2.4")
        assert art is not None and art.kind == "deb-tar"
        assert art.url.endswith("deb/x86_64/LibreOffice_26.2.4_Linux_x86-64_deb.tar.gz")

    @pytest.mark.parametrize(
        ("system", "machine"), [("Linux", "aarch64"), ("Windows", "x86_64"), ("FreeBSD", "amd64")]
    )
    def test_artifact_unsupported(self, system: str, machine: str) -> None:
        assert _soffice.artifact_for(system, machine, "26.2.4") is None

    def test_sha_sidecar_parsed(self, monkeypatch: pytest.MonkeyPatch) -> None:
        url = "https://example.com/x.dmg"
        sha = "a" * 64
        _install_fake_http(monkeypatch, {f"{url}.sha256": f"{sha}  x.dmg\n".encode()})
        assert _soffice._expected_sha256(url) == sha

    def test_sha_sidecar_malformed(self, monkeypatch: pytest.MonkeyPatch) -> None:
        url = "https://example.com/x.dmg"
        _install_fake_http(monkeypatch, {f"{url}.sha256": b"not-a-hash\n"})
        with pytest.raises(_soffice.ProvisionError):
            _soffice._expected_sha256(url)


# ---------------------------------------------------------------------------
# Download + verify
# ---------------------------------------------------------------------------


class TestDownload:
    def test_download_verified_ok(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        payload = b"the-binary-bytes" * 1000
        url = "https://example.com/a.bin"
        _install_fake_http(monkeypatch, {url: payload})
        dest = tmp_path / "a.bin"
        _soffice.download_verified(url, dest, hashlib.sha256(payload).hexdigest())
        assert dest.read_bytes() == payload
        assert not dest.with_suffix(".bin.part").exists()

    def test_download_checksum_mismatch(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        url = "https://example.com/a.bin"
        _install_fake_http(monkeypatch, {url: b"tampered"})
        dest = tmp_path / "a.bin"
        with pytest.raises(_soffice.ProvisionError, match="checksum mismatch"):
            _soffice.download_verified(url, dest, "b" * 64)
        assert not dest.exists()
        assert not dest.with_suffix(".bin.part").exists()

    def test_download_rejects_non_https(self, tmp_path: Path) -> None:
        with pytest.raises(_soffice.ProvisionError, match="non-HTTPS"):
            _soffice.download_verified("http://x/a.bin", tmp_path / "a.bin", "a" * 64)


# ---------------------------------------------------------------------------
# Extraction + location
# ---------------------------------------------------------------------------


def _make_ar(members: dict[str, bytes]) -> bytes:
    out = bytearray(b"!<arch>\n")
    for name, payload in members.items():
        header = f"{name:<16}{'0':<12}{'0':<6}{'0':<6}{'644':<8}{len(payload):<10}`\n"
        out += header.encode("ascii")
        out += payload
        if len(payload) % 2:
            out += b"\n"
    return bytes(out)


def _make_data_tar_gz(files: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w:gz") as tar:
        for name, data in files.items():
            info = tarfile.TarInfo(name)
            info.size = len(data)
            tar.addfile(info, io.BytesIO(data))
    return buf.getvalue()


class TestExtraction:
    def test_extract_deb_roundtrip(self, tmp_path: Path) -> None:
        data_tar = _make_data_tar_gz({"opt/libreoffice26.2/program/soffice": b"#!/bin/sh\n"})
        deb = _make_ar({"debian-binary": b"2.0\n", "data.tar.gz": data_tar})
        dest = tmp_path / "root"
        _soffice.extract_deb(deb, dest)
        soffice = _soffice.locate_soffice(dest)
        assert soffice is not None and soffice.endswith("program/soffice")
        assert Path(soffice).read_bytes() == b"#!/bin/sh\n"

    def test_extract_deb_no_data_member(self, tmp_path: Path) -> None:
        deb = _make_ar({"debian-binary": b"2.0\n"})
        with pytest.raises(_soffice.ProvisionError, match="no data.tar"):
            _soffice.extract_deb(deb, tmp_path / "root")

    def test_locate_soffice_macos_layout(self, tmp_path: Path) -> None:
        app = tmp_path / "LibreOffice.app" / "Contents" / "MacOS"
        app.mkdir(parents=True)
        (app / "soffice").write_bytes(b"bin")
        assert _soffice.locate_soffice(tmp_path) == str(app / "soffice")

    def test_locate_soffice_absent(self, tmp_path: Path) -> None:
        assert _soffice.locate_soffice(tmp_path) is None

    def test_install_deb_tarball(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        data_tar = _make_data_tar_gz({"opt/libreoffice26.2/program/soffice": b"bin"})
        deb = _make_ar({"data.tar.gz": data_tar})
        outer = io.BytesIO()
        with tarfile.open(fileobj=outer, mode="w:gz") as tar:
            info = tarfile.TarInfo("DEBS/core.deb")
            info.size = len(deb)
            tar.addfile(info, io.BytesIO(deb))
        archive = tmp_path / "bundle.tar.gz"
        archive.write_bytes(outer.getvalue())
        dest = tmp_path / "root"
        _soffice._install_deb_tarball(archive, dest)
        assert _soffice.locate_soffice(dest) is not None


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------


class TestOrchestration:
    def test_disabled_returns_none(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        monkeypatch.setenv("DOCXENGINE_AUTO_FETCH_SOFFICE", "0")
        monkeypatch.setenv("DOCXENGINE_SOFFICE_CACHE", str(tmp_path))
        assert _soffice.provision_if_enabled() is None

    def test_cached_hit(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        monkeypatch.setenv("DOCXENGINE_AUTO_FETCH_SOFFICE", "1")
        monkeypatch.setenv("DOCXENGINE_SOFFICE_CACHE", str(tmp_path))
        version_dir = tmp_path / "libreoffice" / "26.2.4"
        version_dir.mkdir(parents=True)
        soffice = version_dir / "program" / "soffice"
        soffice.parent.mkdir(parents=True)
        soffice.write_bytes(b"bin")
        (version_dir / ".soffice-path").write_text(str(soffice))
        assert _soffice.cached_soffice() == str(soffice)
        assert _soffice.provision_if_enabled() == str(soffice)

    def test_provision_failure_records_error(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        monkeypatch.setenv("DOCXENGINE_AUTO_FETCH_SOFFICE", "1")
        monkeypatch.setenv("DOCXENGINE_SOFFICE_CACHE", str(tmp_path))

        def boom() -> str:
            raise _soffice.ProvisionError("network down")

        monkeypatch.setattr(_soffice, "provision", boom)
        assert _soffice.provision_if_enabled() is None
        assert _soffice.last_error() == "network down"

    def test_provision_success(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        monkeypatch.setenv("DOCXENGINE_AUTO_FETCH_SOFFICE", "1")
        monkeypatch.setenv("DOCXENGINE_SOFFICE_CACHE", str(tmp_path))
        monkeypatch.setattr(_soffice, "provision", lambda: "/cached/soffice")
        assert _soffice.provision_if_enabled() == "/cached/soffice"
        assert _soffice.last_error() is None
