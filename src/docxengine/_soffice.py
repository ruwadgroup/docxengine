"""Auto-provision a portable LibreOffice so rendering needs no manual install.

``_render`` renders DOCX to PDF/PNG by driving LibreOffice (``soffice``). When no
``soffice`` is detected on the machine, this module downloads an official
LibreOffice build from The Document Foundation, verifies it against the
publisher's own ``.sha256`` sidecar, extracts it into a per-user cache, and
returns the ``soffice`` path - so the first render "just works" with no separate
LibreOffice install.

Trust model: artifacts and their checksums are fetched over HTTPS from
``download.documentfoundation.org``; the download is verified against the
publisher-provided SHA-256 before anything is executed. Nothing is fetched unless
auto-fetch is enabled (it is by default) and no local ``soffice`` was found.

Knobs (all optional):

- ``DOCXENGINE_AUTO_FETCH_SOFFICE`` - set to ``0``/``false`` to disable fetching.
- ``DOCXENGINE_SOFFICE_CACHE`` - cache dir (default ``$XDG_CACHE_HOME/docxengine``
  or ``~/.cache/docxengine``).
- ``DOCXENGINE_SOFFICE_VERSION`` - pin a version instead of resolving the latest.
- ``DOCXENGINE_SOFFICE_MIRROR`` - base URL of a mirror of the TDF ``stable`` tree.

Platform support: macOS (``.dmg``) and Linux x86-64 (``.deb`` tarball) are
auto-fetched. Other platforms fall back to detection with an actionable error.
"""

from __future__ import annotations

import hashlib
import io
import os
import platform
import re
import shutil
import stat
import subprocess
import tarfile
import tempfile
import urllib.request
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import cast
from urllib.error import URLError

AUTO_FETCH_ENV = "DOCXENGINE_AUTO_FETCH_SOFFICE"
CACHE_ENV = "DOCXENGINE_SOFFICE_CACHE"
VERSION_ENV = "DOCXENGINE_SOFFICE_VERSION"
MIRROR_ENV = "DOCXENGINE_SOFFICE_MIRROR"

DEFAULT_MIRROR = "https://download.documentfoundation.org/libreoffice/stable"

#: Hard ceiling on a download so a hostile/broken mirror cannot fill the disk.
_MAX_DOWNLOAD_BYTES = 3 * 1024 * 1024 * 1024
_NET_TIMEOUT = 60
_VERSION_RE = re.compile(r"\b([0-9]+\.[0-9]+\.[0-9]+)\b")

#: soffice locations relative to an extracted install tree, most-specific first.
_SOFFICE_GLOBS = (
    "LibreOffice.app/Contents/MacOS/soffice",  # macOS .dmg
    "opt/libreoffice*/program/soffice",  # Linux .deb
    "**/program/soffice",  # defensive catch-all
)

#: The last provisioning failure, surfaced by ``_render`` in its error message.
_last_error: str | None = None


class ProvisionError(RuntimeError):
    """A download / verification / extraction step failed."""


@dataclass(frozen=True, slots=True)
class Artifact:
    """One downloadable LibreOffice build for a specific OS/arch."""

    url: str
    kind: str  # "dmg" | "deb-tar"


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------


def auto_fetch_enabled() -> bool:
    """Whether auto-provisioning is allowed (default ``True``)."""
    raw = os.environ.get(AUTO_FETCH_ENV)
    if raw is None:
        return True
    return raw.strip().lower() not in {"0", "false", "no", "off", ""}


def cache_root() -> Path:
    """The base cache directory for provisioned LibreOffice builds."""
    override = os.environ.get(CACHE_ENV)
    if override:
        return Path(override).expanduser()
    xdg = os.environ.get("XDG_CACHE_HOME")
    base = Path(xdg).expanduser() if xdg else Path.home() / ".cache"
    return base / "docxengine"


def last_error() -> str | None:
    """The reason the most recent provisioning attempt failed, if any."""
    return _last_error


# ---------------------------------------------------------------------------
# Version + artifact resolution
# ---------------------------------------------------------------------------


def _mirror() -> str:
    return os.environ.get(MIRROR_ENV, DEFAULT_MIRROR).rstrip("/")


def _http_get(url: str) -> bytes:
    if not url.lower().startswith("https://"):
        raise ProvisionError(f"refusing non-HTTPS URL: {url}")
    req = urllib.request.Request(url, headers={"User-Agent": "docxengine"})
    try:
        with urllib.request.urlopen(req, timeout=_NET_TIMEOUT) as resp:  # noqa: S310 - https enforced
            data: bytes = resp.read(_MAX_DOWNLOAD_BYTES + 1)
    except (URLError, OSError, TimeoutError) as exc:
        raise ProvisionError(f"could not fetch {url}: {exc}") from exc
    if len(data) > _MAX_DOWNLOAD_BYTES:
        raise ProvisionError(f"response from {url} exceeds the size cap")
    return data


def resolve_version() -> str:
    """The pinned (``DOCXENGINE_SOFFICE_VERSION``) or latest-stable version."""
    pinned = os.environ.get(VERSION_ENV)
    if pinned:
        return pinned.strip()
    listing = _http_get(f"{_mirror()}/").decode("utf-8", "replace")
    versions = sorted(
        {m.group(1) for m in _VERSION_RE.finditer(listing)},
        key=lambda v: tuple(int(p) for p in v.split(".")),
    )
    if not versions:
        raise ProvisionError("could not determine the latest LibreOffice version")
    return versions[-1]


def artifact_for(system: str, machine: str, version: str) -> Artifact | None:
    """The download for this OS/arch, or ``None`` when auto-fetch is unsupported."""
    base = f"{_mirror()}/{version}"
    mach = machine.lower()
    if system == "Darwin":
        arch = "aarch64" if mach in {"arm64", "aarch64"} else "x86-64"
        return Artifact(f"{base}/mac/{arch}/LibreOffice_{version}_MacOS_{arch}.dmg", "dmg")
    if system == "Linux" and mach in {"x86_64", "amd64"}:
        name = f"LibreOffice_{version}_Linux_x86-64_deb.tar.gz"
        return Artifact(f"{base}/deb/x86_64/{name}", "deb-tar")
    return None


def _expected_sha256(artifact_url: str) -> str:
    body = _http_get(f"{artifact_url}.sha256").decode("utf-8", "replace").strip()
    token = body.split()[0] if body else ""
    if not re.fullmatch(r"[0-9a-fA-F]{64}", token):
        raise ProvisionError(f"malformed checksum sidecar for {artifact_url}")
    return token.lower()


# ---------------------------------------------------------------------------
# Download (streamed, checksum-verified)
# ---------------------------------------------------------------------------


def download_verified(url: str, dest: Path, expected_sha256: str) -> None:
    """Stream ``url`` to ``dest``, aborting unless the SHA-256 matches."""
    if not url.lower().startswith("https://"):
        raise ProvisionError(f"refusing non-HTTPS URL: {url}")
    digest = hashlib.sha256()
    total = 0
    tmp = dest.with_suffix(dest.suffix + ".part")
    req = urllib.request.Request(url, headers={"User-Agent": "docxengine"})
    try:
        with urllib.request.urlopen(req, timeout=_NET_TIMEOUT) as resp, tmp.open("wb") as fh:  # noqa: S310
            while chunk := resp.read(1024 * 256):
                total += len(chunk)
                if total > _MAX_DOWNLOAD_BYTES:
                    raise ProvisionError(f"download of {url} exceeds the size cap")
                digest.update(chunk)
                fh.write(chunk)
    except (URLError, OSError, TimeoutError) as exc:
        tmp.unlink(missing_ok=True)
        raise ProvisionError(f"download failed for {url}: {exc}") from exc
    actual = digest.hexdigest()
    if actual != expected_sha256:
        tmp.unlink(missing_ok=True)
        raise ProvisionError(
            f"checksum mismatch for {url}: expected {expected_sha256}, got {actual}"
        )
    tmp.replace(dest)


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------


def locate_soffice(tree: Path) -> str | None:
    """Find the ``soffice`` executable inside an extracted install tree."""
    for pattern in _SOFFICE_GLOBS:
        for hit in sorted(tree.glob(pattern)):
            if hit.is_file():
                return str(hit)
    return None


def _make_executable(path: str) -> None:
    mode = os.stat(path).st_mode
    os.chmod(path, mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)


def _iter_ar_members(data: bytes) -> Iterator[tuple[str, bytes]]:
    """Yield ``(name, payload)`` for each member of a Unix ``ar`` archive (a .deb)."""
    if not data.startswith(b"!<arch>\n"):
        raise ProvisionError("not an ar archive")
    offset = 8
    while offset + 60 <= len(data):
        header = data[offset : offset + 60]
        name = header[0:16].decode("ascii", "replace").strip()
        size = int(header[48:58].decode("ascii", "replace").strip() or "0")
        start = offset + 60
        yield name.rstrip("/"), data[start : start + size]
        offset = start + size + (size & 1)  # members are 2-byte aligned


def _open_data_tar(member_name: str, payload: bytes) -> tarfile.TarFile:
    suffix = member_name.rsplit(".", 1)[-1]
    modes = {"xz": "r:xz", "gz": "r:gz", "bz2": "r:bz2", "zst": "r:zst"}
    mode = modes.get(suffix)
    if mode is None:
        raise ProvisionError(f"unsupported deb data compression: {member_name}")
    try:
        opened = tarfile.open(fileobj=io.BytesIO(payload), mode=mode)  # type: ignore[call-overload]  # noqa: SIM115
        return cast("tarfile.TarFile", opened)
    except (tarfile.TarError, RuntimeError, OSError) as exc:
        # zstd needs Python 3.14+ (or a zstandard backend); surface it clearly.
        raise ProvisionError(f"could not open deb payload {member_name}: {exc}") from exc


def extract_deb(deb_bytes: bytes, dest: Path) -> None:
    """Extract the ``data.tar.*`` payload of a ``.deb`` into ``dest``."""
    for name, payload in _iter_ar_members(deb_bytes):
        if name.startswith("data.tar"):
            with _open_data_tar(name, payload) as tar:
                tar.extractall(dest, filter="data")  # noqa: S202 - trusted TDF artifact
            return
    raise ProvisionError("deb archive had no data.tar member")


def _install_deb_tarball(archive: Path, dest: Path) -> None:
    """Explode a TDF ``*_deb.tar.gz`` (a bundle of .deb files) into ``dest``."""
    with tempfile.TemporaryDirectory(prefix="docxengine-deb-") as tmp:
        with tarfile.open(archive, "r:gz") as tar:
            tar.extractall(tmp, filter="data")  # noqa: S202 - trusted TDF artifact
        debs = sorted(Path(tmp).rglob("*.deb"))
        if not debs:
            raise ProvisionError("deb tarball contained no .deb packages")
        for deb in debs:
            extract_deb(deb.read_bytes(), dest)


def _install_dmg(archive: Path, dest: Path) -> None:
    """Mount a macOS ``.dmg`` and copy ``LibreOffice.app`` into ``dest``."""
    mount = tempfile.mkdtemp(prefix="docxengine-dmg-")
    try:
        res = subprocess.run(  # noqa: S603 - fixed argv, operator-triggered
            ["/usr/bin/hdiutil", "attach", "-nobrowse", "-readonly",
             "-mountpoint", mount, str(archive)],
            capture_output=True, text=True, timeout=180, check=False,
        )
        if res.returncode != 0:
            raise ProvisionError(f"hdiutil attach failed: {(res.stderr or '').strip()[:200]}")
        apps = sorted(Path(mount).glob("*.app"))
        if not apps:
            raise ProvisionError("no .app bundle inside the dmg")
        dest.mkdir(parents=True, exist_ok=True)
        target = dest / apps[0].name
        shutil.rmtree(target, ignore_errors=True)
        shutil.copytree(apps[0], target, symlinks=True)
    finally:
        subprocess.run(  # noqa: S603 - fixed argv
            ["/usr/bin/hdiutil", "detach", "-force", mount],
            capture_output=True, text=True, timeout=60, check=False,
        )
        shutil.rmtree(mount, ignore_errors=True)


def _install(artifact: Artifact, archive: Path, dest: Path) -> None:
    if artifact.kind == "dmg":
        _install_dmg(archive, dest)
    elif artifact.kind == "deb-tar":
        _install_deb_tarball(archive, dest)
    else:  # pragma: no cover - guarded by artifact_for
        raise ProvisionError(f"unknown artifact kind: {artifact.kind}")


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------


def _install_dir(version: str) -> Path:
    return cache_root() / "libreoffice" / version


def _ready_marker(version: str) -> Path:
    return _install_dir(version) / ".soffice-path"


def cached_soffice() -> str | None:
    """Return a previously provisioned ``soffice`` for a ready version, else ``None``."""
    root = cache_root() / "libreoffice"
    if not root.is_dir():
        return None
    for version_dir in sorted(root.iterdir(), reverse=True):
        marker = version_dir / ".soffice-path"
        if marker.is_file():
            path = marker.read_text(encoding="utf-8").strip()
            if path and os.path.isfile(path):
                return path
    return None


def provision() -> str:
    """Download, verify, and install LibreOffice; return the ``soffice`` path.

    Raises :class:`ProvisionError` on any failure. Callers that want a
    non-raising, opt-in path should use :func:`provision_if_enabled`.
    """
    system = platform.system()
    version = resolve_version()
    artifact = artifact_for(system, platform.machine(), version)
    if artifact is None:
        raise ProvisionError(
            f"auto-fetch is not supported on {system}/{platform.machine()}; "
            "install LibreOffice or set DOCXENGINE_SOFFICE"
        )
    install_dir = _install_dir(version)
    install_dir.mkdir(parents=True, exist_ok=True)
    sha = _expected_sha256(artifact.url)
    with tempfile.TemporaryDirectory(prefix="docxengine-dl-") as tmp:
        archive = Path(tmp) / artifact.url.rsplit("/", 1)[-1]
        download_verified(artifact.url, archive, sha)
        _install(artifact, archive, install_dir)
    soffice = locate_soffice(install_dir)
    if soffice is None:
        raise ProvisionError("provisioned LibreOffice but could not locate soffice")
    _make_executable(soffice)
    _ready_marker(version).write_text(soffice, encoding="utf-8")
    return soffice


def provision_if_enabled() -> str | None:
    """Return a cached or freshly provisioned ``soffice``, or ``None``.

    Never raises: honours the auto-fetch flag, reuses a cached install, and on a
    provisioning failure records the reason in :func:`last_error` and returns
    ``None`` (so preview degrades to the structural fallback).
    """
    global _last_error
    if not auto_fetch_enabled():
        return None
    cached = cached_soffice()
    if cached is not None:
        return cached
    try:
        soffice = provision()
    except ProvisionError as exc:
        _last_error = str(exc)
        return None
    _last_error = None
    return soffice
