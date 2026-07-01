---
id: 001
title: Zero-install rendering - auto-provision LibreOffice by default
slug: 001-zero-install-rendering
status: done
tags: [area:render, type:feat, ux]
priority: P1
severity: medium
effort: L
risk: downloads/executes a third-party binary; must be checksum-verified, opt-outable, and never break the no-renderer fallback
planned_at: { commit: 3a14c01, date: 2026-07-01 }
depends_on: []
mockups:
  interface: null
  architecture: arch.md
research: null
---

# Spec 001: Zero-install rendering - auto-provision LibreOffice by default

> **Executor instructions**: This spec is portable - everything you need is in
> this file and `arch.md`. Follow it top to bottom. Run every command in the AI
> verification checklist and confirm the expected result before reporting done.
> If a STOP condition fires, stop and report - do not improvise.

## Problem

`docx_convert` (to `pdf`/`png`) and `docx_render_preview` render by shelling out
to LibreOffice (`soffice`). Today, if LibreOffice is not already installed,
`docx_convert` fails with `render_unavailable` and the user is told to "install
LibreOffice or set `DOCXENGINE_SOFFICE`". For a `pip install docxengine` tool
whose whole promise is that agents can turn a Word doc into an image to check
their work, a mandatory ~300 MB manual system install is a hassle that makes the
feature effectively off-by-default.

Goal: **rendering works out of the box.** On the first render with no local
LibreOffice, docxengine downloads an official LibreOffice build, verifies it
against the publisher's checksum, caches it per-user, and uses it - no manual
install. This is the **default UX**, opt-out via one env var. The structural
fallback stays intact for the opt-out / unsupported / offline cases.

## Mockups

- **Interface** - `null`. No user-visible UI; the surface is MCP tool behaviour
  and error/notice text, covered under Instructions.
- **Architecture** - [`arch.md`](arch.md): the `ensure_soffice()` resolution
  order (detect -> cache -> fetch), the first-render provisioning sequence
  (resolve version -> fetch sha256 -> stream+verify -> extract -> cache), and the
  verified per-platform artifact table.

## Context (self-contained)

Python-only package; source under `src/docxengine/`, tests under `tests/`,
contract under `spec/`, docs under `docs/`. Verified via `ruff` + `pytest` (see
checklist). Zero runtime dependencies (`pyproject.toml` `dependencies = []`) -
**keep it that way**; use only the standard library (`urllib`, `hashlib`,
`tarfile`, `subprocess`).

Render adapter as it is now, `src/docxengine/_render.py`:

- `detect_soffice() -> str | None` probes env `DOCXENGINE_SOFFICE`, then `PATH`,
  then `PLATFORM_DEFAULTS` (`/Applications/LibreOffice.app/Contents/MacOS/soffice`,
  `/usr/bin/soffice`).
- `render_to_file(doc, fmt, dest)` raised `render_unavailable` when
  `detect_soffice()` is `None`.
- `render_preview(doc, pages, doc_id)` returned the structural fallback
  (`renderer: "structural"`, never errors) when `detect_soffice()` is `None`.

Existing env-var conventions: `DOCXENGINE_MAX_*` (`_limits.py`),
`DOCXENGINE_ROOT` (`_paths.py`), `DOCXENGINE_SOFFICE`, `DOCXENGINE_AUTHOR`,
`DOCXENGINE_FIXED_DATE`. New vars follow the same `DOCXENGINE_*` style.

Verified live against `download.documentfoundation.org` on 2026-07-01: the
`stable/` index lists versions (latest 26.2.4); per-version artifacts and their
`.sha256` sidecars resolve for macOS (`mac/<arch>/...dmg`), Linux
(`deb/x86_64/...deb.tar.gz`), Windows (`win/x86_64/...msi`). The sidecar body is
`"<hex64>  <filename>"`. See `arch.md` for the table.

Trust model: HTTPS-only fetch from TDF; the download is SHA-256-verified against
the publisher's own sidecar before any extraction or execution. This is the same
model as any signed release download.

**Current progress (already in the working tree on branch `main`):**
`src/docxengine/_soffice.py` (the provisioner) is written; `_render.py` is wired
to `ensure_soffice()` with an adaptive `_unavailable_reason()`; `tests/test_render.py`
`_no_soffice` helper now also sets `DOCXENGINE_AUTO_FETCH_SOFFICE=0`. Remaining:
provisioner tests, docs/spec cleanup, and the real E2E. Treat the Instructions as
the full end state and reconcile the working tree to it.

## Non-goals

- Bundling LibreOffice in the wheel, or any pip-installable renderer dependency.
- Windows auto-fetch (msi/`msiexec`) - resolution returns `None` -> actionable
  message; may be a follow-up spec. macOS + Linux x86-64 are in scope.
- A pure-Python / non-LibreOffice rasterizer. Fidelity stays LibreOffice-grade.
- Changing PDF/PNG output format, the `docx_convert`/`docx_render_preview` result
  shapes, or the resource-link scheme.
- Auto-installing the PDF->PNG rasterizer (`pdftoppm`/`sips`) - unchanged.

## Instructions

Units A-C are largely done in the working tree; verify each matches this end
state and finish it. Units D (docs) and E (E2E) remain. A/B are sequential
(B wires A); C, D independent; E is last (needs A-C green).

### Unit A - provisioner module `src/docxengine/_soffice.py`

A stdlib-only module that resolves/downloads/verifies/extracts LibreOffice.
Public surface used elsewhere:

- `auto_fetch_enabled() -> bool` - `True` unless `DOCXENGINE_AUTO_FETCH_SOFFICE`
  is `0/false/no/off/""` (case-insensitive).
- `cache_root() -> Path` - `DOCXENGINE_SOFFICE_CACHE`, else `$XDG_CACHE_HOME/docxengine`,
  else `~/.cache/docxengine`.
- `resolve_version() -> str` - `DOCXENGINE_SOFFICE_VERSION` if set, else the max
  version parsed from the TDF `stable/` listing.
- `artifact_for(system, machine, version) -> Artifact | None` - macOS dmg
  (arm64->`aarch64`, else `x86-64`), Linux x86_64 deb tarball, else `None`.
- `download_verified(url, dest, expected_sha256)` - HTTPS-only, streamed to a
  `.part` file, SHA-256 enforced, size-capped, atomic rename; abort + cleanup on
  mismatch.
- `extract_deb(bytes, dest)` / `_iter_ar_members` / `_install_deb_tarball` -
  pure-Python `ar` + `data.tar.{xz,gz,bz2,zst}` extraction of TDF debs.
- `_install_dmg(archive, dest)` - `hdiutil attach` -> copy `.app` -> detach.
- `locate_soffice(tree) -> str | None` - glob known relative soffice paths.
- `provision() -> str` - orchestrate; raise `ProvisionError` on failure.
- `provision_if_enabled() -> str | None` - honour the flag, reuse cache, catch
  `ProvisionError` -> record `last_error()`, return `None` (never raises).
- `cached_soffice() -> str | None` - return a ready cached soffice via the
  `.soffice-path` marker.

Constraints: no non-stdlib imports; HTTPS-only (reject other schemes);
`_MAX_DOWNLOAD_BYTES` cap; `tarfile.extractall(..., filter="data")`.

### Unit B - wire into `src/docxengine/_render.py` (depends on A)

- Add `ensure_soffice() -> str | None`: `detect_soffice()` first, else
  `_soffice.provision_if_enabled()`.
- `render_to_file` and `render_preview` call `ensure_soffice()` (not
  `detect_soffice()`).
- `_unavailable_reason() -> str`: auto-fetch off -> mention install /
  `DOCXENGINE_SOFFICE` / enable auto-fetch (must contain the literal string
  `DOCXENGINE_SOFFICE`); else if `_soffice.last_error()` -> include it; else the
  plain "not detected" line. Used by the `render_unavailable` error and the
  preview structural note.
- Refresh the module docstring to describe the new resolution order; remove the
  stale "TypeScript twin" reference.
- `detect_soffice()` and `PLATFORM_DEFAULTS` stay defined in `_render.py`
  (tests monkeypatch `_render.PLATFORM_DEFAULTS`).

### Unit C - tests

- `tests/test_render.py`: `_no_soffice` helper also sets
  `DOCXENGINE_AUTO_FETCH_SOFFICE=0` so existing fallback/`render_unavailable`
  tests never hit the network. Existing assertions keep passing.
- New `tests/test_soffice.py`, no network (monkeypatch `urllib.request.urlopen`
  with an in-memory fake; build a synthetic `.deb`/tarball in a tmp dir). Cover:
  enable-flag parsing; `cache_root` precedence; `artifact_for` per platform incl.
  `None` cases; `resolve_version` (pinned + parsed); sha sidecar parse + reject
  malformed; `download_verified` success and checksum-mismatch (aborts, no dest,
  `.part` cleaned); `_iter_ar_members`/`extract_deb`/`locate_soffice` round-trip;
  `cached_soffice` marker; `provision_if_enabled` disabled->None, cached->path,
  failure->None+`last_error` set, success (primitives monkeypatched)->path.

### Unit D - docs + contract cleanup (the "no more install hassle" messaging)

Reframe every "install LibreOffice" instruction to "automatic, opt-out". Update:

- `docs/core/render-adapter.md` - lead with auto-provisioning; document the four
  env vars (`DOCXENGINE_AUTO_FETCH_SOFFICE`, `DOCXENGINE_SOFFICE_CACHE`,
  `DOCXENGINE_SOFFICE_VERSION`, `DOCXENGINE_SOFFICE_MIRROR`), the cache location,
  and the trust model (HTTPS + official sha256).
- `spec/algorithms.md` §24 - resolution order now includes cache + auto-fetch.
- `spec/tools/docx_convert.json`, `spec/tools/docx_render_preview.json`,
  `spec/errors.json` (`render_unavailable`) - text reflects auto-fetch default;
  then run `python scripts/sync_spec.py` to re-sync `_specdata` and commit it.
- `docs/reference/error-codes.md`, `docs/tools/errors.md`,
  `docs/conformance/fidelity.md` - align the `render_unavailable`/soffice wording.
- `SECURITY.md` - add the new env vars and the download-a-binary trust model.
- Confirm `README.md` render mention needs no "install" caveat (it already
  doesn't); add env vars to any env-var table if present.

### Unit E - real E2E (macOS, this machine)

With a clean cache and no `DOCXENGINE_SOFFICE`/PATH soffice, drive a real
`provision()` + render of a created doc to PDF and PNG; assert a real `%PDF` /
PNG file is produced and a second render hits the cache (no re-download). Run the
download in the background (large). Record the outcome; do not commit the cache.

## STOP conditions

- The working tree diverged from the cited `_render.py` shape beyond the "current
  progress" note (e.g. `detect_soffice`/`PLATFORM_DEFAULTS` moved out of
  `_render.py`) - report the drift.
- Any fix would add a non-stdlib runtime dependency, or require pinning a
  hardcoded checksum - stop; the design forbids both.
- TDF artifact URL scheme changed such that `artifact_for`/sidecar no longer
  resolve - stop and report (the live shape is cited above).
- A change would alter `docx_convert`/`docx_render_preview` result shapes or the
  structural-fallback guarantee (preview never errors) - stop.

## AI verification checklist (automatable)

- [ ] `ruff check src tests scripts` - clean
- [ ] `pytest -q` - all pass (existing 476 + new `test_soffice.py`)
- [ ] `pytest -q tests/test_render.py tests/test_soffice.py` - pass; no network
      access during the run (fully monkeypatched)
- [ ] `mypy src/docxengine/_soffice.py src/docxengine/_render.py` - clean under
      strict (config in `pyproject.toml`)
- [ ] `python scripts/sync_spec.py` - runs; `_specdata` re-synced and committed
- [ ] `grep -rn "install LibreOffice" docs spec | wc -l` - only auto-fetch-framed
      mentions remain (no bare "you must install" instruction)
- [ ] `python -c "import docxengine._soffice"` - imports with no extra deps

## Human verification checklist (judgment calls)

- [ ] First render with no local LibreOffice actually provisions + renders (Unit E
      output), and a second render reuses the cache with no download.
- [ ] `DOCXENGINE_AUTO_FETCH_SOFFICE=0` cleanly restores the old behaviour
      (structural preview / `render_unavailable`), message is actionable.
- [ ] Trust-model wording in `SECURITY.md` is accurate and sufficient for review.
- [ ] Docs read as "rendering just works", not "install this first".
- [ ] Download size / first-render latency is acceptable to ship as default-on.
