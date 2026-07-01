# Releasing

DocxEngine releases are tag-driven and automated by the [release workflow](.github/workflows/release.yml). `docxengine` is published to **PyPI** via trusted publishing (OIDC) — no long-lived tokens.

## Versioning

- [SemVer](https://semver.org).
- The tool contract in [`spec/`](spec/) carries its own schema version; a breaking contract change forces at least a minor bump and a changelog deprecation note.

## Process

1. Ensure `main` is green: CI + test suite.
2. Update [CHANGELOG.md](CHANGELOG.md) — move `Unreleased` items under the new version with the date.
3. Bump the version in `pyproject.toml` (and the runtime `__version__` in `src/docxengine/__init__.py`).
4. Commit: `chore(release): vX.Y.Z`, then tag: `git tag vX.Y.Z && git push origin vX.Y.Z`.
5. Run the release workflow (manual `workflow_dispatch`, or re-enable the tag trigger). It runs the full test suite, builds the sdist + wheel, publishes to PyPI via trusted publishing, and drafts the GitHub release from the changelog.
6. Verify the published artifact installs cleanly (`pip install docxengine==X.Y.Z`) and the server runs (`uvx docxengine-mcp`).

## One-time PyPI trusted-publishing setup

Configure a [Trusted Publisher](https://docs.pypi.org/trusted-publishers/) on the `docxengine` PyPI project pointing at: repo `ruwadgroup/docxengine`, workflow `release.yml`, environment `release`. No API token is stored anywhere.

## Release checklist

- [ ] CI green on `main`
- [ ] CHANGELOG updated, version bumped
- [ ] No unresolved security advisories
- [ ] Contract changes (if any) documented with migration notes
