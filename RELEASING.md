# Releasing

DocxEngine releases are tag-driven and automated by the [release workflow](.github/workflows/release.yml). Python (`docxengine` on PyPI) and JS (`@docxengine/core` on npm) are released **in lockstep with the same version** — the conformance suite guarantees they implement the same contract, so they share a version number.

## Versioning

- [SemVer](https://semver.org). Pre-1.0: minor bumps may break; patch bumps never do.
- The tool contract in [`spec/`](spec/) carries its own schema version; a breaking contract change forces at least a minor bump and a changelog deprecation note.

## Process

1. Ensure `main` is green: CI, conformance harness, benchmark smoke run.
2. Update [CHANGELOG.md](CHANGELOG.md) — move `Unreleased` items under the new version with the date.
3. Bump versions: `python/pyproject.toml`, `js/package.json` (same number).
4. Commit: `chore(release): vX.Y.Z`, then tag: `git tag vX.Y.Z && git push origin vX.Y.Z`.
5. The release workflow builds both packages, runs the full test + conformance matrix, publishes to PyPI/npm with provenance/attestations, and drafts the GitHub release from the changelog.
6. Verify the published artifacts install cleanly (`pip install docxengine==X.Y.Z`, `npm i @docxengine/core@X.Y.Z`).

## Release checklist

- [ ] CI green on `main`
- [ ] Conformance suite passes in both implementations
- [ ] CHANGELOG updated, versions bumped in lockstep
- [ ] No unresolved security advisories
- [ ] Contract changes (if any) documented with migration notes
