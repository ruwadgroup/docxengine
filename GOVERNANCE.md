# Governance

DocxEngine is an open-source project maintained under the [Ruwad Group](https://github.com/ruwadgroup) organization.

## Roles

- **Maintainers** — own technical direction, review/merge PRs, cut releases, and steward the tool contract in [`spec/`](spec/). Current maintainers are listed in [.github/CODEOWNERS](.github/CODEOWNERS).
- **Contributors** — anyone with a merged PR, a corpus document, a benchmark task, or substantive design review. Sustained, high-quality contribution is the path to maintainership (invited by existing maintainers).

## Decision making

- **Lazy consensus** for routine changes: a PR with maintainer approval and no objections merges.
- **Proposals** for larger changes — anything touching the tool contract (`spec/`), the anchor scheme, the validation gate, or a project invariant starts as a GitHub Discussion or a `docs/` design PR, open for comment before implementation.
- Maintainers make the final call on technical merit; disagreements are resolved by the maintainer group, not by seniority of opinion.

## The tool contract is a stability surface

The JSON Schemas in [`spec/`](spec/) are the public contract consumed by MCP clients, OpenAI function-calling, and both SDKs:

- Breaking changes require a versioned schema revision and a deprecation note in the changelog.
- Additive changes (new optional fields, new tools) are preferred over mutations.
- Both implementations and the conformance harness must land in the same release as any contract change.

## Releases

Tag-driven, coordinated by maintainers — see [RELEASING.md](RELEASING.md).

## Code of Conduct

Enforcement of the [Code of Conduct](CODE_OF_CONDUCT.md) is the maintainers' responsibility; reports are handled confidentially (see [SECURITY.md](SECURITY.md) for the private channel).
