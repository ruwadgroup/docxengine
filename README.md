<div align="center">

# DocxEngine

**The fast, reliable way for AI agents to create and edit Word documents.**

DocxEngine lets an AI agent open a `.docx` file, read it back as clean text, and make precise edits without corrupting the file or dropping the tracked changes, comments, and formatting that other tools throw away.
It runs as an [MCP](https://modelcontextprotocol.io) server, so it plugs straight into Claude Code, Codex, Cursor, and any other MCP client.

[![License: Apache-2.0](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![CI](https://github.com/ruwadgroup/docxengine/actions/workflows/ci.yml/badge.svg)](https://github.com/ruwadgroup/docxengine/actions/workflows/ci.yml)
[![Python ≥3.12](https://img.shields.io/badge/python-%E2%89%A53.12-blue)](https://pypi.org/project/docxengine/)
[![MCP](https://img.shields.io/badge/MCP-stdio%20%2B%20streamable--http-8A2BE2)](docs/mcp/server.md)
[![Release](https://img.shields.io/github/v/tag/ruwadgroup/docxengine?sort=semver&label=release&color=blue)](https://github.com/ruwadgroup/docxengine/releases)

[Install](#install) · [Quickstart](docs/start/quickstart.md) · [Tool reference](docs/tools/index.md) · [MCP server](docs/mcp/server.md) · [Docs](docs/README.md)

</div>

---

## Why agents need this

Ask an agent to "accept Jane's edits and update the effective date" and most tooling falls apart.
python-docx silently drops tracked changes, plain find-and-replace misses text because Word splits it across hidden boundaries, and a single bad write makes Word pop a "repair" dialog when the file is opened.

DocxEngine avoids all of that by editing the document format directly and checking every change before it saves.
What the agent gets:

- **Clean, token-light reads.** Documents come back as Markdown-like text with stable line IDs, never raw XML, so the agent spends its context on content instead of markup.
- **Edits that stick.** Replace, insert, delete, or rewrite any paragraph without disturbing tracked changes, comments, footnotes, styles, or images.
- **Real redlines.** Write genuine tracked changes as a named author, then accept or reject them filtered by author or date.
- **Safe by default.** Every edit is validated against the Word format before it is written, so files open cleanly and Word never "repairs" them.
- **A way to check its work.** Render any document to PDF or PNG so the agent can visually confirm the result.

The full surface covers comments, tables, styles, sections, lists, images, fields, and mustache-style template fill, plus Markdown ↔ docx conversion.
24 focused tools in total, backed by 476 tests.

## Install

DocxEngine ships as an MCP server.
Point your agent at it and go: there are no document handles to track and no save step, because every tool takes a file path and saves automatically.
Pick your client below.

<details open>
<summary><b>Claude Code</b></summary>

```bash
claude mcp add docx -- uvx docxengine-mcp
```

</details>

<details>
<summary><b>Codex CLI</b></summary>

```bash
codex mcp add docx -- uvx docxengine-mcp
```

Or add it to `~/.codex/config.toml`:

```toml
[mcp_servers.docx]
command = "uvx"
args = ["docxengine-mcp"]
```

</details>

<details>
<summary><b>Cursor</b></summary>

Add to `.cursor/mcp.json` (project) or `~/.cursor/mcp.json` (global):

```json
{ "mcpServers": { "docx": { "command": "uvx", "args": ["docxengine-mcp"] } } }
```

</details>

<details>
<summary><b>Claude Desktop</b></summary>

Add to `claude_desktop_config.json` (Settings → Developer → Edit Config):

```json
{ "mcpServers": { "docx": { "command": "uvx", "args": ["docxengine-mcp"] } } }
```

</details>

<details>
<summary><b>Windsurf</b></summary>

Add to `~/.codeium/windsurf/mcp_config.json`:

```json
{ "mcpServers": { "docx": { "command": "uvx", "args": ["docxengine-mcp"] } } }
```

</details>

<details>
<summary><b>VS Code (GitHub Copilot)</b></summary>

Add to `.vscode/mcp.json`:

```json
{ "servers": { "docx": { "command": "uvx", "args": ["docxengine-mcp"] } } }
```

</details>

<details>
<summary><b>Any other MCP client</b></summary>

Run `docxengine-mcp` over stdio, or point the client at this server block:

```json
{ "mcpServers": { "docx": { "command": "uvx", "args": ["docxengine-mcp"] } } }
```

</details>

`uvx` runs the server with zero install ([install uv](https://docs.astral.sh/uv/)).
Prefer pip? Run `pip install docxengine` and use `docxengine-mcp` as the command.

## What the agent sees

The agent never touches raw OOXML.
A read returns Markdown-like text with stable anchors and only the formatting that matters:

```
[P1#a7b2  H1]            Master Services Agreement
[P2#f3c1]                This Agreement is entered into as of {{EffectiveDate}}...
[P3#b2c4  H2]            1. Definitions
[P4#d4e5]                "Confidential Information" means... [comment:C1 by J.Doe]
[T1  3×4 @after:P5]      | Term | Value | ... |
[P12#e7f8  List:ol L1]   First obligation
```

Each line has a stable ID (`P4#d4e5`) the agent uses to target edits, so it never has to re-read the whole document mid-task.
A typical edit is one call:

```json
→ docx_revision {"path":"contract.docx","op":"accept","filter":{"author":"Jane Doe"}}
← {"accepted":12,"remaining_by_author":{"Bob":3},"note":"Accepted Jane Doe's tracked changes; Bob's 3 revisions untouched."}
```

See [Concepts](docs/start/concepts.md) for how anchors, reads, and validation fit together, and the [tool reference](docs/tools/index.md) for every tool.

## How it works

DocxEngine is a deterministic core with no LLM inside: the same call on the same document always produces the same bytes.
It edits the OOXML directly (unzip, patch the XML, rezip), coalesces the split runs Word leaves behind, writes real `w:ins`/`w:del` redlines, and validates the result before saving.
The full design, including the addressing scheme and tool surface, is in [ARCHITECTURE.md](ARCHITECTURE.md).

## Documentation

| Lane                                      | What you'll find                                                            |
| ----------------------------------------- | --------------------------------------------------------------------------- |
| [Start](docs/start/quickstart.md)         | Installation, quickstart flows, core concepts                               |
| [Core](docs/core/ooxml-pitfalls.md)       | OOXML pitfalls, anchors, projection, tracked changes, validation, rendering |
| [Tools](docs/tools/index.md)              | The full tool interface, group by group, plus error design                  |
| [MCP](docs/mcp/server.md)                 | Transports, resources, session state, scaling                               |
| [Conformance](docs/conformance/corpus.md) | Round-trip fidelity corpus, agent task benchmark                            |
| [Reference](docs/reference/glossary.md)   | Glossary, tool schemas, error codes                                         |

Start at [docs/README.md](docs/README.md).

## Repository layout

```
docxengine/
├── src/            # docxengine package + MCP server
├── tests/          # Test suite
├── spec/           # Language-agnostic JSON tool contract (the source of truth)
├── conformance/    # Shared corpus + renderer fidelity harness
├── examples/       # End-to-end agent flows
├── docs/           # Design docs, tool reference, guides
└── .github/        # CI, release, security scanning, templates
```

## Status

**Stable (v1.0.0).**
All 24 tools are implemented and tested: 476 Python tests, plus a 10-task agent benchmark that runs end to end over the MCP server with zero tool errors and zero Word-repair events.
Hostile input is handled out of the box (zip-bomb caps, `<!DOCTYPE`/`<!ENTITY` rejection, XML depth caps, path-traversal clamping, all tunable via `DOCXENGINE_MAX_*`; see [SECURITY.md](SECURITY.md)), alongside a large-document perf benchmark (`make perf`) and a cross-renderer fidelity harness (`make fidelity`).
Full plan: [ROADMAP.md](ROADMAP.md).

## Contributing

Contributions are welcome, especially conformance corpus documents, OOXML edge-case reports, and benchmark tasks.
Read [CONTRIBUTING.md](CONTRIBUTING.md) for the ground rules, development setup, and commit conventions ([Conventional Commits](https://www.conventionalcommits.org) with enforced scopes).

## Community & support

- **Bugs & features** - [GitHub issues](https://github.com/ruwadgroup/docxengine/issues) (structured templates)
- **Security reports** - privately, per [SECURITY.md](SECURITY.md)
- **Governance** - [GOVERNANCE.md](GOVERNANCE.md)

## License

[Apache-2.0](LICENSE).
DocxEngine optionally shells out to external renderers/converters under their own licenses; see [LICENSING.md](LICENSING.md).
</content>
</invoke>
