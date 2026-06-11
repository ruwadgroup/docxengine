# MCP server

The MCP face exposes the full tool surface to any MCP client (Claude Desktop, Claude Code, IDEs, gateways). It is a thin translation layer — all behavior lives in the core.

## Transports

| Transport           | When                                                 | Notes                                                                                                            |
| ------------------- | ---------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| **stdio** (default) | Local clients — Claude Desktop, Claude Code, editors | Run the binary; no auth surface, no ports                                                                        |
| **Streamable HTTP** | Hosted/multi-user deployments (Phase 2)              | Current MCP standard (2025-03-26 spec): single endpoint, `Mcp-Session-Id` header, gateway/load-balancer friendly |

Legacy HTTP+SSE is **not** supported.

```bash
docxengine-mcp            # stdio
docxengine-mcp --http --port 8080   # Streamable HTTP (Phase 2)
```

### Claude Desktop config

```json
{
  "mcpServers": {
    "docx": { "command": "docxengine-mcp" }
  }
}
```

## Tools

All tools from the [tool reference](../tools/index.md), registered from the same JSON Schemas in [`spec/`](../../spec/) that feed the SDKs — `tools/list` is generated, never hand-maintained.

## Resources

Open documents and generated artifacts are exposed as MCP resources, so clients can surface them without burning a tool call:

```
docx://d1/outline              # current outline
docx://d1/preview/page-1.png   # rendered preview pages
docx://d1/export/contract.md   # conversion outputs
```

## Response discipline

- The whole document is **never** returned; reads paginate under the ~25k-token cap.
- `convert`/`render` outputs stream as resource links, not inline blobs.
- Edits batch naturally against a `doc_id` with a single validate+save at the end.

## Concurrency

- Per-`doc_id` **write lock**; concurrent edits to one document serialize.
- Reads are lock-free against an immutable snapshot.
- Idempotent edit semantics make client retries safe (a retried `accept_all` no-ops).

State lifecycle, sessions, and horizontal scaling: [state-and-scaling.md](state-and-scaling.md).
