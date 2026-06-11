# MCP state & scaling

## The `doc_id` lifecycle

`docx_open` parses the package once and holds it server-side — the in-memory OPC model plus the anchor index — keyed by `doc_id`. Every subsequent call reuses that state, so multi-step edit sessions never re-parse:

```
docx_open → doc_id "d1"          (parse once)
docx_search / docx_replace / …   (operate on in-memory model)
docx_save                        (validate gate → write bytes)
docx_close (optional)            (drop state; ids also expire with the session)
```

State is **per-session**:

- **stdio**: the process is the session; `doc_id`s live until the process exits.
- **Streamable HTTP**: `doc_id` lifecycle binds to the `Mcp-Session-Id` header. A new session sees no documents; an expired session's ids return `doc_not_found` with the corrective suggestion to re-open.

## Memory discipline

- Per-document and per-session memory caps; oversized documents stream parts lazily rather than holding every story inflated.
- Unsaved changes are flagged: the server warns on session end with dirty documents.

## Horizontal scaling (Streamable HTTP)

Shared in-memory state is the known scaling pain point for HTTP MCP servers. The escalation path:

1. **Single instance** (default) — in-memory store, simplest and fastest.
2. **Sticky sessions** — a gateway pins each `Mcp-Session-Id` to one instance; still in-memory.
3. **Externalized doc store** — serialized OPC models in an object store (or shared cache), instances stateless; needed only when sessions must survive instance restarts or move between instances.

Per [ROADMAP.md](../../ROADMAP.md) thresholds: don't build (3) until a hosted deployment with more than a handful of concurrent users exists.

## Crash safety

- The engine never mutates the source file in place; `docx_save` writes atomically (temp file + rename).
- A crashed session loses unsaved in-memory edits — by design; agents are encouraged to save at checkpoints.
- Idempotent semantics mean a replayed edit batch after reconnect converges instead of double-applying.
