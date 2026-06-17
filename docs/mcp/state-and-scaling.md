# MCP state & scaling

## The per-call lifecycle

The MCP server is stateless across calls — the file on disk is the only state. Each `tools/call` opens the file, runs the tool, and saves it back when the edit changed it (algorithms.md §26):

```
docx_create report.docx          (build → validate → write the file)
docx_search / docx_replace / …   (each: open report.docx → edit → validate → save back)
docx_outline / docx_read         (each: open report.docx → render, no write)
```

There is no `doc_id` to track and no explicit save: an edit in one call is on disk for the next. (Internally the engine and the §11 CLI keep an in-memory `doc_id` handle, but the server does not expose it.)

## Sessions

- **stdio**: the process serves one client; document state is purely the filesystem.
- **Streamable HTTP**: the `Mcp-Session-Id` is protocol bookkeeping only (initialize mints it; an expired id returns HTTP 410). Sessions carry **no** document state, so a request is served identically regardless of which session — or instance — handles it.

## Horizontal scaling (Streamable HTTP)

Because the filesystem is the state, the in-memory scaling problem that dogs handle-based MCP servers disappears:

1. **Single instance** (default) — the per-path write lock prevents concurrent self-conflicts.
2. **Multiple instances over shared storage** — instances are stateless, so any instance can serve any request; sticky sessions are not needed for correctness. The only cross-instance concern is concurrent writes to the _same_ file from different instances, which needs shared-filesystem locking (out of scope here; see [ROADMAP.md](../../ROADMAP.md)).

Set `DOCXENGINE_ROOT` to confine all paths to one directory (sandbox; escapes → `path_denied`).

## Crash safety

- The engine never mutates a source file in place; every write is atomic (temp file + rename), so a crash mid-write leaves the original intact.
- There is no unsaved in-memory state to lose: each call persists before it returns.
- Idempotent semantics mean a replayed edit after reconnect converges instead of double-applying.
