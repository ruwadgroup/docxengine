"""Markdown projection of text-first tool results for the MCP ``tools/call`` surface.

The core tools return structured dicts — the CLI / native-``Document`` API contract,
and what the conformance suite checks. But for tools whose payload is *fundamentally
a text document*, JSON-encoding it is wasteful double-encoding: the agent receives a
JSON string wrapping an escaped markdown string (every newline ``\\n`` → ``\\\\n``) and
must parse JSON to recover text that was markdown all along.

So the MCP server (algorithms.md §26) emits markdown directly for these tools instead
of a JSON envelope — the same principle the resource endpoints already use
(``docx://{path}/outline|projection`` serve ``text/markdown``). This is a presentation
rule of the MCP transport only; the structured result contract is unchanged.

Metadata that the markdown body cannot carry (the read pagination cursor, the convert
note, the structural page estimate, the echoed path) is preserved in a single trailing
HTML comment so nothing is lost and the trailer stays greppable:

    <!-- docxengine: ~12 pages | structural projection, no render adapter | path=… -->

:func:`project_markdown` returns the markdown string for a text-first tool, or ``None``
to tell the caller to fall back to JSON (renderer image links, pdf/png file results,
and every structured/op tool).
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

_TRAILER_OPEN = "<!-- docxengine:"


def _trailer(parts: list[str], payload: dict[str, Any]) -> str:
    """One trailing HTML comment carrying the dropped metadata, or ``""`` if none."""
    path = payload.get("path")
    if path is not None:
        parts = [*parts, f"path={path}"]
    if not parts:
        return ""
    return f"\n\n{_TRAILER_OPEN} {' | '.join(parts)} -->"


def _render_preview(payload: dict[str, Any]) -> str | None:
    # Only the structural fallback is text; the renderer path returns image links (JSON).
    if payload.get("renderer") != "structural":
        return None
    parts: list[str] = []
    count = payload.get("page_count")
    if count is not None:
        parts.append(f"~{count} page{'' if count == 1 else 's'}")
    parts.append("structural projection, no render adapter")
    return str(payload.get("structural", "")) + _trailer(parts, payload)


def _convert(payload: dict[str, Any]) -> str | None:
    # md/html targets carry inline `content`; pdf/png return a file path (JSON).
    if "content" not in payload:
        return None
    parts: list[str] = []
    note = payload.get("note")
    if note:
        parts.append(str(note))
    return str(payload["content"]) + _trailer(parts, payload)


def _read(payload: dict[str, Any]) -> str | None:
    if "content" not in payload:
        return None
    parts: list[str] = []
    cursor = payload.get("continuation")
    if cursor is not None:
        parts.append(f"continuation={cursor}")
    return str(payload["content"]) + _trailer(parts, payload)


def _outline(payload: dict[str, Any]) -> str | None:
    lines: list[str] = []
    for entry in payload.get("outline", []):
        level = int(entry["level"])
        lines.append(f"{'#' * level} {entry['text']} [{entry['anchor']}]")
    for table in payload.get("tables", []):
        after = table.get("after")
        where = f" @after:{after}" if after is not None else ""
        lines.append(f"- table {table['anchor']} {table['dims']}{where}")
    body = "\n".join(lines) if lines else "_(no headings or tables)_"
    return body + _trailer([], payload)


def _search(payload: dict[str, Any]) -> str | None:
    matches = payload.get("matches", [])
    lines: list[str] = []
    for match in matches:
        context = match.get("context")
        suffix = f"  (under: {context})" if context else ""
        lines.append(f"- [{match['anchor']}] \"{match['snippet']}\"{suffix}")
    body = "\n".join(lines) if lines else "_(no matches)_"
    n = payload.get("n_matches", len(matches))
    return body + _trailer([f"{n} matches"], payload)


_PROJECTORS: dict[str, Callable[[dict[str, Any]], str | None]] = {
    "docx_render_preview": _render_preview,
    "docx_convert": _convert,
    "docx_read": _read,
    "docx_outline": _outline,
    "docx_search": _search,
}


def project_markdown(tool: str, payload: dict[str, Any]) -> str | None:
    """The markdown rendering of a text-first tool result, or ``None`` for JSON.

    ``payload`` is the successful tool result (never an error payload — errors are
    always JSON). Returns ``None`` for any tool not in the text-first set, and for
    text-first tools whose particular result is structured rather than text (the
    render_preview renderer path, convert to pdf/png).
    """
    projector = _PROJECTORS.get(tool)
    return projector(payload) if projector is not None else None
