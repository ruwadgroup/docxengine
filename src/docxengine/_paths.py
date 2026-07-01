"""Filesystem path resolution for the MCP file facade (``mcp.py`` only).

The facade is the one surface that addresses documents by filesystem path; the
wire contract, the CLI, and the SDKs stay ``doc_id``/bytes based. Relative paths
resolve against the server's working directory — or against ``DOCXENGINE_ROOT``
when it is set, in which case any path escaping that root is refused with
``path_denied``.
"""

from __future__ import annotations

import os
from pathlib import Path

from ._errors import ToolError

ROOT_ENV = "DOCXENGINE_ROOT"


def server_root() -> Path | None:
    """The configured sandbox root (``DOCXENGINE_ROOT``), resolved, or ``None``."""
    raw = os.environ.get(ROOT_ENV)
    if not raw:
        return None
    return Path(raw).expanduser().resolve()


def resolve_path(raw: object, root: Path | None) -> Path:
    """Resolve a user-supplied path; enforce the sandbox when a root is set.

    ``~`` expands; relative paths join ``root`` (when set) else the current
    working directory; symlinks collapse via :meth:`Path.resolve`. With a root
    configured, the resolved target must stay inside it, else ``path_denied``.
    """
    if not isinstance(raw, str) or not raw:
        raise ToolError(
            "path_denied",
            "A non-empty file path is required.",
            ["Pass a filesystem path to the .docx file."],
        )
    candidate = Path(raw).expanduser()
    if not candidate.is_absolute():
        candidate = (root if root is not None else Path.cwd()) / candidate
    resolved = candidate.resolve()
    if root is not None and not resolved.is_relative_to(root):
        raise ToolError(
            "path_denied",
            f"Path escapes DOCXENGINE_ROOT: {raw}.",
            [f"Use a path inside {root}."],
        )
    return resolved
