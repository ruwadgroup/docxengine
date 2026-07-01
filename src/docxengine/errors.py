"""Public error surface: :class:`ToolError` plus the spec error-code set.

Every failure raised by the engine is a :class:`ToolError` whose ``code`` is
one of ``ERROR_CODES`` (the packaged copy of ``spec/errors.json``). At the
CLI/MCP boundary the exception serializes to
``{"error": code, "message": str, "suggestions": [str]}``.
"""

from __future__ import annotations

from ._errors import ToolError
from ._spec import error_codes

ERROR_CODES: frozenset[str] = error_codes()

__all__ = ["ERROR_CODES", "ToolError"]
