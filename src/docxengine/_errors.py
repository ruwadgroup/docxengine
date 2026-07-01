"""Structured engine errors.

Every failure surfaced to a tool caller is a :class:`ToolError` carrying one of the
codes from ``spec/errors.json``. In-language callers catch the exception; the CLI/MCP
boundary serializes :meth:`ToolError.to_payload` as one JSON object.
"""

from __future__ import annotations


class ToolError(Exception):
    """An error with a stable code, human message, and recovery suggestions."""

    def __init__(self, code: str, message: str, suggestions: list[str] | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.suggestions: list[str] = list(suggestions or [])

    def to_payload(self) -> dict[str, object]:
        """The wire shape: ``{"error": code, "message": str, "suggestions": [str]}``."""
        return {"error": self.code, "message": self.message, "suggestions": self.suggestions}

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return f"ToolError(code={self.code!r}, message={self.message!r})"
