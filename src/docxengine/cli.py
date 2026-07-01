"""Line-oriented JSON CLI (algorithms.md §11): ``python -m docxengine.cli``.

One JSON object per stdin line — ``{"tool": "docx_replace", "args": {…}}`` —
answered by exactly one JSON object per stdout line, in request order: the
tool's result, or ``{"error": code, "message": …, "suggestions": […]}``.
doc_ids persist for the process lifetime; EOF exits 0; stderr is free-form.
"""

from __future__ import annotations

import json
import sys
from typing import IO, Any

from ._dispatch import call
from ._errors import ToolError
from ._session import Session


def _respond(line: str, session: Session) -> dict[str, Any]:
    try:
        request = json.loads(line)
    except json.JSONDecodeError as exc:
        return ToolError(
            "invalid_args",
            f"Request line is not valid JSON: {exc}.",
            ['Send one {"tool": …, "args": {…}} object per line.'],
        ).to_payload()
    if not isinstance(request, dict) or not isinstance(request.get("tool"), str):
        return ToolError(
            "invalid_args",
            "Request must be a JSON object with a string 'tool'.",
            ['Send one {"tool": …, "args": {…}} object per line.'],
        ).to_payload()
    try:
        return call(request["tool"], request.get("args"), session=session)
    except ToolError as exc:
        return exc.to_payload()


def serve(stdin: IO[str], stdout: IO[str], *, session: Session | None = None) -> int:
    """Run the request/response loop until EOF; one session per call."""
    session = session if session is not None else Session()
    for raw in stdin:
        line = raw.strip()
        if not line:
            continue
        response = _respond(line, session)
        stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
        stdout.flush()
    return 0


def main(argv: list[str] | None = None) -> int:
    return serve(sys.stdin, sys.stdout)


if __name__ == "__main__":
    raise SystemExit(main())
