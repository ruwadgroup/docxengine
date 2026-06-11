"""Lifecycle tools: docx_validate, docx_repair, docx_save (algorithms.md §8/§9).

``docx_save`` is the always-on gate: it refuses to write any package carrying
an error-severity validation issue (``validation_failed``, suggesting
``docx_repair``). Warnings never block. Saving never closes the doc_id.
"""

from __future__ import annotations

import os

from ._errors import ToolError
from ._session import Session
from ._validate import is_valid, repair_package, validate_package


def docx_validate(
    session: Session, *, doc_id: str, response_format: str = "concise"
) -> dict[str, object]:
    """Run the §8 package checks; issues carry severity/part/message/fix_hint."""
    doc = session.get(doc_id)
    issues = validate_package(doc.package)
    return {"valid": is_valid(issues), "issues": [issue.to_payload() for issue in issues]}


def docx_repair(session: Session, *, doc_id: str) -> dict[str, object]:
    """Apply the §8a mechanical fixes; reports what was fixed and what remains."""
    doc = session.get(doc_id)
    fixed, remaining = repair_package(doc.package)
    if fixed:
        doc.mark_dirty()
    return {"fixed": fixed, "remaining": remaining}


def docx_save(session: Session, *, doc_id: str, path: str) -> dict[str, object]:
    """Validate (§8), then write atomically (§9). Refuses on error-severity issues."""
    doc = session.get(doc_id)
    errors = [issue for issue in validate_package(doc.package) if issue.severity == "error"]
    if errors:
        raise ToolError(
            "validation_failed",
            "Package would trigger Word repair; save refused.",
            ["Run docx_repair, then re-validate.", *(issue.message for issue in errors)],
        )
    doc.package.save(path)
    doc.mark_saved()
    return {"ok": True, "validated": True, "bytes": os.path.getsize(path)}
