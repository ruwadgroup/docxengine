"""Read-surface tools: docx_open, docx_outline, docx_read, docx_search.

Each function takes the :class:`~docxengine._session.Session` plus keyword
arguments named exactly as in ``spec/tools/<tool>.json`` and returns the result
object in exactly that schema's shape. Failures raise
:class:`~docxengine._errors.ToolError` with a ``spec/errors.json`` code.

``response_format`` is accepted everywhere per the schemas; in this phase the
``detailed`` view is identical to ``concise`` (run-level formatting lands later).
"""

from __future__ import annotations

import base64
import binascii
import re

from . import _projector
from ._errors import ToolError
from ._opc import Package
from ._session import OpenDocument, Session

_TRACKED_RE = re.compile(rb"<w:(?:ins|del)[ />]")
_COMMENT_REF_RE = re.compile(rb"<w:commentReference[ />]")
_SECT_PR_RE = re.compile(rb"<w:sectPr[ />]")

_Bytes = bytes  # the builtin, reachable where a schema-pinned parameter shadows it


def _has_tracked_changes(package: Package) -> bool:
    return _TRACKED_RE.search(package.part(package.main_document_part())) is not None


def _has_comments(package: Package) -> bool:
    return _COMMENT_REF_RE.search(package.part(package.main_document_part())) is not None


def _plural(n: int, noun: str) -> str:
    return f"{n} {noun}" if n == 1 else f"{n} {noun}s"


def _summary(package: Package, blocks: list[_projector.ProjectedBlock]) -> tuple[str, int]:
    """§2a: ``"{title} — {p} paragraphs, {s} sections, {t} tables"``."""
    paragraphs = [b for b in blocks if b.kind == "paragraph"]
    n_tables = sum(1 for b in blocks if b.kind == "table")
    n_sections = len(_SECT_PR_RE.findall(package.part(package.main_document_part())))
    title = next(
        (b.normalized for b in paragraphs if b.heading_level is not None and b.normalized),
        next((b.normalized for b in paragraphs if b.normalized), "Untitled"),
    )
    facts = [
        _plural(len(paragraphs), "paragraph"),
        _plural(n_sections, "section"),
        _plural(n_tables, "table"),
    ]
    return f"{title} — {', '.join(facts)}", len(paragraphs)


def docx_open(
    session: Session,
    *,
    path: str | None = None,
    bytes: str | None = None,  # noqa: A002 - wire name pinned by the tool schema
    response_format: str = "concise",
) -> dict[str, object]:
    """Open a .docx from a path or base64 bytes; returns the doc_id handle + summary."""
    if path is None and bytes is None:
        raise ToolError(
            "open_failed",
            "Cannot open document: provide either path or bytes.",
            ["Pass a filesystem path or base64-encoded .docx content."],
        )
    source: str | _Bytes
    if path is not None:
        source = path
    else:
        assert bytes is not None
        try:
            source = base64.b64decode(bytes, validate=True)
        except (binascii.Error, ValueError) as exc:
            raise ToolError(
                "open_failed",
                f"Cannot open <bytes>: invalid base64 content ({exc}).",
                ["Pass the .docx file content base64-encoded."],
            ) from exc
    doc = session.open_doc(source)
    summary, n_paragraphs = _summary(doc.package, _projector.project_body(doc.package))
    return {
        "doc_id": doc.doc_id,
        "summary": summary,
        "n_paragraphs": n_paragraphs,
        "has_tracked_changes": _has_tracked_changes(doc.package),
        "has_comments": _has_comments(doc.package),
    }


def docx_outline(
    session: Session, *, doc_id: str, response_format: str = "concise"
) -> dict[str, object]:
    """The heading tree and table list with anchors, style cascade resolved."""
    doc: OpenDocument = session.get(doc_id)
    return dict(_projector.project_outline(doc.package))


def docx_read(
    session: Session,
    *,
    doc_id: str,
    anchor: str | None = None,
    range: str | None = None,  # noqa: A002 - wire name pinned by the tool schema
    window: int = 0,
    scope: str = "body",
    response_format: str = "concise",
) -> dict[str, object]:
    """The Markdown projection of an anchor window, a range, or a whole story."""
    doc = session.get(doc_id)
    return dict(
        _projector.project_read(doc.package, anchor=anchor, range=range, window=window, scope=scope)
    )


def docx_search(
    session: Session,
    *,
    doc_id: str,
    query: str,
    regex: bool = False,
    scope: str = "body",
    response_format: str = "concise",
) -> dict[str, object]:
    """Coalesced-text search: matching anchors with snippets and heading context."""
    doc = session.get(doc_id)
    return dict(_projector.project_search(doc.package, query, regex=regex, scope=scope))
