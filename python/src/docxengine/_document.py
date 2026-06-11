"""``Document``: a thin Pythonic wrapper over the same tool core.

For callers who want a native API instead of the tool surface. Every method
delegates to the exact code paths the tools use (sessions, anchors, splicing,
the validation gate), so behavior is identical to the agent surface.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any

from . import _convert, _projector, _tools_edit, _tools_lifecycle
from ._edits import paragraph_entries
from ._session import Session


@dataclass(frozen=True, slots=True)
class Paragraph:
    """One body paragraph: its current anchor and normalized text."""

    anchor: str
    text: str


class Document:
    """One open .docx over a private session.

    >>> doc = Document.open("contract.docx")
    >>> doc.replace("five (5) years", "three (3) years")
    >>> doc.save("contract-amended.docx")
    """

    def __init__(self, source: str | os.PathLike[str] | bytes | bytearray) -> None:
        self._session = Session()
        self._doc = self._session.open_doc(source)

    @classmethod
    def open(cls, source: str | os.PathLike[str] | bytes | bytearray) -> Document:
        return cls(source)

    @property
    def doc_id(self) -> str:
        return self._doc.doc_id

    @property
    def dirty(self) -> bool:
        """True iff the document has unsaved edits."""
        return self._doc.dirty

    def paragraphs(self) -> list[Paragraph]:
        """Body paragraphs in document order, freshly anchored."""
        return [Paragraph(e.anchor, e.normalized) for e in paragraph_entries(self._doc.package)]

    def find(self, query: str, *, regex: bool = False) -> list[dict[str, Any]]:
        """Search matches (anchor, snippet, heading context) per algorithms.md §2a."""
        result = _projector.project_search(self._doc.package, query, regex=regex, scope="body")
        matches = result["matches"]
        assert isinstance(matches, list)
        return matches

    def replace(
        self,
        old: str,
        new: str,
        *,
        anchor: str | None = None,
        all: bool = False,  # noqa: A002 - mirrors the tool argument name
        track_changes: bool = False,
        author: str | None = None,
    ) -> dict[str, Any]:
        """Replace text (optionally tracked); returns the docx_replace result."""
        return _tools_edit.docx_replace(
            self._session,
            doc_id=self.doc_id,
            old=old,
            new=new,
            anchor=anchor,
            all=all,
            track_changes=track_changes,
            author=author,
        )

    def convert(self, to: str, path: str | os.PathLike[str] | None = None) -> dict[str, Any]:
        """Render the document to md/html (in-engine) or pdf/png (render adapter)."""
        return _convert.convert_document(
            self._doc, to, os.fspath(path) if path is not None else None
        )

    def save(self, path: str | os.PathLike[str]) -> dict[str, Any]:
        """Write to ``path`` through the full validation gate (docx_save)."""
        return _tools_lifecycle.docx_save(self._session, doc_id=self.doc_id, path=os.fspath(path))
