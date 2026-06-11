"""In-memory document store: doc_id handles over open packages.

``docx_open`` parses a package once and holds it here keyed by ``doc_id``
(``"d1"``, ``"d2"``, … per process); every subsequent tool call reuses that
state. The session also tracks whether a document carries unsaved edits —
edit tools call :meth:`OpenDocument.mark_dirty`, ``docx_save`` calls
:meth:`OpenDocument.mark_saved`.
"""

from __future__ import annotations

import os

from ._errors import ToolError
from ._opc import Package


class OpenDocument:
    """One open document: the package plus session bookkeeping."""

    __slots__ = ("doc_id", "package", "_dirty")

    def __init__(self, doc_id: str, package: Package) -> None:
        self.doc_id = doc_id
        self.package = package
        self._dirty = False

    @property
    def dirty(self) -> bool:
        """True iff the document has been edited since open (or since last save)."""
        return self._dirty

    def mark_dirty(self) -> None:
        self._dirty = True

    def mark_saved(self) -> None:
        self._dirty = False


class Session:
    """The doc_id → document map for one process (or one MCP session)."""

    def __init__(self) -> None:
        self._docs: dict[str, OpenDocument] = {}
        self._next_id = 1

    def open_doc(self, source: str | os.PathLike[str] | bytes | bytearray) -> OpenDocument:
        """Open a package from a path or raw zip bytes and register it as ``d{n}``."""
        package = Package.open(source)
        doc_id = f"d{self._next_id}"
        self._next_id += 1
        doc = OpenDocument(doc_id, package)
        self._docs[doc_id] = doc
        return doc

    def get(self, doc_id: str) -> OpenDocument:
        """The open document for ``doc_id``; ``doc_not_found`` if unknown/expired."""
        try:
            return self._docs[doc_id]
        except KeyError:
            raise ToolError(
                "doc_not_found",
                f"Unknown or expired doc_id: {doc_id}.",
                ["Call docx_open again."],
            ) from None

    def close(self, doc_id: str) -> None:
        """Forget ``doc_id``; ``doc_not_found`` if it was never open."""
        if doc_id not in self._docs:
            raise ToolError(
                "doc_not_found",
                f"Unknown or expired doc_id: {doc_id}.",
                ["Call docx_open again."],
            )
        del self._docs[doc_id]

    def doc_ids(self) -> list[str]:
        """Every currently-open doc_id, in open order."""
        return list(self._docs)

    def __contains__(self, doc_id: str) -> bool:
        return doc_id in self._docs

    def __len__(self) -> int:
        return len(self._docs)
