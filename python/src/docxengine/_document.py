"""``Document``: a Pythonic handle over the same tool core.

A native, storage-agnostic API for programmatic manipulation — open from a path
or raw bytes, edit in memory through the exact contract tools the agent surface
uses, then persist explicitly with :meth:`Document.save` (to a path) or
:meth:`Document.to_bytes` (raw bytes, e.g. for an HTTP response). Every method
delegates to :func:`~docxengine._dispatch.call`, so anchors, tracked changes,
and the validation gate behave identically to the MCP and ``call()`` surfaces.

This is deliberately *not* file-first like the MCP server: embedding software —
including browser JS with no filesystem, and bytes-in/bytes-out servers — needs
an in-memory handle and caller-controlled persistence.
"""

from __future__ import annotations

import os
from collections.abc import Sequence

from . import _projector
from ._anchors import build_anchor_index
from ._dispatch import call
from ._errors import ToolError
from ._session import OpenDocument, Session
from ._template import fill_document
from ._tools_lifecycle import export_bytes

_Source = str | os.PathLike[str] | bytes | bytearray
_Result = dict[str, object]


class Paragraph:
    """One body paragraph: a throwaway view of (anchor, normalized text, styleId).

    Anchors are content-addressed, so a held :class:`Paragraph` goes stale after
    any edit to its paragraph — re-fetch via :meth:`Document.paragraphs` /
    :meth:`Document.find`. Operating through a stale anchor raises ``anchor_stale``,
    the spec's normal recovery signal.
    """

    __slots__ = ("anchor", "text", "style", "_doc")

    def __init__(self, doc: Document, anchor: str, text: str, style: str | None) -> None:
        self._doc = doc
        self.anchor = anchor
        self.text = text
        self.style = style  # the w:pStyle styleId (e.g. "Heading1"), or None

    def replace(
        self,
        old: str,
        new: str,
        *,
        all: bool = False,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._doc.replace(
            old, new, anchor=self.anchor, all=all, track_changes=track_changes, author=author
        )

    def edit(self, text: str, *, track_changes: bool = False, author: str | None = None) -> _Result:
        return self._doc.edit_paragraph(
            self.anchor, text, track_changes=track_changes, author=author
        )

    def insert_after(
        self,
        content: str,
        *,
        style: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._doc.insert(
            content, after=self.anchor, style=style, track_changes=track_changes, author=author
        )

    def insert_before(
        self,
        content: str,
        *,
        style: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._doc.insert(
            content, before=self.anchor, style=style, track_changes=track_changes, author=author
        )

    def delete(self, *, track_changes: bool = False, author: str | None = None) -> _Result:
        return self._doc.delete(anchor=self.anchor, track_changes=track_changes, author=author)


class Document:
    """One open .docx over a private (or caller-supplied) session.

    >>> doc = Document.open("contract.docx")
    >>> doc.replace("five (5) years", "three (3) years", all=True)
    >>> doc.save("contract-amended.docx")
    """

    def __init__(self, source: _Source, *, session: Session | None = None) -> None:
        self._session = session if session is not None else Session()
        self._doc = self._session.open_doc(source)

    # -- construction ---------------------------------------------------------

    @classmethod
    def open(cls, source: _Source, *, session: Session | None = None) -> Document:
        """Open a .docx from a filesystem path or raw zip bytes."""
        return cls(source, session=session)

    @classmethod
    def _from_handle(cls, session: Session, doc: OpenDocument) -> Document:
        self = object.__new__(cls)
        self._session = session
        self._doc = doc
        return self

    @classmethod
    def create(
        cls,
        content_md: str | None = None,
        *,
        spec: dict[str, object] | None = None,
        session: Session | None = None,
    ) -> Document:
        """Create a new document from Markdown or a structured spec (§22)."""
        session = session if session is not None else Session()
        result = call("docx_create", {"content_md": content_md, "spec": spec}, session=session)
        return cls._from_handle(session, session.get(str(result["doc_id"])))

    @classmethod
    def fill_template(
        cls,
        template: _Source,
        data: dict[str, object] | None,
        *,
        syntax: str = "mustache",
        strict: bool = False,
        session: Session | None = None,
    ) -> Document:
        """Open a template (path or bytes), fill it (§21), and return the filled Document."""
        if syntax != "mustache":
            raise ToolError(
                "template_syntax",
                f"Unsupported template syntax: {syntax}.",
                ["Only the mustache subset is supported."],
            )
        session = session if session is not None else Session()
        doc = session.open_doc(template)
        fill_document(doc, data, strict=strict)
        return cls._from_handle(session, doc)

    @classmethod
    def attach(cls, session: Session, doc_id: str) -> Document:
        """Wrap an already-open ``doc_id`` in ``session`` (shares its state)."""
        return cls._from_handle(session, session.get(doc_id))

    # -- properties -----------------------------------------------------------

    @property
    def doc_id(self) -> str:
        return self._doc.doc_id

    @property
    def session(self) -> Session:
        """The backing session — pass it to ``call()`` or hand off to an agent loop."""
        return self._session

    @property
    def dirty(self) -> bool:
        """True iff the document has unsaved edits."""
        return self._doc.dirty

    # -- read -----------------------------------------------------------------

    def outline(self) -> _Result:
        return self._call("docx_outline")

    def read(
        self,
        *,
        anchor: str | None = None,
        range: str | None = None,
        window: int = 0,
        scope: str = "body",
    ) -> _Result:
        return self._call("docx_read", anchor=anchor, range=range, window=window, scope=scope)

    def search(self, query: str, *, regex: bool = False, scope: str = "body") -> _Result:
        return self._call("docx_search", query=query, regex=regex, scope=scope)

    def paragraphs(self) -> list[Paragraph]:
        """Body paragraphs in document order, freshly anchored, with styleId."""
        package = self._doc.package
        data = package.part(package.main_document_part())
        out: list[Paragraph] = []
        for entry in build_anchor_index(package):
            if entry.kind != "paragraph":
                continue
            style_id, _, _ = _projector._paragraph_props(data, entry.span)
            out.append(Paragraph(self, entry.anchor, entry.normalized, style_id))
        return out

    def find(self, text: str) -> Paragraph | None:
        """First paragraph whose normalized text contains ``text``, or ``None``."""
        for paragraph in self.paragraphs():
            if text in paragraph.text:
                return paragraph
        return None

    # -- edit -----------------------------------------------------------------

    def replace(
        self,
        old: str,
        new: str,
        *,
        anchor: str | None = None,
        all: bool = False,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_replace",
            old=old,
            new=new,
            anchor=anchor,
            all=all,
            track_changes=track_changes,
            author=author,
        )

    def edit_paragraph(
        self, anchor: str, text: str, *, track_changes: bool = False, author: str | None = None
    ) -> _Result:
        return self._call(
            "docx_edit_paragraph",
            anchor=anchor,
            text=text,
            track_changes=track_changes,
            author=author,
        )

    def insert(
        self,
        content: str,
        *,
        after: str | None = None,
        before: str | None = None,
        style: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_insert",
            content=content,
            after=after,
            before=before,
            style=style,
            track_changes=track_changes,
            author=author,
        )

    def delete(
        self,
        *,
        anchor: str | None = None,
        range: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_delete", anchor=anchor, range=range, track_changes=track_changes, author=author
        )

    def revision(
        self, op: str, *, id: str | None = None, filter: dict[str, str] | None = None
    ) -> _Result:
        return self._call("docx_revision", op=op, id=id, filter=filter)

    def comment(
        self,
        op: str,
        *,
        anchor: str | None = None,
        comment_id: str | None = None,
        text: str | None = None,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_comment", op=op, anchor=anchor, comment_id=comment_id, text=text, author=author
        )

    def table(
        self,
        op: str,
        *,
        anchor: str | None = None,
        after: str | None = None,
        rows: int | None = None,
        cols: int | None = None,
        data: list[list[str]] | None = None,
        header: bool | None = None,
        cells: object = None,
        at: str | None = None,
        range: str | None = None,
        style: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_table",
            op=op,
            anchor=anchor,
            after=after,
            rows=rows,
            cols=cols,
            data=data,
            header=header,
            cells=cells,
            at=at,
            range=range,
            style=style,
            track_changes=track_changes,
            author=author,
        )

    def style(
        self,
        op: str,
        *,
        anchor: str | None = None,
        style: str | None = None,
        name: str | None = None,
        based_on: str | None = None,
        props: dict[str, object] | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_style",
            op=op,
            anchor=anchor,
            style=style,
            name=name,
            based_on=based_on,
            props=props,
            track_changes=track_changes,
            author=author,
        )

    def format(
        self,
        props: dict[str, object],
        *,
        anchor: str | None = None,
        range: str | None = None,
        style_selector: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_format",
            props=props,
            anchor=anchor,
            range=range,
            style_selector=style_selector,
            track_changes=track_changes,
            author=author,
        )

    def list(
        self,
        op: str,
        *,
        anchor: str | None = None,
        range: str | None = None,
        after: str | None = None,
        kind: str | None = None,
        items: list[str] | None = None,
        at: str | None = None,
        level: int | None = None,
        to: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_list",
            op=op,
            anchor=anchor,
            range=range,
            after=after,
            kind=kind,
            items=items,
            at=at,
            level=level,
            to=to,
            track_changes=track_changes,
            author=author,
        )

    def section(
        self,
        op: str,
        *,
        section: int | None = None,
        page_size: str | None = None,
        orientation: str | None = None,
        margins: dict[str, object] | None = None,
        columns: int | None = None,
        content: str | None = None,
        variant: str | None = None,
        after: str | None = None,
        break_type: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_section",
            op=op,
            section=section,
            page_size=page_size,
            orientation=orientation,
            margins=margins,
            columns=columns,
            content=content,
            variant=variant,
            after=after,
            break_type=break_type,
            track_changes=track_changes,
            author=author,
        )

    def media(
        self,
        op: str,
        *,
        after: str | None = None,
        before: str | None = None,
        image: str | None = None,
        width_cm: float | None = None,
        height_cm: float | None = None,
        media_id: str | None = None,
        path: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_media",
            op=op,
            after=after,
            before=before,
            image=image,
            width_cm=width_cm,
            height_cm=height_cm,
            media_id=media_id,
            path=path,
            track_changes=track_changes,
            author=author,
        )

    def field(
        self,
        op: str,
        *,
        after: str | None = None,
        levels: object = None,
        scope: str | None = None,
        track_changes: bool = False,
        author: str | None = None,
    ) -> _Result:
        return self._call(
            "docx_field",
            op=op,
            after=after,
            levels=levels,
            scope=scope,
            track_changes=track_changes,
            author=author,
        )

    # -- lifecycle ------------------------------------------------------------

    def validate(self) -> _Result:
        return self._call("docx_validate")

    def repair(self) -> _Result:
        return self._call("docx_repair")

    # `pages` uses Sequence, not list[int]: the `list` method above shadows the
    # builtin `list` in class-body annotations that follow it.
    def render_preview(self, *, pages: Sequence[int] | None = None) -> _Result:
        return self._call("docx_render_preview", pages=pages)

    def convert(self, to: str, *, path: str | os.PathLike[str] | None = None) -> _Result:
        """Render to md/html (in-engine) or pdf/png (render adapter)."""
        return self._call("docx_convert", to=to, path=os.fspath(path) if path is not None else None)

    def save(self, path: str | os.PathLike[str]) -> _Result:
        """Write to ``path`` through the full validation gate (``docx_save``)."""
        return call(
            "docx_save", {"doc_id": self.doc_id, "path": os.fspath(path)}, session=self._session
        )

    def to_bytes(self) -> bytes:
        """The validated .docx bytes, no filesystem (e.g. for an HTTP response)."""
        return export_bytes(self._session, doc_id=self.doc_id)

    # -- internals ------------------------------------------------------------

    def _call(self, tool: str, **kwargs: object) -> _Result:
        """Dispatch ``tool`` against this doc; omit None kwargs so handler defaults apply."""
        args: dict[str, object] = {k: v for k, v in kwargs.items() if v is not None}
        args["doc_id"] = self._doc.doc_id
        return call(tool, args, session=self._session)
