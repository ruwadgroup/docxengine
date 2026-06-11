"""Lists & numbering (``docx_list``) — algorithms.md §17.

Creates ``word/numbering.xml`` on demand (content-type Override + document rel),
allocates abstractNum/num ids (max + 1), wires each item's ``w:numPr`` as the first
``w:pPr`` child, and ensures the ``ListParagraph`` style. Ops: create, restart,
set_level, convert (ol/ul/paragraphs). Edits splice raw bytes per §3.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence

from . import _edits, _parts, _xml
from ._errors import ToolError
from ._opc import Package
from ._session import Session

_NUMBERING_PART = "word/numbering.xml"
_W_NS_DECL = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

#: §17 ol cascade of numFmt per level (0-8); level 9 reuses the level-8 pattern start.
_OL_FORMATS = ["decimal", "lowerLetter", "lowerRoman"]
#: §17 ul bullet glyphs cascading per level.
_UL_GLYPHS = ["•", "◦", "▪"]


def _list_invalid(detail: str) -> ToolError:
    return ToolError("anchor_invalid", detail, ["Check the anchor/range and op arguments."])


def _first(data: bytes, parent: _xml.Span, name: str) -> _xml.Span | None:
    """The first direct child named ``name`` of ``parent`` (``None`` when absent)."""
    if parent.empty:
        return None
    return next(
        _xml.iter_elements(data, parent.inner_start, parent.inner_end, names=(name,), max_depth=1),
        None,
    )


# ---------------------------------------------------------------------------
# numbering.xml creation + id allocation (§17)
# ---------------------------------------------------------------------------


def _ensure_numbering(package: Package) -> bytes:
    return _parts.ensure_part(
        package,
        _NUMBERING_PART,
        root="w:numbering",
        content_type=f"{_parts.CT_BASE}.numbering+xml",
        rel_type=f"{_parts.REL_BASE}/numbering",
    )


def _max_id(data: bytes, name: str, attr: str) -> int:
    max_id = 0
    pattern = re.compile(rf'<{re.escape(name)}[^>]*{re.escape(attr)}="(\d+)"'.encode())
    for m in pattern.finditer(data):
        max_id = max(max_id, int(m.group(1)))
    return max_id


def _next_abstract_id(data: bytes) -> int:
    return _max_id(data, "w:abstractNum", "w:abstractNumId") + 1


def _next_num_id(data: bytes) -> int:
    return _max_id(data, "w:num", "w:numId") + 1


def _ol_level(ilvl: int) -> str:
    fmt = _OL_FORMATS[ilvl % len(_OL_FORMATS)]
    left = 720 * (ilvl + 1)
    return (
        f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/>'
        f'<w:numFmt w:val="{fmt}"/><w:lvlText w:val="%{ilvl + 1}."/>'
        f'<w:pPr><w:ind w:left="{left}" w:hanging="360"/></w:pPr></w:lvl>'
    )


def _ul_level(ilvl: int) -> str:
    glyph = _UL_GLYPHS[ilvl % len(_UL_GLYPHS)]
    left = 720 * (ilvl + 1)
    return (
        f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/>'
        f'<w:numFmt w:val="bullet"/><w:lvlText w:val="{_xml.escape_attr(glyph)}"/>'
        f'<w:pPr><w:ind w:left="{left}" w:hanging="360"/></w:pPr></w:lvl>'
    )


def _abstract_num_xml(abstract_id: int, kind: str) -> str:
    level = _ol_level if kind == "ol" else _ul_level
    levels = "".join(level(i) for i in range(9))
    return f'<w:abstractNum w:abstractNumId="{abstract_id}">{levels}</w:abstractNum>'


def _num_xml(num_id: int, abstract_id: int, *, start_override: int | None = None) -> str:
    override = ""
    if start_override is not None:
        override = (
            f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="{start_override}"/></w:lvlOverride>'
        )
    return f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/>{override}</w:num>'


def _append_numbering(package: Package, fragment: str) -> None:
    data = package.part(_NUMBERING_PART)
    package.set_part(
        _NUMBERING_PART, _parts.append_before_close(data, b"</w:numbering>", fragment)
    )


def _create_run(package: Package, kind: str) -> tuple[int, int]:
    """Allocate a fresh abstractNum (``kind``) + num pointing at it; returns ``(num_id, _)``."""
    _ensure_numbering(package)
    data = package.part(_NUMBERING_PART)
    abstract_id = _next_abstract_id(data)
    _append_numbering(package, _abstract_num_xml(abstract_id, kind))
    data = package.part(_NUMBERING_PART)
    num_id = _next_num_id(data)
    _append_numbering(package, _num_xml(num_id, abstract_id))
    return num_id, abstract_id


# ---------------------------------------------------------------------------
# numPr splicing on a paragraph (§17)
# ---------------------------------------------------------------------------


def _num_pr(num_id: int, level: int) -> str:
    return f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>'


def _set_paragraph_num(data: bytes, p: _xml.Span, num_id: int, level: int) -> bytes:
    """Set the paragraph's ``w:numPr`` (first ``w:pPr`` child) + ensure pStyle ListParagraph."""
    pstyle = '<w:pStyle w:val="ListParagraph"/>'
    numpr = _num_pr(num_id, level)
    ppr = _first(data, p, "w:pPr")
    if ppr is None:
        block = f"<w:pPr>{pstyle}{numpr}</w:pPr>".encode()
        return _xml.splice(data, [(p.inner_start, p.inner_start, block)])
    edits: list[tuple[int, int, bytes]] = []
    existing_style = _first(data, ppr, "w:pStyle")
    existing_num = _first(data, ppr, "w:numPr")
    head = b""
    if existing_style is None:
        head += pstyle.encode("utf-8")
    if existing_num is not None:
        edits.append((existing_num.start, existing_num.end, numpr.encode("utf-8")))
    else:
        head += numpr.encode("utf-8")
    if head:
        edits.append((ppr.inner_start, ppr.inner_start, head))
    return _xml.splice(data, edits)


def _remove_num_pr(data: bytes, p: _xml.Span) -> bytes:
    ppr = _first(data, p, "w:pPr")
    if ppr is None:
        return data
    numpr = _first(data, ppr, "w:numPr")
    if numpr is None:
        return data
    return _xml.splice(data, [(numpr.start, numpr.end, b"")])


def _set_level(data: bytes, p: _xml.Span, level: int) -> bytes:
    ppr = _first(data, p, "w:pPr")
    if ppr is None:
        raise _list_invalid("Paragraph is not a list item (no w:pPr).")
    numpr = _first(data, ppr, "w:numPr")
    if numpr is None or numpr.empty:
        raise _list_invalid("Paragraph is not a list item (no w:numPr).")
    ilvl = _first(data, numpr, "w:ilvl")
    new_ilvl = f'<w:ilvl w:val="{level}"/>'.encode()
    if ilvl is None:
        return _xml.splice(data, [(numpr.inner_start, numpr.inner_start, new_ilvl)])
    return _xml.splice(data, [(ilvl.start, ilvl.end, new_ilvl)])


def _paragraph_num_id(data: bytes, p: _xml.Span) -> int | None:
    ppr = _first(data, p, "w:pPr")
    if ppr is None:
        return None
    numpr = _first(data, ppr, "w:numPr")
    if numpr is None or numpr.empty:
        return None
    numid = _first(data, numpr, "w:numId")
    if numid is None:
        return None
    end = numid.end if numid.empty else numid.inner_start
    m = re.search(rb'w:val="(\d+)"', data[numid.start : end])
    return int(m.group(1)) if m else None


def _num_to_abstract(package: Package, num_id: int) -> int | None:
    if not package.has_part(_NUMBERING_PART):
        return None
    data = package.part(_NUMBERING_PART)
    for num in _xml.iter_elements(data, names=("w:num",)):
        end = num.end if num.empty else num.inner_start
        if f'w:numId="{num_id}"'.encode() not in data[num.start : end]:
            continue
        abstract = _first(data, num, "w:abstractNumId")
        if abstract is None:
            return None
        aend = abstract.end if abstract.empty else abstract.inner_start
        m = re.search(rb'w:val="(\d+)"', data[abstract.start : aend])
        return int(m.group(1)) if m else None
    return None


# ---------------------------------------------------------------------------
# Target resolution (anchor or range)
# ---------------------------------------------------------------------------


def _target_ordinals(package: Package, anchor: str | None, range_ref: str | None) -> list[int]:
    entries = _edits.paragraph_entries(package)
    if anchor is not None:
        return [_edits.require_paragraph(entries, anchor).ordinal]
    if range_ref is not None:
        m = _edits.RANGE_RE.match(range_ref)
        if not m:
            raise _list_invalid(f"Malformed range string: {range_ref}.")
        start, end = int(m.group(1)), int(m.group(3))
        if start > end:
            raise _list_invalid(f"Inverted range: {range_ref}.")
        for ordinal, hash_part in ((start, m.group(2)), (end, m.group(4))):
            entry = _edits.entry_at(entries, ordinal, f"P{ordinal}")
            if hash_part is not None and entry.anchor != f"P{ordinal}#{hash_part}":
                raise _edits.anchor_stale_error(f"P{ordinal}#{hash_part}")
        return list(range(start, end + 1))
    raise _list_invalid("Provide an anchor or a range.")


# ---------------------------------------------------------------------------
# Ops (§17)
# ---------------------------------------------------------------------------


def _create_list(
    package: Package, after: str | None, kind: str, items: Sequence[Mapping[str, object]]
) -> list[str]:
    if after is None:
        raise _list_invalid("create requires an 'after' paragraph anchor.")
    if not items:
        return []
    _parts.ensure_style(package, "ListParagraph")
    num_id, _ = _create_run(package, kind)
    main = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    entry = _edits.require_paragraph(entries, after)
    pieces: list[str] = []
    for item in items:
        text = str(item.get("text", ""))
        level = int(item.get("level", 0))  # type: ignore[call-overload]
        run = f"<w:r>{_xml.emit_text_element(text)}</w:r>" if text else ""
        ppr = f'<w:pPr><w:pStyle w:val="ListParagraph"/>{_num_pr(num_id, level)}</w:pPr>'
        pieces.append(f"<w:p>{ppr}{run}</w:p>")
    data = package.part(main)
    pos = entry.span.end
    package.set_part(main, _xml.splice(data, [(pos, pos, "".join(pieces).encode("utf-8"))]))
    fresh = _edits.paragraph_entries(package)
    base = entry.ordinal + 1
    return [fresh[base - 1 + i].anchor for i in range(len(items))]


def _restart_list(package: Package, anchor: str | None, at: int) -> list[str]:
    ordinals = _target_ordinals(package, anchor, None)
    main = package.main_document_part()
    data = package.part(main)
    p = _edits.paragraph_span_at(data, ordinals[0])
    num_id = _paragraph_num_id(data, p)
    if num_id is None:
        raise _list_invalid("Target paragraph is not a list item.")
    abstract_id = _num_to_abstract(package, num_id)
    if abstract_id is None:
        raise _list_invalid("List numbering is unresolvable.")
    _ensure_numbering(package)
    new_num_id = _next_num_id(package.part(_NUMBERING_PART))
    _append_numbering(package, _num_xml(new_num_id, abstract_id, start_override=at))
    # repoint the paragraph's numId.
    data = package.part(main)
    p = _edits.paragraph_span_at(data, ordinals[0])
    level = _current_level(data, p)
    package.set_part(main, _set_paragraph_num(data, p, new_num_id, level))
    fresh = _edits.paragraph_entries(package)
    return [fresh[ordinals[0] - 1].anchor]


def _current_level(data: bytes, p: _xml.Span) -> int:
    ppr = _first(data, p, "w:pPr")
    numpr = _first(data, ppr, "w:numPr") if ppr is not None else None
    ilvl = _first(data, numpr, "w:ilvl") if numpr is not None else None
    if ilvl is None:
        return 0
    end = ilvl.end if ilvl.empty else ilvl.inner_start
    m = re.search(rb'w:val="(\d+)"', data[ilvl.start : end])
    return int(m.group(1)) if m else 0


def _set_level_op(
    package: Package, anchor: str | None, range_ref: str | None, level: int
) -> list[str]:
    ordinals = _target_ordinals(package, anchor, range_ref)
    main = package.main_document_part()
    for ordinal in ordinals:
        data = package.part(main)
        p = _edits.paragraph_span_at(data, ordinal)
        package.set_part(main, _set_level(data, p, level))
    fresh = _edits.paragraph_entries(package)
    return [fresh[o - 1].anchor for o in ordinals]


def _convert(
    package: Package, anchor: str | None, range_ref: str | None, to: str
) -> list[str]:
    ordinals = _target_ordinals(package, anchor, range_ref)
    main = package.main_document_part()
    if to == "paragraphs":
        for ordinal in ordinals:
            data = package.part(main)
            p = _edits.paragraph_span_at(data, ordinal)
            package.set_part(main, _remove_num_pr(data, p))
        fresh = _edits.paragraph_entries(package)
        return [fresh[o - 1].anchor for o in ordinals]
    if to not in ("ol", "ul"):
        raise _list_invalid(f"Unknown convert target: {to}.")
    _parts.ensure_style(package, "ListParagraph")
    num_id, _ = _create_run(package, to)  # one abstractNum reused for the run
    for ordinal in ordinals:
        data = package.part(main)
        p = _edits.paragraph_span_at(data, ordinal)
        level = _current_level(data, p)
        package.set_part(main, _set_paragraph_num(data, p, num_id, level))
    fresh = _edits.paragraph_entries(package)
    return [fresh[o - 1].anchor for o in ordinals]


# ---------------------------------------------------------------------------
# docx_list
# ---------------------------------------------------------------------------


def docx_list(
    session: Session,
    *,
    doc_id: str,
    op: str,
    anchor: str | None = None,
    range: str | None = None,  # noqa: A002 - wire name pinned by the tool schema
    after: str | None = None,
    kind: str = "ul",
    items: Sequence[Mapping[str, object]] | None = None,
    at: int = 1,
    level: int | None = None,
    to: str | None = None,
    track_changes: bool = False,
    author: str | None = None,
) -> dict[str, object]:
    """Create lists, restart/re-level numbering, convert paragraphs ↔ list items (§17)."""
    doc = session.get(doc_id)
    package = doc.package
    if op == "create":
        new_anchors = _create_list(package, after, kind, items or [])
        if new_anchors:
            doc.mark_dirty()
        return {"new_anchors": new_anchors, "n_affected": len(new_anchors)}
    if op == "restart":
        anchors = _restart_list(package, anchor, at)
        doc.mark_dirty()
        return {"new_anchors": anchors, "n_affected": len(anchors)}
    if op == "set_level":
        if level is None:
            raise _list_invalid("set_level requires a level.")
        anchors = _set_level_op(package, anchor, range, level)
        doc.mark_dirty()
        return {"new_anchors": anchors, "n_affected": len(anchors)}
    if op == "convert":
        if to is None:
            raise _list_invalid("convert requires a 'to' target.")
        anchors = _convert(package, anchor, range, to)
        doc.mark_dirty()
        return {"new_anchors": anchors, "n_affected": len(anchors)}
    raise _list_invalid(f"Unknown list op: {op}.")
