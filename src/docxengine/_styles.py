"""Styles (``docx_style``) and formatting (``docx_format``) — algorithms.md §16.

``docx_style`` lists/defines/applies named styles in ``word/styles.xml``;
``docx_format`` either merges the §16 closed prop set into a *style definition*
(``style_selector`` — one document-wide edit) or splices the props as *direct*
formatting onto every run/paragraph of an anchor or range. Both share the
property emission of :mod:`._props`. Edits splice raw bytes per §3.
"""

from __future__ import annotations

from collections.abc import Mapping
from xml.etree import ElementTree as ET

from . import _edits, _props, _xml
from ._errors import ToolError
from ._opc import Package
from ._session import Session

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_STYLES_PART = "word/styles.xml"


def _style_unknown(style: str) -> ToolError:
    return ToolError(
        "style_unknown",
        f"Named style {style} does not exist.",
        ['Call docx_style {op: "list"} to see available styles.'],
    )


def _inclusive(start: int, end: int) -> list[int]:
    """``[start, end]`` as a list — defined where the ``range`` builtin is not shadowed
    by :func:`docx_format`'s schema-pinned ``range`` parameter name."""
    return list(range(start, end + 1))


# ---------------------------------------------------------------------------
# Style table reads
# ---------------------------------------------------------------------------


def _styles_root(package: Package) -> ET.Element | None:
    if not package.has_part(_STYLES_PART):
        return None
    try:
        return ET.fromstring(package.part(_STYLES_PART))
    except ET.ParseError:
        return None


def list_styles(package: Package) -> list[dict[str, object]]:
    """§16 ``list``: id, name, type, based_on, in_use (effective-style count)."""
    from ._projector import _based_on_map, _paragraph_props  # late: import cycle

    root = _styles_root(package)
    if root is None:
        return []
    based_on = _based_on_map(package)
    in_use: dict[str, int] = {}
    main_data = package.part(package.main_document_part())
    for p in _xml.iter_body_children(main_data):
        if p.name != "w:p":
            continue
        style_id, _, _ = _paragraph_props(main_data, p)
        current: str | None = style_id if style_id is not None else "Normal"
        seen: set[str] = set()
        while current and current not in seen:
            in_use[current] = in_use.get(current, 0) + 1
            seen.add(current)
            current = based_on.get(current)
    out: list[dict[str, object]] = []
    for style in root.iter(f"{{{_W_NS}}}style"):
        style_id = style.get(f"{{{_W_NS}}}styleId")
        if not style_id:
            continue
        name_el = style.find(f"{{{_W_NS}}}name")
        based = style.find(f"{{{_W_NS}}}basedOn")
        entry: dict[str, object] = {
            "id": style_id,
            "name": name_el.get(f"{{{_W_NS}}}val", "") if name_el is not None else "",
            "type": style.get(f"{{{_W_NS}}}type", "paragraph"),
            "in_use": in_use.get(style_id, 0),
        }
        if based is not None:
            based_val = based.get(f"{{{_W_NS}}}val")
            if based_val:
                entry["based_on"] = based_val
        out.append(entry)
    return out


def resolve_style_id(package: Package, name_or_id: str) -> str:
    """Resolve a style reference to its styleId (§16).

    Match order, mirroring the §6a ``docx_insert`` style idiom: the styleId verbatim,
    then the styleId with whitespace removed (``"Heading 2"`` → ``Heading2``), then a
    style whose ``w:name`` equals the argument; otherwise ``style_unknown``.
    """
    styles = list_styles(package)
    by_id = {s["id"] for s in styles}
    if name_or_id in by_id:
        return name_or_id
    compact = "".join(ch for ch in name_or_id if ch not in _xml.WHITESPACE)
    if compact in by_id:
        return compact
    for style in styles:
        if style["name"] == name_or_id:
            return str(style["id"])
    raise _style_unknown(name_or_id)


# ---------------------------------------------------------------------------
# define (§16)
# ---------------------------------------------------------------------------


def _define_style(
    package: Package, name: str, based_on: str | None, props: Mapping[str, object]
) -> str:
    existing = {str(s["id"]) for s in list_styles(package)}
    base_id = "".join(ch for ch in name if ch not in _xml.WHITESPACE) or "Style"
    style_id = base_id
    suffix = 2
    while style_id in existing:
        style_id = f"{base_id}{suffix}"
        suffix += 1
    canonical = _props.canonical_props(props)
    children = [f'<w:name w:val="{_xml.escape_attr(name)}"/>']
    if based_on is not None:
        based_id = resolve_style_id(package, based_on)
        children.append(f'<w:basedOn w:val="{_xml.escape_attr(based_id)}"/>')
    ppr = _props.ppr_children_xml(canonical)
    if ppr:
        children.append(f"<w:pPr>{ppr}</w:pPr>")
    rpr = _props.rpr_xml(canonical)
    if rpr:
        children.append(rpr)
    definition = (
        f'<w:style w:type="paragraph" w:styleId="{_xml.escape_attr(style_id)}">'
        f"{''.join(children)}</w:style>"
    )
    if not package.has_part(_STYLES_PART):
        body = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f"{definition}</w:styles>"
        ).encode()
        package.set_part(_STYLES_PART, body)
    else:
        data = package.part(_STYLES_PART)
        close = data.rfind(b"</w:styles>")
        package.set_part(
            _STYLES_PART, _xml.splice(data, [(close, close, definition.encode("utf-8"))])
        )
    return style_id


# ---------------------------------------------------------------------------
# apply (§16)
# ---------------------------------------------------------------------------


def _apply_style(package: Package, anchor: str, style: str) -> str:
    """Splice ``<w:pStyle w:val="{id}"/>`` as the first ``w:pPr`` child of the paragraph."""
    style_id = resolve_style_id(package, style)
    main = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    entry = _edits.require_paragraph(entries, anchor)
    data = package.part(main)
    p = entry.span
    pstyle = f'<w:pStyle w:val="{_xml.escape_attr(style_id)}"/>'
    ppr = next(
        _xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:pPr",), max_depth=1), None
    )
    if ppr is None:
        block = f"<w:pPr>{pstyle}</w:pPr>".encode()
        new = _xml.splice(data, [(p.inner_start, p.inner_start, block)])
    else:
        existing = next(
            _xml.iter_elements(
                data, ppr.inner_start, ppr.inner_end, names=("w:pStyle",), max_depth=1
            ),
            None,
        )
        if existing is not None:
            new = _xml.splice(data, [(existing.start, existing.end, pstyle.encode("utf-8"))])
        else:
            new = _xml.splice(data, [(ppr.inner_start, ppr.inner_start, pstyle.encode("utf-8"))])
    package.set_part(main, new)
    fresh = _edits.paragraph_entries(package)
    return fresh[entry.ordinal - 1].anchor


# ---------------------------------------------------------------------------
# Property merge into an existing rPr / pPr (§16 style-selector + direct)
# ---------------------------------------------------------------------------


def _merge_into_container(
    data: bytes, container: _xml.Span, props_xml: list[tuple[str, str]]
) -> bytes:
    """Replace same-named children of ``container`` in place; append the rest at its end."""
    edits: list[tuple[int, int, bytes]] = []
    present = {
        child.name: child
        for child in _xml.iter_elements(
            data, container.inner_start, container.inner_end, max_depth=1
        )
    }
    appended: list[str] = []
    for tag, xml in props_xml:
        if tag in present:
            el = present[tag]
            edits.append((el.start, el.end, xml.encode("utf-8")))
        else:
            appended.append(xml)
    if appended:
        edits.append((container.inner_end, container.inner_end, "".join(appended).encode("utf-8")))
    return _xml.splice(data, edits)


def _first_child(data: bytes, parent: _xml.Span, name: str) -> _xml.Span | None:
    if parent.empty:
        return None
    return next(
        _xml.iter_elements(data, parent.inner_start, parent.inner_end, names=(name,), max_depth=1),
        None,
    )


def _splice_rpr(data: bytes, run: _xml.Span, rpr_xml: list[tuple[str, str]]) -> bytes:
    """Merge run props into ``w:rPr`` (created as the first ``w:r`` child) (§16)."""
    if not rpr_xml or run.empty:
        return data
    rpr = _first_child(data, run, "w:rPr")
    if rpr is None:
        block = "<w:rPr>" + "".join(xml for _, xml in rpr_xml) + "</w:rPr>"
        return _xml.splice(data, [(run.inner_start, run.inner_start, block.encode("utf-8"))])
    return _merge_into_container(data, rpr, rpr_xml)


def _splice_ppr(data: bytes, p: _xml.Span, ppr_xml: list[tuple[str, str]]) -> bytes:
    """Merge paragraph props into ``w:pPr`` (created when absent) (§16)."""
    if not ppr_xml or p.empty:
        return data
    ppr = _first_child(data, p, "w:pPr")
    if ppr is None:
        block = "<w:pPr>" + "".join(xml for _, xml in ppr_xml) + "</w:pPr>"
        return _xml.splice(data, [(p.inner_start, p.inner_start, block.encode("utf-8"))])
    return _merge_into_container(data, ppr, ppr_xml)


def _find_style_span(data: bytes, style_id: str) -> _xml.Span | None:
    for style in _xml.iter_elements(data, names=("w:style",)):
        end = style.end if style.empty else style.inner_start
        if f'w:styleId="{style_id}"'.encode() in data[style.start : end]:
            return style
    return None


def _merge_into_style(package: Package, style_id: str, props: Mapping[str, object]) -> None:
    """§16 style-selector: merge the props into the style's ``w:rPr``/``w:pPr``."""
    canonical = _props.canonical_props(props)
    ppr_xml = _props.ppr_children(canonical)
    rpr_xml = _props.rpr_children(canonical)
    # Re-locate the style span after every splice (offsets shift).
    for kind, want in (("w:pPr", ppr_xml), ("w:rPr", rpr_xml)):
        if not want:
            continue
        data = package.part(_STYLES_PART)
        style = _find_style_span(data, style_id)
        if style is None or style.empty:
            continue
        container = _first_child(data, style, kind)
        if container is None:
            block = f"<{kind}>" + "".join(xml for _, xml in want) + f"</{kind}>"
            data = _xml.splice(data, [(style.inner_end, style.inner_end, block.encode("utf-8"))])
        else:
            data = _merge_into_container(data, container, want)
        package.set_part(_STYLES_PART, data)


# ---------------------------------------------------------------------------
# docx_style
# ---------------------------------------------------------------------------


def docx_style(
    session: Session,
    *,
    doc_id: str,
    op: str,
    anchor: str | None = None,
    style: str | None = None,
    name: str | None = None,
    based_on: str | None = None,
    props: Mapping[str, object] | None = None,
    track_changes: bool = False,
    author: str | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """List, define, or apply named styles (§16)."""
    doc = session.get(doc_id)
    package = doc.package
    if op == "list":
        return {"styles": list_styles(package)}
    if op == "define":
        if name is None:
            raise _edits.anchor_invalid_error("define requires a name.")
        style_id = _define_style(package, name, based_on, props or {})
        doc.mark_dirty()
        return {"style_id": style_id, "note": f"Defined style {style_id}."}
    if op == "apply":
        if anchor is None or style is None:
            raise _edits.anchor_invalid_error("apply requires anchor and style.")
        new_anchor = _apply_style(package, anchor, style)
        doc.mark_dirty()
        return {"new_anchor": new_anchor, "note": f"Applied {style}."}
    raise _edits.anchor_invalid_error(f"Unknown style op: {op}.")


# ---------------------------------------------------------------------------
# docx_format
# ---------------------------------------------------------------------------


def docx_format(
    session: Session,
    *,
    doc_id: str,
    props: Mapping[str, object],
    anchor: str | None = None,
    range: str | None = None,  # noqa: A002 - wire name pinned by the tool schema
    style_selector: Mapping[str, object] | None = None,
    track_changes: bool = False,
    author: str | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """Edit a style definition (style_selector) or apply direct formatting (§16)."""
    doc = session.get(doc_id)
    package = doc.package
    if style_selector is not None:
        style_name = style_selector.get("style")
        if not isinstance(style_name, str):
            raise _edits.anchor_invalid_error("style_selector requires a style name.")
        style_id = resolve_style_id(package, style_name)
        _merge_into_style(package, style_id, props)
        doc.mark_dirty()
        return {"affected": 0, "anchors": [], "note": f"Edited style {style_id}."}
    main = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    if anchor is not None:
        ordinals = [_edits.require_paragraph(entries, anchor).ordinal]
    elif range is not None:
        m = _edits.RANGE_RE.match(range)
        if not m:
            raise _edits.anchor_invalid_error(f"Malformed range string: {range}.")
        start, end = int(m.group(1)), int(m.group(3))
        if start > end:
            raise _edits.anchor_invalid_error(f"Inverted range: {range}.")
        for ordinal, hash_part in ((start, m.group(2)), (end, m.group(4))):
            entry = _edits.entry_at(entries, ordinal, f"P{ordinal}")
            if hash_part is not None and entry.anchor != f"P{ordinal}#{hash_part}":
                raise _edits.anchor_stale_error(f"P{ordinal}#{hash_part}")
        ordinals = _inclusive(start, end)
    else:
        raise _edits.anchor_invalid_error("Provide anchor, range, or style_selector.")
    canonical = _props.canonical_props(props)
    rpr_xml = _props.rpr_children(canonical)
    ppr_xml = _props.ppr_children(canonical)
    for ordinal in ordinals:
        data = package.part(main)
        p = _edits.paragraph_span_at(data, ordinal)
        if rpr_xml and not p.empty:
            runs = sorted(
                _xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:r",), max_depth=1),
                key=lambda r: r.start,
                reverse=True,  # splice right-to-left so earlier offsets stay valid
            )
            for run in runs:
                data = _splice_rpr(data, run, rpr_xml)
            package.set_part(main, data)
        if ppr_xml:
            data = package.part(main)
            p = _edits.paragraph_span_at(data, ordinal)
            data = _splice_ppr(data, p, ppr_xml)
            package.set_part(main, data)
    doc.mark_dirty()
    fresh = _edits.paragraph_entries(package)
    return {
        "affected": len(ordinals),
        "anchors": [fresh[o - 1].anchor for o in sorted(ordinals)],
    }
