"""Sections (``docx_section``) — algorithms.md §15.

Page size/margins/orientation/columns geometry on a ``w:sectPr`` (preset twips,
1 cm = 567, 1 in = 1440), section breaks (clone the body ``sectPr`` into a
paragraph's ``w:pPr``), and per-section headers/footers (create the part, wire the
rel + content-type Override, splice the reference). Sections are ``S{n}`` over
``w:sectPr`` in document order; the body trailing ``sectPr`` is the last ``S`` (§13).
All edits splice raw bytes per §3.
"""

from __future__ import annotations

import re
from collections.abc import Mapping

from . import _edits, _parts, _xml
from ._errors import ToolError
from ._opc import Package
from ._session import Session

#: Portrait page presets, (width, height) in twips (§15).
PAGE_SIZES: dict[str, tuple[int, int]] = {
    "A4": (11906, 16838),
    "Letter": (12240, 15840),
    "A3": (16838, 23811),
    "A5": (8391, 11906),
    "Legal": (12240, 20160),
    "Tabloid": (15840, 24480),
}
_DEFAULT_MARGINS = {
    "top": 1440,
    "right": 1440,
    "bottom": 1440,
    "left": 1440,
    "header": 708,
    "footer": 708,
    "gutter": 0,
}
_CM_TWIPS = 567

_VARIANT_TYPE = {"default": "default", "first": "first", "even": "even"}
_SECTION_ID_RE = re.compile(r"^S([1-9][0-9]*)$")
_W_NS_DECL = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

_HEADER_REL_TYPE = f"{_parts.REL_BASE}/header"
_FOOTER_REL_TYPE = f"{_parts.REL_BASE}/footer"
_HEADER_CT = f"{_parts.CT_BASE}.header+xml"
_FOOTER_CT = f"{_parts.CT_BASE}.footer+xml"


def _section_invalid(detail: str) -> ToolError:
    return ToolError(
        "anchor_invalid", detail, ["Check the section id (e.g. 'S2') and op arguments."]
    )


def _first(data: bytes, parent: _xml.Span, name: str) -> _xml.Span | None:
    if parent.empty:
        return None
    return next(
        _xml.iter_elements(data, parent.inner_start, parent.inner_end, names=(name,), max_depth=1),
        None,
    )


# ---------------------------------------------------------------------------
# Section enumeration (§13)
# ---------------------------------------------------------------------------


def _section_spans(data: bytes) -> list[_xml.Span]:
    """Every ``w:sectPr`` in document order (paragraph-embedded then the body trailer)."""
    return sorted(_xml.iter_elements(data, names=("w:sectPr",)), key=lambda s: s.start)


def _section_span_for(data: bytes, section_id: str) -> _xml.Span:
    m = _SECTION_ID_RE.match(section_id)
    if not m:
        raise _section_invalid(f"Malformed section id: {section_id}.")
    index = int(m.group(1)) - 1
    spans = _section_spans(data)
    if index < 0 or index >= len(spans):
        raise ToolError(
            "anchor_not_found",
            f"Section {section_id} not found ({len(spans)} sections).",
            ['Call docx_section {op: "list"} to map section ids.'],
        )
    return spans[index]


def _attr(data: bytes, span: _xml.Span | None, attr: str) -> str | None:
    if span is None:
        return None
    return _edits.start_tag_attrs(data, span).get(attr)


def _list_sections(package: Package) -> list[dict[str, object]]:
    data = package.part(package.main_document_part())
    out: list[dict[str, object]] = []
    spans = _section_spans(data)
    for i, sect in enumerate(spans):
        pg_sz = _first(data, sect, "w:pgSz")
        orient = _attr(data, pg_sz, "w:orient") or "portrait"
        width = _attr(data, pg_sz, "w:w") or "12240"
        height = _attr(data, pg_sz, "w:h") or "15840"
        page_size = _page_size_name(width, height)
        cols = _first(data, sect, "w:cols")
        columns = int(_attr(data, cols, "w:num") or "1") if cols is not None else 1
        type_el = _first(data, sect, "w:type")
        break_type = _attr(data, type_el, "w:val") if type_el is not None else None
        has_header = _first(data, sect, "w:headerReference") is not None
        has_footer = _first(data, sect, "w:footerReference") is not None
        out.append(
            {
                "id": f"S{i + 1}",
                "break_type": break_type or "nextPage",
                "page_size": page_size,
                "orientation": "landscape" if orient == "landscape" else "portrait",
                "columns": columns,
                "has_header": has_header,
                "has_footer": has_footer,
            }
        )
    return out


def _page_size_name(width: str, height: str) -> str:
    """Preset name matching ``width``×``height`` in either orientation, else ``custom``."""
    try:
        w, h = int(width), int(height)
    except ValueError:
        return "custom"
    for name, (pw, ph) in PAGE_SIZES.items():
        if (pw, ph) == (w, h) or (pw, ph) == (h, w):
            return name
    return "custom"


# ---------------------------------------------------------------------------
# set_geometry (§15)
# ---------------------------------------------------------------------------


#: Canonical child order inside ``w:sectPr`` for insert positioning (matches the JS engine).
_SECTPR_ORDER = ["w:headerReference", "w:footerReference", "w:pgSz", "w:pgMar", "w:cols"]


def _apply_sectpr_child(data: bytes, sect: _xml.Span, name: str, value: str) -> bytes:
    """Replace ``sect``'s ``name`` child in place, else insert in §15 canonical order."""
    existing = _first(data, sect, name)
    if existing is not None:
        return _xml.splice(data, [(existing.start, existing.end, value.encode("utf-8"))])
    my_idx = _SECTPR_ORDER.index(name)
    insert_at = sect.inner_end
    if not sect.empty:
        for kid in _xml.iter_elements(
            data, sect.inner_start, sect.inner_end, max_depth=1
        ):
            k_idx = _SECTPR_ORDER.index(kid.name) if kid.name in _SECTPR_ORDER else -1
            if k_idx > my_idx or k_idx == -1:
                insert_at = kid.start
                break
    return _xml.splice(data, [(insert_at, insert_at, value.encode("utf-8"))])


def _set_geometry(
    package: Package,
    section_id: str,
    page_size: str | None,
    orientation: str | None,
    margins: Mapping[str, object] | None,
    columns: int | None,
) -> None:
    main = package.main_document_part()
    data = package.part(main)
    sect = _section_span_for(data, section_id)

    pg_sz = _first(data, sect, "w:pgSz")
    width = int(_attr(data, pg_sz, "w:w") or "12240")
    height = int(_attr(data, pg_sz, "w:h") or "15840")
    orient = _attr(data, pg_sz, "w:orient") or "portrait"
    if page_size is not None:
        if page_size not in PAGE_SIZES:
            raise _section_invalid(f"Unknown page size: {page_size}.")
        width, height = PAGE_SIZES[page_size]
    if orientation is not None:
        orient = orientation
    portrait_w, portrait_h = min(width, height), max(width, height)
    if orient == "landscape":
        out_w, out_h = portrait_h, portrait_w
        orient_attr = ' w:orient="landscape"'
    else:
        out_w, out_h = portrait_w, portrait_h
        orient_attr = ""
    pg_sz_xml = f'<w:pgSz w:w="{out_w}" w:h="{out_h}"{orient_attr}/>'

    pg_mar_xml: str | None = None
    if margins is not None:
        existing = _first(data, sect, "w:pgMar")
        ea = _edits.start_tag_attrs(data, existing) if existing is not None else {}

        def attr_or(key: str, default: int) -> int:
            if key in ("top", "right", "bottom", "left"):
                val = margins.get(key)
                if val is not None:
                    return round(float(val) * _CM_TWIPS)  # type: ignore[arg-type]
            existing_val = ea.get(f"w:{key}")
            if existing_val and existing_val.lstrip("-").isdigit():
                return int(existing_val)
            return default

        merged = {k: attr_or(k, _DEFAULT_MARGINS[k]) for k in _DEFAULT_MARGINS}
        order = ["top", "right", "bottom", "left", "header", "footer", "gutter"]
        attrs = " ".join(f'w:{k}="{merged[k]}"' for k in order)
        pg_mar_xml = f"<w:pgMar {attrs}/>"

    cols_xml: str | None = None
    if columns is not None and columns > 1:
        cols_xml = f'<w:cols w:num="{int(columns)}" w:space="708"/>'

    # Splice each child, re-resolving the section span after every edit.
    data = _apply_sectpr_child(data, sect, "w:pgSz", pg_sz_xml)
    package.set_part(main, data)
    if pg_mar_xml is not None:
        data = package.part(main)
        sect = _section_span_for(data, section_id)
        data = _apply_sectpr_child(data, sect, "w:pgMar", pg_mar_xml)
        package.set_part(main, data)
    if cols_xml is not None:
        data = package.part(main)
        sect = _section_span_for(data, section_id)
        data = _apply_sectpr_child(data, sect, "w:cols", cols_xml)
        package.set_part(main, data)


# ---------------------------------------------------------------------------
# Headers / footers (§15)
# ---------------------------------------------------------------------------


def _markdown_paragraphs(content: str) -> str:
    """§22 plain-paragraph mapping (no lists/tables in headers MVP)."""
    pieces: list[str] = []
    for raw_line in content.split("\n"):
        line = raw_line.rstrip("\r")
        if not line.strip():
            continue
        run = f"<w:r>{_xml.emit_text_element(line)}</w:r>"
        pieces.append(f"<w:p>{run}</w:p>")
    return "".join(pieces) or "<w:p/>"


def _next_part_number(package: Package, prefix: str) -> int:
    pattern = re.compile(rf"^word/{prefix}([0-9]+)\.xml$")
    max_n = 0
    for name in package.part_names:
        m = pattern.match(name)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def _set_header_footer(
    package: Package, section_id: str, content: str, variant: str, *, is_header: bool
) -> str:
    var_type = _VARIANT_TYPE.get(variant)
    if var_type is None:
        raise _section_invalid(f"Unknown variant: {variant}.")
    main = package.main_document_part()
    prefix = "header" if is_header else "footer"
    root = "w:hdr" if is_header else "w:ftr"
    rel_type = _HEADER_REL_TYPE if is_header else _FOOTER_REL_TYPE
    content_type = _HEADER_CT if is_header else _FOOTER_CT
    ref_name = "w:headerReference" if is_header else "w:footerReference"

    number = _next_part_number(package, prefix)
    part_name = f"word/{prefix}{number}.xml"
    body = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f"<{root} {_W_NS_DECL}>{_markdown_paragraphs(content)}</{root}>"
    ).encode()
    package.set_part(part_name, body)
    _parts.ensure_content_type_override(package, part_name, content_type)
    rel_id = _parts.next_rel_id(package, main)
    _parts.add_relationship(package, main, rel_id, rel_type, f"{prefix}{number}.xml")

    # Splice the reference as the FIRST child of the target sectPr (references precede
    # everything else).
    data = package.part(main)
    sect = _section_span_for(data, section_id)
    reference = f'<{ref_name} w:type="{var_type}" r:id="{rel_id}"/>'
    new = _xml.splice(data, [(sect.inner_start, sect.inner_start, reference.encode("utf-8"))])
    package.set_part(main, new)
    return f"word/{prefix}{number}.xml"


def _find_reference(
    data: bytes, sect: _xml.Span, ref_name: str, var_type: str
) -> _xml.Span | None:
    for ref in _xml.iter_elements(
        data, sect.inner_start, sect.inner_end, names=(ref_name,), max_depth=1
    ):
        if _edits.start_tag_attrs(data, ref).get("w:type", "default") == var_type:
            return ref
    return None


# ---------------------------------------------------------------------------
# insert_break (§15)
# ---------------------------------------------------------------------------


def _body_sectpr(data: bytes) -> _xml.Span:
    """The trailing body ``w:sectPr`` (last section); raises when absent."""
    for child in _xml.iter_body_children(data):
        if child.name == "w:sectPr":
            return child
    raise _section_invalid("Document has no body-level w:sectPr to clone.")


def _insert_break(package: Package, after: str, break_type: str) -> str:
    if after is None:
        raise _section_invalid("insert_break requires an 'after' anchor.")
    main = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    entry = _edits.require_paragraph(entries, after)
    data = package.part(main)
    body_sect = _body_sectpr(data)
    # Clone the body sectPr's children; w:type (the requested break) comes first.
    inner = "" if body_sect.empty else data[body_sect.inner_start : body_sect.inner_end].decode(
        "utf-8"
    )
    inner_no_type = re.sub(r"<w:type\b[^>]*/>", "", inner)
    cloned = f'<w:sectPr><w:type w:val="{break_type}"/>{inner_no_type}</w:sectPr>'
    p = entry.span
    ppr = _first(data, p, "w:pPr")
    if ppr is not None:
        # Splice the sectPr as the first w:pPr child.
        new = _xml.splice(data, [(ppr.inner_start, ppr.inner_start, cloned.encode("utf-8"))])
    elif p.empty:
        block = f"<w:p><w:pPr>{cloned}</w:pPr></w:p>".encode()
        new = _xml.splice(data, [(p.start, p.end, block)])
    else:
        block = f"<w:pPr>{cloned}</w:pPr>".encode()
        new = _xml.splice(data, [(p.inner_start, p.inner_start, block)])
    package.set_part(main, new)
    fresh = _edits.paragraph_entries(package)
    return fresh[entry.ordinal - 1].anchor if entry.ordinal - 1 < len(fresh) else after


# ---------------------------------------------------------------------------
# docx_section
# ---------------------------------------------------------------------------


def docx_section(
    session: Session,
    *,
    doc_id: str,
    op: str,
    section: str = "S1",
    page_size: str | None = None,
    orientation: str | None = None,
    margins: Mapping[str, object] | None = None,
    columns: int | None = None,
    content: str | None = None,
    variant: str = "default",
    after: str | None = None,
    break_type: str = "nextPage",
    track_changes: bool = False,
    author: str | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """Inspect/modify section geometry, headers/footers, and section breaks (§15)."""
    doc = session.get(doc_id)
    package = doc.package
    if op == "list":
        return {"sections": _list_sections(package)}
    if op == "set_geometry":
        _set_geometry(package, section, page_size, orientation, margins, columns)
        doc.mark_dirty()
        return {"section": section, "note": f"Updated geometry of {section}."}
    if op == "set_header":
        part_name = _set_header_footer(package, section, content or "", variant, is_header=True)
        doc.mark_dirty()
        var_type = _VARIANT_TYPE.get(variant, variant)
        return {"section": section, "note": f"Set {var_type} header on {section} ({part_name})."}
    if op == "set_footer":
        part_name = _set_header_footer(package, section, content or "", variant, is_header=False)
        doc.mark_dirty()
        var_type = _VARIANT_TYPE.get(variant, variant)
        return {"section": section, "note": f"Set {var_type} footer on {section} ({part_name})."}
    if op == "insert_break":
        if after is None:
            raise _section_invalid("insert_break requires an 'after' anchor.")
        new_anchor = _insert_break(package, after, break_type)
        doc.mark_dirty()
        return {
            "new_anchor": new_anchor,
            "note": f"Inserted a {break_type} section break after {after}.",
        }
    raise _section_invalid(f"Unknown section op: {op}.")
