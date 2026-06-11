"""Fields (``docx_field``) — algorithms.md §20.

``insert_toc`` emits a paragraph holding the TOC run-triple field after an anchor
(never ``w:fldSimple``). ``insert_page_number`` ensures the body section's footer
(or header) exists and appends a ``PAGE`` field run-triple to it. ``update`` flips
``<w:updateFields w:val="true"/>`` in ``word/settings.xml`` (creating the part on
demand). All edits splice raw bytes per §3; values materialize only at render.
"""

from __future__ import annotations

import re

from . import _edits, _parts, _xml
from ._errors import ToolError
from ._opc import Package, resolve_rel_target
from ._session import Session

SETTINGS_PART = "word/settings.xml"
_SETTINGS_REL_TYPE = f"{_parts.REL_BASE}/settings"
_SETTINGS_CT = f"{_parts.CT_BASE}.settings+xml"
_HEADER_REL_TYPE = f"{_parts.REL_BASE}/header"
_FOOTER_REL_TYPE = f"{_parts.REL_BASE}/footer"
_HEADER_CT = f"{_parts.CT_BASE}.header+xml"
_FOOTER_CT = f"{_parts.CT_BASE}.footer+xml"
_W_NS_DECL = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

_WS_EDGE_RE = re.compile(f"^[{re.escape(_xml.WHITESPACE)}]|[{re.escape(_xml.WHITESPACE)}]$")


def _field_invalid(detail: str) -> ToolError:
    return ToolError("anchor_invalid", detail, ["Check the op arguments."])


# ---------------------------------------------------------------------------
# Field run-triples (§20)
# ---------------------------------------------------------------------------


def _field_run_triple(instr: str, placeholder: str) -> str:
    """A field run-triple: begin / instrText / separate / placeholder / end."""
    space = ' xml:space="preserve"' if _WS_EDGE_RE.search(instr) else ""
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f"<w:r><w:instrText{space}>{instr}</w:instrText></w:r>"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f"<w:r><w:t>{placeholder}</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


# ---------------------------------------------------------------------------
# insert_toc (§20)
# ---------------------------------------------------------------------------


def _insert_toc(package: Package, after: str, levels: int) -> tuple[str, int]:
    main = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    entry = _edits.require_paragraph(entries, after)
    safe_levels = max(1, int(levels))
    instr = f' TOC \\o "1-{safe_levels}" \\h \\z \\u '
    para = f"<w:p>{_field_run_triple(instr, 'Right-click to update field.')}</w:p>"
    data = package.part(main)
    position = entry.span.end
    package.set_part(main, _xml.splice(data, [(position, position, para.encode("utf-8"))]))
    fresh = _edits.paragraph_entries(package)
    new_anchor = fresh[entry.ordinal].anchor if entry.ordinal < len(fresh) else after
    return new_anchor, safe_levels


# ---------------------------------------------------------------------------
# insert_page_number (§20)
# ---------------------------------------------------------------------------


def _body_sectpr(data: bytes) -> _xml.Span | None:
    for child in _xml.iter_body_children(data):
        if child.name == "w:sectPr":
            return child
    return None


def _next_header_footer_index(package: Package) -> int:
    pattern = re.compile(r"^word/(?:header|footer)([0-9]+)\.xml$")
    max_n = 0
    for name in package.part_names:
        m = pattern.match(name)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def _ensure_scope_part(package: Package, scope: str) -> str:
    """The body section's default ``scope`` part, creating it (+ wiring) when absent."""
    main = package.main_document_part()
    ref_name = "w:headerReference" if scope == "header" else "w:footerReference"
    data = package.part(main)
    body_sect = _body_sectpr(data)
    if body_sect is not None and not body_sect.empty:
        for ref in _xml.iter_elements(
            data, body_sect.inner_start, body_sect.inner_end, names=(ref_name,), max_depth=1
        ):
            attrs = _edits.start_tag_attrs(data, ref)
            if attrs.get("w:type", "default") == "default":
                rel_id = attrs.get("r:id")
                if rel_id:
                    for rel in package.rels(main):
                        if rel.rel_id == rel_id and not rel.is_external:
                            target = resolve_rel_target(main, rel.target)
                            if package.has_part(target):
                                return target
    # Create a fresh empty part and wire it.
    idx = _next_header_footer_index(package)
    part_name = f"word/{scope}{idx}.xml"
    root = "w:hdr" if scope == "header" else "w:ftr"
    body = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f"<{root} {_W_NS_DECL}></{root}>"
    ).encode()
    package.set_part(part_name, body)
    _parts.ensure_content_type_override(
        package, part_name, _HEADER_CT if scope == "header" else _FOOTER_CT
    )
    rel_id = _parts.next_rel_id(package, main)
    _parts.add_relationship(
        package,
        main,
        rel_id,
        _HEADER_REL_TYPE if scope == "header" else _FOOTER_REL_TYPE,
        f"{scope}{idx}.xml",
    )
    # Splice the reference as the first child of the body sectPr.
    data = package.part(main)
    body_sect = _body_sectpr(data)
    if body_sect is None:
        raise ToolError("anchor_not_found", "Document has no body section.", [])
    ref_tag = f'<{ref_name} w:type="default" r:id="{rel_id}"/>'
    if body_sect.empty:
        new = _xml.splice(
            data, [(body_sect.start, body_sect.end, f"<w:sectPr>{ref_tag}</w:sectPr>".encode())]
        )
    else:
        new = _xml.splice(
            data, [(body_sect.inner_start, body_sect.inner_start, ref_tag.encode("utf-8"))]
        )
    package.set_part(main, new)
    return part_name


def _insert_page_number(package: Package, scope: str) -> str:
    part_name = _ensure_scope_part(package, scope)
    data = package.part(part_name)
    root = "w:hdr" if scope == "header" else "w:ftr"
    para = f"<w:p>{_field_run_triple(' PAGE ', '1')}</w:p>"
    close_tag = f"</{root}>".encode()
    close = data.rfind(close_tag)
    if close >= 0:
        new = _xml.splice(data, [(close, close, para.encode("utf-8"))])
    else:
        text = data.decode("utf-8")
        self_close = text.rfind("/>")
        new = (text[:self_close] + f">{para}</{root}>" + text[self_close + 2 :]).encode("utf-8")
    package.set_part(part_name, new)
    return part_name


# ---------------------------------------------------------------------------
# update (§20)
# ---------------------------------------------------------------------------


def _ensure_settings(package: Package) -> None:
    if package.has_part(SETTINGS_PART):
        return
    body = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f"<w:settings {_W_NS_DECL}></w:settings>"
    ).encode()
    package.set_part(SETTINGS_PART, body)
    _parts.ensure_content_type_override(package, SETTINGS_PART, _SETTINGS_CT)
    main = package.main_document_part()
    _parts.add_relationship(
        package, main, _parts.next_rel_id(package, main), _SETTINGS_REL_TYPE, "settings.xml"
    )


def _update_fields(package: Package) -> tuple[int, str]:
    _ensure_settings(package)
    data = package.part(SETTINGS_PART)
    if _xml.find_element(data, "w:updateFields") is not None:
        return 1, "Fields already flagged for update."
    settings = _xml.find_element(data, "w:settings")
    if settings is None:
        return 1, "Flagged all fields for update on next render."
    fragment = '<w:updateFields w:val="true"/>'
    if settings.empty:
        replacement = f"<w:settings {_W_NS_DECL}>{fragment}</w:settings>".encode()
        new = _xml.splice(data, [(settings.start, settings.end, replacement)])
    else:
        new = _xml.splice(
            data, [(settings.inner_start, settings.inner_start, fragment.encode("utf-8"))]
        )
    package.set_part(SETTINGS_PART, new)
    return 1, "Flagged all fields for update on next render."


# ---------------------------------------------------------------------------
# docx_field
# ---------------------------------------------------------------------------


def docx_field(
    session: Session,
    *,
    doc_id: str,
    op: str,
    after: str | None = None,
    levels: int = 3,
    scope: str = "footer",
    track_changes: bool = False,
    author: str | None = None,
) -> dict[str, object]:
    """Insert TOC/page-number field codes or flag fields for update (§20)."""
    doc = session.get(doc_id)
    package = doc.package
    if op == "insert_toc":
        if after is None:
            raise _field_invalid("op 'insert_toc' requires after.")
        new_anchor, safe_levels = _insert_toc(package, after, levels)
        doc.mark_dirty()
        return {
            "new_anchor": new_anchor,
            "note": f"Inserted a TOC field (levels 1-{safe_levels}) after {after}.",
        }
    if op == "insert_page_number":
        part_name = _insert_page_number(package, scope)
        doc.mark_dirty()
        return {"note": f"Inserted a PAGE field in the {scope} ({part_name})."}
    if op == "update":
        updated, note = _update_fields(package)
        doc.mark_dirty()
        return {"updated": updated, "note": note}
    raise _field_invalid(f"Unknown field op: {op}.")
