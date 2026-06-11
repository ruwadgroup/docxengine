"""Package-mutation helpers shared across the Phase 2 writers (algorithms.md §13-§22).

Adding a relationship, registering a content type (``Override``/``Default``),
allocating the next ``rId``, and ensuring an on-demand style exists in
``word/styles.xml`` are the same five-place wiring every structural tool repeats
(tables, lists, comments, media, sections, fields). All edits splice raw bytes
per §3 — these helpers never re-serialize an untouched region.
"""

from __future__ import annotations

import re

from . import _xml
from ._opc import Package, rels_part_for

_W_NS_DECL = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_BASE = "application/vnd.openxmlformats-officedocument.wordprocessingml"

_RELS_OPEN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
)
_RID_RE = re.compile(r'Id="rId(\d+)"')


# ---------------------------------------------------------------------------
# Relationships
# ---------------------------------------------------------------------------


def next_rel_id(package: Package, source_part: str | None) -> str:
    """``rId{max+1}`` over the relationship ids already in ``source_part``'s rels."""
    rels_name = rels_part_for(source_part)
    max_id = 0
    if package.has_part(rels_name):
        for m in _RID_RE.finditer(package.part(rels_name).decode("utf-8")):
            max_id = max(max_id, int(m.group(1)))
    return f"rId{max_id + 1}"


def add_relationship(
    package: Package,
    source_part: str | None,
    rel_id: str,
    rel_type: str,
    target: str,
    *,
    mode: str | None = None,
) -> None:
    """Append a ``<Relationship>`` to ``source_part``'s rels, creating the rels part."""
    rels_name = rels_part_for(source_part)
    mode_attr = f' TargetMode="{mode}"' if mode else ""
    rel = (
        f'<Relationship Id="{rel_id}" Type="{_xml.escape_attr(rel_type)}"'
        f' Target="{_xml.escape_attr(target)}"{mode_attr}/>'
    )
    if not package.has_part(rels_name):
        package.set_part(rels_name, (_RELS_OPEN + rel + "</Relationships>").encode("utf-8"))
        return
    data = package.part(rels_name)
    close = data.rfind(b"</Relationships>")
    package.set_part(rels_name, _xml.splice(data, [(close, close, rel.encode("utf-8"))]))


# ---------------------------------------------------------------------------
# Content types
# ---------------------------------------------------------------------------


def ensure_content_type_override(package: Package, part_name: str, content_type: str) -> None:
    """Add an ``Override`` for ``part_name`` if ``[Content_Types].xml`` lacks one."""
    name = part_name.lstrip("/")
    if name in package.content_types().overrides:
        return
    data = package.part("[Content_Types].xml")
    close = data.rfind(b"</Types>")
    override = f'<Override PartName="/{name}" ContentType="{_xml.escape_attr(content_type)}"/>'
    package.set_part("[Content_Types].xml", _xml.splice(data, [(close, close, override.encode())]))


def ensure_content_type_default(package: Package, ext: str, content_type: str) -> None:
    """Add a ``Default`` for ``ext`` if ``[Content_Types].xml`` lacks one."""
    ext = ext.lower()
    if ext in package.content_types().defaults:
        return
    data = package.part("[Content_Types].xml")
    close = data.rfind(b"</Types>")
    default = f'<Default Extension="{ext}" ContentType="{_xml.escape_attr(content_type)}"/>'
    package.set_part("[Content_Types].xml", _xml.splice(data, [(close, close, default.encode())]))


# ---------------------------------------------------------------------------
# On-demand parts (numbering, comments, settings, …)
# ---------------------------------------------------------------------------


def ensure_part(
    package: Package,
    part_name: str,
    *,
    root: str,
    content_type: str,
    rel_type: str,
) -> bytes:
    """Return ``part_name``'s bytes, creating an empty ``<root/>`` part + wiring first.

    Creating wires the document relationship (``rel_type``) and the content-type
    ``Override`` (``content_type``); the part body is ``<root …w-ns…></root>``.
    """
    if package.has_part(part_name):
        return package.part(part_name)
    main = package.main_document_part()
    body = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f"<{root} {_W_NS_DECL}></{root}>"
    ).encode()
    package.set_part(part_name, body)
    ensure_content_type_override(package, part_name, content_type)
    target = part_name[len("word/") :] if part_name.startswith("word/") else f"/{part_name}"
    add_relationship(package, main, next_rel_id(package, main), rel_type, target)
    return body


def append_before_close(data: bytes, close_tag: bytes, fragment: str) -> bytes:
    """Splice ``fragment`` immediately before the last ``close_tag`` in ``data``."""
    close = data.rfind(close_tag)
    return _xml.splice(data, [(close, close, fragment.encode("utf-8"))])


# ---------------------------------------------------------------------------
# Ensure a style exists (§16, used by tables/lists/comments)
# ---------------------------------------------------------------------------

_STYLES_PART = "word/styles.xml"

#: Canonical definitions for styles the structural tools ensure on demand
#: (§14/§17/§18). These are byte-parity with the TypeScript engine's
#: ``phase2common`` definitions; the conformance harness compares produced
#: ``word/styles.xml`` across implementations.
_ENSURED_STYLES: dict[str, str] = {
    "TableGrid": (
        '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
        "<w:tblPr><w:tblBorders>"
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        "</w:tblBorders></w:tblPr></w:style>"
    ),
    "ListParagraph": (
        '<w:style w:type="paragraph" w:styleId="ListParagraph">'
        '<w:name w:val="List Paragraph"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
    ),
    "CommentReference": (
        '<w:style w:type="character" w:styleId="CommentReference">'
        '<w:name w:val="annotation reference"/>'
        '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>'
    ),
}


def style_ids(package: Package) -> set[str]:
    """Every ``w:styleId`` declared in ``word/styles.xml`` (empty when absent)."""
    if not package.has_part(_STYLES_PART):
        return set()
    data = package.part(_STYLES_PART)
    out: set[str] = set()
    for el in _xml.iter_elements(data, names=("w:style",)):
        end = el.end if el.empty else el.inner_start
        m = re.search(rb'w:styleId="([^"]*)"', data[el.start : end])
        if m:
            out.add(m.group(1).decode("utf-8"))
    return out


def ensure_style(package: Package, style_id: str) -> None:
    """Ensure ``style_id`` is defined in ``word/styles.xml`` (§16 ensure-style)."""
    if style_id in style_ids(package):
        return
    definition = _ENSURED_STYLES.get(style_id)
    if definition is None:
        return
    if not package.has_part(_STYLES_PART):
        body = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f"<w:styles {_W_NS_DECL}>{definition}</w:styles>"
        ).encode()
        package.set_part(_STYLES_PART, body)
        ensure_content_type_override(
            package, _STYLES_PART, f"{CT_BASE}.styles+xml"
        )
        add_relationship(
            package,
            package.main_document_part(),
            next_rel_id(package, package.main_document_part()),
            f"{REL_BASE}/styles",
            "styles.xml",
        )
        return
    data = package.part(_STYLES_PART)
    package.set_part(_STYLES_PART, append_before_close(data, b"</w:styles>", definition))
