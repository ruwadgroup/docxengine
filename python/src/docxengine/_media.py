"""Media (``docx_media``) — algorithms.md §19.

``insert`` writes ``word/media/image{k}.{ext}`` (k = max + 1), a document rel
(type ``…/image``), a content-type ``Default`` for the extension, and splices an
inline drawing run after/before an anchor; aspect-preserving sizing parses PNG/JPEG
pixel dimensions. ``extract`` copies an ``M{id}`` part's bytes to a path; ``replace``
overwrites the part keeping its rel/rId. ``M{ordinal}`` = document order of drawing
references (``a:blip``). All edits splice raw bytes per §3.
"""

from __future__ import annotations

import posixpath
import re

from . import _edits, _parts, _xml
from ._errors import ToolError
from ._opc import Package
from ._session import Session

_MEDIA_REL_TYPE = f"{_parts.REL_BASE}/image"
_MEDIA_ID_RE = re.compile(r"^M([1-9][0-9]*)$")

#: Extension → content-type for the media Default (§8 repair map; else octet-stream).
_IMAGE_CONTENT_TYPES = {
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
}

_EMU_PER_CM = 360000


def _media_invalid(detail: str) -> ToolError:
    return ToolError("anchor_invalid", detail, ["Check the media id (e.g. 'M2') and op arguments."])


def _media_not_found(media_id: str) -> ToolError:
    return ToolError(
        "not_found",
        f"Media {media_id} does not exist.",
        ["Call docx_outline or a projection to see media ids."],
    )


# ---------------------------------------------------------------------------
# Pixel-dimension parsing (§19)
# ---------------------------------------------------------------------------


def _png_dimensions(data: bytes) -> tuple[int, int] | None:
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        return None
    if data[12:16] != b"IHDR":
        return None
    width = int.from_bytes(data[16:20], "big")
    height = int.from_bytes(data[20:24], "big")
    return width, height


def _jpeg_dimensions(data: bytes) -> tuple[int, int] | None:
    if not data.startswith(b"\xff\xd8"):
        return None
    pos = 2
    n = len(data)
    while pos + 1 < n:
        if data[pos] != 0xFF:
            pos += 1
            continue
        marker = data[pos + 1]
        if marker in (0xC0, 0xC1, 0xC2):
            # SOFn: length(2) precision(1) height(2) width(2)
            height = int.from_bytes(data[pos + 5 : pos + 7], "big")
            width = int.from_bytes(data[pos + 7 : pos + 9], "big")
            return width, height
        if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
            pos += 2
            continue
        length = int.from_bytes(data[pos + 2 : pos + 4], "big")
        pos += 2 + length
    return None


def _pixel_dimensions(data: bytes, ext: str) -> tuple[int, int] | None:
    if ext == "png":
        return _png_dimensions(data)
    if ext in ("jpg", "jpeg"):
        return _jpeg_dimensions(data)
    return None


def _compute_emu(
    data: bytes, ext: str, width_cm: float | None, height_cm: float | None
) -> tuple[int, int]:
    """EMU (cx, cy) honoring aspect when one dimension is given (§19)."""
    if width_cm is not None and height_cm is not None:
        return round(width_cm * _EMU_PER_CM), round(height_cm * _EMU_PER_CM)
    if width_cm is not None:
        dims = _pixel_dimensions(data, ext)
        cx = round(width_cm * _EMU_PER_CM)
        if dims is not None and dims[0] > 0:
            return cx, round(cx * (dims[1] / dims[0]))
        return cx, cx
    if height_cm is not None:
        dims = _pixel_dimensions(data, ext)
        cy = round(height_cm * _EMU_PER_CM)
        if dims is not None and dims[1] > 0:
            return round(cy * (dims[0] / dims[1])), cy
        return cy, cy
    # Neither given → the source's native pixels at 96 dpi (914400 EMU/in), else 4cm².
    dims = _pixel_dimensions(data, ext)
    if dims is not None and dims[0] > 0 and dims[1] > 0:
        return round(dims[0] / 96 * 914400), round(dims[1] / 96 * 914400)
    fallback = round(4 * _EMU_PER_CM)
    return fallback, fallback


# ---------------------------------------------------------------------------
# Drawing XML (§19)
# ---------------------------------------------------------------------------


def _drawing_xml(rel_id: str, cx: int, cy: int, name: str) -> str:
    return (
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="1" name="{_xml.escape_attr(name)}"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="1" name="{_xml.escape_attr(name)}"/>'
        "<pic:cNvPicPr/></pic:nvPicPr>"
        f'<pic:blipFill><a:blip r:embed="{rel_id}"/>'
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )


# ---------------------------------------------------------------------------
# Media-part bookkeeping
# ---------------------------------------------------------------------------


def _next_image_number(package: Package) -> int:
    pattern = re.compile(r"^word/media/image([0-9]+)\.")
    max_n = 0
    for name in package.part_names:
        m = pattern.match(name)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def _extension_of(part_name: str) -> str:
    """Lowercased extension of a path/part basename (``""`` when none) (JS parity)."""
    base = posixpath.basename(part_name)
    return base.rpartition(".")[2].lower() if "." in base else ""


def _read_image(path: str) -> tuple[bytes, str]:
    try:
        with open(path, "rb") as fh:
            return fh.read(), _extension_of(path)
    except OSError as exc:
        raise ToolError(
            "not_found",
            f"Cannot read image {path}: {exc.strerror or exc}.",
            ["Check the path to the source image."],
        ) from exc


# ---------------------------------------------------------------------------
# insert (§19)
# ---------------------------------------------------------------------------


def _insert(
    package: Package,
    after: str | None,
    before: str | None,
    image: str,
    width_cm: float | None,
    height_cm: float | None,
) -> tuple[str, str | None, str]:
    if (after is None) == (before is None):
        raise _media_invalid("op 'insert' requires exactly one of after or before.")
    img_bytes, ext = _read_image(image)
    main = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    anchor = after if after is not None else before
    assert anchor is not None
    entry = _edits.require_paragraph(entries, anchor)

    number = _next_image_number(package)
    name = f"image{number}"
    part_name = f"word/media/image{number}.{ext}"
    cx, cy = _compute_emu(img_bytes, ext, width_cm, height_cm)

    package.set_part(part_name, img_bytes)
    _parts.ensure_content_type_default(
        package, ext, _IMAGE_CONTENT_TYPES.get(ext, "application/octet-stream")
    )
    rel_id = _parts.next_rel_id(package, main)
    _parts.add_relationship(
        package, main, rel_id, _MEDIA_REL_TYPE, f"media/image{number}.{ext}"
    )

    data = package.part(main)
    run_p = f"<w:p>{_drawing_xml(rel_id, cx, cy, name)}</w:p>"
    position = entry.span.end if after is not None else entry.span.start
    package.set_part(main, _xml.splice(data, [(position, position, run_p.encode("utf-8"))]))

    media_id = _media_id_at(package, rel_id)
    fresh = _edits.paragraph_entries(package)
    new_ord = entry.ordinal + 1 if after is not None else entry.ordinal
    new_anchor = fresh[new_ord - 1].anchor if new_ord - 1 < len(fresh) else anchor
    note = f"Inserted {name}.{ext} ({cx}×{cy} EMU)."
    return media_id, new_anchor, note


# ---------------------------------------------------------------------------
# M{ordinal} resolution
# ---------------------------------------------------------------------------


def _blip_embeds(data: bytes) -> list[str]:
    """rIds of every ``a:blip`` (``r:embed`` then ``r:link``) in document order (§19/§13)."""
    out: list[str] = []
    for m in re.finditer(rb"<a:blip\b[^>]*>", data):
        tag = m.group(0)
        rid = re.search(rb'\br:embed="([^"]*)"', tag) or re.search(rb'\br:link="([^"]*)"', tag)
        out.append(rid.group(1).decode("utf-8") if rid else "")
    return out


def _media_id_at(package: Package, rel_id: str) -> str:
    main = package.main_document_part()
    embeds = _blip_embeds(package.part(main))
    for i, rid in enumerate(embeds, start=1):
        if rid == rel_id:
            return f"M{i}"
    return f"M{len(embeds)}"


def _resolve_media(package: Package, media_id: str) -> tuple[str, str]:
    """Return ``(part_name, rel_id)`` for an ``M{ordinal}`` reference (§19)."""
    m = _MEDIA_ID_RE.match(media_id)
    if not m:
        raise _media_invalid(f"Malformed media id: {media_id}.")
    ordinal = int(m.group(1))
    main = package.main_document_part()
    embeds = _blip_embeds(package.part(main))
    if ordinal < 1 or ordinal > len(embeds):
        raise _media_not_found(media_id)
    rel_id = embeds[ordinal - 1]
    for rel in package.rels(main):
        if rel.rel_id == rel_id and not rel.is_external:
            from ._opc import resolve_rel_target

            return resolve_rel_target(main, rel.target), rel_id
    raise _media_not_found(media_id)


# ---------------------------------------------------------------------------
# extract / replace (§19)
# ---------------------------------------------------------------------------


def _extract(package: Package, media_id: str, path: str) -> str:
    part_name, _ = _resolve_media(package, media_id)
    if not package.has_part(part_name):
        raise _media_not_found(media_id)
    data = package.part(part_name)
    try:
        with open(path, "wb") as fh:
            fh.write(data)
    except OSError as exc:
        raise ToolError(
            "save_failed",
            f"I/O failure writing media to {path}: {exc.strerror or exc}.",
            ["Check the path and permissions."],
        ) from exc
    return path


def _replace(package: Package, media_id: str, image: str) -> str:
    part_name, _ = _resolve_media(package, media_id)
    new_bytes, new_ext = _read_image(image)
    old_ext = _extension_of(part_name)
    package.set_part(part_name, new_bytes)
    if new_ext != old_ext:
        _parts.ensure_content_type_default(
            package, new_ext, _IMAGE_CONTENT_TYPES.get(new_ext, "application/octet-stream")
        )
    return media_id


# ---------------------------------------------------------------------------
# docx_media
# ---------------------------------------------------------------------------


def docx_media(
    session: Session,
    *,
    doc_id: str,
    op: str,
    after: str | None = None,
    before: str | None = None,
    image: str | None = None,
    width_cm: float | None = None,
    height_cm: float | None = None,
    media_id: str | None = None,
    path: str | None = None,
    track_changes: bool = False,
    author: str | None = None,
) -> dict[str, object]:
    """Insert, extract, or replace images (§19)."""
    doc = session.get(doc_id)
    package = doc.package
    if op == "insert":
        if image is None:
            raise _media_invalid("op 'insert' requires image.")
        mid, new_anchor, note = _insert(package, after, before, image, width_cm, height_cm)
        doc.mark_dirty()
        result: dict[str, object] = {"media_id": mid, "note": note}
        if new_anchor is not None:
            result["new_anchor"] = new_anchor
        return result
    if op == "extract":
        if media_id is None:
            raise _media_invalid("op 'extract' requires media_id.")
        if path is None:
            raise _media_invalid("op 'extract' requires path.")
        out_path = _extract(package, media_id, path)
        return {"media_id": media_id, "path": out_path, "note": f"Extracted {media_id}."}
    if op == "replace":
        if media_id is None:
            raise _media_invalid("op 'replace' requires media_id.")
        if image is None:
            raise _media_invalid("op 'replace' requires image.")
        mid = _replace(package, media_id, image)
        doc.mark_dirty()
        return {"media_id": mid, "note": f"Replaced {mid} bytes."}
    raise _media_invalid(f"Unknown media op: {op}.")
