"""OPC package layer: lazy zip access, raw part bytes, rels/content-types, save.

Fidelity invariant: every part is held as the exact bytes read from the source zip;
only parts explicitly replaced via :meth:`Package.set_part` ever change. Saving
streams the source entries in their original order — untouched parts byte-for-byte
(re-stored with normalized metadata), new parts appended — per algorithms.md §9.

Rels and ``[Content_Types].xml`` are parsed read-only with the stdlib ElementTree;
the §3 no-DOM rule applies to *editing* (those parts are never re-serialized from a
tree here).
"""

from __future__ import annotations

import contextlib
import io
import os
import posixpath
import tempfile
import zipfile
from dataclasses import dataclass, field
from xml.etree import ElementTree as ET

from ._errors import ToolError
from ._limits import check_archive, forbid_doctype, is_xml_part, max_part_bytes

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_DOCUMENT_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)

CONTENT_TYPES_PART = "[Content_Types].xml"


def _normalize_part_name(name: str) -> str:
    """OPC part names may be written with a leading ``/``; zip entries are not."""
    return name.lstrip("/")


@dataclass(frozen=True, slots=True)
class Relationship:
    """One ``<Relationship>`` from a ``.rels`` part, in document order."""

    rel_id: str
    rel_type: str
    target: str
    target_mode: str = "Internal"

    @property
    def is_external(self) -> bool:
        return self.target_mode == "External"


@dataclass(frozen=True, slots=True)
class ContentTypes:
    """Parsed ``[Content_Types].xml``: extension defaults and part-name overrides."""

    defaults: dict[str, str] = field(default_factory=dict)  # lowercase extension -> type
    overrides: dict[str, str] = field(default_factory=dict)  # part name (no leading /) -> type

    def content_type_of(self, part_name: str) -> str | None:
        name = _normalize_part_name(part_name)
        if name in self.overrides:
            return self.overrides[name]
        _, _, ext = name.rpartition(".")
        return self.defaults.get(ext.lower())


def rels_part_for(part_name: str | None) -> str:
    """The ``.rels`` part that holds relationships sourced from ``part_name``.

    ``None`` (or ``""``) means the package root: ``_rels/.rels``.
    """
    if not part_name:
        return "_rels/.rels"
    name = _normalize_part_name(part_name)
    directory, base = posixpath.split(name)
    rels = f"_rels/{base}.rels"
    return f"{directory}/{rels}" if directory else rels


def resolve_rel_target(source_part: str | None, target: str) -> str:
    """Resolve an Internal relationship target to a package part name.

    ``..`` segments are collapsed and *clamped at the package root*: a target can
    never resolve to a name with leading ``..`` or an absolute path, so a hostile
    relationship cannot escape the package (matches the TypeScript engine).
    """
    if target.startswith("/"):
        return _normalize_part_name(target)
    base_dir = posixpath.dirname(_normalize_part_name(source_part or ""))
    segments: list[str] = []
    for segment in posixpath.join(base_dir, target).split("/"):
        if segment in ("", "."):
            continue
        if segment == "..":
            if segments:
                segments.pop()  # else: drop — cannot climb above the root
        else:
            segments.append(segment)
    return "/".join(segments)


class Package:
    """An OPC package over raw part bytes, opened from a path or bytes."""

    def __init__(self, archive: bytes, source_path: str | None = None) -> None:
        try:
            self._zip = zipfile.ZipFile(io.BytesIO(archive))
        except zipfile.BadZipFile as exc:
            raise ToolError(
                "open_failed",
                f"Cannot open {source_path or '<bytes>'}: not a zip archive ({exc}).",
                ["Check the path; the file is not a .docx package."],
            ) from exc
        infolist = self._zip.infolist()
        check_archive([(i.filename, i.file_size, i.compress_size) for i in infolist])
        self._original_order: list[str] = [info.filename for info in infolist]
        self._original: frozenset[str] = frozenset(self._original_order)
        self._cache: dict[str, bytes] = {}  # lazily decompressed original parts
        self._dirty: dict[str, bytes] = {}  # modified + new parts, insertion-ordered
        self._new_order: list[str] = []  # new part names in creation order
        self.source_path = source_path
        if CONTENT_TYPES_PART not in self._original:
            raise ToolError(
                "open_failed",
                f"Cannot open {source_path or '<bytes>'}: zip has no {CONTENT_TYPES_PART} "
                "(not an OPC package).",
                ["Check the path; the message says what the file actually is."],
            )

    # -- opening ---------------------------------------------------------------

    @classmethod
    def open(cls, source: str | os.PathLike[str] | bytes | bytearray) -> Package:
        """Open a package from a filesystem path or in-memory zip bytes."""
        if isinstance(source, bytes | bytearray):
            return cls(bytes(source))
        path = os.fspath(source)
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except OSError as exc:
            raise ToolError(
                "open_failed",
                f"Cannot open {path}: {exc.strerror or exc}.",
                ["Check the path and permissions."],
            ) from exc
        return cls(data, source_path=path)

    # -- parts -----------------------------------------------------------------

    @property
    def part_names(self) -> list[str]:
        """All entry names: source order first, then new parts in creation order."""
        return [*self._original_order, *self._new_order]

    @property
    def dirty_part_names(self) -> tuple[str, ...]:
        """Names whose bytes differ from the source (modified or new), in edit order."""
        return tuple(self._dirty)

    def has_part(self, name: str) -> bool:
        name = _normalize_part_name(name)
        return name in self._dirty or name in self._original

    def part(self, name: str) -> bytes:
        """The part's current raw bytes (lazily decompressed on first access)."""
        name = _normalize_part_name(name)
        if name in self._dirty:
            return self._dirty[name]
        if name not in self._cache:
            if name not in self._original:
                raise KeyError(name)
            self._cache[name] = self._read_original(name)
        return self._cache[name]

    def _read_original(self, name: str) -> bytes:
        """Decompress an original part through a bounded reader and screen its XML.

        The §27 metadata pre-check already refused declared bombs at open; this
        bounds memory against a central directory that lies about its sizes, and
        rejects a DTD/entity declaration in any XML part on first read.
        """
        if name.endswith("/"):
            return b""
        cap = max_part_bytes()
        with self._zip.open(name) as fh:
            data = fh.read(cap + 1)
        if len(data) > cap:
            raise ToolError(
                "doc_too_large",
                f"Part {name} decompresses past the {cap}-byte cap.",
                ["Raise DOCXENGINE_MAX_PART_BYTES if the file is trusted."],
            )
        if is_xml_part(name):
            forbid_doctype(name, data)
        return data

    def set_part(self, name: str, data: bytes) -> None:
        """Replace (or create) a part's bytes; the part is marked dirty."""
        name = _normalize_part_name(name)
        if name not in self._original and name not in self._dirty:
            self._new_order.append(name)
        self._dirty[name] = bytes(data)

    def is_dirty(self, name: str) -> bool:
        return _normalize_part_name(name) in self._dirty

    # -- package metadata --------------------------------------------------------

    def content_types(self) -> ContentTypes:
        """Parse ``[Content_Types].xml`` (current bytes, including pending edits)."""
        root = ET.fromstring(self.part(CONTENT_TYPES_PART))
        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        for el in root.iter():
            if el.tag == f"{{{_CT_NS}}}Default":
                defaults[el.attrib["Extension"].lower()] = el.attrib["ContentType"]
            elif el.tag == f"{{{_CT_NS}}}Override":
                part_name = _normalize_part_name(el.attrib["PartName"])
                overrides[part_name] = el.attrib["ContentType"]
        return ContentTypes(defaults, overrides)

    def rels(self, part_name: str | None = None) -> list[Relationship]:
        """Relationships sourced from ``part_name`` (``None`` = package root)."""
        rels_name = rels_part_for(part_name)
        if not self.has_part(rels_name):
            return []
        root = ET.fromstring(self.part(rels_name))
        out: list[Relationship] = []
        for el in root:
            if el.tag != f"{{{_RELS_NS}}}Relationship":
                continue
            out.append(
                Relationship(
                    rel_id=el.attrib["Id"],
                    rel_type=el.attrib["Type"],
                    target=el.attrib["Target"],
                    target_mode=el.attrib.get("TargetMode", "Internal"),
                )
            )
        return out

    def main_document_part(self) -> str:
        """The officeDocument target from the root rels (fallback ``word/document.xml``)."""
        for rel in self.rels(None):
            if rel.rel_type == OFFICE_DOCUMENT_REL_TYPE and not rel.is_external:
                return resolve_rel_target(None, rel.target)
        return "word/document.xml"

    # -- save (algorithms.md §9) --------------------------------------------------

    def save(self, path: str | os.PathLike[str]) -> None:
        """Write the package atomically with normalized zip metadata.

        Source entries stream in their original order; untouched parts pass through
        with identical decompressed content; new parts append in creation order.
        Validation gating (§9 step 1) is layered on by ``docx_save``, not here.
        """
        dest = os.fspath(path)
        archive = self.to_bytes()
        dest_dir = os.path.dirname(dest) or "."
        tmp_path: str | None = None
        try:
            fd, tmp_path = tempfile.mkstemp(dir=dest_dir, prefix=".docxengine-", suffix=".tmp")
            with os.fdopen(fd, "wb") as fh:
                fh.write(archive)
            os.replace(tmp_path, dest)
        except OSError as exc:
            if tmp_path is not None:
                with contextlib.suppress(OSError):
                    os.unlink(tmp_path)
            raise ToolError(
                "save_failed",
                f"I/O failure writing output to {dest}: {exc.strerror or exc}.",
                ["Check the path and permissions."],
            ) from exc

    def to_bytes(self) -> bytes:
        """The package serialized to zip bytes with the §9 normalization."""
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zout:
            for name in self.part_names:
                data = b"" if name.endswith("/") else self.part(name)
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.create_system = 0  # normalized: DOS, no platform variance
                info.external_attr = 0
                # ZipInfo defaults: no extra field, no comment.
                zout.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=6)
        return buf.getvalue()
