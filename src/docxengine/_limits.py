"""Resource and content-safety limits for parsing untrusted packages.

DocxEngine opens documents from untrusted sources, so the OPC layer bounds a
hostile package's cost *before* it is paid (zip bombs) and refuses hostile XML
(DTD/entity declarations). Every cap is read from the environment on each open so
deployments can tune them; the defaults are generous for real documents and tight
against abuse. The TypeScript engine enforces the identical checks (parity).

See ``SECURITY.md``, ``ROADMAP.md`` Phase 3, and ``spec/algorithms.md`` §27.
"""

from __future__ import annotations

import os

from ._errors import ToolError

# Defaults — generous for genuine documents, bounded against abuse.
_DEFAULT_MAX_PARTS = 10_000
_DEFAULT_MAX_TOTAL_BYTES = 512 * 1024 * 1024  # 512 MiB uncompressed, whole package
_DEFAULT_MAX_PART_BYTES = 128 * 1024 * 1024  # 128 MiB uncompressed, single part
_DEFAULT_MAX_COMPRESSION_RATIO = 200  # uncompressed / compressed, per part
_DEFAULT_MAX_XML_DEPTH = 1_000  # element nesting depth

# Below this uncompressed size a part is never flagged by the ratio check:
# small, highly compressible parts are not a decompression bomb.
RATIO_FLOOR_BYTES = 64 * 1024

_XML_SUFFIXES = (".xml", ".rels")


def _int_env(name: str, default: int) -> int:
    """A positive integer from the environment, or ``default`` if unset/invalid."""
    raw = os.environ.get(name)
    if raw is None or raw.strip() == "":
        return default
    try:
        value = int(raw)
    except ValueError:
        return default
    return value if value > 0 else default


def max_parts() -> int:
    return _int_env("DOCXENGINE_MAX_PARTS", _DEFAULT_MAX_PARTS)


def max_total_bytes() -> int:
    return _int_env("DOCXENGINE_MAX_TOTAL_BYTES", _DEFAULT_MAX_TOTAL_BYTES)


def max_part_bytes() -> int:
    return _int_env("DOCXENGINE_MAX_PART_BYTES", _DEFAULT_MAX_PART_BYTES)


def max_compression_ratio() -> int:
    return _int_env("DOCXENGINE_MAX_COMPRESSION_RATIO", _DEFAULT_MAX_COMPRESSION_RATIO)


def max_xml_depth() -> int:
    return _int_env("DOCXENGINE_MAX_XML_DEPTH", _DEFAULT_MAX_XML_DEPTH)


def is_xml_part(name: str) -> bool:
    """True for parts parsed as XML (where DTD/entity declarations are a threat)."""
    return name.lower().endswith(_XML_SUFFIXES)


def check_archive(entries: list[tuple[str, int, int]]) -> None:
    """Refuse a package whose declared sizes exceed the configured caps.

    ``entries`` is ``(name, uncompressed_size, compressed_size)`` per zip entry,
    taken from the central directory — no decompression has happened yet.
    """
    cap_parts = max_parts()
    cap_total = max_total_bytes()
    cap_part = max_part_bytes()
    cap_ratio = max_compression_ratio()

    if len(entries) > cap_parts:
        raise ToolError(
            "doc_too_large",
            f"Package has {len(entries)} parts, over the {cap_parts}-part cap.",
            ["Split the document, or raise DOCXENGINE_MAX_PARTS if the file is trusted."],
        )

    total = 0
    for name, uncompressed, compressed in entries:
        total += uncompressed
        if uncompressed > cap_part:
            raise ToolError(
                "doc_too_large",
                f"Part {name} is {uncompressed} bytes uncompressed, over the {cap_part}-byte cap.",
                ["Raise DOCXENGINE_MAX_PART_BYTES if the file is trusted."],
            )
        if (
            uncompressed > RATIO_FLOOR_BYTES
            and compressed > 0
            and uncompressed / compressed > cap_ratio
        ):
            raise ToolError(
                "doc_too_large",
                f"Part {name} compresses {uncompressed / compressed:.0f}:1, over the "
                f"{cap_ratio}:1 ratio cap (possible zip bomb).",
                ["Raise DOCXENGINE_MAX_COMPRESSION_RATIO if the file is trusted."],
            )

    if total > cap_total:
        raise ToolError(
            "doc_too_large",
            f"Package is {total} bytes uncompressed, over the {cap_total}-byte cap.",
            ["Split the document, or raise DOCXENGINE_MAX_TOTAL_BYTES if the file is trusted."],
        )


def forbid_doctype(name: str, data: bytes) -> None:
    """Reject a DTD/entity declaration in an XML part (XXE / billion-laughs).

    XML keywords are case-sensitive; a conformant ``DOCTYPE``/``ENTITY`` is
    uppercase, so a substring scan is exact and cheap.
    """
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        raise ToolError(
            "malicious_content",
            f"Refusing {name}: contains a DTD/entity declaration (DOCTYPE/ENTITY).",
            ["Conformant Word documents never declare a DTD; treat this file as untrusted."],
        )
