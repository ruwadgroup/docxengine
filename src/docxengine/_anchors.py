"""Anchors: normalized text, hashes, and the package anchor index (algorithms.md §1).

A paragraph anchor is ``P{ordinal}#{hash}``: 1-based position among body-level
``w:p`` elements plus the first 4 lowercase hex chars of the SHA-256 of the UTF-8
encoding of the paragraph's normalized text. Tables anchor as ``T{ordinal}`` over
body-level ``w:tbl`` elements. The trailing ``w:sectPr`` is not a paragraph.
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from dataclasses import dataclass

from . import _xml
from ._opc import Package

_WS_RUN = re.compile(f"[{re.escape(_xml.WHITESPACE)}]+")

HASH_LENGTH = 4
EMPTY_HASH = "e3b0"  # SHA-256 of b"" begins e3b0c442…


def normalized_text(raw: str) -> str:
    """§1 normalization: NFC, collapse White_Space runs to one space, strip.

    ``raw`` is the already-concatenated ``w:t`` character data of the paragraph
    (``w:delText`` excluded — the hash sees the document as-if-accepted).
    """
    s = unicodedata.normalize("NFC", raw)
    s = _WS_RUN.sub(" ", s)
    return s.strip(" ")


def anchor_hash(normalized: str) -> str:
    """First 4 lowercase hex chars of SHA-256 over the UTF-8 normalized text."""
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:HASH_LENGTH]


def paragraph_anchor(ordinal: int, normalized: str) -> str:
    return f"P{ordinal}#{anchor_hash(normalized)}"


def table_anchor(ordinal: int) -> str:
    return f"T{ordinal}"


@dataclass(frozen=True, slots=True)
class AnchorEntry:
    """One body-level block: its anchor, kind, ordinal, byte span, normalized text."""

    anchor: str
    kind: str  # "paragraph" | "table"
    ordinal: int
    span: _xml.Span
    normalized: str  # "" for tables (cell text is a §2 projection concern)


def paragraph_normalized_text(data: bytes, p: _xml.Span) -> str:
    """Normalized text of one ``w:p`` span within a document part's bytes."""
    raw, _ = _xml.paragraph_text(data, p)
    return normalized_text(raw)


def build_anchor_index(package: Package, part_name: str | None = None) -> list[AnchorEntry]:
    """Anchor every body-level ``w:p``/``w:tbl`` of the main document part, in order."""
    name = part_name if part_name is not None else package.main_document_part()
    data = package.part(name)
    entries: list[AnchorEntry] = []
    p_ordinal = 0
    t_ordinal = 0
    for child in _xml.iter_body_children(data):
        if child.name == "w:p":
            p_ordinal += 1
            norm = paragraph_normalized_text(data, child)
            entries.append(
                AnchorEntry(paragraph_anchor(p_ordinal, norm), "paragraph", p_ordinal, child, norm)
            )
        elif child.name == "w:tbl":
            t_ordinal += 1
            entries.append(AnchorEntry(table_anchor(t_ordinal), "table", t_ordinal, child, ""))
    return entries
