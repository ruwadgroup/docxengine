"""Lightweight XML scanner and splice helpers (algorithms.md §3/§4).

Every part is held as raw UTF-8 bytes. This module tokenizes a part's XML *just
enough* to locate element boundaries as byte offsets in the original buffer; edits
replace exactly the spliced byte ranges. There is no DOM build-then-serialize step:
attribute order, namespace prefixes, inter-element whitespace, and rsid attributes
in untouched regions survive verbatim.

Element names are matched as written (the conventional ``w:`` prefix); multi-byte
UTF-8 sequences never contain ASCII bytes, so scanning the raw bytes is safe.
"""

from __future__ import annotations

import re
from collections.abc import Collection, Iterable, Iterator
from dataclasses import dataclass

# The Unicode White_Space=Yes set, exactly as pinned by algorithms.md §1 step 3.
WHITESPACE = (
    "\t\n\x0b\x0c\r \x85\xa0\u1680"
    "\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a"
    "\u2028\u2029\u202f\u205f\u3000"
)

_NAME_TERMINATORS = frozenset(b" \t\r\n/>")


@dataclass(frozen=True, slots=True)
class Span:
    """An element located in a byte buffer.

    ``start``/``end`` delimit the whole element (``<`` of the start tag through the
    ``>`` of the end tag); ``inner_start``/``inner_end`` delimit its content. For an
    empty element (``<w:t/>``) all inner offsets equal ``end``.
    """

    name: str
    start: int
    end: int
    inner_start: int
    inner_end: int
    empty: bool


@dataclass(frozen=True, slots=True)
class TextPiece:
    """One ``w:t``'s decoded character data within a paragraph's concatenated text.

    Maps text indices ``[start, start + len(text))`` back to the ``w:t`` element and
    its enclosing ``w:r`` (algorithms.md §4 offset map).
    """

    text: str
    start: int
    t: Span
    run: Span | None


def _scan_tag(data: bytes, pos: int) -> tuple[str, str, int]:
    """Scan the markup starting at ``data[pos] == b"<"``.

    Returns ``(kind, name, end)`` where ``end`` is just past the closing ``>``.
    ``kind`` is one of ``start``, ``empty``, ``end``, ``comment``, ``cdata``, ``pi``,
    ``doctype``.
    """
    if data.startswith(b"<!--", pos):
        close = data.find(b"-->", pos + 4)
        if close < 0:
            raise ValueError(f"unterminated comment at byte {pos}")
        return "comment", "", close + 3
    if data.startswith(b"<![CDATA[", pos):
        close = data.find(b"]]>", pos + 9)
        if close < 0:
            raise ValueError(f"unterminated CDATA section at byte {pos}")
        return "cdata", "", close + 3
    if data.startswith(b"<!", pos):
        close = data.find(b">", pos + 2)
        if close < 0:
            raise ValueError(f"unterminated declaration at byte {pos}")
        return "doctype", "", close + 1
    if data.startswith(b"<?", pos):
        close = data.find(b"?>", pos + 2)
        if close < 0:
            raise ValueError(f"unterminated processing instruction at byte {pos}")
        return "pi", "", close + 2
    if data.startswith(b"</", pos):
        close = data.find(b">", pos + 2)
        if close < 0:
            raise ValueError(f"unterminated end tag at byte {pos}")
        return "end", data[pos + 2 : close].strip().decode("utf-8"), close + 1

    # Start or empty-element tag: parse the name, then skip attributes respecting quotes.
    i = pos + 1
    n = len(data)
    while i < n and data[i] not in _NAME_TERMINATORS:
        i += 1
    name = data[pos + 1 : i].decode("utf-8")
    while i < n:
        c = data[i]
        if c == 0x22 or c == 0x27:  # " or '
            close = data.find(data[i : i + 1], i + 1)
            if close < 0:
                raise ValueError(f"unterminated attribute value at byte {i}")
            i = close + 1
        elif c == 0x3E:  # >
            kind = "empty" if data[i - 1] == 0x2F else "start"  # preceded by /
            return kind, name, i + 1
        else:
            i += 1
    raise ValueError(f"unterminated tag at byte {pos}")


def _tokens(data: bytes, lo: int, hi: int) -> Iterator[tuple[str, str, int, int]]:
    """Yield ``(kind, name, start, end)`` tokens for ``data[lo:hi]``."""
    pos = lo
    while pos < hi:
        lt = data.find(b"<", pos, hi)
        if lt < 0:
            if pos < hi:
                yield "text", "", pos, hi
            return
        if lt > pos:
            yield "text", "", pos, lt
        kind, name, end = _scan_tag(data, lt)
        yield kind, name, lt, end
        pos = end


def iter_elements(
    data: bytes,
    lo: int = 0,
    hi: int | None = None,
    names: Collection[str] | None = None,
    max_depth: int | None = None,
) -> Iterator[Span]:
    """Yield element :class:`Span`\\ s found in ``data[lo:hi]``.

    ``names`` filters by qualified name as written; ``max_depth=1`` restricts to
    elements opened at the top level of the scanned range (direct children when the
    range is a parent's content). Same-name, non-nesting elements (the OOXML cases
    this engine scans for) are yielded in document order.
    """
    if hi is None:
        hi = len(data)
    stack: list[tuple[str, int, int]] = []  # (name, start, inner_start)
    for kind, name, a, b in _tokens(data, lo, hi):
        if kind == "start":
            stack.append((name, a, b))
        elif kind == "empty":
            depth = len(stack)
            if (names is None or name in names) and (max_depth is None or depth < max_depth):
                yield Span(name, a, b, b, b, True)
        elif kind == "end":
            if not stack:
                raise ValueError(f"unbalanced end tag </{name}> at byte {a}")
            open_name, start, inner_start = stack.pop()
            if open_name != name:
                raise ValueError(f"mismatched end tag </{name}> at byte {a} (open <{open_name}>)")
            depth = len(stack)
            if (names is None or name in names) and (max_depth is None or depth < max_depth):
                yield Span(name, start, b, inner_start, a, False)
    if stack:
        raise ValueError(f"unclosed element <{stack[-1][0]}> at byte {stack[-1][1]}")


def find_element(data: bytes, name: str, lo: int = 0, hi: int | None = None) -> Span | None:
    """The first element named ``name`` in ``data[lo:hi]``, or ``None``."""
    return next(iter_elements(data, lo, hi, names=(name,)), None)


def find_body(data: bytes) -> Span:
    """Locate the ``w:body`` element of a document part."""
    body = find_element(data, "w:body")
    if body is None:
        raise ValueError("no w:body element in document part")
    return body


def iter_body_children(data: bytes) -> Iterator[Span]:
    """Direct children of ``w:body`` in document order (``w:p``, ``w:tbl``, ``w:sectPr``…)."""
    body = find_body(data)
    yield from iter_elements(data, body.inner_start, body.inner_end, max_depth=1)


_ENTITY = re.compile(r"&(#x[0-9A-Fa-f]+|#[0-9]+|amp|lt|gt|quot|apos);")
_NAMED_ENTITIES = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}


def _entity_repl(m: re.Match[str]) -> str:
    ref = m.group(1)
    if ref.startswith("#x"):
        return chr(int(ref[2:], 16))
    if ref.startswith("#"):
        return chr(int(ref[1:]))
    return _NAMED_ENTITIES[ref]


def unescape(raw: str) -> str:
    """Decode the five XML entities and numeric character references."""
    return _ENTITY.sub(_entity_repl, raw)


def escape_text(text: str) -> str:
    """Escape character data per §3: exactly ``&``, ``<``, ``>``."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def escape_attr(value: str) -> str:
    """Escape an attribute value per §3: text escapes plus ``"``."""
    return escape_text(value).replace('"', "&quot;")


def emit_text_element(text: str, tag: str = "w:t") -> str:
    """Emit a ``w:t``/``w:delText`` with ``xml:space="preserve"`` iff needed (§3)."""
    preserve = bool(text) and (text[0] in WHITESPACE or text[-1] in WHITESPACE)
    attr = ' xml:space="preserve"' if preserve else ""
    return f"<{tag}{attr}>{escape_text(text)}</{tag}>"


def element_text(data: bytes, span: Span) -> str:
    """Decoded character data of a text-only element such as ``w:t``."""
    if span.empty:
        return ""
    return unescape(data[span.inner_start : span.inner_end].decode("utf-8"))


def paragraph_text(data: bytes, p: Span) -> tuple[str, list[TextPiece]]:
    """Concatenated ``w:t`` text of a paragraph plus the §4 offset map.

    ``w:delText`` is a distinct element name and is therefore excluded; ``w:tab``,
    ``w:br``, and every other element contribute nothing.
    """
    if p.empty:
        return "", []
    runs = sorted(
        iter_elements(data, p.inner_start, p.inner_end, names=("w:r",)),
        key=lambda r: r.start,
    )
    pieces: list[TextPiece] = []
    parts: list[str] = []
    offset = 0
    for t in iter_elements(data, p.inner_start, p.inner_end, names=("w:t",)):
        text = element_text(data, t)
        run = None
        for r in runs:  # runs cannot nest; at most one contains this w:t
            if r.start < t.start and t.end <= r.end:
                run = r
        pieces.append(TextPiece(text, offset, t, run))
        parts.append(text)
        offset += len(text)
    return "".join(parts), pieces


def locate(pieces: list[TextPiece], index: int) -> tuple[TextPiece, int]:
    """Map a concatenated-text index to ``(piece, char offset within that w:t)``."""
    for piece in pieces:
        if piece.start <= index < piece.start + len(piece.text):
            return piece, index - piece.start
    raise IndexError(f"text index {index} out of range")


def splice(data: bytes, edits: Iterable[tuple[int, int, bytes]]) -> bytes:
    """Replace byte ranges ``(start, end, replacement)``; ranges must not overlap."""
    out = bytearray()
    pos = 0
    for start, end, replacement in sorted(edits, key=lambda e: (e[0], e[1])):
        if start < pos or end < start:
            raise ValueError("overlapping or inverted splice ranges")
        out += data[pos:start]
        out += replacement
        pos = end
    out += data[pos:]
    return bytes(out)
