"""Edit-surface helpers (algorithms.md §4–§7, §6a).

Anchor validation for edits, §4 offset-map replace splices, §5 tracked-change
emission, the §6 LCS word diff, and §7 revision scan/resolve plus the run-merge
post-pass. Everything operates on raw part bytes via :mod:`._xml` splicing —
untouched byte regions survive verbatim.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from datetime import UTC, datetime
from itertools import pairwise
from xml.etree import ElementTree as ET

from . import _xml
from ._anchors import AnchorEntry, build_anchor_index
from ._errors import ToolError
from ._opc import Package

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: Full edit-grade anchor: edits always validate the hash (§6a).
FULL_ANCHOR_RE = re.compile(r"^P([1-9][0-9]*)#([0-9a-f]{4})$")
#: Paragraph range; endpoint hashes are validated when present (§6a).
RANGE_RE = re.compile(r"^P([1-9][0-9]*)(?:#([0-9a-f]{4}))?\.\.P([1-9][0-9]*)(?:#([0-9a-f]{4}))?$")

_ATTR_RE = re.compile(r'([^\s=/>"\']+)\s*=\s*(?:"([^"]*)"|\'([^\']*)\')')
_RSID_ATTR_RE = re.compile(rb'\s+w:rsid[A-Za-z]*="[^"]*"')
_TOKEN_RE = re.compile(f"[{re.escape(_xml.WHITESPACE)}]+|[^{re.escape(_xml.WHITESPACE)}]+")

REVISION_NAMES = ("w:ins", "w:del")


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


def anchor_invalid_error(detail: str) -> ToolError:
    return ToolError(
        "anchor_invalid", detail, ["Check the format 'P{index}#{hash}' (ranges: 'P10..P24')."]
    )


def anchor_not_found_error(label: str) -> ToolError:
    return ToolError(
        "anchor_not_found",
        f"Anchor {label} not found: index out of range.",
        ["Call docx_outline to re-map anchors."],
    )


def anchor_stale_error(anchor: str) -> ToolError:
    return ToolError(
        "anchor_stale",
        f"Anchor {anchor} is stale: the hash no longer matches the paragraph content.",
        ["Call docx_read {anchor, window} and retry with the fresh anchor."],
    )


# ---------------------------------------------------------------------------
# Anchor validation (§1, §6a)
# ---------------------------------------------------------------------------


def paragraph_entries(package: Package) -> list[AnchorEntry]:
    """Body paragraphs only, indexable by ``ordinal - 1``."""
    return [e for e in build_anchor_index(package) if e.kind == "paragraph"]


def entry_at(entries: list[AnchorEntry], ordinal: int, label: str) -> AnchorEntry:
    if ordinal > len(entries):
        raise anchor_not_found_error(label)
    return entries[ordinal - 1]


def require_paragraph(entries: list[AnchorEntry], anchor: str) -> AnchorEntry:
    """§6a validation order: parse → ordinal in range → hash match."""
    m = FULL_ANCHOR_RE.match(anchor)
    if not m:
        raise anchor_invalid_error(f"Malformed anchor string: {anchor}.")
    entry = entry_at(entries, int(m.group(1)), anchor)
    if entry.anchor != anchor:
        raise anchor_stale_error(anchor)
    return entry


def paragraph_span_at(data: bytes, ordinal: int) -> _xml.Span:
    """The current byte span of body paragraph ``ordinal`` (recomputed per splice)."""
    count = 0
    for child in _xml.iter_body_children(data):
        if child.name == "w:p":
            count += 1
            if count == ordinal:
                return child
    raise anchor_not_found_error(f"P{ordinal}")


# ---------------------------------------------------------------------------
# Tracked-change metadata (§5)
# ---------------------------------------------------------------------------


def resolve_author(author: str | None) -> str:
    """§5: the ``author`` argument, else env ``DOCXENGINE_AUTHOR``, else ``DocxEngine``."""
    if author is not None:
        return author
    return os.environ.get("DOCXENGINE_AUTHOR") or "DocxEngine"


def revision_date() -> str:
    """§5: ``DOCXENGINE_FIXED_DATE`` verbatim, else current UTC ISO-8601 seconds + Z."""
    fixed = os.environ.get("DOCXENGINE_FIXED_DATE")
    if fixed:
        return fixed
    return datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def start_tag_attrs(data: bytes, span: _xml.Span) -> dict[str, str]:
    """Attributes of an element's start tag, decoded, as written."""
    end = span.end if span.empty else span.inner_start
    tag = data[span.start : end].decode("utf-8")
    out: dict[str, str] = {}
    for m in _ATTR_RE.finditer(tag):
        out[m.group(1)] = _xml.unescape(m.group(2) if m.group(2) is not None else m.group(3))
    return out


def next_revision_id(data: bytes) -> int:
    """§5 id allocation: max existing ``w:ins``/``w:del`` id (counted together) + 1."""
    max_id = 0
    for el in _xml.iter_elements(data, names=REVISION_NAMES):
        value = start_tag_attrs(data, el).get("w:id")
        if value is not None and value.isdigit():
            max_id = max(max_id, int(value))
    return max_id + 1


def revision_open(kind: str, rev_id: int, author: str, date: str) -> str:
    """A ``w:ins``/``w:del`` start tag with attributes in §5 order: id, author, date."""
    return (
        f'<w:{kind} w:id="{rev_id}" w:author="{_xml.escape_attr(author)}"'
        f' w:date="{_xml.escape_attr(date)}">'
    )


def run_rpr(data: bytes, span: _xml.Span) -> str:
    """The raw ``w:rPr`` of a run, decoded (``""`` when absent or not a run)."""
    if span.name != "w:r" or span.empty:
        return ""
    child = next(
        _xml.iter_elements(data, span.inner_start, span.inner_end, names=("w:rPr",), max_depth=1),
        None,
    )
    return data[child.start : child.end].decode("utf-8") if child is not None else ""


def emit_run(rpr: str, text: str, tag: str = "w:t") -> str:
    return f"<w:r>{rpr}{_xml.emit_text_element(text, tag)}</w:r>"


# ---------------------------------------------------------------------------
# §4 replace (plain) and §5 replace (tracked)
# ---------------------------------------------------------------------------


def count_occurrences(text: str, needle: str) -> int:
    """Literal, case-sensitive, non-overlapping, left-to-right (§2a/§6a)."""
    count = 0
    pos = 0
    while True:
        hit = text.find(needle, pos)
        if hit < 0:
            return count
        count += 1
        pos = hit + len(needle)


def _overlapping(pieces: list[_xml.TextPiece], start: int, end: int) -> list[_xml.TextPiece]:
    return [pc for pc in pieces if pc.start < end and pc.start + len(pc.text) > start]


def splice_replace_plain(
    data: bytes, pieces: list[_xml.TextPiece], start: int, end: int, new: str
) -> bytes:
    """§4: prefix + replacement into the first overlapping ``w:t``; suffixes after."""
    hits = _overlapping(pieces, start, end)
    run_t_counts: dict[int, int] = {}
    for pc in pieces:
        if pc.run is not None:
            run_t_counts[pc.run.start] = run_t_counts.get(pc.run.start, 0) + 1
    edits: list[tuple[int, int, bytes]] = []
    for i, pc in enumerate(hits):
        lo = max(0, start - pc.start)
        hi = min(len(pc.text), max(0, end - pc.start))
        text = pc.text[:lo] + new + pc.text[hi:] if i == 0 else pc.text[hi:]
        if text:
            edits.append((pc.t.start, pc.t.end, _xml.emit_text_element(text).encode("utf-8")))
        elif pc.run is not None and run_t_counts[pc.run.start] == 1:
            edits.append((pc.run.start, pc.run.end, b""))  # §4 rule 4: drop the emptied run
        else:
            edits.append((pc.t.start, pc.t.end, b""))
    return _xml.splice(data, edits)


def splice_replace_tracked(
    data: bytes,
    pieces: list[_xml.TextPiece],
    start: int,
    end: int,
    new: str,
    author: str,
    date: str,
) -> bytes:
    """§5: rebuild the matched run region as prefix-run + w:del + w:ins + suffix-run."""
    hits = _overlapping(pieces, start, end)
    containers: list[_xml.Span] = []
    for pc in hits:
        span = pc.run if pc.run is not None else pc.t
        if not containers or containers[-1].start != span.start:
            containers.append(span)
    first, last = containers[0], containers[-1]
    container_starts = {c.start for c in containers}
    grouped: dict[int, list[_xml.TextPiece]] = {}
    for pc in pieces:
        span = pc.run if pc.run is not None else pc.t
        if span.start in container_starts:
            grouped.setdefault(span.start, []).append(pc)

    def clamp(pc: _xml.TextPiece, index: int) -> int:
        return max(0, min(len(pc.text), index - pc.start))

    prefix = "".join(pc.text[: clamp(pc, start)] for pc in grouped[first.start])
    suffix = "".join(pc.text[clamp(pc, end) :] for pc in grouped[last.start])
    rpr_first = run_rpr(data, first)
    rev_id = next_revision_id(data)
    parts: list[str] = []
    if prefix:
        parts.append(emit_run(rpr_first, prefix))
    parts.append(revision_open("del", rev_id, author, date))
    rev_id += 1
    for container in containers:
        matched = "".join(
            pc.text[clamp(pc, start) : clamp(pc, end)] for pc in grouped[container.start]
        )
        parts.append(emit_run(run_rpr(data, container), matched, "w:delText"))
    parts.append("</w:del>")
    if new:
        parts.append(revision_open("ins", rev_id, author, date))
        parts.append(emit_run(rpr_first, new))
        parts.append("</w:ins>")
    if suffix:
        parts.append(emit_run(run_rpr(data, last), suffix))
    return _xml.splice(data, [(first.start, last.end, "".join(parts).encode("utf-8"))])


# ---------------------------------------------------------------------------
# §6 word-level diff
# ---------------------------------------------------------------------------


def diff_units(text: str) -> list[str]:
    """§6 step 1: units of word + following whitespace; leading whitespace → unit 1."""
    units: list[str] = []
    leading = ""
    for m in _TOKEN_RE.finditer(text):
        token = m.group()
        if token[0] in _xml.WHITESPACE:
            if units:
                units[-1] += token
            else:
                leading += token
        else:
            units.append(leading + token)
            leading = ""
    if leading:
        units.append(leading)
    return units


def word_diff(old: list[str], new: list[str]) -> list[tuple[str, str]]:
    """§6 step 2: LCS over units with the pinned deterministic forward backtrack."""
    n, m = len(old), len(new)
    lcs = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        row, below = lcs[i], lcs[i + 1]
        for j in range(m - 1, -1, -1):
            row[j] = below[j + 1] + 1 if old[i] == new[j] else max(below[j], row[j + 1])
    ops: list[tuple[str, str]] = []
    i = j = 0
    while i < n and j < m:
        if old[i] == new[j] and lcs[i][j] == lcs[i + 1][j + 1] + 1:
            ops.append(("keep", old[i]))
            i += 1
            j += 1
        elif lcs[i + 1][j] >= lcs[i][j + 1]:
            ops.append(("del", old[i]))
            i += 1
        else:
            ops.append(("ins", new[j]))
            j += 1
    ops.extend(("del", unit) for unit in old[i:])
    ops.extend(("ins", unit) for unit in new[j:])
    return ops


def diff_blocks(ops: list[tuple[str, str]]) -> list[tuple[str, str]]:
    """§6 step 3: maximal del/ins runs as one span each; del precedes ins at a position."""
    blocks: list[tuple[str, str]] = []
    dels: list[str] = []
    inss: list[str] = []

    def flush() -> None:
        if dels:
            blocks.append(("del", "".join(dels)))
            dels.clear()
        if inss:
            blocks.append(("ins", "".join(inss)))
            inss.clear()

    for op, unit in ops:
        if op == "keep":
            flush()
            if blocks and blocks[-1][0] == "keep":
                blocks[-1] = ("keep", blocks[-1][1] + unit)
            else:
                blocks.append(("keep", unit))
        elif op == "del":
            dels.append(unit)
        else:
            inss.append(unit)
    flush()
    return blocks


def _run_portions(
    data: bytes, pieces: list[_xml.TextPiece], start: int, end: int
) -> list[tuple[str, str]]:
    """``(rPr, portion)`` per overlapped run for ``[start, end)`` of the §4 text.

    Consecutive ``w:t`` pieces of the same run concatenate into one portion (§6a).
    """
    out: list[tuple[int, str, str]] = []  # (container start, rPr, portion)
    for pc in pieces:
        if pc.start >= end or pc.start + len(pc.text) <= start:
            continue
        lo = max(0, start - pc.start)
        hi = min(len(pc.text), end - pc.start)
        container = pc.run if pc.run is not None else pc.t
        if out and out[-1][0] == container.start:
            key, rpr, portion = out[-1]
            out[-1] = (key, rpr, portion + pc.text[lo:hi])
        else:
            out.append((container.start, run_rpr(data, container), pc.text[lo:hi]))
    return [(rpr, portion) for _, rpr, portion in out]


def _rpr_at_offset(data: bytes, pieces: list[_xml.TextPiece], offset: int) -> str:
    """§6a insert-only spans: the rPr of the run containing the insertion offset."""
    if not pieces:
        return ""  # empty paragraph yields no rPr
    for pc in pieces:
        if pc.start <= offset < pc.start + len(pc.text):
            return run_rpr(data, pc.run if pc.run is not None else pc.t)
    last = pieces[-1]  # end-of-paragraph insertion takes the last run's rPr
    return run_rpr(data, last.run if last.run is not None else last.t)


def rebuild_paragraph(
    data: bytes,
    p: _xml.Span,
    blocks: list[tuple[str, str]],
    *,
    tracked: bool,
    author: str,
    date: str,
) -> bytes:
    """Replace a paragraph's content after ``w:pPr`` from §6 diff blocks (§6a)."""
    ppr: _xml.Span | None = None
    first_run: _xml.Span | None = None
    if not p.empty:
        ppr = next(
            _xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:pPr",), max_depth=1),
            None,
        )
        first_run = next(_xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:r",)), None)
    ppr_xml = data[ppr.start : ppr.end].decode("utf-8") if ppr is not None else ""
    parts: list[str] = [ppr_xml]
    if not tracked:
        text = "".join(t for kind, t in blocks if kind != "del")
        if text:
            rpr = run_rpr(data, first_run) if first_run is not None else ""
            parts.append(emit_run(rpr, text))
    else:
        _, pieces = _xml.paragraph_text(data, p)
        rev_id = next_revision_id(data)
        pos = 0  # offset into the old §4 concatenated text
        replace_rpr: str | None = None  # first deleted run's rPr, for del→ins pairs
        for kind, text in blocks:
            if kind == "keep":
                for rpr, portion in _run_portions(data, pieces, pos, pos + len(text)):
                    parts.append(emit_run(rpr, portion))
                pos += len(text)
                replace_rpr = None
            elif kind == "del":
                portions = _run_portions(data, pieces, pos, pos + len(text))
                parts.append(revision_open("del", rev_id, author, date))
                rev_id += 1
                for rpr, portion in portions:
                    parts.append(emit_run(rpr, portion, "w:delText"))
                parts.append("</w:del>")
                pos += len(text)
                replace_rpr = portions[0][0] if portions else ""
            else:
                rpr = replace_rpr if replace_rpr is not None else _rpr_at_offset(data, pieces, pos)
                parts.append(revision_open("ins", rev_id, author, date))
                rev_id += 1
                parts.append(emit_run(rpr, text))
                parts.append("</w:ins>")
                replace_rpr = None
    inner = "".join(parts).encode("utf-8")
    if p.empty:
        open_tag = data[p.start : p.end - 2] + b">"  # reopen "<w:p …/>"
        return _xml.splice(data, [(p.start, p.end, open_tag + inner + b"</w:p>")])
    return _xml.splice(data, [(p.inner_start, p.inner_end, inner)])


def resolve_style_id(package: Package, style: str) -> str:
    """§6a: the styleId verbatim if defined, else with whitespace removed, else error."""
    ids: set[str] = set()
    if package.has_part("word/styles.xml"):
        try:
            root: ET.Element | None = ET.fromstring(package.part("word/styles.xml"))
        except ET.ParseError:
            root = None
        if root is not None:
            for st in root.iter(f"{{{_W_NS}}}style"):
                style_id = st.get(f"{{{_W_NS}}}styleId")
                if style_id:
                    ids.add(style_id)
    if style in ids:
        return style
    compact = "".join(ch for ch in style if ch not in _xml.WHITESPACE)
    if compact in ids:
        return compact
    raise ToolError(
        "style_unknown",
        f"Named style {style} does not exist.",
        ['Call docx_style {op: "list"} to see available styles.'],
    )


# ---------------------------------------------------------------------------
# §7 revisions
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class Revision:
    """One ``w:ins``/``w:del`` element of the document part."""

    span: _xml.Span
    kind: str  # "ins" | "del"
    rev_id: str  # "R{w:id}"
    author: str
    date: str
    anchor: str | None  # containing body block's anchor, when there is one
    ordinal: int | None  # containing body *paragraph* ordinal (merge post-pass)
    text: str  # the wrapper's own raw text (§6a)


def scan_revisions(data: bytes, blocks: list[AnchorEntry]) -> list[Revision]:
    """Every ``w:ins``/``w:del`` of the part, in document order (§6a)."""
    out: list[Revision] = []
    for el in _xml.iter_elements(data, names=REVISION_NAMES):
        attrs = start_tag_attrs(data, el)
        kind = "ins" if el.name == "w:ins" else "del"
        block = next((b for b in blocks if b.span.start < el.start and el.end <= b.span.end), None)
        tag = "w:t" if kind == "ins" else "w:delText"
        raw = (
            ""
            if el.empty
            else "".join(
                _xml.element_text(data, t)
                for t in _xml.iter_elements(data, el.inner_start, el.inner_end, names=(tag,))
            )
        )
        out.append(
            Revision(
                el,
                kind,
                f"R{attrs.get('w:id', '')}",
                attrs.get("w:author", "unknown"),
                attrs.get("w:date", ""),
                block.anchor if block is not None else None,
                block.ordinal if block is not None and block.kind == "paragraph" else None,
                raw,
            )
        )
    out.sort(key=lambda rev: rev.span.start)
    return out


def revision_matches(rev: Revision, flt: dict[str, str]) -> bool:
    """§7/§6a filters: author exact; date prefix; after ≤ w:date < before (strings)."""
    author = flt.get("author")
    if author is not None and rev.author != author:
        return False
    date = flt.get("date")
    if date is not None and not rev.date.startswith(date):
        return False
    after = flt.get("after")
    if after is not None and rev.date < after:
        return False
    before = flt.get("before")
    return before is None or rev.date < before


def t_to_deltext(content: bytes) -> bytes:
    return (
        content.replace(b"<w:t>", b"<w:delText>")
        .replace(b"<w:t ", b"<w:delText ")
        .replace(b"<w:t/>", b"<w:delText/>")
        .replace(b"</w:t>", b"</w:delText>")
    )


def deltext_to_t(content: bytes) -> bytes:
    return (
        content.replace(b"<w:delText>", b"<w:t>")
        .replace(b"<w:delText ", b"<w:t ")
        .replace(b"<w:delText/>", b"<w:t/>")
        .replace(b"</w:delText>", b"</w:t>")
    )


def resolve_revisions(data: bytes, candidates: list[Revision], accept: bool) -> bytes:
    """§7: accept ins / reject del → unwrap; accept del / reject ins → remove."""
    edits: list[tuple[int, int, bytes]] = []
    for rev in candidates:
        span = rev.span
        if (rev.kind == "ins") == accept:
            inner = b"" if span.empty else data[span.inner_start : span.inner_end]
            if rev.kind == "del":
                inner = deltext_to_t(inner)
            edits.append((span.start, span.end, inner))
        else:
            edits.append((span.start, span.end, b""))
    return _xml.splice(data, edits)


def _mergeable_parts(data: bytes, run: _xml.Span) -> tuple[bytes, str] | None:
    """``(raw rPr, concatenated text)`` iff the run is ``rPr? + w:t*`` only."""
    rpr = b""
    texts: list[str] = []
    if not run.empty:
        for child in _xml.iter_elements(data, run.inner_start, run.inner_end, max_depth=1):
            if child.name == "w:rPr" and not rpr and not texts:
                rpr = data[child.start : child.end]
            elif child.name == "w:t":
                texts.append(_xml.element_text(data, child))
            else:
                return None
    return rpr, "".join(texts)


def _rpr_key(rpr: bytes) -> bytes:
    """§7 post-pass comparison key: drop ``rsid*`` attributes; empty rPr ≡ absent."""
    stripped = _RSID_ATTR_RE.sub(b"", rpr)
    return b"" if stripped in (b"<w:rPr/>", b"<w:rPr></w:rPr>") else stripped


def merge_paragraph_runs(data: bytes, ordinal: int) -> bytes:
    """§7 post-pass: merge adjacent sibling runs with identical rPr (rsid-blind)."""
    while True:
        p = paragraph_span_at(data, ordinal)
        if p.empty:
            return data
        runs = [
            c
            for c in _xml.iter_elements(data, p.inner_start, p.inner_end, max_depth=1)
            if c.name == "w:r"
        ]
        merged: tuple[int, int, bytes] | None = None
        for first, second in pairwise(runs):
            if first.end != second.start:
                continue
            a = _mergeable_parts(data, first)
            b = _mergeable_parts(data, second)
            if a is None or b is None or _rpr_key(a[0]) != _rpr_key(b[0]):
                continue
            open_tag = (
                data[first.start : first.inner_start]
                if not first.empty
                else data[first.start : first.end - 2] + b">"
            )
            replacement = (
                open_tag + a[0] + _xml.emit_text_element(a[1] + b[1]).encode("utf-8") + b"</w:r>"
            )
            merged = (first.start, second.end, replacement)
            break
        if merged is None:
            return data
        data = _xml.splice(data, [merged])
