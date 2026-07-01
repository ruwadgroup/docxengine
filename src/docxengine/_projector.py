"""Projection: outline, windowed reads, and search over the agent view (algorithms.md §2/§2a).

One line per body-level block — ``[{anchor}{annotations}] {text}`` — with headings
resolved through the ``styles.xml`` ``basedOn`` cascade, list annotations resolved
through ``numbering.xml``, tables as markdown grids, and tracked changes shown
as-if-accepted with ``[ins by …]``/``[del by …]``/``[comment:C… by …]`` markers.

Package-level lookups (styles, numbering, comment authors) are parsed read-only
with ElementTree, like the rels/content-types readers in ``_opc``; the §3 no-DOM
rule applies to editing, which never happens here.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import cast
from xml.etree import ElementTree as ET

from . import _xml
from ._anchors import build_anchor_index, normalized_text, paragraph_anchor
from ._errors import ToolError
from ._opc import Package

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: §2a pagination: the content-character budget per page.
PAGE_CHAR_BUDGET = 24_000
#: §2a search: characters of raw context kept on each side of a match.
SNIPPET_RADIUS = 40

STORY_SCOPES = ("body", "footnotes", "comments", "headers", "footers")

_STYLES_PART = "word/styles.xml"
_NUMBERING_PART = "word/numbering.xml"
_COMMENTS_PART = "word/comments.xml"

_HEADING_ID = re.compile(r"^Heading([1-9])$")
_ANCHOR_RE = re.compile(r"^P([1-9][0-9]*)(?:#([0-9a-f]{4}))?$")
_RANGE_RE = re.compile(r"^P([1-9][0-9]*)(?:#[0-9a-f]{4})?\.\.P([1-9][0-9]*)(?:#[0-9a-f]{4})?$")
_ATTR_RE = re.compile(r'([^\s=/>"\']+)\s*=\s*(?:"([^"]*)"|\'([^\']*)\')')

# Story scope → part-name pattern (§2a: well-known names, headers/footers by number).
_STORY_PART_RE: dict[str, re.Pattern[str]] = {
    "footnotes": re.compile(r"^word/footnotes\.xml$"),
    "comments": re.compile(r"^word/comments\.xml$"),
    "headers": re.compile(r"^word/header(\d*)\.xml$"),
    "footers": re.compile(r"^word/footer(\d*)\.xml$"),
}


@dataclass(frozen=True, slots=True)
class ProjectedBlock:
    """One projection-ready block: its anchor, kind, rendered line(s), and metadata."""

    kind: str  # "paragraph" | "table"
    ordinal: int
    anchor: str
    raw: str  # §4 concatenated w:t text ("" for tables)
    text: str  # display text with markers ("" for tables)
    normalized: str  # §1 normalized text without markers ("" for tables)
    heading_level: int | None
    lines: str  # full projection line(s), tables span several


def _anchor_error(code: str, detail: str) -> ToolError:
    suggestions = {
        "anchor_invalid": ["Check the format 'P{index}#{hash}' (ranges: 'P10..P24')."],
        "anchor_not_found": ["Call docx_outline to re-map anchors."],
    }
    return ToolError(code, detail, suggestions.get(code, []))


# ---------------------------------------------------------------------------
# Package-level lookups (read-only)
# ---------------------------------------------------------------------------


def _start_tag_attrs(data: bytes, span: _xml.Span) -> dict[str, str]:
    """Attributes of an element's start tag, decoded, as written."""
    end = span.end if span.empty else span.inner_start
    tag = data[span.start : end].decode("utf-8")
    out: dict[str, str] = {}
    for m in _ATTR_RE.finditer(tag):
        out[m.group(1)] = _xml.unescape(m.group(2) if m.group(2) is not None else m.group(3))
    return out


def _parse_part(package: Package, name: str) -> ET.Element | None:
    if not package.has_part(name):
        return None
    try:
        return ET.fromstring(package.part(name))
    except ET.ParseError:
        return None


def _based_on_map(package: Package) -> dict[str, str]:
    """styleId → basedOn styleId from ``word/styles.xml`` (empty when absent)."""
    root = _parse_part(package, _STYLES_PART)
    if root is None:
        return {}
    out: dict[str, str] = {}
    for style in root.iter(f"{{{_W_NS}}}style"):
        style_id = style.get(f"{{{_W_NS}}}styleId")
        based_on = style.find(f"{{{_W_NS}}}basedOn")
        if style_id and based_on is not None:
            val = based_on.get(f"{{{_W_NS}}}val")
            if val:
                out[style_id] = val
    return out


def _numbering_maps(package: Package) -> tuple[dict[str, str], dict[tuple[str, str], str]]:
    """``numId → abstractNumId`` and ``(abstractNumId, ilvl) → numFmt``."""
    root = _parse_part(package, _NUMBERING_PART)
    if root is None:
        return {}, {}
    num_to_abstract: dict[str, str] = {}
    for num in root.iter(f"{{{_W_NS}}}num"):
        num_id = num.get(f"{{{_W_NS}}}numId")
        abstract = num.find(f"{{{_W_NS}}}abstractNumId")
        if num_id and abstract is not None:
            val = abstract.get(f"{{{_W_NS}}}val")
            if val:
                num_to_abstract[num_id] = val
    level_fmt: dict[tuple[str, str], str] = {}
    for abstract_num in root.iter(f"{{{_W_NS}}}abstractNum"):
        abstract_id = abstract_num.get(f"{{{_W_NS}}}abstractNumId")
        if not abstract_id:
            continue
        for lvl in abstract_num.iter(f"{{{_W_NS}}}lvl"):
            ilvl = lvl.get(f"{{{_W_NS}}}ilvl")
            fmt = lvl.find(f"{{{_W_NS}}}numFmt")
            if ilvl is not None and fmt is not None:
                level_fmt[(abstract_id, ilvl)] = fmt.get(f"{{{_W_NS}}}val", "")
    return num_to_abstract, level_fmt


def _comment_authors(package: Package) -> dict[str, str]:
    """comment id → author from ``word/comments.xml`` (empty when absent)."""
    root = _parse_part(package, _COMMENTS_PART)
    if root is None:
        return {}
    out: dict[str, str] = {}
    for comment in root.iter(f"{{{_W_NS}}}comment"):
        comment_id = comment.get(f"{{{_W_NS}}}id")
        if comment_id is not None:
            out[comment_id] = comment.get(f"{{{_W_NS}}}author", "unknown")
    return out


@dataclass(frozen=True, slots=True)
class _Context:
    """Per-package lookups shared by every projected paragraph."""

    based_on: dict[str, str]
    num_to_abstract: dict[str, str]
    level_fmt: dict[tuple[str, str], str]
    comment_authors: dict[str, str]


def _context(package: Package) -> _Context:
    num_to_abstract, level_fmt = _numbering_maps(package)
    return _Context(_based_on_map(package), num_to_abstract, level_fmt, _comment_authors(package))


# ---------------------------------------------------------------------------
# Paragraph projection
# ---------------------------------------------------------------------------


def heading_level(style_id: str | None, based_on: dict[str, str]) -> int | None:
    """§2: ``Heading1``…``Heading9`` as the styleId itself or via its basedOn chain."""
    seen: set[str] = set()
    current = style_id
    while current and current not in seen:
        m = _HEADING_ID.match(current)
        if m:
            return int(m.group(1))
        seen.add(current)
        current = based_on.get(current)
    return None


def _paragraph_props(data: bytes, p: _xml.Span) -> tuple[str | None, str | None, str]:
    """``(pStyle val, numId, ilvl)`` from the paragraph's direct ``w:pPr`` (§2a)."""
    style_id: str | None = None
    num_id: str | None = None
    ilvl = "0"
    if p.empty:
        return style_id, num_id, ilvl
    ppr = next(
        _xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:pPr",), max_depth=1), None
    )
    if ppr is None or ppr.empty:
        return style_id, num_id, ilvl
    props = _xml.iter_elements(data, ppr.inner_start, ppr.inner_end, names=("w:pStyle", "w:numPr"))
    for el in props:
        if el.name == "w:pStyle":
            style_id = _start_tag_attrs(data, el).get("w:val")
        elif not el.empty:
            for sub in _xml.iter_elements(
                data, el.inner_start, el.inner_end, names=("w:ilvl", "w:numId")
            ):
                val = _start_tag_attrs(data, sub).get("w:val")
                if val is None:
                    continue
                if sub.name == "w:ilvl":
                    ilvl = val
                else:
                    num_id = val
    return style_id, num_id, ilvl


def _marked_text(data: bytes, p: _xml.Span, ctx: _Context) -> tuple[str, str]:
    """``(raw text, display text)``: as-if-accepted, §2 markers spliced in.

    Walking the paragraph in document order, a marker is inserted at each
    ``w:ins``/``w:del`` wrapper end and at each ``w:commentReference``, with one
    space on each side; §1 normalization then absorbs doubled/edge spaces.
    """
    raw, pieces = _xml.paragraph_text(data, p)
    if p.empty:
        return raw, ""
    markers: list[tuple[int, int, str]] = []  # (text offset, byte position, marker)
    names = ("w:ins", "w:del", "w:commentReference")
    for el in _xml.iter_elements(data, p.inner_start, p.inner_end, names=names):
        attrs = _start_tag_attrs(data, el)
        if el.name == "w:commentReference":
            comment_id = attrs.get("w:id")
            if comment_id is None:
                continue
            author = ctx.comment_authors.get(comment_id, "unknown")
            marker = f"[comment:C{comment_id} by {author}]"
        else:
            kind = "ins" if el.name == "w:ins" else "del"
            marker = f"[{kind} by {attrs.get('w:author', 'unknown')}]"
        offset = sum(len(piece.text) for piece in pieces if piece.t.start < el.end)
        markers.append((offset, el.end, marker))
    out: list[str] = []
    pos = 0
    for offset, _, marker in sorted(markers, key=lambda m: (m[0], m[1])):
        out.append(raw[pos:offset])
        out.append(f" {marker} ")
        pos = offset
    out.append(raw[pos:])
    return raw, normalized_text("".join(out))


def _project_paragraph(
    data: bytes, anchor: str, ordinal: int, normalized: str, p: _xml.Span, ctx: _Context
) -> ProjectedBlock:
    style_id, num_id, ilvl = _paragraph_props(data, p)
    level = heading_level(style_id, ctx.based_on)
    annotations = ""
    if level is not None:
        annotations += f" H{level}"
    if num_id is not None and num_id != "0":
        abstract = ctx.num_to_abstract.get(num_id)
        fmt = ctx.level_fmt.get((abstract, ilvl), "") if abstract is not None else ""
        list_kind = "ul" if fmt == "bullet" else "ol"
        try:
            n = int(ilvl) + 1
        except ValueError:
            n = 1
        annotations += f" List:{list_kind} L{n}"
    raw, text = _marked_text(data, p, ctx)
    bracket = f"[{anchor}{annotations}]"
    line = f"{bracket} {text}" if text else bracket
    return ProjectedBlock("paragraph", ordinal, anchor, raw, text, normalized, level, line)


# ---------------------------------------------------------------------------
# Table projection
# ---------------------------------------------------------------------------


def _cell_text(data: bytes, tc: _xml.Span) -> str:
    """Cell paragraphs' normalized texts joined with a single space; ``|`` escaped."""
    if tc.empty:
        return ""
    texts = []
    for p in _xml.iter_elements(data, tc.inner_start, tc.inner_end, names=("w:p",)):
        raw, _ = _xml.paragraph_text(data, p)
        texts.append(normalized_text(raw))
    return " ".join(t for t in texts if t).replace("|", "\\|")


def _project_table(
    data: bytes, anchor: str, ordinal: int, tbl: _xml.Span, prev_anchor: str | None
) -> ProjectedBlock:
    rows: list[list[str]] = []
    n_grid_cols = 0
    if not tbl.empty:
        grid = next(
            _xml.iter_elements(
                data, tbl.inner_start, tbl.inner_end, names=("w:tblGrid",), max_depth=1
            ),
            None,
        )
        if grid is not None and not grid.empty:
            n_grid_cols = sum(
                1
                for _ in _xml.iter_elements(
                    data, grid.inner_start, grid.inner_end, names=("w:gridCol",), max_depth=1
                )
            )
        for tr in _xml.iter_elements(
            data, tbl.inner_start, tbl.inner_end, names=("w:tr",), max_depth=1
        ):
            if tr.empty:
                rows.append([])
                continue
            rows.append(
                [
                    _cell_text(data, tc)
                    for tc in _xml.iter_elements(
                        data, tr.inner_start, tr.inner_end, names=("w:tc",), max_depth=1
                    )
                ]
            )
    cols = n_grid_cols or max((len(r) for r in rows), default=0)
    where = f"@after:{prev_anchor}" if prev_anchor is not None else "@start"
    dims = f"{len(rows)}×{cols}"
    lines = [f"[{anchor} {dims} {where}]"]
    if rows and cols:
        lines.append("| " + " | ".join(rows[0]) + " |")
        lines.append("| " + " | ".join(["---"] * cols) + " |")
        for row in rows[1:]:
            lines.append("| " + " | ".join(row) + " |")
    return ProjectedBlock("table", ordinal, anchor, "", "", "", None, "\n".join(lines))


# ---------------------------------------------------------------------------
# Stories → block lists
# ---------------------------------------------------------------------------


def project_body(package: Package) -> list[ProjectedBlock]:
    """Every body-level block of the main document part, projection-ready."""
    part = package.main_document_part()
    data = package.part(part)
    ctx = _context(package)
    blocks: list[ProjectedBlock] = []
    prev_anchor: str | None = None
    for entry in build_anchor_index(package, part):
        if entry.kind == "paragraph":
            blocks.append(
                _project_paragraph(
                    data, entry.anchor, entry.ordinal, entry.normalized, entry.span, ctx
                )
            )
            prev_anchor = entry.anchor
        else:
            blocks.append(
                _project_table(data, entry.anchor, entry.ordinal, entry.span, prev_anchor)
            )
    return blocks


def _story_part_names(package: Package, scope: str) -> list[str]:
    """§2a story parts: well-known names; headers/footers ascending by number."""
    pattern = _STORY_PART_RE[scope]
    matched: list[tuple[int, str]] = []
    for name in package.part_names:
        m = pattern.match(name)
        if m:
            number = int(m.group(1) or 0) if m.groups() else 0
            matched.append((number, name))
    return [name for _, name in sorted(matched)]


def story_blocks(package: Package, scope: str) -> list[ProjectedBlock]:
    """Blocks for a story scope; non-body stories anchor per story (§2a)."""
    if scope == "body":
        return project_body(package)
    if scope not in _STORY_PART_RE:
        raise _anchor_error(
            "anchor_invalid", f"Unknown scope: {scope} (use {', '.join(STORY_SCOPES)})."
        )
    ctx = _context(package)
    blocks: list[ProjectedBlock] = []
    ordinal = 0
    for part_name in _story_part_names(package, scope):
        data = package.part(part_name)
        for p in _xml.iter_elements(data, names=("w:p",)):
            ordinal += 1
            raw, _ = _xml.paragraph_text(data, p)
            normalized = normalized_text(raw)
            anchor = paragraph_anchor(ordinal, normalized)
            blocks.append(_project_paragraph(data, anchor, ordinal, normalized, p, ctx))
    return blocks


# ---------------------------------------------------------------------------
# Read: windows, ranges, pagination (§2a)
# ---------------------------------------------------------------------------


def _parse_anchor(anchor: str) -> int:
    m = _ANCHOR_RE.match(anchor)
    if not m:
        raise _anchor_error("anchor_invalid", f"Malformed anchor string: {anchor}.")
    return int(m.group(1))


def _paragraph_index(blocks: list[ProjectedBlock], ordinal: int, label: str) -> int:
    for i, block in enumerate(blocks):
        if block.kind == "paragraph" and block.ordinal == ordinal:
            return i
    raise _anchor_error("anchor_not_found", f"Anchor {label} not found: index out of range.")


def _paginate(blocks: list[ProjectedBlock], char_budget: int) -> tuple[str, str | None]:
    """§2a: cut only before a paragraph (tables ride); never an empty first page."""
    pieces: list[str] = []
    total = 0
    cut: int | None = None
    for i, block in enumerate(blocks):
        added = len(block.lines) + (1 if pieces else 0)
        if block.kind == "paragraph" and pieces and total + added > char_budget:
            cut = i
            break
        pieces.append(block.lines)
        total += added
    content = "\n".join(pieces)
    if cut is None:
        return content, None
    last = next(b for b in reversed(blocks) if b.kind == "paragraph")
    return content, f"P{blocks[cut].ordinal}..P{last.ordinal}"


def project_read(
    package: Package,
    *,
    anchor: str | None = None,
    range: str | None = None,
    window: int = 0,
    scope: str = "body",
    char_budget: int = PAGE_CHAR_BUDGET,
) -> dict[str, str]:
    """The ``docx_read`` projection: anchor window, range, or whole story.

    The hash half of ``anchor`` is not validated (§2a) — a read with a stale anchor
    is exactly how a caller refreshes it. ``window`` counts body-level blocks on
    each side of the anchor block; ``anchor`` wins when both it and ``range`` are
    given.
    """
    blocks = story_blocks(package, scope)
    if anchor is not None:
        idx = _paragraph_index(blocks, _parse_anchor(anchor), anchor)
        window = max(0, window)
        selection = blocks[max(0, idx - window) : idx + window + 1]
    elif range is not None:
        m = _RANGE_RE.match(range)
        if not m:
            raise _anchor_error("anchor_invalid", f"Malformed range string: {range}.")
        start, end = int(m.group(1)), int(m.group(2))
        if start > end:
            raise _anchor_error("anchor_invalid", f"Inverted range: {range}.")
        start_idx = _paragraph_index(blocks, start, f"P{start}")
        end_idx = _paragraph_index(blocks, end, f"P{end}")
        selection = blocks[start_idx : end_idx + 1]
    else:
        selection = blocks
    content, continuation = _paginate(selection, char_budget)
    result = {"content": content}
    if continuation is not None:
        result["continuation"] = continuation
    return result


# ---------------------------------------------------------------------------
# Outline
# ---------------------------------------------------------------------------


def project_outline(package: Package) -> dict[str, list[dict[str, object]]]:
    """The ``docx_outline`` result: heading tree + table list with anchors."""
    outline: list[dict[str, object]] = []
    tables: list[dict[str, object]] = []
    prev_anchor: str | None = None
    for block in project_body(package):
        if block.kind == "paragraph":
            if block.heading_level is not None:
                outline.append(
                    {"anchor": block.anchor, "level": block.heading_level, "text": block.normalized}
                )
            prev_anchor = block.anchor
        else:
            dims = block.lines.split(" ", 2)[1]  # from the [T{n} {rows}×{cols} …] header
            entry: dict[str, object] = {"anchor": block.anchor, "dims": dims}
            if prev_anchor is not None:
                entry["after"] = prev_anchor
            tables.append(entry)
    return {"outline": outline, "tables": tables}


# ---------------------------------------------------------------------------
# Resource renderings (algorithms.md §25): text/markdown for MCP resources/read
# ---------------------------------------------------------------------------


def render_projection_markdown(package: Package) -> str:
    """The full §2 projection of the body as a markdown text block.

    One line per body paragraph and several per table — exactly the block
    ``lines`` the §2 projection emits — joined with newlines.
    """
    return "\n".join(block.lines for block in project_body(package))


def render_outline_markdown(package: Package) -> str:
    """The §2a ``docx_outline`` rendering as markdown (headings tree + tables)."""
    outline = project_outline(package)
    lines: list[str] = []
    for entry in outline["outline"]:
        level = cast("int", entry["level"])
        lines.append(f"{'#' * level} {entry['text']} [{entry['anchor']}]")
    for entry in outline["tables"]:
        after = entry.get("after")
        where = f" @after:{after}" if after is not None else ""
        lines.append(f"- table {entry['anchor']} {entry['dims']}{where}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Search
# ---------------------------------------------------------------------------


def _snippet(raw: str, start: int, end: int) -> str:
    lo = max(0, start - SNIPPET_RADIUS)
    hi = min(len(raw), end + SNIPPET_RADIUS)
    prefix = "…" if lo > 0 else ""
    suffix = "…" if hi < len(raw) else ""
    return f"{prefix}{normalized_text(raw[lo:hi])}{suffix}"


def _match_spans(text: str, query: str, pattern: re.Pattern[str] | None) -> list[tuple[int, int]]:
    if pattern is not None:
        return [m.span() for m in pattern.finditer(text) if m.start() != m.end()]
    spans: list[tuple[int, int]] = []
    pos = 0
    while True:
        hit = text.find(query, pos)
        if hit < 0:
            break
        spans.append((hit, hit + len(query)))
        pos = hit + len(query)
    return spans


def project_search(
    package: Package, query: str, *, regex: bool = False, scope: str = "body"
) -> dict[str, object]:
    """The ``docx_search`` result: §4 coalesced-text matches with snippet + context.

    Matching runs over each paragraph's raw concatenated ``w:t`` text (§2a); the
    nearest-heading context scan ignores any range scope filter.
    """
    if not query:
        raise ToolError(
            "not_found",
            "Text not found: the query is empty.",
            ["Provide a non-empty query."],
        )
    range_match = _RANGE_RE.match(scope)
    if range_match is not None:
        blocks = project_body(package)
        start, end = int(range_match.group(1)), int(range_match.group(2))
    elif scope in STORY_SCOPES:
        blocks = story_blocks(package, scope)
        start, end = 1, 0  # 0 = unbounded
    else:
        raise ToolError(
            "not_found",
            f"Unknown scope: {scope}.",
            [f"Use a story name ({', '.join(STORY_SCOPES)}) or a range like 'P10..P24'."],
        )
    pattern: re.Pattern[str] | None = None
    if regex:
        try:
            pattern = re.compile(query)
        except re.error as exc:
            raise ToolError(
                "not_found",
                f"Invalid regular expression: {exc}.",
                ["Fix the pattern or set regex: false for a literal search."],
            ) from exc
    matches: list[dict[str, str]] = []
    nearest_heading: str | None = None
    for block in blocks:
        if block.kind != "paragraph":
            continue
        if block.heading_level is not None:
            nearest_heading = block.normalized
        if block.ordinal < start or (end and block.ordinal > end):
            continue
        for span_start, span_end in _match_spans(block.raw, query, pattern):
            match: dict[str, str] = {
                "anchor": block.anchor,
                "snippet": _snippet(block.raw, span_start, span_end),
            }
            if nearest_heading is not None:
                match["context"] = nearest_heading
            matches.append(match)
    return {"matches": matches, "n_matches": len(matches)}
