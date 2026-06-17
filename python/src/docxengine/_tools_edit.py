"""Edit-surface tools: docx_replace, docx_edit_paragraph, docx_insert, docx_delete,
docx_revision (algorithms.md §4–§7, §6a).

Each function takes the :class:`~docxengine._session.Session` plus keyword arguments
named exactly as in ``spec/tools/<tool>.json`` and returns the result object in that
schema's shape. Every edit validates its anchor hash first (``anchor_stale``) and
returns fresh anchors for everything it touched. Failures raise
:class:`~docxengine._errors.ToolError` with a ``spec/errors.json`` code.
"""

from __future__ import annotations

from . import _create, _edits, _xml
from ._anchors import build_anchor_index
from ._errors import ToolError
from ._session import Session

_REVISION_OPS = frozenset({"list", "accept", "reject", "accept_all", "reject_all"})

#: The §22 horizontal-rule paragraph (mirrors ``_create._emit_rule``).
_RULE_PARAGRAPH = (
    "<w:p><w:pPr><w:pBdr>"
    '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>'
    "</w:pBdr></w:pPr></w:p>"
)


def _plural(n: int, noun: str) -> str:
    return f"{n} {noun}" if n == 1 else f"{n} {noun}s"


def _parse_insert_blocks(content: str) -> list[_create.LineBlock]:
    """§6a insert markdown: classify each non-blank line (shares the §22 grammar).

    Headings, quotes, rules, task lists, ``-``/``*``/``1.`` list items, else plain;
    multi-line tables are not recognized here. Blank/whitespace-only lines emit nothing.
    """
    blocks: list[_create.LineBlock] = []
    for raw_line in content.split("\n"):
        line = raw_line.removesuffix("\r")
        if not line.strip(_xml.WHITESPACE):
            continue
        blocks.append(_create.classify_line(line))
    return blocks


def _block_style(block: _create.LineBlock) -> str | None:
    """The pStyle a §6a insert block carries (``None`` = no pStyle; ``numPr`` is Phase 2)."""
    if block.kind == "heading":
        return f"Heading{block.level}"
    if block.kind == "quote":
        return "Quote"
    if block.kind in ("task", "ul", "ol"):
        return "ListParagraph"
    return None


# ---------------------------------------------------------------------------
# docx_replace (§4/§5, §6a)
# ---------------------------------------------------------------------------


def docx_replace(
    session: Session,
    *,
    doc_id: str,
    old: str,
    new: str,
    anchor: str | None = None,
    all: bool = False,  # noqa: A002 - wire name pinned by the tool schema
    track_changes: bool = False,
    author: str | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """Replace text in one anchored paragraph or across the whole body."""
    doc = session.get(doc_id)
    package = doc.package
    part = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    target = _edits.require_paragraph(entries, anchor) if anchor is not None else None
    if not old:
        raise ToolError(
            "not_found",
            "Text not found: the search text is empty.",
            ["Provide non-empty old text."],
        )
    data = package.part(part)
    ordinals = [target.ordinal] if target is not None else [e.ordinal for e in entries]
    total = 0
    for ordinal in ordinals:
        raw, _ = _xml.paragraph_text(data, entries[ordinal - 1].span)
        total += _edits.count_occurrences(raw, old)
    if not all:
        if total == 0:
            raise ToolError(
                "not_found",
                f"Text not found: {old}.",
                ["Broaden the query; check the projection for the exact text."],
            )
        if total > 1:
            raise ToolError(
                "ambiguous_target",
                f"{old} matches {total} times without all: true.",
                ["Add all: true or narrow with an anchor."],
            )
    author_name = _edits.resolve_author(author)
    date = _edits.revision_date()
    n_replaced = 0
    affected: list[int] = []
    for ordinal in ordinals:
        pos = 0
        while True:
            data = package.part(part)
            span = _edits.paragraph_span_at(data, ordinal)
            raw, pieces = _xml.paragraph_text(data, span)
            start = raw.find(old, pos)
            if start < 0:
                break
            end = start + len(old)
            if track_changes:
                new_data = _edits.splice_replace_tracked(
                    data, pieces, start, end, new, author_name, date
                )
            else:
                new_data = _edits.splice_replace_plain(data, pieces, start, end, new)
            package.set_part(part, new_data)
            n_replaced += 1
            pos = start + len(new)
            if ordinal not in affected:
                affected.append(ordinal)
            if not all:
                break
        if not all and n_replaced:
            break
    if n_replaced:
        doc.mark_dirty()
    fresh = _edits.paragraph_entries(package)
    if all:  # §6a: all → anchors of affected paragraphs ascending; otherwise new_anchor
        return {
            "n_replaced": n_replaced,
            "anchors": [fresh[ordinal - 1].anchor for ordinal in sorted(affected)],
        }
    ordinal = target.ordinal if target is not None else affected[0]
    return {"n_replaced": n_replaced, "new_anchor": fresh[ordinal - 1].anchor}


# ---------------------------------------------------------------------------
# docx_edit_paragraph (§6, §6a)
# ---------------------------------------------------------------------------


def docx_edit_paragraph(
    session: Session,
    *,
    doc_id: str,
    anchor: str,
    text: str,
    track_changes: bool = False,
    author: str | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """Rewrite a paragraph's full text; tracking emits the §6 minimal word-level redline."""
    doc = session.get(doc_id)
    package = doc.package
    part = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    entry = _edits.require_paragraph(entries, anchor)
    data = package.part(part)
    raw, _ = _xml.paragraph_text(data, entry.span)
    ops = _edits.word_diff(_edits.diff_units(raw), _edits.diff_units(text))
    blocks = _edits.diff_blocks(ops)
    new_data = _edits.rebuild_paragraph(
        data,
        entry.span,
        blocks,
        tracked=track_changes,
        author=_edits.resolve_author(author),
        date=_edits.revision_date(),
    )
    package.set_part(part, new_data)
    doc.mark_dirty()
    fresh = _edits.paragraph_entries(package)
    changed = max(
        sum(1 for op, _ in ops if op == "del"),
        sum(1 for op, _ in ops if op == "ins"),
    )
    return {
        "new_anchor": fresh[entry.ordinal - 1].anchor,
        "diff": f"~{_plural(changed, 'word')} changed",
    }


# ---------------------------------------------------------------------------
# docx_insert (§6a)
# ---------------------------------------------------------------------------


def docx_insert(
    session: Session,
    *,
    doc_id: str,
    content: str,
    after: str | None = None,
    before: str | None = None,
    style: str | None = None,
    track_changes: bool = False,
    author: str | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """Insert new paragraphs (plain text or minimal Markdown) after/before an anchor."""
    doc = session.get(doc_id)
    package = doc.package
    part = package.main_document_part()
    if (after is None) == (before is None):
        raise _edits.anchor_invalid_error("Provide exactly one of after or before.")
    entries = _edits.paragraph_entries(package)
    entry = _edits.require_paragraph(entries, after if after is not None else before or "")
    blocks = _parse_insert_blocks(content)
    style_id = _edits.resolve_style_id(package, style) if style is not None else None
    if not blocks:
        return {"new_anchors": []}
    data = package.part(part)
    author_name = _edits.resolve_author(author)
    date = _edits.revision_date()
    rev_id = _edits.next_revision_id(data) if track_changes else 0
    pieces: list[str] = []
    for block in blocks:
        # A horizontal rule has no runs; a style override makes it a styled empty paragraph.
        if block.kind == "rule" and style_id is None:
            pieces.append(_RULE_PARAGRAPH)
            continue
        effective_style = style_id if style_id is not None else _block_style(block)
        ppr = (
            f'<w:pPr><w:pStyle w:val="{_xml.escape_attr(effective_style)}"/></w:pPr>'
            if effective_style is not None
            else ""
        )
        text = "" if block.kind == "rule" else block.text
        runs = _create.emit_inline(text) if text else ""
        if track_changes and runs:
            runs = _edits.revision_open("ins", rev_id, author_name, date) + runs + "</w:ins>"
            rev_id += 1
        pieces.append(f"<w:p>{ppr}{runs}</w:p>")
    position = entry.span.end if after is not None else entry.span.start
    new_data = _xml.splice(data, [(position, position, "".join(pieces).encode("utf-8"))])
    package.set_part(part, new_data)
    doc.mark_dirty()
    fresh = _edits.paragraph_entries(package)
    base = entry.ordinal + 1 if after is not None else entry.ordinal
    return {"new_anchors": [fresh[base - 1 + i].anchor for i in range(len(blocks))]}


# ---------------------------------------------------------------------------
# docx_delete (§6a)
# ---------------------------------------------------------------------------


def docx_delete(
    session: Session,
    *,
    doc_id: str,
    anchor: str | None = None,
    range: str | None = None,  # noqa: A002 - wire name pinned by the tool schema
    track_changes: bool = False,
    author: str | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """Delete one paragraph or a contiguous range; tracked deletion is a redline."""
    doc = session.get(doc_id)
    package = doc.package
    part = package.main_document_part()
    if (anchor is None) == (range is None):
        raise _edits.anchor_invalid_error("Provide exactly one of anchor or range.")
    entries = _edits.paragraph_entries(package)
    if anchor is not None:
        targets = [_edits.require_paragraph(entries, anchor)]
    else:
        assert range is not None
        m = _edits.RANGE_RE.match(range)
        if not m:
            raise _edits.anchor_invalid_error(f"Malformed range string: {range}.")
        start, end = int(m.group(1)), int(m.group(3))
        if start > end:
            raise _edits.anchor_invalid_error(f"Inverted range: {range}.")
        for ordinal, hash_part in ((start, m.group(2)), (end, m.group(4))):
            entry = _edits.entry_at(entries, ordinal, f"P{ordinal}")
            if hash_part is not None and entry.anchor != f"P{ordinal}#{hash_part}":
                raise _edits.anchor_stale_error(f"P{ordinal}#{hash_part}")
        targets = entries[start - 1 : end]
    data = package.part(part)
    if not track_changes:
        new_data = _xml.splice(data, [(t.span.start, t.span.end, b"") for t in targets])
    else:
        author_name = _edits.resolve_author(author)
        date = _edits.revision_date()
        rev_id = _edits.next_revision_id(data)
        edits: list[tuple[int, int, bytes]] = []
        for target in targets:
            p = target.span
            if p.empty:
                continue
            ppr = next(
                _xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:pPr",), max_depth=1),
                None,
            )
            content_start = ppr.end if ppr is not None else p.inner_start
            if content_start >= p.inner_end:
                continue
            wrapped = (
                _edits.revision_open("del", rev_id, author_name, date).encode("utf-8")
                + _edits.t_to_deltext(data[content_start : p.inner_end])
                + b"</w:del>"
            )
            edits.append((content_start, p.inner_end, wrapped))
            rev_id += 1
        new_data = _xml.splice(data, edits)
    package.set_part(part, new_data)
    doc.mark_dirty()
    return {"ok": True, "deleted": len(targets)}


# ---------------------------------------------------------------------------
# docx_revision (§7, §6a)
# ---------------------------------------------------------------------------


def docx_revision(
    session: Session,
    *,
    doc_id: str,
    op: str,
    id: str | None = None,  # noqa: A002 - wire name pinned by the tool schema
    filter: dict[str, str] | None = None,  # noqa: A002 - wire name pinned by the tool schema
    response_format: str = "concise",
) -> dict[str, object]:
    """List, accept, or reject tracked changes, individually, in bulk, or filtered."""
    doc = session.get(doc_id)
    package = doc.package
    part = package.main_document_part()
    if op not in _REVISION_OPS:
        raise ToolError(
            "not_found",
            f"Unknown revision op: {op}.",
            ["Use list, accept, reject, accept_all, or reject_all."],
        )
    blocks = build_anchor_index(package)
    data = package.part(part)
    revisions = _edits.scan_revisions(data, blocks)
    flt = dict(filter or {})
    if op == "list":
        listed: list[dict[str, str]] = []
        for rev in revisions:
            if not _edits.revision_matches(rev, flt):
                continue
            item = {"id": rev.rev_id, "type": rev.kind, "author": rev.author, "date": rev.date}
            if rev.anchor is not None:
                item["anchor"] = rev.anchor
            item["text"] = rev.text
            listed.append(item)
        return {"revisions": listed}
    accept = op in ("accept", "accept_all")
    if op in ("accept_all", "reject_all"):  # §6a: _all ops ignore id/filter
        candidates = list(revisions)
    elif id is not None:
        # An id selecting nothing resolves nothing (§7 idempotency), not an error.
        candidates = [rev for rev in revisions if rev.rev_id == id]
    else:
        candidates = [rev for rev in revisions if _edits.revision_matches(rev, flt)]
    candidates = [  # a candidate nested in another candidate resolves with its container
        rev
        for rev in candidates
        if not any(
            other.span.start < rev.span.start and rev.span.end < other.span.end
            for other in candidates
        )
    ]
    ordinals = sorted({rev.ordinal for rev in candidates if rev.ordinal is not None})
    if candidates:
        new_data = _edits.resolve_revisions(data, candidates, accept)
        for ordinal in ordinals:
            new_data = _edits.merge_paragraph_runs(new_data, ordinal)
        package.set_part(part, new_data)
        doc.mark_dirty()
    fresh_blocks = build_anchor_index(package)
    fresh = [b for b in fresh_blocks if b.kind == "paragraph"]
    remaining = _edits.scan_revisions(package.part(part), fresh_blocks)
    by_author: dict[str, int] = {}
    for rev in remaining:
        by_author[rev.author] = by_author.get(rev.author, 0) + 1
    return {
        "accepted" if accept else "rejected": len(candidates),
        "remaining_by_author": dict(sorted(by_author.items())),
        "anchors": [fresh[ordinal - 1].anchor for ordinal in ordinals],
    }
