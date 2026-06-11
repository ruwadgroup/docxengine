"""Templates (``docx_template_fill``) — algorithms.md §21/§23a.

A mustache subset — ``{{var}}``, ``{{#s}}…{{/s}}`` (loop / render-once on truthy),
``{{^s}}…{{/s}}`` (inverted), ``{{!c}}`` (comment, dropped) — is matched against the
§4 *coalesced* paragraph text so split-run placeholders resolve. A ``{{var}}`` is
written into the first overlapping ``w:t`` and the rest are trimmed (§4 first-overlap).
Loop/inverted sections whose open and close tags sit in whole paragraphs of one body
region clone the spanned paragraphs per array element (substituting ``{{.}}``/
``{{key}}``); when both tags sit in cells of exactly one table row, the row is cloned.
Missing vars stay verbatim and are listed in ``unfilled`` (dedup, document order);
``strict:true`` raises ``placeholder_unfilled``. Emission is XML-escaping only (§3).

The TypeScript twin (``template.ts``) is the byte-parity reference; the engine operates
purely on the document-part bytes so the splice output is deterministic across
languages. Offsets are byte offsets in Python and UTF-16 indices in JS, but the emitted
markup text is identical.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import _xml
from ._errors import ToolError
from ._session import Session

_TAG_RE = re.compile(r"\{\{([#^/!]?)\s*([^{}]*?)\s*\}\}")


# ---------------------------------------------------------------------------
# Tag tokens
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class TagToken:
    kind: str  # "var" | "section" | "inverted" | "close" | "comment"
    key: str
    start: int  # start of `{{` within the scanned string
    end: int  # one past `}}` within the scanned string


def _classify(sigil: str) -> str:
    return {"#": "section", "^": "inverted", "/": "close", "!": "comment"}.get(sigil, "var")


def scan_tags(text: str) -> list[TagToken]:
    """Scan a coalesced string for mustache tags, in order."""
    out: list[TagToken] = []
    for m in _TAG_RE.finditer(text):
        out.append(TagToken(_classify(m.group(1) or ""), m.group(2) or "", m.start(), m.end()))
    return out


# ---------------------------------------------------------------------------
# Value model
# ---------------------------------------------------------------------------

Scope = dict


def _lookup(scopes: list[object], key: str) -> object:
    if key == ".":
        top = scopes[-1] if scopes else None
        if isinstance(top, dict) and "." in top:
            return top["."]
        return top
    for s in reversed(scopes):
        if isinstance(s, dict) and key in s:
            return s[key]
    return _MISSING


_MISSING = object()


def _is_truthy(value: object) -> bool:
    if value is _MISSING or value is None or value is False or value == "":
        return False
    if isinstance(value, (list, tuple)):
        return len(value) > 0
    return True


def _stringify(value: object) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        # Shortest round-trip; integral floats render without a trailing ".0".
        return repr(value) if value != int(value) else str(int(value))
    if isinstance(value, str):
        return value
    return ""


def _as_scope(item: object) -> object:
    """Wrap a loop element so ``{{.}}`` resolves to a scalar element."""
    if isinstance(item, dict):
        return item
    return {".": item}


# ---------------------------------------------------------------------------
# Fill state
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class FillState:
    filled: int = 0
    loops_expanded: dict[str, int] = field(default_factory=dict)
    unfilled: list[str] = field(default_factory=list)
    unfilled_set: set[str] = field(default_factory=set)
    strict: bool = False


def _note_unfilled(state: FillState, key: str) -> None:
    if key not in state.unfilled_set:
        state.unfilled_set.add(key)
        state.unfilled.append(key)


# ---------------------------------------------------------------------------
# Element enumeration within an arbitrary markup buffer
# ---------------------------------------------------------------------------


def _paragraphs_in(data: bytes) -> list[_xml.Span]:
    """Top-level ``w:p`` spans in ``data`` (document order, non-nesting walk)."""
    out: list[_xml.Span] = []
    i = 0
    n = len(data)
    while i < n:
        nxt = data.find(b"<w:p", i)
        if nxt < 0:
            break
        # Confirm this is a `w:p` element start (not w:pPr, w:pStyle, etc).
        after = data[nxt + 4 : nxt + 5]
        if after in (b">", b" ", b"/"):
            p = next(_xml.iter_elements(data, nxt, names=("w:p",)), None)
            if p is not None and p.start == nxt:
                out.append(p)
                i = p.end
                continue
        i = nxt + 4
    return out


def _tables_in(data: bytes) -> list[_xml.Span]:
    out: list[_xml.Span] = []
    i = 0
    n = len(data)
    while i < n:
        nxt = data.find(b"<w:tbl", i)
        if nxt < 0:
            break
        after = data[nxt + 6 : nxt + 7]
        if after in (b">", b" ", b"/"):
            tbl = next(_xml.iter_elements(data, nxt, names=("w:tbl",)), None)
            if tbl is not None and tbl.start == nxt:
                out.append(tbl)
                i = tbl.end
                continue
        i = nxt + 6
    return out


def _coalesced_text(data: bytes, scope: _xml.Span) -> str:
    """Coalesced ``w:t`` text of a scope (paragraph or row)."""
    raw, _ = _xml.paragraph_text(data, scope)
    return raw


def _t_pieces(data: bytes, scope: _xml.Span) -> list[_xml.TextPiece]:
    _, pieces = _xml.paragraph_text(data, scope)
    return pieces


# ---------------------------------------------------------------------------
# Var substitution within one paragraph (§4 first-overlap)
# ---------------------------------------------------------------------------


def _emit_t(text: str) -> bytes:
    return _xml.emit_text_element(text).encode("utf-8")


def _fill_paragraph_vars(
    data: bytes, para: _xml.Span, scopes: list[object], state: FillState
) -> bytes:
    """Substitute ``{{var}}``/``{{!comment}}`` in one paragraph slice; sections left alone."""
    pieces = _t_pieces(data, para)
    if not pieces:
        return data
    text = "".join(p.text for p in pieces)
    tags = [t for t in scan_tags(text) if t.kind in ("var", "comment")]
    if not tags:
        return data

    resolved: list[tuple[int, int, str]] = []  # (start, end, value)
    for t in tags:
        if t.kind == "comment":
            resolved.append((t.start, t.end, ""))
            continue
        v = _lookup(scopes, t.key)
        if v is _MISSING:
            _note_unfilled(state, t.key)
            continue  # leave verbatim
        resolved.append((t.start, t.end, _stringify(v)))
        state.filled += 1
    if not resolved:
        return data
    resolved.sort(key=lambda r: r[0])

    # Compute the new text of each piece.
    new_texts: list[str] = []
    for piece in pieces:
        range_start = piece.start
        range_end = piece.start + len(piece.text)
        out = ""
        cursor = range_start
        while cursor < range_end:
            r = next((e for e in resolved if e[0] <= cursor < e[1]), None)
            if r is not None:
                if cursor == r[0]:
                    out += r[2]
                cursor = min(r[1], range_end)
                continue
            out += piece.text[cursor - range_start]
            cursor += 1
        new_texts.append(out)

    edits = [
        (pieces[i].t.start, pieces[i].t.end, _emit_t(new_texts[i])) for i in range(len(pieces))
    ]
    return _xml.splice(data, edits)


def _drop_section_tags(data: bytes, para: _xml.Span) -> bytes:
    """Erase residual section/close/comment tags from a paragraph's ``w:t`` pieces."""
    pieces = _t_pieces(data, para)
    if not pieces:
        return data
    text = "".join(p.text for p in pieces)
    tags = [t for t in scan_tags(text) if t.kind in ("section", "inverted", "close")]
    if not tags:
        return data
    new_texts: list[str] = []
    for piece in pieces:
        range_start = piece.start
        range_end = piece.start + len(piece.text)
        out = ""
        cursor = range_start
        while cursor < range_end:
            tag = next((t for t in tags if t.start <= cursor < t.end), None)
            if tag is not None:
                cursor = min(tag.end, range_end)
                continue
            out += piece.text[cursor - range_start]
            cursor += 1
        new_texts.append(out)
    edits = [
        (pieces[i].t.start, pieces[i].t.end, _emit_t(new_texts[i])) for i in range(len(pieces))
    ]
    return _xml.splice(data, edits)


# ---------------------------------------------------------------------------
# Section spans
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class SectionSpan:
    kind: str  # "paragraph" | "row"
    token: TagToken
    start: int
    end: int
    inner: bytes


def _find_para_at(data: bytes, start: int) -> _xml.Span | None:
    p = next(_xml.iter_elements(data, start, names=("w:p",)), None)
    if p is None or p.start != start:
        return None
    return p


def _first_paragraph_span(data: bytes) -> SectionSpan | None:
    paras = _paragraphs_in(data)
    open_idx = -1
    open_tok: TagToken | None = None
    for i, para in enumerate(paras):
        tags = scan_tags(_coalesced_text(data, para))
        open_t = next((t for t in tags if t.kind in ("section", "inverted")), None)
        if open_t is not None:
            open_idx = i
            open_tok = open_t
            break
    if open_idx < 0 or open_tok is None:
        return None

    depth = 0
    close_idx = -1
    for i in range(open_idx, len(paras)):
        for t in scan_tags(_coalesced_text(data, paras[i])):
            if t.kind in ("section", "inverted") and t.key == open_tok.key:
                if i == open_idx and t.start == open_tok.start:
                    continue
                depth += 1
            elif t.kind == "close" and t.key == open_tok.key:
                if depth == 0:
                    close_idx = i
                    break
                depth -= 1
        if close_idx >= 0:
            break
    if close_idx < 0:
        raise ToolError(
            "template_syntax",
            f"Unclosed section {{{{#{open_tok.key}}}}}.",
            ["Every {{#section}} needs a matching {{/section}}."],
        )
    open_para = paras[open_idx]
    close_para = paras[close_idx]
    has_inner = close_idx - open_idx >= 2
    inner = (
        data[paras[open_idx + 1].start : paras[close_idx - 1].end] if has_inner else b""
    )
    return SectionSpan("paragraph", open_tok, open_para.start, close_para.end, inner)


def _first_row_span(data: bytes) -> SectionSpan | None:
    for tbl in _tables_in(data):
        rows = [
            tr
            for tr in _xml.iter_elements(
                data, tbl.inner_start, tbl.inner_end, names=("w:tr",), max_depth=1
            )
        ]
        for row in rows:
            tags = scan_tags(_coalesced_text(data, row))
            open_t = next((t for t in tags if t.kind in ("section", "inverted")), None)
            close_t = next((t for t in tags if t.kind == "close"), None)
            if open_t is not None and close_t is not None and open_t.key == close_t.key:
                return SectionSpan(
                    "row", open_t, row.start, row.end, data[row.start : row.end]
                )
    return None


def _first_section_span(data: bytes) -> SectionSpan | None:
    row_span = _first_row_span(data)
    para_span = _first_paragraph_span(data)
    if row_span is not None and para_span is not None:
        return row_span if row_span.start <= para_span.start else para_span
    return row_span if row_span is not None else para_span


# ---------------------------------------------------------------------------
# Region expansion
# ---------------------------------------------------------------------------


def _render_plain(fragment: bytes, scopes: list[object], state: FillState) -> bytes:
    """Fill vars in a section-free markup slice (paragraphs only)."""
    work = fragment
    for p in reversed(_paragraphs_in(work)):
        para = _find_para_at(work, p.start) or p
        work = _fill_paragraph_vars(work, para, scopes, state)
    return work


def _render_row(row_xml: bytes, scopes: list[object], state: FillState) -> bytes:
    """Render one ``w:tr``: drop section tags + fill vars in its cell paragraphs."""
    work = row_xml
    for p in reversed(_paragraphs_in(work)):
        para = _find_para_at(work, p.start) or p
        work = _drop_section_tags(work, para)
        para2 = _find_para_at(work, p.start) or para
        work = _fill_paragraph_vars(work, para2, scopes, state)
    return work


def _expand_span(span: SectionSpan, scopes: list[object], state: FillState) -> bytes:
    """Expand one located section span into its rendered replacement markup."""
    render = _render_row if span.kind == "row" else _render_fragment
    value = _lookup(scopes, span.token.key)
    if span.token.kind == "inverted":
        if _is_truthy(value):
            return b""
        return render(span.inner, scopes, state)
    if isinstance(value, (list, tuple)):
        items = list(value)
        out = b""
        for item in items:
            out += render(span.inner, [*scopes, _as_scope(item)], state)
        state.loops_expanded[span.token.key] = len(items)
        return out
    if _is_truthy(value):
        state.loops_expanded[span.token.key] = 1
        return render(span.inner, [*scopes, _as_scope(value)], state)
    state.loops_expanded[span.token.key] = 0
    return b""


def _render_fragment(fragment: bytes, scopes: list[object], state: FillState) -> bytes:
    """Render a markup fragment under the scope chain (single-pass, §21)."""
    span = _first_section_span(fragment)
    if span is None:
        work = fragment
        for p in reversed(_paragraphs_in(work)):
            para = _find_para_at(work, p.start) or p
            work = _drop_section_tags(work, para)
            para2 = _find_para_at(work, p.start) or para
            work = _fill_paragraph_vars(work, para2, scopes, state)
        return work
    before = _render_plain(fragment[: span.start], scopes, state)
    expanded = _expand_span(span, scopes, state)
    rest = _render_fragment(fragment[span.end :], scopes, state)
    return before + expanded + rest


def _process_document(xml: bytes, data: object, state: FillState) -> bytes:
    body = _xml.find_body(xml)
    before = xml[: body.inner_start]
    after = xml[body.inner_end :]
    rendered = _render_fragment(xml[body.inner_start : body.inner_end], [data], state)
    return before + rendered + after


# ---------------------------------------------------------------------------
# docx_template_fill
# ---------------------------------------------------------------------------


def docx_template_fill(
    session: Session,
    *,
    template: str,
    data: dict[str, object],
    syntax: str = "mustache",
    strict: bool = False,
) -> dict[str, object]:
    """Fill a mustache template and register the filled doc as the next ``d{n}`` (§21)."""
    if syntax != "mustache":
        raise ToolError(
            "template_syntax",
            f"Unsupported template syntax: {syntax}.",
            ["Only the mustache subset is supported."],
        )
    doc = session.open_doc(template)
    payload: object = data if data is not None else {}
    state = FillState(strict=strict is True)

    main = doc.package.main_document_part()
    filled = _process_document(doc.package.part(main), payload, state)
    doc.package.set_part(main, filled)
    doc.mark_dirty()

    if state.strict and state.unfilled:
        raise ToolError(
            "placeholder_unfilled",
            f"Unfilled placeholders: {', '.join(state.unfilled)}.",
            ["Supply data for every placeholder or set strict: false."],
        )

    plural = "" if len(state.unfilled) == 1 else "s"
    note = (
        "All placeholders resolved."
        if not state.unfilled
        else f"{len(state.unfilled)} placeholder{plural} left unfilled."
    )
    return {
        "doc_id": doc.doc_id,
        "filled": state.filled,
        "loops_expanded": dict(state.loops_expanded),
        "unfilled": list(state.unfilled),
        "note": note,
    }
