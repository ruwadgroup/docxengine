"""Convert (``docx_convert``) — algorithms.md §23/§23a.

``md``/``html`` are produced in-engine from the §2 projection model: headings
``#``×level, ordered/unordered list items indented per ilvl, GitHub tables,
``**bold**``/``*italic*`` reconstructed from run ``w:b``/``w:i``, comments inline as
``<!-- comment:{author}: {text} -->``, revisions in accepted view (ins shown, del
omitted) with ``[ins]``/``[del]`` markers. ``md`` keeps ``&``,``<``,``>`` literal;
``html`` HTML-escapes and adds inline styles for alignment and color. ``pdf``/``png``
go through the §24 render adapter.

Cross-language parity is conformance-tested on ``md``/``html`` content — this module
mirrors ``convert.ts`` to the byte. ``pdf``/``png`` parity is not required.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import _projector, _render, _xml
from ._anchors import build_anchor_index, normalized_text
from ._errors import ToolError
from ._opc import Package
from ._session import OpenDocument, Session

# ---------------------------------------------------------------------------
# Run-level model for one paragraph (accepted view)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ConvRun:
    text: str
    bold: bool
    italic: bool
    ins: bool  # a tracked-insertion run (for the `[ins]` marker)


@dataclass(slots=True)
class ParaModel:
    heading: int | None
    list_info: tuple[str, int] | None  # (kind "ol"|"ul", level)
    alignment: str | None
    color: str | None
    runs: list[ConvRun] = field(default_factory=list)
    comments: list[tuple[str, str]] = field(default_factory=list)  # (author, text)
    has_deletion: bool = False


@dataclass(frozen=True, slots=True)
class _ConvCtx:
    based_on: dict[str, str]
    num_to_abstract: dict[str, str]
    level_fmt: dict[tuple[str, str], str]
    comment_authors: dict[str, str]
    comment_texts: dict[str, str]


def _build_context(package: Package) -> _ConvCtx:
    num_to_abstract, level_fmt = _projector._numbering_maps(package)
    return _ConvCtx(
        _projector._based_on_map(package),
        num_to_abstract,
        level_fmt,
        _projector._comment_authors(package),
        _comment_texts(package),
    )


def _comment_texts(package: Package) -> dict[str, str]:
    """comment id → §1 normalized body text from ``word/comments.xml``."""
    if not package.has_part(_projector._COMMENTS_PART):
        return {}
    data = package.part(_projector._COMMENTS_PART)
    out: dict[str, str] = {}
    for comment in _xml.iter_elements(data, names=("w:comment",)):
        cid = _projector._start_tag_attrs(data, comment).get("w:id")
        if cid is None:
            continue
        raw = "".join(
            _xml.element_text(data, t)
            for t in _xml.iter_elements(
                data, comment.inner_start, comment.inner_end, names=("w:t",)
            )
        )
        out[cid] = normalized_text(raw)
    return out


def _first(data: bytes, parent: _xml.Span, name: str) -> _xml.Span | None:
    if parent.empty:
        return None
    return next(
        _xml.iter_elements(data, parent.inner_start, parent.inner_end, names=(name,), max_depth=1),
        None,
    )


def _paragraph_style_id(data: bytes, p: _xml.Span) -> str | None:
    ppr = _first(data, p, "w:pPr")
    if ppr is None or ppr.empty:
        return None
    pstyle = _first(data, ppr, "w:pStyle")
    if pstyle is None:
        return None
    return _projector._start_tag_attrs(data, pstyle).get("w:val")


def _run_text(data: bytes, run: _xml.Span) -> str:
    """Concatenated ``w:t`` text of a run (decoded; ``delText`` excluded)."""
    if run.empty:
        return ""
    return "".join(
        _xml.element_text(data, t)
        for t in _xml.iter_elements(data, run.inner_start, run.inner_end, names=("w:t",))
    )


def _run_formatting(data: bytes, run: _xml.Span) -> tuple[bool, bool]:
    """bold/italic from a run's ``w:rPr`` (toggle present and not ``w:val="0"``)."""
    rpr = _first(data, run, "w:rPr")
    if rpr is None or rpr.empty or rpr.start != run.inner_start:
        return False, False
    kids = list(_xml.iter_elements(data, rpr.inner_start, rpr.inner_end, max_depth=1))

    def on(name: str) -> bool:
        el = next((k for k in kids if k.name == name), None)
        if el is None:
            return False
        val = _projector._start_tag_attrs(data, el).get("w:val")
        return val not in ("0", "false")

    return on("w:b"), on("w:i")


def _paragraph_formatting(data: bytes, p: _xml.Span) -> tuple[str | None, str | None]:
    """Alignment (``jc``) and color from a paragraph's ``w:pPr`` / first run ``w:rPr``."""
    alignment: str | None = None
    color: str | None = None
    ppr = _first(data, p, "w:pPr")
    if ppr is not None and not ppr.empty:
        jc = _first(data, ppr, "w:jc")
        if jc is not None:
            alignment = _projector._start_tag_attrs(data, jc).get("w:val")
    first_run = next(
        _xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:r",)), None
    ) if not p.empty else None
    if first_run is not None:
        rpr = _first(data, first_run, "w:rPr")
        if rpr is not None and not rpr.empty:
            c = _first(data, rpr, "w:color")
            if c is not None:
                v = _projector._start_tag_attrs(data, c).get("w:val")
                if v and v != "auto":
                    color = v
    return alignment, color


def _list_annotation(data: bytes, p: _xml.Span, ctx: _ConvCtx) -> tuple[str, int] | None:
    """List annotation (ol/ul + level) from ``numPr``, mirroring §2."""
    ppr = _first(data, p, "w:pPr")
    if ppr is None or ppr.empty:
        return None
    numpr = _first(data, ppr, "w:numPr")
    if numpr is None or numpr.empty:
        return None
    num_id_el = _first(data, numpr, "w:numId")
    num_id = (
        _projector._start_tag_attrs(data, num_id_el).get("w:val")
        if num_id_el is not None
        else None
    )
    if num_id is None or num_id == "0":
        return None
    level = 0
    ilvl_el = _first(data, numpr, "w:ilvl")
    if ilvl_el is not None:
        raw = _projector._start_tag_attrs(data, ilvl_el).get("w:val")
        if raw is not None and raw.isdigit():
            level = int(raw)
    abstract = ctx.num_to_abstract.get(num_id)
    fmt = ctx.level_fmt.get((abstract, str(level)), "") if abstract is not None else ""
    return ("ul" if fmt == "bullet" else "ol", level)


def _parse_paragraph(data: bytes, p: _xml.Span, ctx: _ConvCtx) -> ParaModel:
    """Parse one body paragraph into the §23 conversion model (accepted view)."""
    style_id = _paragraph_style_id(data, p)
    heading = _projector.heading_level(style_id, ctx.based_on)
    lst = _list_annotation(data, p, ctx)
    alignment, color = _paragraph_formatting(data, p)
    runs: list[ConvRun] = []
    comments: list[tuple[str, str]] = []
    has_deletion = False
    if p.empty:
        return ParaModel(heading, lst, alignment, color, runs, comments, has_deletion)

    # Walk the paragraph in document order, tracking ins/del depth.
    names = ("w:ins", "w:del", "w:r", "w:commentReference")
    # Build depth maps from wrapper spans (ins/del cannot interleave runs by name only,
    # so we resolve enclosure by span containment).
    ins_spans = list(_xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:ins",)))
    del_spans = list(_xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:del",)))
    has_deletion = len(del_spans) > 0
    ref_seen: set[int] = set()
    for el in _xml.iter_elements(data, p.inner_start, p.inner_end, names=names):
        if el.name == "w:commentReference":
            cid = _projector._start_tag_attrs(data, el).get("w:id", "?")
            author = ctx.comment_authors.get(cid, "unknown")
            comments.append((author, ctx.comment_texts.get(cid, "")))
            ref_seen.add(el.start)
            continue
        if el.name != "w:r":
            continue
        in_del = any(d.start < el.start and el.end <= d.end for d in del_spans)
        in_ins = any(i.start < el.start and el.end <= i.end for i in ins_spans)
        if not in_del:
            text = _run_text(data, el)
            if text != "":
                bold, italic = _run_formatting(data, el)
                runs.append(ConvRun(text, bold, italic, in_ins))
        # A comment reference often sits inside its own run: capture it at the range end.
        ref = _first(data, el, "w:commentReference")
        if ref is not None and ref.start not in ref_seen:
            cid = _projector._start_tag_attrs(data, ref).get("w:id", "?")
            comments.append(
                (ctx.comment_authors.get(cid, "unknown"), ctx.comment_texts.get(cid, ""))
            )
            ref_seen.add(ref.start)
    return ParaModel(heading, lst, alignment, color, runs, comments, has_deletion)


# ---------------------------------------------------------------------------
# Block iteration
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class _Block:
    kind: str  # "p" | "tbl"
    span: _xml.Span


def _body_blocks(package: Package) -> tuple[bytes, list[_Block]]:
    data = package.part(package.main_document_part())
    blocks = [
        _Block("p" if e.kind == "paragraph" else "tbl", e.span)
        for e in build_anchor_index(package)
    ]
    return data, blocks


# ---------------------------------------------------------------------------
# Markdown rendering (§23)
# ---------------------------------------------------------------------------


def _runs_to_markdown(runs: list[ConvRun], markers: bool) -> str:
    out = ""
    for run in runs:
        text = run.text
        if run.bold:
            text = f"**{text}**"
        if run.italic:
            text = f"*{text}*"
        if run.ins and markers:
            text = f"[ins]{text}"
        out += text
    return out


def _paragraph_to_markdown(model: ParaModel, markers: bool) -> str:
    text = _runs_to_markdown(model.runs, markers)
    if model.has_deletion and markers:
        text += "[del]"
    for author, ctext in model.comments:
        text += f" <!-- comment:{author}: {ctext} -->"
    if model.heading is not None:
        return f'{"#" * model.heading} {text}'
    if model.list_info is not None:
        kind, level = model.list_info
        indent = "  " * level
        bullet = "1. " if kind == "ol" else "- "
        return f"{indent}{bullet}{text}"
    return text


def _md_cell_text(data: bytes, tc: _xml.Span) -> str:
    if tc.empty:
        return ""
    texts = []
    for p in _xml.iter_elements(data, tc.inner_start, tc.inner_end, names=("w:p",)):
        raw, _ = _xml.paragraph_text(data, p)
        texts.append(normalized_text(raw))
    return " ".join(texts).replace("|", "\\|")


def _table_to_markdown(data: bytes, tbl: _xml.Span) -> str:
    cols = _table_cols(data, tbl)
    lines: list[str] = []
    rows = (
        []
        if tbl.empty
        else [
            tr
            for tr in _xml.iter_elements(
                data, tbl.inner_start, tbl.inner_end, names=("w:tr",), max_depth=1
            )
        ]
    )
    for idx, tr in enumerate(rows):
        cells = (
            []
            if tr.empty
            else list(
                _xml.iter_elements(
                    data, tr.inner_start, tr.inner_end, names=("w:tc",), max_depth=1
                )
            )
        )
        cell_text = [_md_cell_text(data, tc) for tc in cells]
        lines.append("| " + " | ".join(cell_text) + " |")
        if idx == 0:
            lines.append("|" + " --- |" * max(cols, 1))
    return "\n".join(lines)


def _table_cols(data: bytes, tbl: _xml.Span) -> int:
    if tbl.empty:
        return 0
    grid = _first(data, tbl, "w:tblGrid")
    n_grid = 0
    if grid is not None and not grid.empty:
        n_grid = sum(
            1
            for _ in _xml.iter_elements(
                data, grid.inner_start, grid.inner_end, names=("w:gridCol",), max_depth=1
            )
        )
    if n_grid:
        return n_grid
    max_tc = 0
    rows = _xml.iter_elements(
        data, tbl.inner_start, tbl.inner_end, names=("w:tr",), max_depth=1
    )
    for tr in rows:
        if tr.empty:
            continue
        count = sum(
            1
            for _ in _xml.iter_elements(
                data, tr.inner_start, tr.inner_end, names=("w:tc",), max_depth=1
            )
        )
        max_tc = max(max_tc, count)
    return max_tc


def _to_markdown(package: Package, ctx: _ConvCtx, markers: bool) -> str:
    data, blocks = _body_blocks(package)
    out = ""
    prev_was_list_item = False
    for b in blocks:
        if b.kind == "tbl":
            line = _table_to_markdown(data, b.span)
            is_list_item = False
        else:
            model = _parse_paragraph(data, b.span, ctx)
            is_list_item = model.list_info is not None
            line = _paragraph_to_markdown(model, markers)
        if out == "":
            out = line
        else:
            out += ("\n" if prev_was_list_item and is_list_item else "\n\n") + line
        prev_was_list_item = is_list_item
    return out


# ---------------------------------------------------------------------------
# HTML rendering (§23)
# ---------------------------------------------------------------------------


def _html_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _runs_to_html(runs: list[ConvRun]) -> str:
    out = ""
    for run in runs:
        text = _html_escape(run.text)
        if run.bold:
            text = f"<strong>{text}</strong>"
        if run.italic:
            text = f"<em>{text}</em>"
        out += text
    return out


def _style_attr(model: ParaModel) -> str:
    styles: list[str] = []
    if model.alignment is not None:
        if model.alignment == "both":
            css: str | None = "justify"
        elif model.alignment in ("center", "right", "left"):
            css = model.alignment
        else:
            css = None
        if css:
            styles.append(f"text-align:{css}")
    if model.color is not None:
        styles.append(f"color:#{model.color}")
    return f' style="{";".join(styles)}"' if styles else ""


def _paragraph_to_html(model: ParaModel) -> str:
    inner = _runs_to_html(model.runs)
    sa = _style_attr(model)
    if model.heading is not None:
        h = min(6, model.heading)
        return f"<h{h}{sa}>{inner}</h{h}>"
    return f"<p{sa}>{inner}</p>"


def _table_to_html(data: bytes, tbl: _xml.Span) -> str:
    rows = (
        []
        if tbl.empty
        else list(
            _xml.iter_elements(data, tbl.inner_start, tbl.inner_end, names=("w:tr",), max_depth=1)
        )
    )
    trs: list[str] = []
    for tr in rows:
        cells = (
            []
            if tr.empty
            else list(
                _xml.iter_elements(
                    data, tr.inner_start, tr.inner_end, names=("w:tc",), max_depth=1
                )
            )
        )
        tds = [
            f"<td>{_html_escape(_md_cell_text(data, tc).replace(chr(92) + '|', '|'))}</td>"
            for tc in cells
        ]
        trs.append(f'<tr>{"".join(tds)}</tr>')
    return f'<table>{"".join(trs)}</table>'


def _to_html(package: Package, ctx: _ConvCtx) -> str:
    data, blocks = _body_blocks(package)
    out: list[str] = []
    list_open: str | None = None

    def close_list() -> None:
        nonlocal list_open
        if list_open is not None:
            out.append(f"</{list_open}>")
            list_open = None

    for b in blocks:
        if b.kind == "tbl":
            close_list()
            out.append(_table_to_html(data, b.span))
            continue
        model = _parse_paragraph(data, b.span, ctx)
        if model.list_info is not None:
            kind = model.list_info[0]
            if list_open != kind:
                close_list()
                out.append(f"<{kind}>")
                list_open = kind
            out.append(f"<li>{_runs_to_html(model.runs)}</li>")
            continue
        close_list()
        out.append(_paragraph_to_html(model))
    close_list()
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Convert note (free-form; masked in parity)
# ---------------------------------------------------------------------------


def _count_occurrences(s: str, sub: str) -> int:
    n = 0
    i = s.find(sub)
    while i >= 0:
        n += 1
        i = s.find(sub, i + len(sub))
    return n


def _convert_note(package: Package) -> str:
    xml = package.part(package.main_document_part()).decode("utf-8")
    comments = _count_occurrences(xml, "<w:commentReference")
    revisions = _count_occurrences(xml, "<w:ins ") + _count_occurrences(xml, "<w:del ")
    if comments == 0 and revisions == 0:
        return "Converted (no comments or tracked changes)."
    parts: list[str] = []
    if comments > 0:
        parts.append(f"{comments} comment{'' if comments == 1 else 's'}")
    if revisions > 0:
        parts.append(f"{revisions} tracked change{'' if revisions == 1 else 's'}")
    return f"{' and '.join(parts)} annotated inline"


# ---------------------------------------------------------------------------
# docx_convert
# ---------------------------------------------------------------------------


def docx_convert(
    session: Session,
    *,
    doc_id: str,
    to: str,
    path: str | None = None,
) -> dict[str, object]:
    """Convert an open document to md/html (in-engine) or pdf/png (render adapter)."""
    doc = session.get(doc_id)
    if to not in ("md", "html", "pdf", "png"):
        raise ToolError(
            "unsupported_format",
            f"Unsupported conversion target: {to}.",
            ["Use to: md, html, pdf or png."],
        )
    if to in ("md", "html"):
        ctx = _build_context(doc.package)
        content = _to_markdown(doc.package, ctx, True) if to == "md" else _to_html(doc.package, ctx)
        return {"content": content, "note": _convert_note(doc.package)}
    if path is None:
        raise ToolError(
            "unsupported_format",
            f"'{to}' requires an output path.",
            ["Pass path for pdf/png targets."],
        )
    return _render.render_to_file(doc, to, path)


def convert_document(doc: OpenDocument, to: str, path: str | None = None) -> dict[str, object]:
    """In-language helper used by :class:`Document` (no session lookup)."""
    if to not in ("md", "html", "pdf", "png"):
        raise ToolError(
            "unsupported_format",
            f"Unsupported conversion target: {to}.",
            ["Use to: md, html, pdf or png."],
        )
    if to in ("md", "html"):
        ctx = _build_context(doc.package)
        content = _to_markdown(doc.package, ctx, True) if to == "md" else _to_html(doc.package, ctx)
        return {"content": content, "note": _convert_note(doc.package)}
    if path is None:
        raise ToolError("unsupported_format", f"'{to}' requires an output path.", [])
    return _render.render_to_file(doc, to, path)
