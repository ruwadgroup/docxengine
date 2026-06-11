"""Create a document from Markdown (``docx_create content_md``) — algorithms.md §22/§23a.

Deterministic skeleton parts in the §22 creation order: ``word/document.xml``,
``word/styles.xml``, ``[Content_Types].xml``, ``_rels/.rels``,
``word/_rels/document.xml.rels``, ``docProps/core.xml`` (``dcterms:created``/
``dcterms:modified`` = ``DOCXENGINE_FIXED_DATE`` or its default ``2026-01-01T00:00:00Z``).
Block mapping: ATX headings, quotes, ``---``/``***`` rules, ``-``/``*``/``1.`` list
items (via §17 numbering.xml), GitHub pipe tables (§14), else plain paragraphs.
Inline ``**bold**``/``*italic*``/`` `code` `` split the text into runs at marker
boundaries (§3 escaping). Every emitted body paragraph is a bare ``<w:p>`` with the
trailing ``<w:sectPr>`` carrying the §15 A4 default geometry.

The TypeScript twin (``create.ts``) is the byte-parity reference: ``word/document.xml``
and ``word/styles.xml`` are conformance-compared deep-equal after normalization.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

from . import _xml
from ._errors import ToolError
from ._session import Session
from ._validate import is_valid, validate_package

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

DEFAULT_DATE = "2026-01-01T00:00:00Z"
DEFAULT_CONTENT_WIDTH = 9026  # §15 A4 default content width (twips)

_OL_FORMATS = ("decimal", "lowerLetter", "lowerRoman")
_UL_GLYPHS = ("•", "◦", "▪")


def _core_date() -> str:
    """The created/modified timestamp: ``DOCXENGINE_FIXED_DATE`` or the §22 default."""
    fixed = os.environ.get("DOCXENGINE_FIXED_DATE")
    return fixed if fixed else DEFAULT_DATE


# ---------------------------------------------------------------------------
# Inline markdown → runs (§22)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class InlineRun:
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False


def parse_inline(text: str) -> list[InlineRun]:
    """Split inline markdown into runs at marker boundaries (§22).

    Supports ``**x**``/``__x__`` (bold), ``*x*``/``_x_`` (italic), `` `x` `` (code).
    Markers do not nest in MVP; an unmatched marker is literal text.
    """
    runs: list[InlineRun] = []
    buf = ""
    i = 0

    def flush() -> None:
        nonlocal buf
        if buf != "":
            runs.append(InlineRun(buf))
            buf = ""

    n = len(text)
    while i < n:
        two = text[i : i + 2]
        if two in ("**", "__"):
            close = text.find(two, i + 2)
            if close >= 0:
                flush()
                runs.append(InlineRun(text[i + 2 : close], bold=True))
                i = close + 2
                continue
        ch = text[i]
        if ch in ("*", "_"):
            close = text.find(ch, i + 1)
            if close >= 0 and close > i + 1:
                flush()
                runs.append(InlineRun(text[i + 1 : close], italic=True))
                i = close + 1
                continue
        if ch == "`":
            close = text.find("`", i + 1)
            if close >= 0:
                flush()
                runs.append(InlineRun(text[i + 1 : close], code=True))
                i = close + 1
                continue
        buf += ch
        i += 1
    flush()
    return [r for r in runs if r.text != "" or len(runs) == 1]


def _emit_inline_run(run: InlineRun) -> str:
    rpr_parts: list[str] = []
    if run.bold:
        rpr_parts.append("<w:b/>")
    if run.italic:
        rpr_parts.append("<w:i/>")
    if run.code:
        rpr_parts.append('<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>')
    rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>" if rpr_parts else ""
    return f"<w:r>{rpr}{_xml.emit_text_element(run.text)}</w:r>"


#: HTML line breaks accepted in markdown (``<br>``/``<br/>``/``<br />``) → ``<w:br/>``.
_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
#: A break-only run — valid between styled runs (a bare ``<w:br/>`` may not sit in ``w:p``).
_LINE_BREAK_RUN = "<w:r><w:br/></w:r>"


def emit_inline(text: str) -> str:
    lines = _BR_RE.sub("\n", text).split("\n")
    parts = ["".join(_emit_inline_run(r) for r in parse_inline(line)) for line in lines]
    return _LINE_BREAK_RUN.join(parts)


# ---------------------------------------------------------------------------
# Block parsing (§22)
# ---------------------------------------------------------------------------

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_QUOTE_RE = re.compile(r"^>\s?(.*)$")
_HR_RE = re.compile(r"^(?:-{3,}|\*{3,})$")
_UL_RE = re.compile(r"^[-*]\s+(.*)$")
_OL_RE = re.compile(r"^\d+\.\s+(.*)$")
_TABLE_ROW_RE = re.compile(r"^\|(.+)\|\s*$")
_TABLE_SEP_RE = re.compile(r"^\|?[\s:|-]+\|?$")


def _split_row(line: str) -> list[str]:
    m = _TABLE_ROW_RE.match(line)
    inner = m.group(1) if m else re.sub(r"\|$", "", re.sub(r"^\|", "", line))
    return [c.strip() for c in inner.split("|")]


def _is_separator_row(line: str) -> bool:
    if "-" not in line:
        return False
    return bool(_TABLE_SEP_RE.match(line.strip())) and "|" in line


@dataclass(slots=True)
class BuildState:
    paragraphs: list[str] = field(default_factory=list)
    numbering: list[str] = field(default_factory=list)
    next_num_id: int = 1
    plan_ol: int | None = None
    plan_ul: int | None = None
    body_paragraph_count: int = 0


def _ensure_numbering(state: BuildState, kind: str) -> int:
    """Build the §17 abstractNum + num markup for ol/ul; returns the numId."""
    if kind == "ol" and state.plan_ol is not None:
        return state.plan_ol
    if kind == "ul" and state.plan_ul is not None:
        return state.plan_ul
    nid = state.next_num_id
    state.next_num_id += 1
    levels: list[str] = []
    for ilvl in range(9):
        left = 720 * (ilvl + 1)
        ind = f'<w:ind w:left="{left}" w:hanging="360"/>'
        if kind == "ol":
            fmt = _OL_FORMATS[ilvl % len(_OL_FORMATS)]
            levels.append(
                f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="{_xml.escape_attr(f"%{ilvl + 1}.")}"/>'
                f"<w:pPr>{ind}</w:pPr></w:lvl>"
            )
        else:
            glyph = _UL_GLYPHS[ilvl % len(_UL_GLYPHS)]
            levels.append(
                f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
                f'<w:lvlText w:val="{_xml.escape_attr(glyph)}"/>'
                f"<w:pPr>{ind}</w:pPr></w:lvl>"
            )
    joined = "".join(levels)
    state.numbering.append(f'<w:abstractNum w:abstractNumId="{nid}">{joined}</w:abstractNum>')
    state.numbering.append(f'<w:num w:numId="{nid}"><w:abstractNumId w:val="{nid}"/></w:num>')
    if kind == "ol":
        state.plan_ol = nid
    else:
        state.plan_ul = nid
    return nid


def _emit_list_item(state: BuildState, text: str, kind: str) -> None:
    num_id = _ensure_numbering(state, kind)
    num_pr = f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'
    ppr = f'<w:pPr><w:pStyle w:val="ListParagraph"/>{num_pr}</w:pPr>'
    state.paragraphs.append(f"<w:p>{ppr}{emit_inline(text)}</w:p>")
    state.body_paragraph_count += 1


def _distribute_widths(total: int, cols: int) -> list[int]:
    if cols <= 0:
        return []
    base = total // cols
    widths = [base] * cols
    widths[-1] = total - base * (cols - 1)
    return widths


def _cell_paragraph(text: str, header: bool) -> str:
    if text == "":
        return "<w:p/>"
    if header:
        runs = _xml.emit_text_runs(_BR_RE.sub("\n", text), "<w:rPr><w:b/></w:rPr>")
        return f"<w:p>{runs}</w:p>"
    return f"<w:p>{emit_inline(text)}</w:p>"


def _emit_table(state: BuildState, rows: list[list[str]], header: bool) -> None:
    cols = max((len(r) for r in rows), default=0)
    cols = max(cols, 1)
    widths = _distribute_widths(DEFAULT_CONTENT_WIDTH, cols)
    grid_cols = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tbl_style = '<w:tblStyle w:val="TableGrid"/>' if header else ""
    tbl_pr = f'<w:tblPr>{tbl_style}<w:tblW w:w="0" w:type="auto"/></w:tblPr>'
    trs: list[str] = []
    for r, row in enumerate(rows):
        is_header = header and r == 0
        cells: list[str] = []
        for c in range(cols):
            text = row[c] if c < len(row) else ""
            shd = (
                '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' if is_header else ""
            )
            tc_pr = f'<w:tcPr><w:tcW w:w="{widths[c]}" w:type="dxa"/>{shd}</w:tcPr>'
            cells.append(f"<w:tc>{tc_pr}{_cell_paragraph(text, is_header)}</w:tc>")
        trs.append(f'<w:tr>{"".join(cells)}</w:tr>')
    state.paragraphs.append(
        f'<w:tbl>{tbl_pr}<w:tblGrid>{grid_cols}</w:tblGrid>{"".join(trs)}</w:tbl>'
    )


def _emit_rule(state: BuildState) -> None:
    state.paragraphs.append(
        "<w:p><w:pPr><w:pBdr>"
        '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>'
        "</w:pBdr></w:pPr></w:p>"
    )
    state.body_paragraph_count += 1


def _emit_paragraph(state: BuildState, text: str, style: str | None) -> None:
    ppr = ""
    if style is not None:
        ppr = f'<w:pPr><w:pStyle w:val="{_xml.escape_attr(style)}"/></w:pPr>'
    state.paragraphs.append(f"<w:p>{ppr}{emit_inline(text)}</w:p>")
    state.body_paragraph_count += 1


def build_body(md: str) -> BuildState:
    """Parse the markdown body into the §22 block sequence."""
    state = BuildState()
    lines = [(line[:-1] if line.endswith("\r") else line) for line in md.split("\n")]
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.strip() == "":
            i += 1
            continue
        if _TABLE_ROW_RE.match(line):
            rows: list[list[str]] = []
            j = i
            header = False
            rows.append(_split_row(lines[j]))
            j += 1
            if j < len(lines) and _is_separator_row(lines[j]):
                header = True
                j += 1
            while (
                j < len(lines)
                and _TABLE_ROW_RE.match(lines[j])
                and not _is_separator_row(lines[j])
            ):
                rows.append(_split_row(lines[j]))
                j += 1
            _emit_table(state, rows, header)
            i = j
            continue
        heading = _HEADING_RE.match(line)
        if heading:
            level = len(heading.group(1))
            _emit_paragraph(state, heading.group(2), f"Heading{level}")
            i += 1
            continue
        quote = _QUOTE_RE.match(line)
        if quote:
            _emit_paragraph(state, quote.group(1), "Quote")
            i += 1
            continue
        if _HR_RE.match(line.strip()):
            _emit_rule(state)
            i += 1
            continue
        ul = _UL_RE.match(line)
        if ul:
            _emit_list_item(state, ul.group(1), "ul")
            i += 1
            continue
        ol = _OL_RE.match(line)
        if ol:
            _emit_list_item(state, ol.group(1), "ol")
            i += 1
            continue
        _emit_paragraph(state, line, None)
        i += 1
    return state


# ---------------------------------------------------------------------------
# Skeleton parts (§22/§23a)
# ---------------------------------------------------------------------------


def _styles_xml() -> str:
    headings: list[str] = []
    for n in range(1, 7):
        headings.append(
            f'<w:style w:type="paragraph" w:styleId="Heading{n}">'
            f'<w:name w:val="heading {n}"/><w:basedOn w:val="Normal"/>'
            f'<w:pPr><w:outlineLvl w:val="{n - 1}"/></w:pPr>'
            f"<w:rPr><w:b/></w:rPr></w:style>"
        )
    normal = (
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style>'
    )
    list_paragraph = (
        '<w:style w:type="paragraph" w:styleId="ListParagraph">'
        '<w:name w:val="List Paragraph"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
    )
    table_grid = (
        '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
        "<w:tblPr><w:tblBorders>"
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        "</w:tblBorders></w:tblPr></w:style>"
    )
    quote = (
        '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr>'
        "<w:rPr><w:i/></w:rPr></w:style>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{_W_NS}">'
        + normal
        + "".join(headings)
        + list_paragraph
        + table_grid
        + quote
        + "</w:styles>"
    )


def _document_xml(body_paragraphs: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{_W_NS}">'
        f"<w:body>{body_paragraphs}<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>"
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
    )


def _numbering_xml(entries: list[str]) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:numbering xmlns:w="{_W_NS}">{"".join(entries)}</w:numbering>'
    )


def _content_types_xml(has_numbering: bool) -> str:
    overrides = [
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.wordprocessingml.document.main+xml"/>',
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.wordprocessingml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-'
        'package.core-properties+xml"/>',
    ]
    if has_numbering:
        overrides.append(
            '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-'
            'officedocument.wordprocessingml.numbering+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Types xmlns="{_CT_NS}">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    )


def _root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/'
        'relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        "</Relationships>"
    )


def _document_rels_xml(has_numbering: bool) -> str:
    rels = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/styles" Target="styles.xml"/>'
    ]
    if has_numbering:
        rels.append(
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/numbering" Target="numbering.xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{_REL_NS}">{"".join(rels)}</Relationships>'
    )


def _core_props_xml(date: str) -> str:
    d = _xml.escape_text(date)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/'
        'metadata/core-properties"'
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
        ' xmlns:dcterms="http://purl.org/dc/terms/"'
        ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{d}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{d}</dcterms:modified>'
        "</cp:coreProperties>"
    )


# ---------------------------------------------------------------------------
# Structured-spec → markdown shim
# ---------------------------------------------------------------------------


def _spec_to_markdown(spec: dict[str, object]) -> str:
    """Lower a ``{blocks:[{type,text,level}]}`` spec to markdown (degrade to empty)."""
    blocks = spec.get("blocks")
    if not isinstance(blocks, list):
        return ""
    lines: list[str] = []
    for b in blocks:
        if not isinstance(b, dict):
            continue
        text = b["text"] if isinstance(b.get("text"), str) else ""
        block_type = b["type"] if isinstance(b.get("type"), str) else "paragraph"
        if block_type == "heading":
            raw_level = b.get("level")
            level = min(6, max(1, int(raw_level))) if isinstance(raw_level, int) else 1
            lines.append(f'{"#" * level} {text}')
        elif block_type in ("list_item", "bullet"):
            lines.append(f"- {text}")
        else:
            lines.append(str(text))
        lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# docx_create
# ---------------------------------------------------------------------------


def docx_create(
    session: Session,
    *,
    content_md: str | None = None,
    spec: dict[str, object] | None = None,
) -> dict[str, object]:
    """Create a new document from Markdown (or a structured spec) and register it (§22)."""
    if content_md is not None and spec is not None:
        raise ToolError(
            "invalid_args",
            "Provide exactly one of content_md or spec.",
            ["content_md and spec are mutually exclusive."],
        )
    if content_md is not None:
        md = content_md
    elif spec is not None:
        md = _spec_to_markdown(spec)
    else:
        md = ""
    built = build_body(md)
    has_numbering = len(built.numbering) > 0
    date = _core_date()

    # §22 creation order: document, styles, [numbering,] [Content_Types], _rels/.rels,
    # word/_rels/document.xml.rels, docProps/core.xml.
    parts: dict[str, str] = {
        "word/document.xml": _document_xml("".join(built.paragraphs)),
        "word/styles.xml": _styles_xml(),
    }
    if has_numbering:
        parts["word/numbering.xml"] = _numbering_xml(built.numbering)
    parts["[Content_Types].xml"] = _content_types_xml(has_numbering)
    parts["_rels/.rels"] = _root_rels_xml()
    parts["word/_rels/document.xml.rels"] = _document_rels_xml(has_numbering)
    parts["docProps/core.xml"] = _core_props_xml(date)

    bytes_ = _zip_package(parts)
    doc = session.open_doc(bytes_)

    issues = validate_package(doc.package)
    if not is_valid(issues):
        errors = [issue for issue in issues if issue.severity == "error"]
        raise ToolError(
            "validation_failed",
            "Created document failed validation.",
            [issue.message for issue in errors],
        )
    return {"doc_id": doc.doc_id, "n_paragraphs": built.body_paragraph_count}


def _zip_package(parts: dict[str, str]) -> bytes:
    """Build the zip bytes for the skeleton parts (stored — re-stored on save)."""
    import io
    import zipfile

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name, xml in parts.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.create_system = 0
            zf.writestr(info, xml.encode("utf-8"), compress_type=zipfile.ZIP_STORED)
    return buf.getvalue()
