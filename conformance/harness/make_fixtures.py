#!/usr/bin/env python3
"""Generate the deterministic conformance corpus (conformance/corpus/<name>/).

Stdlib only. Output is byte-deterministic across runs and platforms:

- fixed revision dates inside the XML,
- fixed ZIP entry metadata (DOS timestamp 1980-01-01 00:00:00, create_system=3,
  external_attr 0644, deflate level 6, no extra fields/comments),
- fixed entry order,
- meta.json serialized with sorted keys and a trailing newline.

Each fixture directory gets input.docx + meta.json (producer, features, notes,
and the computed paragraph anchors per spec/algorithms.md §1 for case authoring).

Usage: python3 conformance/harness/make_fixtures.py [corpus_dir]
"""

from __future__ import annotations

import base64
import hashlib
import json
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

CT_DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_PNG = "image/png"

RT_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
RT_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
RT_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
RT_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
RT_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
RT_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

# Drawing namespaces (spec §19 inline drawing).
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

PRODUCER = "DocxEngine make_fixtures.py 1.0 (hand-assembled synthetic OOXML)"

# Unicode White_Space=Yes, exactly the set pinned by spec/algorithms.md §1.
WS_CHARS = frozenset(
    "\t\n\x0b\x0c\r \x85\xa0\u1680"
    + "".join(chr(c) for c in range(0x2000, 0x200B))
    + "\u2028\u2029\u202f\u205f\u3000"
)

SECT_PR = (
    "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
    "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\""
    " w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>"
)


# ---------------------------------------------------------------------------
# Anchor computation (spec/algorithms.md §1)
# ---------------------------------------------------------------------------


def normalize_text(s: str) -> str:
    s = unicodedata.normalize("NFC", s)
    out: list[str] = []
    in_ws = False
    for ch in s:
        if ch in WS_CHARS:
            in_ws = True
            continue
        if in_ws and out:
            out.append(" ")
        in_ws = False
        out.append(ch)
    return "".join(out)


def anchor_hash(normalized: str) -> str:
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:4]


def body_paragraph_anchors(document_xml: bytes) -> dict[str, str]:
    """Map 'P{n}' -> full anchor for every body-level w:p (as-if-accepted)."""
    root = ET.fromstring(document_xml)
    body = root.find(f"{{{W_NS}}}body")
    assert body is not None, "document has no w:body"
    anchors: dict[str, str] = {}
    ordinal = 0
    for child in body:
        if child.tag != f"{{{W_NS}}}p":
            continue
        ordinal += 1
        text = "".join(t.text or "" for t in child.iter(f"{{{W_NS}}}t"))
        anchors[f"P{ordinal}"] = f"P{ordinal}#{anchor_hash(normalize_text(text))}"
    return anchors


# ---------------------------------------------------------------------------
# XML assembly helpers
# ---------------------------------------------------------------------------


def esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def wt(text: str, tag: str = "w:t") -> str:
    """Emit w:t / w:delText with xml:space per spec §3."""
    preserve = bool(text) and (text[0] in WS_CHARS or text[-1] in WS_CHARS)
    attr = ' xml:space="preserve"' if preserve else ""
    return f"<{tag}{attr}>{esc(text)}</{tag}>"


def run(text: str, *, bold: bool = False, rsid: str | None = None, deleted: bool = False) -> str:
    attrs = f' w:rsidR="{rsid}"' if rsid else ""
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    body = wt(text, "w:delText" if deleted else "w:t")
    return f"<w:r{attrs}>{rpr}{body}</w:r>"


def para(inner: str, *, style: str | None = None, rsid: str | None = None) -> str:
    attrs = f' w:rsidR="{rsid}"' if rsid else ""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p{attrs}>{ppr}{inner}</w:p>"


def ins(rev_id: int, author: str, date: str, inner: str) -> str:
    return f'<w:ins w:id="{rev_id}" w:author="{author}" w:date="{date}">{inner}</w:ins>'


def dele(rev_id: int, author: str, date: str, inner: str) -> str:
    return f'<w:del w:id="{rev_id}" w:author="{author}" w:date="{date}">{inner}</w:del>'


def document(body: str, *, sect_pr: str = SECT_PR, extra_ns: dict[str, str] | None = None) -> bytes:
    ns = f'xmlns:w="{W_NS}" xmlns:r="{R_NS}"'
    for prefix, uri in (extra_ns or {}).items():
        ns += f' xmlns:{prefix}="{uri}"'
    return (
        XML_DECL
        + f"<w:document {ns}>"
        + f"<w:body>{body}{sect_pr}</w:body></w:document>"
    ).encode("utf-8")


def content_types(overrides: list[tuple[str, str]]) -> bytes:
    parts = [XML_DECL, f'<Types xmlns="{CT_NS}">']
    parts.append(
        '<Default Extension="rels"'
        ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    )
    parts.append('<Default Extension="xml" ContentType="application/xml"/>')
    for name, ct in overrides:
        parts.append(f'<Override PartName="{name}" ContentType="{ct}"/>')
    parts.append("</Types>")
    return "".join(parts).encode("utf-8")


def content_types_with_defaults(
    overrides: list[tuple[str, str]], extra_defaults: list[tuple[str, str]]
) -> bytes:
    """Like content_types but with additional <Default Extension=…> entries
    (e.g. png for media). Defaults precede Overrides, matching OPC convention."""
    parts = [XML_DECL, f'<Types xmlns="{CT_NS}">']
    parts.append(
        '<Default Extension="rels"'
        ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    )
    parts.append('<Default Extension="xml" ContentType="application/xml"/>')
    for ext, ct in extra_defaults:
        parts.append(f'<Default Extension="{ext}" ContentType="{ct}"/>')
    for name, ct in overrides:
        parts.append(f'<Override PartName="{name}" ContentType="{ct}"/>')
    parts.append("</Types>")
    return "".join(parts).encode("utf-8")


PKG_RELS = (
    XML_DECL + f'<Relationships xmlns="{REL_NS}">'
    f'<Relationship Id="rId1" Type="{RT_DOCUMENT}" Target="word/document.xml"/>'
    "</Relationships>"
).encode("utf-8")


def relationships(rels: list[tuple[str, str, str]]) -> bytes:
    parts = [XML_DECL, f'<Relationships xmlns="{REL_NS}">']
    for rel_id, rel_type, target in rels:
        parts.append(f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/>')
    parts.append("</Relationships>")
    return "".join(parts).encode("utf-8")


def styles_xml(*, headings: bool, extra: str = "") -> bytes:
    parts = [XML_DECL, f'<w:styles xmlns:w="{W_NS}">']
    parts.append(
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style>'
    )
    if headings:
        for n in (1, 2):
            parts.append(
                f'<w:style w:type="paragraph" w:styleId="Heading{n}">'
                f'<w:name w:val="heading {n}"/><w:basedOn w:val="Normal"/>'
                f'<w:pPr><w:outlineLvl w:val="{n - 1}"/></w:pPr></w:style>'
            )
    parts.append(extra)
    parts.append("</w:styles>")
    return "".join(parts).encode("utf-8")


# Style fragments ensured on demand by Phase 2 ops (spec §14/§16/§17).
TABLE_GRID_STYLE = (
    '<w:style w:type="table" w:styleId="TableGrid">'
    '<w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/></w:style>'
)
LIST_PARAGRAPH_STYLE = (
    '<w:style w:type="paragraph" w:styleId="ListParagraph">'
    '<w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>'
)


STD_OVERRIDES = [
    ("/word/document.xml", CT_DOCUMENT),
    ("/word/styles.xml", CT_STYLES),
]
STD_RELS = [("rId1", RT_STYLES, "styles.xml")]

DATE_ALICE = "2026-01-15T09:30:00Z"
DATE_BOB = "2026-02-20T16:45:00Z"
DATE_COMMENT = "2026-03-01T10:00:00Z"


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def fx_minimal() -> tuple[list[tuple[str, bytes]], dict]:
    body = (
        para(run("Master Services Agreement"), style="Heading1")
        + para(run("This Agreement is entered into by the parties as of the Effective Date."))
        + para(run("Definitions"), style="Heading2")
        + para(
            run(
                "Confidential Information means all information disclosed"
                " by one party to the other."
            )
        )
        + para(run("Each party shall protect Confidential Information with reasonable care."))
    )
    doc = document(body)
    parts = [
        ("[Content_Types].xml", content_types(STD_OVERRIDES)),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", doc),
        ("word/_rels/document.xml.rels", relationships(STD_RELS)),
        ("word/styles.xml", styles_xml(headings=True)),
    ]
    meta = {
        "features": ["headings", "styles.xml", "style cascade (basedOn)", "plain paragraphs"],
        "notes": "Smallest clean document: Heading1/Heading2 + three body paragraphs.",
    }
    return parts, meta


def fx_split_runs() -> tuple[list[tuple[str, bytes]], dict]:
    body = (
        para(
            run("Term ", rsid="00A1B2C3") + run("and ", rsid="00D4E5F6")
            + run("Termination", rsid="00A1B2C3"),
            rsid="00A1B2C3",
        )
        # The spec/algorithms.md §4 worked example, verbatim structure.
        + para(run("The term is five (5) ") + run("years from the Effective Date.", bold=True))
        + para(
            run("Payment is due wi", rsid="00112233")
            + run("thin thirty (30) da", rsid="00445566")
            + run("ys of invoice.", rsid="00778899"),
            rsid="00112233",
        )
    )
    parts = [
        ("[Content_Types].xml", content_types(STD_OVERRIDES)),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body)),
        ("word/_rels/document.xml.rels", relationships(STD_RELS)),
        ("word/styles.xml", styles_xml(headings=False)),
    ]
    meta = {
        "features": [
            "rsid-fragmented runs",
            "formatting splits (bold run mid-sentence)",
            "run coalescing targets (spec §4 worked example)",
        ],
        "notes": "P2 is the algorithms.md §4/§5 worked example; matches span run boundaries.",
    }
    return parts, meta


def fx_redlines() -> tuple[list[tuple[str, bytes]], dict]:
    body = (
        para(run("Revision History"))
        + para(
            run("The fee is ")
            + dele(1, "Alice", DATE_ALICE, run("ten percent", deleted=True))
            + ins(2, "Alice", DATE_ALICE, run("twelve percent"))
            + run(" of net revenue.")
        )
        + para(
            run("Notices must be sent ")
            + ins(3, "Bob", DATE_BOB, run("by certified mail "))
            + run("to the address below")
            + dele(4, "Bob", DATE_BOB, run(" within five days", deleted=True))
            + run(".")
        )
    )
    parts = [
        ("[Content_Types].xml", content_types(STD_OVERRIDES)),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body)),
        ("word/_rels/document.xml.rels", relationships(STD_RELS)),
        ("word/styles.xml", styles_xml(headings=False)),
    ]
    meta = {
        "features": ["multi-author tracked changes", "w:ins", "w:del", "w:delText"],
        "notes": (
            "Alice owns revisions 1-2 (P2), Bob owns 3-4 (P3)."
            " Dates: Alice 2026-01-15, Bob 2026-02-20."
        ),
    }
    return parts, meta


def fx_comments_footnotes() -> tuple[list[tuple[str, bytes]], dict]:
    body = (
        para(run("Project Report"))
        + para(
            run("Revenue grew ")
            + '<w:commentRangeStart w:id="1"/>'
            + run("nine percent")
            + '<w:commentRangeEnd w:id="1"/>'
            + '<w:r><w:commentReference w:id="1"/></w:r>'
            + run(" year over year.")
        )
        + para(run("Figures are audited.") + '<w:r><w:footnoteReference w:id="2"/></w:r>')
    )
    comments = (
        XML_DECL + f'<w:comments xmlns:w="{W_NS}">'
        f'<w:comment w:id="1" w:author="J.Doe" w:date="{DATE_COMMENT}" w:initials="JD">'
        "<w:p><w:r><w:t>Please verify this figure.</w:t></w:r></w:p>"
        "</w:comment></w:comments>"
    ).encode("utf-8")
    footnotes = (
        XML_DECL + f'<w:footnotes xmlns:w="{W_NS}">'
        '<w:footnote w:type="separator" w:id="0">'
        "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
        '<w:footnote w:type="continuationSeparator" w:id="1">'
        "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
        '<w:footnote w:id="2"><w:p><w:r><w:t>Source: annual filing.</w:t></w:r></w:p>'
        "</w:footnote></w:footnotes>"
    ).encode("utf-8")
    parts = [
        (
            "[Content_Types].xml",
            content_types(
                STD_OVERRIDES
                + [("/word/comments.xml", CT_COMMENTS), ("/word/footnotes.xml", CT_FOOTNOTES)]
            ),
        ),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body)),
        (
            "word/_rels/document.xml.rels",
            relationships(
                STD_RELS
                + [("rId2", RT_COMMENTS, "comments.xml"), ("rId3", RT_FOOTNOTES, "footnotes.xml")]
            ),
        ),
        ("word/styles.xml", styles_xml(headings=False)),
        ("word/comments.xml", comments),
        ("word/footnotes.xml", footnotes),
    ]
    meta = {
        "features": [
            "comments.xml with commentRangeStart/End + commentReference",
            "footnotes.xml with separator/continuationSeparator + real footnote",
            "passthrough fidelity (untouched parts must survive byte-for-byte)",
        ],
        "notes": "Comment id=1 and footnote id=2 resolve in both directions (validator check e).",
    }
    return parts, meta


def fx_corrupt_orphan_rel() -> tuple[list[tuple[str, bytes]], dict]:
    body = (
        para(run("Quarterly Update"))
        + para(
            run("See the ")
            + '<w:hyperlink r:id="rId8" w:history="1">'
            + run("full report")
            + "</w:hyperlink>"
            + run(" for details.")
        )
    )
    # Defect 1 (check c, auto-repairable): rId7 targets a part missing from the package.
    # Defect 2 (check b, NOT auto-repairable): document.xml references rId8, absent from rels.
    rels = relationships(STD_RELS + [("rId7", RT_IMAGE, "media/image1.png")])
    parts = [
        ("[Content_Types].xml", content_types(STD_OVERRIDES)),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body)),
        ("word/_rels/document.xml.rels", rels),
        ("word/styles.xml", styles_xml(headings=False)),
    ]
    meta = {
        "features": [
            "corrupt-on-purpose: relationship target missing from package (rId7, check c)",
            "corrupt-on-purpose: dangling r:id reference in document.xml (rId8, check b)",
        ],
        "notes": (
            "Validate must report 2 errors. rId7 is mechanically repairable (drop);"
            " rId8 is not, so docx_save must refuse with validation_failed."
        ),
    }
    return parts, meta


def fx_corrupt_dup_ids() -> tuple[list[tuple[str, bytes]], dict]:
    body = (
        para(run("Amendment"))
        + para(
            run("The closing date is ")
            + dele(5, "Alice", DATE_ALICE, run("March 1", deleted=True))
            + ins(5, "Alice", DATE_ALICE, run("April 1"))
            + run(", as amended.")
        )
        + para(
            run("Counsel ")
            + ins(6, "Alice", DATE_ALICE, run("promptly "))
            + run("notified all parties.")
        )
    )
    parts = [
        ("[Content_Types].xml", content_types(STD_OVERRIDES)),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body)),
        ("word/_rels/document.xml.rels", relationships(STD_RELS)),
        ("word/styles.xml", styles_xml(headings=False)),
    ]
    meta = {
        "features": [
            "corrupt-on-purpose: duplicate revision id (w:del id=5 and w:ins id=5, check d)"
        ],
        "notes": (
            "Fully auto-repairable: docx_repair renumbers the second id=5 to max+1 (7);"
            " validate must then be clean."
        ),
    }
    return parts, meta


# ---------------------------------------------------------------------------
# Phase 2 fixture helpers (spec §13–§22)
# ---------------------------------------------------------------------------


def tc(text: str, width: int, *, header: bool = False) -> str:
    """One w:tc per spec §14: w:tcPr (w:tcW + optional shd), then one w:p."""
    shd = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' if header else ""
    tcpr = f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}</w:tcPr>'
    if text:
        rpr = "<w:rPr><w:b/></w:rPr>" if header else ""
        cell_p = f"<w:p><w:r>{rpr}{wt(text)}</w:r></w:p>"
    else:
        cell_p = "<w:p/>"
    return f"<w:tc>{tcpr}{cell_p}</w:tc>"


def table(rows: list[list[str]], *, header: bool, cols: int) -> str:
    """Body-level w:tbl per spec §14: equal int widths summing to 9026."""
    base = 9026 // cols
    widths = [base] * cols
    widths[-1] = 9026 - base * (cols - 1)
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tblpr = (
        '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
    )
    trs = []
    for i, row in enumerate(rows):
        is_header = header and i == 0
        cells = "".join(tc(row[c], widths[c], header=is_header) for c in range(cols))
        trs.append(f"<w:tr>{cells}</w:tr>")
    return f"<w:tbl>{tblpr}<w:tblGrid>{grid}</w:tblGrid>{''.join(trs)}</w:tbl>"


def numpr(ilvl: int, num_id: int) -> str:
    return f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'


def list_item(text: str, *, ilvl: int, num_id: int) -> str:
    """A ListParagraph item carrying numPr as the first w:pPr child (spec §17)."""
    ppr = f'<w:pPr><w:pStyle w:val="ListParagraph"/>{numpr(ilvl, num_id)}</w:pPr>'
    return f"<w:p>{ppr}<w:r>{wt(text)}</w:r></w:p>"


def fx_tables() -> tuple[list[tuple[str, bytes]], dict]:
    """A header-row table + a plain 3x3 table, both body-level (spec §14)."""
    header_tbl = table(
        [["Term", "Value", "Notes"], ["Fee", "$100", "monthly"]],
        header=True,
        cols=3,
    )
    grid_tbl = table(
        [["A1", "B1", "C1"], ["A2", "B2", "C2"], ["A3", "B3", "C3"]],
        header=False,
        cols=3,
    )
    body = (
        para(run("Pricing Schedule"), style="Heading1")
        + para(run("The fee table follows."))
        + header_tbl
        + para(run("A plain three-by-three grid follows."))
        + grid_tbl
    )
    extra = TABLE_GRID_STYLE
    parts = [
        ("[Content_Types].xml", content_types(STD_OVERRIDES)),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body)),
        ("word/_rels/document.xml.rels", relationships(STD_RELS)),
        ("word/styles.xml", styles_xml(headings=True, extra=extra)),
    ]
    meta = {
        "features": [
            "body-level w:tbl with header row (shd D9D9D9 + bold, spec §14)",
            "plain 3x3 w:tbl (T-sequence independent of P-sequence, spec §13)",
            "TableGrid style ensured in styles.xml",
        ],
        "notes": (
            "Two body tables: T1 = header table (Term/Value/Notes), T2 = 3x3 grid."
            " Body paragraphs P1..P3 are not shifted by the tables (spec §13)."
        ),
    }
    return parts, meta


def fx_numbered_lists() -> tuple[list[tuple[str, bytes]], dict]:
    """ol (two levels) + ul, driven by numbering.xml (spec §17)."""
    # abstractNum 0 = ordered (decimal / lowerLetter cascade); 1 = bullet.
    ol_levels = (
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>'
        '<w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
    )
    ul_levels = (
        '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="•"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
    )
    numbering = (
        XML_DECL + f'<w:numbering xmlns:w="{W_NS}">'
        f'<w:abstractNum w:abstractNumId="0">{ol_levels}</w:abstractNum>'
        f'<w:abstractNum w:abstractNumId="1">{ul_levels}</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
        "</w:numbering>"
    ).encode("utf-8")
    body = (
        para(run("Procedures"), style="Heading1")
        + list_item("First step", ilvl=0, num_id=1)
        + list_item("Sub-step under first", ilvl=1, num_id=1)
        + list_item("Second step", ilvl=0, num_id=1)
        + para(run("Materials"), style="Heading2")
        + list_item("Hammer", ilvl=0, num_id=2)
        + list_item("Nails", ilvl=0, num_id=2)
    )
    extra = LIST_PARAGRAPH_STYLE
    parts = [
        (
            "[Content_Types].xml",
            content_types(STD_OVERRIDES + [("/word/numbering.xml", CT_NUMBERING)]),
        ),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body)),
        (
            "word/_rels/document.xml.rels",
            relationships(STD_RELS + [("rId2", RT_NUMBERING, "numbering.xml")]),
        ),
        ("word/styles.xml", styles_xml(headings=True, extra=extra)),
        ("word/numbering.xml", numbering),
    ]
    meta = {
        "features": [
            "numbering.xml with ordered (decimal/lowerLetter) + bullet abstractNums (spec §17)",
            "two-level ol (numId=1) then a ul (numId=2)",
            "ListParagraph style ensured",
        ],
        "notes": (
            "P2..P4 are ol items (P3 is ilvl=1); P6..P7 are ul items."
            " numId=1 -> abstractNum 0 (ordered), numId=2 -> abstractNum 1 (bullet)."
        ),
    }
    return parts, meta


def fx_headers_footers() -> tuple[list[tuple[str, bytes]], dict]:
    """One section carrying a default header part (spec §15)."""
    header_xml = (
        XML_DECL + f'<w:hdr xmlns:w="{W_NS}">'
        "<w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:hdr>"
    ).encode("utf-8")
    # sectPr gains a headerReference (references precede pgSz, spec §15).
    sect_pr = (
        '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    )
    body = (
        para(run("Annual Report"), style="Heading1")
        + para(run("This document has a running header on every page."))
    )
    parts = [
        (
            "[Content_Types].xml",
            content_types(STD_OVERRIDES + [("/word/header1.xml", CT_HEADER)]),
        ),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body, sect_pr=sect_pr)),
        (
            "word/_rels/document.xml.rels",
            relationships(STD_RELS + [("rId2", RT_HEADER, "header1.xml")]),
        ),
        ("word/styles.xml", styles_xml(headings=True)),
        ("word/header1.xml", header_xml),
    ]
    meta = {
        "features": [
            "single section with a default w:headerReference -> word/header1.xml (spec §15)",
            "header part covered by content-type Override + document rel",
        ],
        "notes": (
            "One body section (S1). Letter page size 12240x15840."
            " set_geometry / set_footer / insert_break target S1."
        ),
    }
    return parts, meta


def fx_template() -> tuple[list[tuple[str, bytes]], dict]:
    """Mustache template: split-run placeholder + a loop section (spec §21)."""
    body = (
        para(run("Engagement Letter"), style="Heading1")
        # Split-run placeholder: {{Client}} fragmented across two runs (spec §21).
        + para(run("Client: {{Cli") + run("ent}}"))
        + para(run("Effective Date: {{EffectiveDate}}"))
        + para(run("Obligations:"))
        # Loop section spanning whole paragraphs: open and close on their own.
        + para(run("{{#obligations}}"))
        + para(run("- {{text}}"))
        + para(run("{{/obligations}}"))
        + para(run("Counter-signed by {{Signatory}}."))
    )
    parts = [
        ("[Content_Types].xml", content_types(STD_OVERRIDES)),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body)),
        ("word/_rels/document.xml.rels", relationships(STD_RELS)),
        ("word/styles.xml", styles_xml(headings=True)),
    ]
    meta = {
        "features": [
            "mustache {{Client}} split across two runs (coalesced match, spec §21)",
            "loop section {{#obligations}}…{{/obligations}} over whole paragraphs",
            "unfilled placeholder left verbatim ({{Signatory}} when data omits it)",
        ],
        "notes": (
            "Placeholders: Client (split runs), EffectiveDate, obligations (loop, item key"
            " 'text'), Signatory. Fill with obligations=[…] to clone P5; omit Signatory to"
            " exercise the unfilled list."
        ),
    }
    return parts, meta


# Deterministic 1x1 truecolor PNG (69 bytes); IHDR width/height at bytes 16-24.
PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR42mP4z8AAAAMBAQD3A0FDAAAAAElFTkSuQmCC"
)


def fx_media_doc() -> tuple[list[tuple[str, bytes]], dict]:
    """One inline-drawing PNG (spec §19), EMU extent for ~1.27 cm square."""
    cx = round(1.27 * 360000)  # 457200 EMU
    cy = cx
    drawing = (
        "<w:r><w:drawing>"
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:docPr id="1" name="image1"/>'
        f'<a:graphic xmlns:a="{A_NS}">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="{PIC_NS}"><pic:nvPicPr>'
        '<pic:cNvPr id="1" name="image1"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rId2"/>'
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )
    body = (
        para(run("Company Logo"), style="Heading1")
        + para(drawing)
        + para(run("The logo above is an inline drawing."))
    )
    parts = [
        (
            "[Content_Types].xml",
            content_types_with_defaults(STD_OVERRIDES, [("png", CT_PNG)]),
        ),
        ("_rels/.rels", PKG_RELS),
        ("word/document.xml", document(body, extra_ns={"wp": WP_NS})),
        (
            "word/_rels/document.xml.rels",
            relationships(STD_RELS + [("rId2", RT_IMAGE, "media/image1.png")]),
        ),
        ("word/styles.xml", styles_xml(headings=True)),
        ("word/media/image1.png", PNG_1X1),
    ]
    meta = {
        "features": [
            "inline drawing run -> a:blip r:embed -> word/media/image1.png (spec §19)",
            "png Default content-type + image relationship",
            "M1 = first drawing reference in document order (spec §13)",
        ],
        "notes": (
            "One embedded 1x1 PNG (69 bytes) as an inline drawing in P2."
            " extent 457200x457200 EMU (1.27 cm square). media_id M1 for extract/replace."
        ),
    }
    return parts, meta


FIXTURES = {
    "minimal": fx_minimal,
    "split-runs": fx_split_runs,
    "redlines": fx_redlines,
    "comments-footnotes": fx_comments_footnotes,
    "corrupt-orphan-rel": fx_corrupt_orphan_rel,
    "corrupt-dup-ids": fx_corrupt_dup_ids,
    "tables": fx_tables,
    "numbered-lists": fx_numbered_lists,
    "headers-footers": fx_headers_footers,
    "template": fx_template,
    "media-doc": fx_media_doc,
}


# ---------------------------------------------------------------------------
# Deterministic ZIP + corpus writer
# ---------------------------------------------------------------------------


def write_docx(path: Path, parts: list[tuple[str, bytes]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in parts:
            zi = zipfile.ZipInfo(filename=name, date_time=(1980, 1, 1, 0, 0, 0))
            zi.create_system = 3
            zi.external_attr = 0o644 << 16
            zi.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(zi, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=6)


def build_corpus(corpus_dir: Path, *, quiet: bool = False) -> dict[str, dict[str, str]]:
    """Generate every fixture; returns {fixture: {P-ordinal: anchor}}."""
    all_anchors: dict[str, dict[str, str]] = {}
    for name, builder in FIXTURES.items():
        parts, meta = builder()
        doc_xml = dict(parts)["word/document.xml"]
        anchors = body_paragraph_anchors(doc_xml)
        all_anchors[name] = anchors
        write_docx(corpus_dir / name / "input.docx", parts)
        meta_full = {"producer": PRODUCER, "anchors": anchors, **meta}
        meta_path = corpus_dir / name / "meta.json"
        meta_path.write_text(
            json.dumps(meta_full, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )
        if not quiet:
            ordered = sorted(anchors, key=lambda s: int(s[1:]))
            print(f"{name}: " + " ".join(anchors[k] for k in ordered))
    # Self-check against the worked example pinned in spec/algorithms.md §1.
    assert all_anchors["minimal"]["P1"] == "P1#515a", all_anchors["minimal"]["P1"]
    return all_anchors


def main(argv: list[str]) -> int:
    default = Path(__file__).resolve().parent.parent / "corpus"
    corpus_dir = Path(argv[1]).resolve() if len(argv) > 1 else default
    build_corpus(corpus_dir)
    print(f"corpus written to {corpus_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
