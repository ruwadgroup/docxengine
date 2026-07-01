"""Shared in-test .docx fixture builders (tiny OPC packages built with zipfile)."""

from __future__ import annotations

import io
import zipfile

import pytest

CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)

ROOT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)

DOCUMENT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/styles" Target="styles.xml"/>'
    "</Relationships>"
)

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:styleId="Heading1">'
    '<w:name w:val="heading 1"/></w:style>'
    "</w:styles>"
)

# Paragraph 1 — the algorithms.md §1 worked example: split runs, proofErr noise,
# odd internal whitespace. Normalizes to "Master Services Agreement" -> hash 515a.
PARA_SPLIT_RUN = (
    "<w:p>"
    '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    "<w:r><w:rPr><w:b/></w:rPr><w:t>Master</w:t></w:r>"
    '<w:r><w:t xml:space="preserve"> Services</w:t></w:r>'
    '<w:proofErr w:type="spellStart"/>'
    '<w:r><w:t xml:space="preserve">  Agreement </w:t></w:r>'
    "</w:p>"
)

# Paragraph 2 — rsid-fragmented: Word split one sentence into per-save runs.
PARA_RSID_FRAGMENTED = (
    '<w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD">'
    '<w:r w:rsidR="00AB12CD"><w:t>The term is </w:t></w:r>'
    '<w:r w:rsidR="00EF34AB"><w:t>five (5) </w:t></w:r>'
    '<w:r w:rsidR="00CD56EF"><w:t>years from the </w:t></w:r>'
    '<w:r w:rsidR="00115E6B"><w:t>Effective Date.</w:t></w:r>'
    "</w:p>"
)

# Paragraph 3 — tracked changes: a deletion (w:delText excluded from anchors) and
# an insertion (included: the hash sees the document as-if-accepted).
PARA_TRACKED = (
    "<w:p>"
    '<w:r><w:t xml:space="preserve">Payment due in </w:t></w:r>'
    '<w:del w:id="1" w:author="J.Doe" w:date="2026-01-01T00:00:00Z">'
    "<w:r><w:delText>30</w:delText></w:r></w:del>"
    '<w:ins w:id="2" w:author="J.Doe" w:date="2026-01-01T00:00:00Z">'
    "<w:r><w:t>45</w:t></w:r></w:ins>"
    '<w:r><w:t xml:space="preserve"> days</w:t></w:r>'
    "</w:p>"
)

SECT_PR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'


def document_xml(*body_children: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{''.join(body_children)}</w:body></w:document>"
    )


DOCUMENT_XML = document_xml(PARA_SPLIT_RUN, PARA_RSID_FRAGMENTED, PARA_TRACKED, SECT_PR)

FIXTURE_PARTS: dict[str, str] = {
    "[Content_Types].xml": CONTENT_TYPES_XML,
    "_rels/.rels": ROOT_RELS_XML,
    "word/document.xml": DOCUMENT_XML,
    "word/_rels/document.xml.rels": DOCUMENT_RELS_XML,
    "word/styles.xml": STYLES_XML,
}


def build_docx(parts: dict[str, str] | None = None) -> bytes:
    """Zip the given parts (default: the standard 3-paragraph fixture) into .docx bytes."""
    parts = FIXTURE_PARTS if parts is None else parts
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, text in parts.items():
            zf.writestr(name, text.encode("utf-8"))
    return buf.getvalue()


@pytest.fixture
def docx_bytes() -> bytes:
    return build_docx()
