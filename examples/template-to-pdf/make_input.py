"""Build msa-template.docx: a synthetic MSA template with mustache placeholders.

Mirrors the redline-review fixture style: hand-assembled OOXML, deterministic ZIP
entries, stdlib only. The body carries three kinds of mustache syntax so the
example exercises the whole template surface:

- scalar placeholders   {{EffectiveDate}}, {{Client}}
- a section loop        {{#obligations}} ... {{text}} ... {{/obligations}}

docx_template_fill coalesces placeholders fragmented across runs, so the template
is intentionally simple here; run-splitting is covered by the conformance corpus.
"""

import zipfile
import zlib
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {W}><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Master Services Agreement</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">This Agreement is effective as of {{{{EffectiveDate}}}} between the Provider and {{{{Client}}}}.</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Obligations</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">The Provider shall: {{{{#obligations}}}}</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">- {{{{text}}}}</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">{{{{/obligations}}}}</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Signed for {{{{Client}}}} on {{{{EffectiveDate}}}}.</w:t></w:r></w:p>
<w:sectPr/></w:body></w:document>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
</w:styles>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def main() -> None:
    out = Path(__file__).parent / "msa-template.docx"
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/document.xml": DOCUMENT,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/styles.xml": STYLES,
    }
    with zipfile.ZipFile(out, "w") as z:
        for name, content in parts.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode("utf-8"), compresslevel=zlib.Z_DEFAULT_COMPRESSION)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
