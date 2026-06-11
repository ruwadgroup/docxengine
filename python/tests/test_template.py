"""docx_template_fill tests (algorithms.md §21/§23a).

Covers split-run placeholders (coalesced before matching), loop section expansion
(clone spanned paragraphs / a single table row), inverted sections, missing-var
reporting in `unfilled` (with strict raising), the `filled` count including loop
clones, and XML-only escaping of values.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from docxengine import Session, ToolError, docx_convert, docx_template_fill

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def _p(*runs: str) -> str:
    return "<w:p>" + "".join(runs) + "</w:p>"


def _r(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def _cell(text: str) -> str:
    tc_pr = '<w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr>'
    return f"<w:tc>{tc_pr}{_p(_r(text))}</w:tc>"


def _build_template(body: str, tmp_path) -> str:  # noqa: ANN001
    doc = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<w:document xmlns:w="{_W}"><w:body>{body}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>'
    )
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.wordprocessingml.document.main+xml"/></Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
    path = tmp_path / "template.docx"
    path.write_bytes(buf.getvalue())
    return str(path)


def _document(session: Session, doc_id: str) -> str:
    pkg = session.get(doc_id).package
    return pkg.part(pkg.main_document_part()).decode("utf-8")


# ---------------------------------------------------------------------------
# Split-run placeholders
# ---------------------------------------------------------------------------


class TestSplitRunPlaceholder:
    def test_placeholder_across_runs(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(
            _r("Client: "),
            "<w:r><w:t>{{Cli</w:t></w:r>",
            "<w:r><w:t>ent}}</w:t></w:r>",
        )
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(
            session, template=path, data={"Client": "GlobalTech & Co"}
        )
        doc = _document(session, str(result["doc_id"]))
        # §4 first-overlap: value written into the first overlapping w:t (the `{{Cli` run);
        # the `Client: ` prefix run is untouched; the trailing run is emptied.
        assert "<w:t>GlobalTech &amp; Co</w:t>" in doc
        assert "{{" not in doc  # placeholder fully consumed across runs
        # Convert resolves the coalesced text.
        md = docx_convert(session, doc_id=str(result["doc_id"]), to="md")["content"]
        assert md == "Client: GlobalTech & Co"
        assert result["filled"] == 1
        assert result["unfilled"] == []

    def test_value_xml_escaped_not_html_escaped(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("X = {{v}}"))
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(session, template=path, data={"v": 'a<b>"c"&d'})
        doc = _document(session, str(result["doc_id"]))
        # §3: only &,<,> escaped; quotes stay literal in text.
        assert 'a&lt;b&gt;"c"&amp;d' in doc


# ---------------------------------------------------------------------------
# Loop sections
# ---------------------------------------------------------------------------


class TestLoops:
    def test_loop_clones_inner_paragraphs(self, tmp_path) -> None:  # noqa: ANN001
        body = (
            _p(_r("Client: {{Client}}"))
            + _p(_r("{{#obligations}}"))
            + _p(_r("- {{text}}"))
            + _p(_r("{{/obligations}}"))
        )
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(
            session,
            template=path,
            data={
                "Client": "Acme",
                "obligations": [{"text": "Pay fees"}, {"text": "Keep records"}],
            },
        )
        out = docx_convert(session, doc_id=str(result["doc_id"]), to="md")["content"]
        assert "- Pay fees" in out
        assert "- Keep records" in out
        # §23a: filled counts Client + text×2 = 3.
        assert result["filled"] == 3
        assert result["loops_expanded"] == {"obligations": 2}

    def test_loop_over_table_row(self, tmp_path) -> None:  # noqa: ANN001
        row = f"<w:tr>{_cell('{{#rows}}{{name}}')}{_cell('{{price}}{{/rows}}')}</w:tr>"
        tbl = (
            '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>'
            f"{row}</w:tbl>"
        )
        path = _build_template(tbl, tmp_path)
        session = Session()
        result = docx_template_fill(
            session,
            template=path,
            data={"rows": [{"name": "A", "price": "1"}, {"name": "B", "price": "2"}]},
        )
        doc = _document(session, str(result["doc_id"]))
        # Two cloned rows (the row is cloned, not the paragraph).
        assert doc.count("<w:tr>") == 2
        assert result["loops_expanded"] == {"rows": 2}

    def test_loop_render_once_on_truthy(self, tmp_path) -> None:  # noqa: ANN001
        body = (
            _p(_r("{{#show}}"))
            + _p(_r("Visible {{label}}"))
            + _p(_r("{{/show}}"))
        )
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(
            session, template=path, data={"show": True, "label": "now"}
        )
        out = docx_convert(session, doc_id=str(result["doc_id"]), to="md")["content"]
        assert "Visible now" in out
        assert result["loops_expanded"] == {"show": 1}

    def test_loop_dot_scalar(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("{{#tags}}")) + _p(_r("#{{.}}")) + _p(_r("{{/tags}}"))
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(session, template=path, data={"tags": ["x", "y"]})
        out = docx_convert(session, doc_id=str(result["doc_id"]), to="md")["content"]
        assert "#x" in out and "#y" in out


# ---------------------------------------------------------------------------
# Inverted sections
# ---------------------------------------------------------------------------


class TestInverted:
    def test_inverted_renders_when_falsy(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("{{^waived}}")) + _p(_r("Not waived.")) + _p(_r("{{/waived}}"))
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(session, template=path, data={"waived": False})
        out = docx_convert(session, doc_id=str(result["doc_id"]), to="md")["content"]
        assert "Not waived." in out
        # Inverted sections are conditions, never in loops_expanded.
        assert result["loops_expanded"] == {}

    def test_inverted_hidden_when_truthy(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("before")) + _p(_r("{{^waived}}")) + _p(_r("Not waived.")) + _p(
            _r("{{/waived}}")
        ) + _p(_r("after"))
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(session, template=path, data={"waived": True})
        out = docx_convert(session, doc_id=str(result["doc_id"]), to="md")["content"]
        assert "Not waived." not in out
        assert "before" in out and "after" in out


# ---------------------------------------------------------------------------
# Unfilled reporting
# ---------------------------------------------------------------------------


class TestUnfilled:
    def test_missing_var_left_verbatim_and_listed(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("A: {{a}} B: {{b}} A2: {{a}}"))
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(session, template=path, data={"a": "1"})
        doc = _document(session, str(result["doc_id"]))
        assert "{{b}}" in doc  # left verbatim
        # dedup, document order; only b unfilled.
        assert result["unfilled"] == ["b"]
        assert result["filled"] == 2  # both {{a}} resolved

    def test_strict_raises(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("{{missing}}"))
        path = _build_template(body, tmp_path)
        session = Session()
        with pytest.raises(ToolError) as exc:
            docx_template_fill(session, template=path, data={}, strict=True)
        assert exc.value.code == "placeholder_unfilled"

    def test_comment_tag_dropped(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("Hello{{! internal note }} World"))
        path = _build_template(body, tmp_path)
        session = Session()
        result = docx_template_fill(session, template=path, data={})
        doc = _document(session, str(result["doc_id"]))
        assert "internal note" not in doc
        assert "Hello World" in doc
        assert result["filled"] == 0


# ---------------------------------------------------------------------------
# Syntax / errors
# ---------------------------------------------------------------------------


class TestSyntax:
    def test_unsupported_syntax(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("x"))
        path = _build_template(body, tmp_path)
        session = Session()
        with pytest.raises(ToolError) as exc:
            docx_template_fill(session, template=path, data={}, syntax="jinja")
        assert exc.value.code == "template_syntax"

    def test_unclosed_section(self, tmp_path) -> None:  # noqa: ANN001
        body = _p(_r("{{#open}}")) + _p(_r("inner"))
        path = _build_template(body, tmp_path)
        session = Session()
        with pytest.raises(ToolError) as exc:
            docx_template_fill(session, template=path, data={"open": [{}]})
        assert exc.value.code == "template_syntax"
