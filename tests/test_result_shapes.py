"""Result-shape sweep: every Phase 2 tool's real result matches its
``spec/tools/*.json`` ``result_schema``.

A tiny dependency-free validator (``_validate``) walks the subset of JSON Schema
the contracts use — ``type``, ``required``, ``properties``, ``items``,
``enum``, ``additionalProperties`` — and asserts required keys are present with
the declared types. Real results are produced by driving each tool the way the
behavioral docs do, so this guards the wire contract end to end.
"""

from __future__ import annotations

import base64
import zlib
from typing import Any

import pytest
from conftest import build_docx

from docxengine import (
    Session,
    _spec,
    docx_comment,
    docx_convert,
    docx_create,
    docx_field,
    docx_format,
    docx_list,
    docx_media,
    docx_open,
    docx_render_preview,
    docx_section,
    docx_style,
    docx_table,
    docx_template_fill,
    paragraph_anchor,
)

A1 = "P1#515a"
A2 = paragraph_anchor(2, "The term is five (5) years from the Effective Date.")


# ---------------------------------------------------------------------------
# A minimal, dependency-free JSON-Schema validator (the subset the specs use)
# ---------------------------------------------------------------------------

_TYPE_CHECKS = {
    "object": lambda v: isinstance(v, dict),
    "array": lambda v: isinstance(v, list),
    "string": lambda v: isinstance(v, str),
    "integer": lambda v: isinstance(v, int) and not isinstance(v, bool),
    "number": lambda v: isinstance(v, (int, float)) and not isinstance(v, bool),
    "boolean": lambda v: isinstance(v, bool),
    "null": lambda v: v is None,
}


def _validate(value: Any, schema: dict[str, Any], path: str = "<result>") -> list[str]:
    """Return a list of human-readable violations (empty == valid)."""
    errors: list[str] = []
    expected = schema.get("type")
    if isinstance(expected, str):
        check = _TYPE_CHECKS.get(expected)
        if check is not None and not check(value):
            errors.append(f"{path}: expected type {expected!r}, got {type(value).__name__}")
            return errors  # type mismatch: deeper checks would be noise
    enum = schema.get("enum")
    if isinstance(enum, list) and value not in enum:
        errors.append(f"{path}: {value!r} not in enum {enum}")
    if isinstance(value, dict):
        for key in schema.get("required", []):
            if key not in value:
                errors.append(f"{path}: missing required key {key!r}")
        properties = schema.get("properties", {})
        for key, val in value.items():
            sub = properties.get(key)
            if isinstance(sub, dict):
                errors += _validate(val, sub, f"{path}.{key}")
            else:
                additional = schema.get("additionalProperties")
                if isinstance(additional, dict):
                    errors += _validate(val, additional, f"{path}.{key}")
    if isinstance(value, list):
        items = schema.get("items")
        if isinstance(items, dict):
            for i, item in enumerate(value):
                errors += _validate(item, items, f"{path}[{i}]")
    return errors


def assert_shape(tool: str, result: dict[str, object]) -> None:
    schema = _spec.result_schema(tool)
    assert schema is not None, f"{tool} has no result_schema in the packaged spec"
    violations = _validate(result, schema)
    assert not violations, f"{tool} result violates its schema:\n" + "\n".join(violations)


def test_validator_catches_violations() -> None:
    """The validator helper itself: type, required, enum, nested array items."""
    schema = {
        "type": "object",
        "required": ["n"],
        "properties": {
            "n": {"type": "integer"},
            "tags": {"type": "array", "items": {"type": "string"}},
            "kind": {"type": "string", "enum": ["a", "b"]},
        },
    }
    assert _validate({"n": 1, "tags": ["x"], "kind": "a"}, schema) == []
    assert _validate({}, schema) == ["<result>: missing required key 'n'"]
    assert _validate({"n": "x"}, schema)  # wrong type
    assert _validate({"n": 1, "tags": [1]}, schema)  # bad array item
    assert _validate({"n": 1, "kind": "z"}, schema)  # bad enum
    # booleans are not integers
    assert _validate({"n": True}, schema)


# ---------------------------------------------------------------------------
# Real-result generators per Phase 2 tool
# ---------------------------------------------------------------------------


def _png(width: int, height: int) -> bytes:
    import struct

    sig = b"\x89PNG\r\n\x1a\n"

    def chunk(typ: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + typ
            + data
            + struct.pack(">I", zlib.crc32(typ + data) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    idat = zlib.compress(b"\x00" + b"\x00\x00\x00" * width)
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)
    monkeypatch.delenv("DOCXENGINE_SOFFICE", raising=False)


@pytest.fixture
def opened() -> tuple[Session, str]:
    session = Session()
    doc_id = str(docx_open(session, bytes=base64.b64encode(build_docx()).decode())["doc_id"])
    return session, doc_id


def test_docx_table_shape(opened: tuple[Session, str]) -> None:
    session, doc_id = opened
    result = docx_table(session, doc_id=doc_id, op="create", after=A2, rows=2, cols=2)
    assert_shape("docx_table", result)
    assert "new_anchor" in result


def test_docx_style_list_and_define_shapes(opened: tuple[Session, str]) -> None:
    session, doc_id = opened
    listed = docx_style(session, doc_id=doc_id, op="list")
    assert_shape("docx_style", listed)
    assert isinstance(listed["styles"], list)
    defined = docx_style(session, doc_id=doc_id, op="define", name="Clause")
    assert_shape("docx_style", defined)


def test_docx_format_shape(opened: tuple[Session, str]) -> None:
    session, doc_id = opened
    result = docx_format(session, doc_id=doc_id, anchor=A2, props={"bold": True})
    assert_shape("docx_format", result)
    assert isinstance(result["affected"], int)


def test_docx_list_shape(opened: tuple[Session, str]) -> None:
    session, doc_id = opened
    result = docx_list(
        session, doc_id=doc_id, op="create", after=A2, kind="ol", items=[{"text": "First"}]
    )
    assert_shape("docx_list", result)


def test_docx_section_list_and_set_shapes(opened: tuple[Session, str]) -> None:
    session, doc_id = opened
    listed = docx_section(session, doc_id=doc_id, op="list")
    assert_shape("docx_section", listed)
    assert isinstance(listed["sections"], list)
    changed = docx_section(session, doc_id=doc_id, op="set_geometry", section="S1", columns=2)
    assert_shape("docx_section", changed)


def test_docx_comment_add_and_list_shapes(opened: tuple[Session, str]) -> None:
    session, doc_id = opened
    added = docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="x", author="A")
    assert_shape("docx_comment", added)
    listed = docx_comment(session, doc_id=doc_id, op="list")
    assert_shape("docx_comment", listed)


def test_docx_media_shape(opened: tuple[Session, str], tmp_path: Any) -> None:
    session, doc_id = opened
    img_path = tmp_path / "logo.png"
    img_path.write_bytes(_png(1, 1))
    result = docx_media(
        session, doc_id=doc_id, op="insert", after=A1, image=str(img_path), width_cm=2
    )
    assert_shape("docx_media", result)


def test_docx_field_insert_and_update_shapes(opened: tuple[Session, str]) -> None:
    session, doc_id = opened
    inserted = docx_field(session, doc_id=doc_id, op="insert_toc", after=A1, levels=3)
    assert_shape("docx_field", inserted)
    updated = docx_field(session, doc_id=doc_id, op="update")
    assert_shape("docx_field", updated)


def test_docx_template_fill_shape(tmp_path: Any) -> None:
    import io
    import zipfile

    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = (
        '<w:p><w:r><w:t xml:space="preserve">Hi {{name}}</w:t></w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
    )
    doc = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<w:document xmlns:w="{w}"><w:body>{body}</w:body></w:document>'
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

    session = Session()
    result = docx_template_fill(session, template=str(path), data={"name": "Ada"})
    assert_shape("docx_template_fill", result)


def test_docx_create_shape() -> None:
    session = Session()
    result = docx_create(session, content_md="# Title\n\nBody.")
    assert_shape("docx_create", result)


def test_docx_convert_md_and_html_shapes() -> None:
    session = Session()
    doc_id = str(docx_create(session, content_md="# Title\n\nBody.")["doc_id"])
    md = docx_convert(session, doc_id=doc_id, to="md")
    assert_shape("docx_convert", md)
    assert isinstance(md["content"], str)
    html = docx_convert(session, doc_id=doc_id, to="html")
    assert_shape("docx_convert", html)


def test_docx_render_preview_structural_shape() -> None:
    # LibreOffice is not installed here, so this exercises the structural fallback.
    session = Session()
    doc_id = str(docx_create(session, content_md="# Title\n\nBody.")["doc_id"])
    result = docx_render_preview(session, doc_id=doc_id)
    assert_shape("docx_render_preview", result)
    assert result["renderer"] == "structural"
