"""Media tests: docx_media (algorithms.md §19).

Covers insert (media part + content-type Default + image rel + inline drawing run,
EMU sizing with PNG/JPEG aspect parsing), extract (copy bytes to a path), and
replace (overwrite the part keeping the rel/rId). A 1×1 px PNG and a tiny JPEG are
built inline so the suite needs no on-disk fixtures. The validator stays green.
"""

from __future__ import annotations

import base64
import re
import struct
import zlib

import pytest
from conftest import build_docx

from docxengine import (
    Session,
    ToolError,
    docx_media,
    docx_open,
    docx_validate,
    paragraph_anchor,
)

A1 = "P1#515a"
A2 = paragraph_anchor(2, "The term is five (5) years from the Effective Date.")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def _png(width: int, height: int) -> bytes:
    """A minimal valid PNG of the given pixel dimensions (truecolor)."""
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


#: A 1×1 px PNG — the canonical inline fixture.
PNG_1X1 = _png(1, 1)
#: A 200×100 px PNG to exercise aspect-ratio scaling.
PNG_200X100 = _png(200, 100)


def _jpeg(width: int, height: int) -> bytes:
    """A tiny JPEG carrying a SOF0 with the given dimensions (header only)."""
    soi = b"\xff\xd8"
    # APP0 segment (length 16).
    app0 = b"\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00"
    # SOF0: marker, length(17), precision(8), height(2), width(2), components(1)+9.
    sof = (
        b"\xff\xc0\x00\x11\x08"
        + struct.pack(">H", height)
        + struct.pack(">H", width)
        + b"\x03\x01\x22\x00\x02\x11\x01\x03\x11\x01"
    )
    eoi = b"\xff\xd9"
    return soi + app0 + sof + eoi


def open_docx() -> tuple[Session, str]:
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx()).decode())
    return session, str(result["doc_id"])


def write_image(tmp_path, name: str, data: bytes) -> str:
    path = tmp_path / name
    path.write_bytes(data)
    return str(path)


def main_xml(session: Session, doc_id: str) -> str:
    package = session.get(doc_id).package
    return package.part(package.main_document_part()).decode("utf-8")


def extent(session: Session, doc_id: str) -> tuple[int, int]:
    md = main_xml(session, doc_id)
    m = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"/>', md)
    assert m is not None
    return int(m.group(1)), int(m.group(2))


# ---------------------------------------------------------------------------
# insert (§19)
# ---------------------------------------------------------------------------


class TestInsert:
    def test_insert_writes_part_content_type_rel_and_drawing(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "logo.png", PNG_1X1)
        res = docx_media(session, doc_id=doc_id, op="insert", after=A1, image=img, width_cm=2)
        assert res["media_id"] == "M1"
        package = session.get(doc_id).package
        assert package.has_part("word/media/image1.png")
        assert package.part("word/media/image1.png") == PNG_1X1
        assert "png" in package.content_types().defaults
        rels = package.rels(package.main_document_part())
        rel_types = {r.rel_type.rsplit("/", 1)[-1] for r in rels}
        assert "image" in rel_types
        md = main_xml(session, doc_id)
        assert "<w:drawing>" in md
        assert "<a:blip r:embed=" in md

    def test_both_dims_give_exact_emu(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "i.png", PNG_1X1)
        docx_media(
            session, doc_id=doc_id, op="insert", after=A1,
            image=img, width_cm=3, height_cm=2,
        )
        # EMU = round(cm × 360000): 3 cm → 1080000, 2 cm → 720000.
        assert extent(session, doc_id) == (1080000, 720000)

    def test_width_only_scales_by_png_aspect(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "wide.png", PNG_200X100)
        docx_media(session, doc_id=doc_id, op="insert", after=A1, image=img, width_cm=4)
        cx, cy = extent(session, doc_id)
        # 4 cm → 1440000 EMU; aspect 200:100 → cy = round(1440000 × 100/200).
        assert cx == 1440000
        assert cy == 720000

    def test_height_only_scales_by_jpeg_aspect(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "photo.jpg", _jpeg(200, 100))
        docx_media(session, doc_id=doc_id, op="insert", after=A1, image=img, height_cm=2)
        cx, cy = extent(session, doc_id)
        # 2 cm → 720000 EMU; aspect 200:100 → cx = round(720000 × 200/100).
        assert cy == 720000
        assert cx == 1440000

    def test_insert_before_positions_drawing_first(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "i.png", PNG_1X1)
        docx_media(session, doc_id=doc_id, op="insert", before=A2, image=img, width_cm=1)
        md = main_xml(session, doc_id)
        assert md.index("<w:drawing>") < md.index("The term is")

    def test_insert_requires_exactly_one_of_after_before(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "i.png", PNG_1X1)
        with pytest.raises(ToolError) as err:
            docx_media(session, doc_id=doc_id, op="insert", after=A1, before=A2, image=img)
        assert err.value.code == "anchor_invalid"

    def test_missing_image_is_not_found(self, tmp_path) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_media(
                session, doc_id=doc_id, op="insert", after=A1,
                image=str(tmp_path / "nope.png"), width_cm=1,
            )
        assert err.value.code == "not_found"

    def test_insert_keeps_document_valid(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "i.png", PNG_1X1)
        docx_media(session, doc_id=doc_id, op="insert", after=A1, image=img, width_cm=2)
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# extract / replace (§19)
# ---------------------------------------------------------------------------


class TestExtractReplace:
    def test_extract_writes_bytes_to_path(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "i.png", PNG_1X1)
        docx_media(session, doc_id=doc_id, op="insert", after=A1, image=img, width_cm=1)
        out = tmp_path / "out.png"
        res = docx_media(session, doc_id=doc_id, op="extract", media_id="M1", path=str(out))
        assert res["path"] == str(out)
        assert out.read_bytes() == PNG_1X1

    def test_extract_unknown_media_is_not_found(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_media(session, doc_id=doc_id, op="extract", media_id="M5", path="/tmp/x.png")
        assert err.value.code == "not_found"

    def test_replace_overwrites_bytes_keeping_rel(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "i.png", PNG_1X1)
        docx_media(session, doc_id=doc_id, op="insert", after=A1, image=img, width_cm=1)
        package = session.get(doc_id).package
        rels_before = [r.rel_id for r in package.rels(package.main_document_part())]
        new_img = write_image(tmp_path, "new.png", PNG_200X100)
        res = docx_media(session, doc_id=doc_id, op="replace", media_id="M1", image=new_img)
        assert res["media_id"] == "M1"
        assert package.part("word/media/image1.png") == PNG_200X100
        rels_after = [r.rel_id for r in package.rels(package.main_document_part())]
        assert rels_before == rels_after  # rId preserved

    def test_replace_with_new_extension_adds_content_type(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "i.png", PNG_1X1)
        docx_media(session, doc_id=doc_id, op="insert", after=A1, image=img, width_cm=1)
        jpg = write_image(tmp_path, "p.jpeg", _jpeg(10, 10))
        docx_media(session, doc_id=doc_id, op="replace", media_id="M1", image=jpg)
        defaults = session.get(doc_id).package.content_types().defaults
        assert defaults.get("jpeg") == "image/jpeg"

    def test_extract_then_validate(self, tmp_path) -> None:
        session, doc_id = open_docx()
        img = write_image(tmp_path, "i.png", PNG_1X1)
        docx_media(session, doc_id=doc_id, op="insert", after=A1, image=img, width_cm=1)
        docx_media(session, doc_id=doc_id, op="replace", media_id="M1", image=img)
        assert docx_validate(session, doc_id=doc_id)["valid"] is True
