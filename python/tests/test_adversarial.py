"""Adversarial / fuzz tests for the §27 hardening (SECURITY.md threat model).

Every defense the engine claims against hostile input is exercised here: zip
bombs (part count, total size, per-part size, compression ratio), hostile XML
(DTD/entity declarations → XXE / billion-laughs), pathological XML nesting, and
path-traversal normalization. The TypeScript suite mirrors these
(`js/test/adversarial.test.ts`).
"""

from __future__ import annotations

import io
import zipfile

import pytest
from conftest import FIXTURE_PARTS, build_docx, document_xml

from docxengine import Package, ToolError, resolve_rel_target
from docxengine._xml import iter_elements


def _docx_with(part_overrides: dict[str, str]) -> bytes:
    parts = dict(FIXTURE_PARTS)
    parts.update(part_overrides)
    return build_docx(parts)


class TestDecompressionBombs:
    def test_part_count_cap(self, monkeypatch: pytest.MonkeyPatch) -> None:
        # The standard fixture has 5 parts; cap below that.
        monkeypatch.setenv("DOCXENGINE_MAX_PARTS", "3")
        with pytest.raises(ToolError) as err:
            Package.open(build_docx())
        assert err.value.code == "doc_too_large"

    def test_total_uncompressed_cap(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_MAX_TOTAL_BYTES", "100")
        with pytest.raises(ToolError) as err:
            Package.open(build_docx())
        assert err.value.code == "doc_too_large"

    def test_per_part_uncompressed_cap(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_MAX_PART_BYTES", "50")
        with pytest.raises(ToolError) as err:
            Package.open(build_docx())  # document.xml alone is well over 50 bytes
        assert err.value.code == "doc_too_large"

    def test_compression_ratio_cap_catches_zip_bomb(self) -> None:
        # A part above the 64 KiB ratio floor that deflates far past the 200:1 cap.
        bomb = "A" * (256 * 1024)
        with pytest.raises(ToolError) as err:
            Package.open(_docx_with({"word/bomb.xml": bomb}))
        assert err.value.code == "doc_too_large"
        assert "ratio" in err.value.message.lower() or "zip bomb" in err.value.message.lower()

    def test_small_compressible_part_is_not_flagged(self) -> None:
        # Below the 64 KiB ratio floor, a highly compressible part is fine.
        ok = "A" * (4 * 1024)
        pkg = Package.open(_docx_with({"word/note.xml": ok}))
        assert pkg.part("word/note.xml") == ok.encode("utf-8")

    def test_limits_relax_when_raised(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_MAX_PARTS", "1")  # would normally refuse
        monkeypatch.setenv("DOCXENGINE_MAX_PARTS", "100")  # last write wins
        Package.open(build_docx())  # no raise

    def test_invalid_env_falls_back_to_default(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_MAX_PARTS", "not-a-number")
        Package.open(build_docx())  # default (10000) applies, no raise


class TestHostileXml:
    def _billion_laughs(self) -> str:
        return (
            '<?xml version="1.0"?>\r\n'
            "<!DOCTYPE lolz [<!ENTITY lol \"lol\">"
            '<!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">]>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body><w:p><w:r><w:t>&lol2;</w:t></w:r></w:p></w:body></w:document>"
        )

    def test_doctype_in_document_rejected(self) -> None:
        pkg = Package.open(_docx_with({"word/document.xml": self._billion_laughs()}))
        with pytest.raises(ToolError) as err:
            pkg.part("word/document.xml")
        assert err.value.code == "malicious_content"

    def test_entity_declaration_rejected_before_dom_parse(self) -> None:
        # content_types() runs ET.fromstring; the chokepoint must fire first.
        evil_ct = (
            '<?xml version="1.0"?>\r\n<!DOCTYPE Types [<!ENTITY x "y">]>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/></Types>'
        )
        pkg = Package.open(_docx_with({"[Content_Types].xml": evil_ct}))
        with pytest.raises(ToolError) as err:
            pkg.content_types()
        assert err.value.code == "malicious_content"

    def test_external_entity_xxe_rejected(self) -> None:
        xxe = (
            '<?xml version="1.0"?>\r\n'
            '<!DOCTYPE r [<!ENTITY ext SYSTEM "file:///etc/passwd">]>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body><w:p><w:r><w:t>&ext;</w:t></w:r></w:body></w:document>"
        )
        pkg = Package.open(_docx_with({"word/document.xml": xxe}))
        with pytest.raises(ToolError) as err:
            pkg.part("word/document.xml")
        assert err.value.code == "malicious_content"

    def test_binary_part_with_entity_bytes_not_flagged(self) -> None:
        # A non-XML part containing the literal bytes is not screened (no false positive).
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for name, text in FIXTURE_PARTS.items():
                zf.writestr(name, text.encode("utf-8"))
            zf.writestr("word/media/image1.png", b"\x89PNG<!ENTITY not really xml>")
        pkg = Package.open(buf.getvalue())
        assert pkg.part("word/media/image1.png").startswith(b"\x89PNG")


class TestXmlNestingDepth:
    def test_deep_nesting_refused(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_MAX_XML_DEPTH", "50")
        nested = "<w:x>" * 200 + "</w:x>" * 200
        data = document_xml(f"<w:p>{nested}</w:p>").encode("utf-8")
        with pytest.raises(ToolError) as err:
            list(iter_elements(data))
        assert err.value.code == "doc_too_large"

    def test_normal_nesting_allowed(self) -> None:
        data = document_xml("<w:p><w:r><w:t>ok</w:t></w:r></w:p>").encode("utf-8")
        spans = list(iter_elements(data))
        assert any(s.name == "w:t" for s in spans)


class TestPathTraversal:
    def test_relative_traversal_is_normalized(self) -> None:
        resolved = resolve_rel_target("word/document.xml", "../docProps/core.xml")
        assert resolved == "docProps/core.xml"

    def test_excess_parent_segments_do_not_escape(self) -> None:
        # Even pathological `..` chains resolve to an in-package name, never an
        # absolute or parent-of-root path.
        resolved = resolve_rel_target("word/document.xml", "../../../../etc/passwd")
        assert not resolved.startswith("/")
        assert ".." not in resolved.split("/")

    def test_absolute_target_strips_leading_slash(self) -> None:
        assert resolve_rel_target("word/document.xml", "/word/styles.xml") == "word/styles.xml"


class TestMalformedXml:
    def test_truncated_document_errors_not_hangs(self) -> None:
        # An unterminated tag must raise, not loop forever or corrupt output.
        broken = (
            '<?xml version="1.0"?>\r\n'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body><w:p><w:r><w:t>unterminated"
        )
        data = broken.encode("utf-8")
        with pytest.raises(ValueError):
            list(iter_elements(data))

    def test_unbalanced_end_tag_errors(self) -> None:
        data = document_xml("<w:p></w:r></w:p>").encode("utf-8")
        with pytest.raises(ValueError):
            list(iter_elements(data))
