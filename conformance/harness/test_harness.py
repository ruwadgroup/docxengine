#!/usr/bin/env python3
"""Self-tests for the conformance harness (stdlib unittest, no CLIs needed).

Run: python3 -m unittest discover -s conformance/harness -p 'test_*.py'
"""

from __future__ import annotations

import json
import tempfile
import unittest
import zipfile
from pathlib import Path

import make_fixtures
import run as harness

REQUIRED_PARTS = [
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/styles.xml",
]


class FixtureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.tmp = tempfile.TemporaryDirectory()
        cls.corpus = Path(cls.tmp.name)
        cls.anchors = make_fixtures.build_corpus(cls.corpus, quiet=True)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.tmp.cleanup()

    def test_all_fixtures_have_required_parts(self) -> None:
        for name in make_fixtures.FIXTURES:
            docx = self.corpus / name / "input.docx"
            self.assertTrue(docx.is_file(), name)
            with zipfile.ZipFile(docx) as zf:
                self.assertIsNone(zf.testzip(), name)
                names = zf.namelist()
                for part in REQUIRED_PARTS:
                    self.assertIn(part, names, f"{name} missing {part}")

    def test_meta_json_present_with_producer_and_features(self) -> None:
        for name in make_fixtures.FIXTURES:
            meta = json.loads((self.corpus / name / "meta.json").read_text("utf-8"))
            self.assertIn("producer", meta, name)
            self.assertTrue(meta["features"], name)
            self.assertIn("anchors", meta, name)

    def test_spec_worked_example_anchor(self) -> None:
        # algorithms.md §1: "Master Services Agreement" -> 515a as 1st paragraph.
        self.assertEqual(self.anchors["minimal"]["P1"], "P1#515a")

    def test_empty_paragraph_hash_is_e3b0(self) -> None:
        self.assertEqual(make_fixtures.anchor_hash(make_fixtures.normalize_text("")), "e3b0")

    def test_normalize_collapses_pinned_whitespace_set(self) -> None:
        self.assertEqual(
            make_fixtures.normalize_text("  Master 　Services\t\nAgreement "),
            "Master Services Agreement",
        )

    def test_split_runs_concatenation_spans_runs(self) -> None:
        with zipfile.ZipFile(self.corpus / "split-runs" / "input.docx") as zf:
            doc = zf.read("word/document.xml").decode("utf-8")
        self.assertNotIn("five (5) years", doc)  # split across two runs
        anchors = make_fixtures.body_paragraph_anchors(doc.encode("utf-8"))
        self.assertEqual(anchors["P2"], self.anchors["split-runs"]["P2"])

    def test_corrupt_fixtures_carry_their_defects(self) -> None:
        with zipfile.ZipFile(self.corpus / "corrupt-orphan-rel" / "input.docx") as zf:
            rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")
            doc = zf.read("word/document.xml").decode("utf-8")
            self.assertIn("media/image1.png", rels)  # target missing from package
            self.assertNotIn("media/image1.png", zf.namelist())
            self.assertIn('r:id="rId8"', doc)  # dangling reference
            self.assertNotIn('Id="rId8"', rels)
        with zipfile.ZipFile(self.corpus / "corrupt-dup-ids" / "input.docx") as zf:
            doc = zf.read("word/document.xml").decode("utf-8")
            self.assertEqual(doc.count('w:id="5"'), 2)  # duplicate revision id

    def test_xml_space_preserve_emission(self) -> None:
        self.assertIn('xml:space="preserve"', make_fixtures.wt("trailing "))
        self.assertNotIn("xml:space", make_fixtures.wt("plain"))

    def test_phase2_fixtures_carry_their_parts(self) -> None:
        # tables: two body-level w:tbl, header shading, TableGrid style.
        with zipfile.ZipFile(self.corpus / "tables" / "input.docx") as zf:
            doc = zf.read("word/document.xml").decode("utf-8")
            styles = zf.read("word/styles.xml").decode("utf-8")
            self.assertEqual(doc.count("<w:tbl>"), 2)
            self.assertIn('w:fill="D9D9D9"', doc)
            self.assertIn('w:styleId="TableGrid"', styles)
            # P-sequence excludes table-cell paragraphs (spec §13).
            self.assertEqual(len(self.anchors["tables"]), 3)
        # numbered-lists: numbering.xml with ordered + bullet abstractNums.
        with zipfile.ZipFile(self.corpus / "numbered-lists" / "input.docx") as zf:
            self.assertIn("word/numbering.xml", zf.namelist())
            numbering = zf.read("word/numbering.xml").decode("utf-8")
            self.assertIn('w:numFmt w:val="decimal"', numbering)
            self.assertIn('w:numFmt w:val="lowerLetter"', numbering)
            self.assertIn('w:numFmt w:val="bullet"', numbering)
            ct = zf.read("[Content_Types].xml").decode("utf-8")
            self.assertIn("numbering+xml", ct)
        # headers-footers: a default headerReference to word/header1.xml.
        with zipfile.ZipFile(self.corpus / "headers-footers" / "input.docx") as zf:
            self.assertIn("word/header1.xml", zf.namelist())
            doc = zf.read("word/document.xml").decode("utf-8")
            self.assertIn('w:headerReference w:type="default"', doc)
        # template: a split-run placeholder and a loop section.
        with zipfile.ZipFile(self.corpus / "template" / "input.docx") as zf:
            doc = zf.read("word/document.xml").decode("utf-8")
            self.assertIn("{{Cli", doc)
            self.assertIn("ent}}", doc)  # {{Client}} split across two runs
            self.assertNotIn("{{Client}}", doc)
            self.assertIn("{{#obligations}}", doc)
            self.assertIn("{{/obligations}}", doc)
        # media-doc: a 69-byte 1x1 PNG behind an inline drawing.
        with zipfile.ZipFile(self.corpus / "media-doc" / "input.docx") as zf:
            png = zf.read("word/media/image1.png")
            self.assertEqual(png[:8], b"\x89PNG\r\n\x1a\n")
            self.assertEqual(png[16:24], b"\x00\x00\x00\x01\x00\x00\x00\x01")  # 1x1 IHDR
            doc = zf.read("word/document.xml").decode("utf-8")
            self.assertIn("<w:drawing>", doc)
            self.assertIn('r:embed="rId2"', doc)
            ct = zf.read("[Content_Types].xml").decode("utf-8")
            self.assertIn('Extension="png"', ct)

    def test_deterministic_bytes(self) -> None:
        with tempfile.TemporaryDirectory() as other:
            make_fixtures.build_corpus(Path(other), quiet=True)
            for name in make_fixtures.FIXTURES:
                a = (self.corpus / name / "input.docx").read_bytes()
                b = (Path(other) / name / "input.docx").read_bytes()
                self.assertEqual(a, b, name)


class CanonXmlTests(unittest.TestCase):
    def test_attribute_order_and_interelement_whitespace_normalize(self) -> None:
        a = b'<r b="2" a="1">\n  <c/>\n</r>'
        b = b'<r a="1" b="2"><c/></r>'
        self.assertEqual(harness.canon_xml(a), harness.canon_xml(b))

    def test_whitespace_only_leaf_text_is_kept(self) -> None:
        # A w:t holding a single space is semantic; only ws between elements drops.
        a = b"<p><t> </t></p>"
        b = b"<p><t></t></p>"
        self.assertNotEqual(harness.canon_xml(a), harness.canon_xml(b))

    def test_text_difference_fails(self) -> None:
        self.assertNotEqual(harness.canon_xml(b"<p>x</p>"), harness.canon_xml(b"<p>y</p>"))

    def test_escaping_is_canonical(self) -> None:
        self.assertEqual(
            harness.canon_xml(b"<p>a&amp;b</p>"), harness.canon_xml(b"<p><![CDATA[a&b]]></p>")
        )


class MatchingTests(unittest.TestCase):
    def test_dict_subset(self) -> None:
        self.assertEqual(harness.match_partial({"a": 1}, {"a": 1, "b": 2}), [])
        self.assertTrue(harness.match_partial({"a": 1}, {"a": 2}))
        self.assertTrue(harness.match_partial({"a": 1}, {}))

    def test_list_positional_same_length(self) -> None:
        self.assertEqual(harness.match_partial([{"x": 1}], [{"x": 1, "y": 2}]), [])
        self.assertTrue(harness.match_partial([], [1]))

    def test_contains_requires_distinct_elements(self) -> None:
        exp = ["$contains", {"severity": "error"}, {"severity": "error"}]
        two = [{"severity": "error"}, {"severity": "warning"}, {"severity": "error"}]
        one = [{"severity": "error"}, {"severity": "warning"}]
        self.assertEqual(harness.match_partial(exp, two), [])
        self.assertTrue(harness.match_partial(exp, one))

    def test_substr_matcher(self) -> None:
        exp = ["$substr", "# Title", "## Sub"]
        self.assertEqual(harness.match_partial(exp, "# Title\n\n## Sub\nbody"), [])
        self.assertTrue(harness.match_partial(exp, "# Title only"))
        # A non-string actual is a clear failure, not a silent pass.
        self.assertTrue(harness.match_partial(["$substr", "x"], ["x"]))

    def test_bool_not_confused_with_int(self) -> None:
        self.assertTrue(harness.match_partial(True, 1))
        self.assertEqual(harness.match_partial(True, True), [])

    def test_mask_volatile(self) -> None:
        # Masked keys are dropped (not placeholdered) so presence never matters.
        masked = harness.mask_volatile(
            {"doc_id": "d1", "bytes": 42, "note": "hi", "n": [{"doc_id": "x", "keep": 1}]}
        )
        self.assertEqual(masked, {"n": [{"keep": 1}]})

    def test_mask_volatile_ignores_optional_note(self) -> None:
        # An optional `note` from only one side must not break parity equality.
        with_note = {"new_anchor": "P2#8137", "note": "Applied."}
        without = {"new_anchor": "P2#8137"}
        self.assertEqual(harness.mask_volatile(with_note), harness.mask_volatile(without))


class PackageCompareTests(unittest.TestCase):
    def test_identical_package_compares_clean(self) -> None:
        corpus = Path(__file__).resolve().parent.parent / "corpus"
        docx = corpus / "minimal" / "input.docx"
        self.assertEqual(harness.compare_packages(docx, docx, "a", "b"), [])

    def test_xml_fallback_tolerates_attribute_order(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            a, b = Path(tmp) / "a.docx", Path(tmp) / "b.docx"
            make_fixtures.write_docx(a, [("x/part.xml", b'<r a="1" b="2"><c/></r>')])
            make_fixtures.write_docx(b, [("x/part.xml", b'<r b="2" a="1">\n<c/>\n</r>')])
            self.assertEqual(harness.compare_packages(a, b, "a", "b"), [])
            make_fixtures.write_docx(b, [("x/part.xml", b'<r a="1" b="3"><c/></r>')])
            self.assertTrue(harness.compare_packages(a, b, "a", "b"))

    def test_entry_order_mismatch_fails(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            a, b = Path(tmp) / "a.docx", Path(tmp) / "b.docx"
            parts = [("one.xml", b"<a></a>"), ("two.xml", b"<b></b>")]
            make_fixtures.write_docx(a, parts)
            make_fixtures.write_docx(b, parts[::-1])
            self.assertTrue(harness.compare_packages(a, b, "a", "b"))


class CaseFileTests(unittest.TestCase):
    """Every checked-in case must reference real fixtures and live anchors."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.tmp = tempfile.TemporaryDirectory()
        cls.anchors = make_fixtures.build_corpus(Path(cls.tmp.name), quiet=True)
        cases_dir = Path(__file__).resolve().parent.parent / "cases"
        cls.cases = {
            p.stem: json.loads(p.read_text("utf-8")) for p in sorted(cases_dir.glob("*.json"))
        }

    @classmethod
    def tearDownClass(cls) -> None:
        cls.tmp.cleanup()

    def test_cases_exist(self) -> None:
        self.assertGreaterEqual(len(self.cases), 12)

    def test_case_shape_and_doc_references(self) -> None:
        for name, case in self.cases.items():
            self.assertIn(case["doc"], make_fixtures.FIXTURES, name)
            self.assertTrue(case["tool"].startswith("docx_"), name)
            expect = case["expect"]
            self.assertTrue("result" in expect or "error" in expect, name)
            self.assertNotIn("doc_id", case.get("args", {}), name)  # harness injects

    def test_case_anchor_args_match_generated_fixtures(self) -> None:
        # Validate every P-anchor passed as `anchor`/`after` against the fixture.
        # Table/section/comment/media anchors (T*/S*/C*/M*) are checked by the
        # tool, not the generator, so they are skipped here.
        for name, case in self.cases.items():
            args = case.get("args", {})
            for key in ("anchor", "after"):
                anchor = args.get(key)
                if not isinstance(anchor, str) or "#" not in anchor:
                    continue
                ordinal = anchor.split("#")[0]
                if not ordinal.startswith("P") or not ordinal[1:].isdigit():
                    continue
                self.assertEqual(anchor, self.anchors[case["doc"]][ordinal], f"{name}/{key}")

    def test_setup_steps_are_well_formed(self) -> None:
        for name, case in self.cases.items():
            for step in case.get("setup", []):
                self.assertIn("tool", step, name)
                self.assertTrue(step["tool"].startswith("docx_"), name)
                self.assertNotIn("doc_id", step.get("args", {}), name)


if __name__ == "__main__":
    unittest.main()
