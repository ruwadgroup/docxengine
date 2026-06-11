"""docx_revision tests: list/accept/reject with filters, run-merge post-pass (§7/§6a)."""

from __future__ import annotations

import base64
from pathlib import Path

import pytest
from conftest import FIXTURE_PARTS, SECT_PR, build_docx, document_xml

from docxengine import (
    Session,
    ToolError,
    docx_open,
    docx_read,
    docx_replace,
    docx_revision,
    paragraph_anchor,
)

FIXED_DATE = "2026-06-10T00:00:00Z"
DATE_ALICE = "2026-01-15T09:30:00Z"
DATE_BOB = "2026-02-20T16:45:00Z"

# Mirrors the conformance "redlines" fixture: Alice owns ids 1-2 (P2), Bob 3-4 (P3).
REDLINES_DOCUMENT = document_xml(
    "<w:p><w:r><w:t>Revision History</w:t></w:r></w:p>",
    (
        '<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r>'
        f'<w:del w:id="1" w:author="Alice" w:date="{DATE_ALICE}">'
        "<w:r><w:delText>ten percent</w:delText></w:r></w:del>"
        f'<w:ins w:id="2" w:author="Alice" w:date="{DATE_ALICE}">'
        "<w:r><w:t>twelve percent</w:t></w:r></w:ins>"
        '<w:r><w:t xml:space="preserve"> of net revenue.</w:t></w:r></w:p>'
    ),
    (
        '<w:p><w:r><w:t xml:space="preserve">Notices must be sent </w:t></w:r>'
        f'<w:ins w:id="3" w:author="Bob" w:date="{DATE_BOB}">'
        '<w:r><w:t xml:space="preserve">by certified mail </w:t></w:r></w:ins>'
        "<w:r><w:t>to the address below</w:t></w:r>"
        f'<w:del w:id="4" w:author="Bob" w:date="{DATE_BOB}">'
        '<w:r><w:delText xml:space="preserve"> within five days</w:delText></w:r></w:del>'
        "<w:r><w:t>.</w:t></w:r></w:p>"
    ),
    SECT_PR,
)

A1 = paragraph_anchor(1, "Revision History")
A2 = paragraph_anchor(2, "The fee is twelve percent of net revenue.")
A3 = paragraph_anchor(3, "Notices must be sent by certified mail to the address below.")
A2_REJECTED = paragraph_anchor(2, "The fee is ten percent of net revenue.")
A3_REJECTED = paragraph_anchor(3, "Notices must be sent to the address below within five days.")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCXENGINE_FIXED_DATE", raising=False)
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def open_docx(document: str = REDLINES_DOCUMENT) -> tuple[Session, str]:
    parts = dict(FIXTURE_PARTS)
    parts["word/document.xml"] = document
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())
    return session, str(result["doc_id"])


def part_bytes(session: Session, doc_id: str) -> bytes:
    package = session.get(doc_id).package
    return package.part(package.main_document_part())


class TestList:
    def test_list_all_in_document_order(self) -> None:
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="list")
        assert result == {
            "revisions": [
                {
                    "id": "R1",
                    "type": "del",
                    "author": "Alice",
                    "date": DATE_ALICE,
                    "anchor": A2,
                    "text": "ten percent",
                },
                {
                    "id": "R2",
                    "type": "ins",
                    "author": "Alice",
                    "date": DATE_ALICE,
                    "anchor": A2,
                    "text": "twelve percent",
                },
                {
                    "id": "R3",
                    "type": "ins",
                    "author": "Bob",
                    "date": DATE_BOB,
                    "anchor": A3,
                    "text": "by certified mail ",  # the wrapper's own raw text
                },
                {
                    "id": "R4",
                    "type": "del",
                    "author": "Bob",
                    "date": DATE_BOB,
                    "anchor": A3,
                    "text": " within five days",
                },
            ]
        }

    def test_list_author_filter(self) -> None:
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="list", filter={"author": "Bob"})
        assert [r["id"] for r in list(result["revisions"])] == ["R3", "R4"]

    def test_list_does_not_mark_dirty(self) -> None:
        session, doc_id = open_docx()
        docx_revision(session, doc_id=doc_id, op="list")
        assert session.get(doc_id).dirty is False

    def test_unknown_op_is_not_found(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as exc_info:
            docx_revision(session, doc_id=doc_id, op="merge")
        assert exc_info.value.code == "not_found"


class TestAcceptReject:
    def test_accept_author_filter_leaves_other_authors(self) -> None:
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="accept", filter={"author": "Alice"})
        assert result == {
            "accepted": 2,
            "remaining_by_author": {"Bob": 2},
            "anchors": [A2],  # the as-if-accepted hash is unchanged by accepting
        }
        data = part_bytes(session, doc_id)
        assert b"ten percent" not in data
        # Post-pass merged P2 into a single clean run.
        assert b"<w:r><w:t>The fee is twelve percent of net revenue.</w:t></w:r>" in data
        # Bob's revisions are byte-for-byte untouched.
        assert f'<w:ins w:id="3" w:author="Bob" w:date="{DATE_BOB}">'.encode() in data
        assert f'<w:del w:id="4" w:author="Bob" w:date="{DATE_BOB}">'.encode() in data
        assert session.get(doc_id).dirty is True

    def test_reject_all_restores_original_text(self) -> None:
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="reject_all")
        assert result == {
            "rejected": 4,
            "remaining_by_author": {},
            "anchors": [A2_REJECTED, A3_REJECTED],
        }
        data = part_bytes(session, doc_id)
        assert b"<w:ins" not in data and b"<w:del" not in data
        assert b"<w:r><w:t>The fee is ten percent of net revenue.</w:t></w:r>" in data
        assert (
            b"<w:r><w:t>Notices must be sent to the address below within five days.</w:t></w:r>"
            in data
        )
        content = str(docx_read(session, doc_id=doc_id)["content"])
        assert "twelve percent" not in content and "certified mail" not in content

    def test_accept_all_resolves_everything_and_is_idempotent(self) -> None:
        session, doc_id = open_docx()
        first = docx_revision(session, doc_id=doc_id, op="accept_all")
        assert first == {
            "accepted": 4,
            "remaining_by_author": {},
            "anchors": [A2, A3],
        }
        second = docx_revision(session, doc_id=doc_id, op="accept_all")
        assert second == {"accepted": 0, "remaining_by_author": {}, "anchors": []}
        content = str(docx_read(session, doc_id=doc_id)["content"])
        assert "The fee is twelve percent of net revenue." in content
        assert "Notices must be sent by certified mail to the address below." in content

    def test_accept_all_ignores_filter(self) -> None:
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="accept_all", filter={"author": "Alice"})
        assert result["accepted"] == 4

    def test_accept_single_id(self) -> None:
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="accept", id="R2")
        assert result == {
            "accepted": 1,
            "remaining_by_author": {"Alice": 1, "Bob": 2},
            "anchors": [A2],
        }
        data = part_bytes(session, doc_id)
        assert b'<w:ins w:id="2"' not in data
        assert b'<w:del w:id="1"' in data  # Alice's deletion is still pending

    def test_unknown_id_resolves_nothing(self) -> None:
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="accept", id="R99")
        assert result["accepted"] == 0
        assert result["remaining_by_author"] == {"Alice": 2, "Bob": 2}
        assert session.get(doc_id).dirty is False

    def test_reject_single_del_restores_delText_as_text(self) -> None:
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="reject", id="R4")
        assert result["rejected"] == 1
        assert b" within five days" in part_bytes(session, doc_id)
        content = str(docx_read(session, doc_id=doc_id)["content"])
        assert "to the address below within five days." in content

    def test_date_filters(self) -> None:
        # after: on or after the date; before: strictly before; date: prefix match.
        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="accept", filter={"after": "2026-02-01"})
        assert result["accepted"] == 2
        assert result["remaining_by_author"] == {"Alice": 2}

        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="accept", filter={"before": "2026-02-01"})
        assert result["accepted"] == 2
        assert result["remaining_by_author"] == {"Bob": 2}

        session, doc_id = open_docx()
        result = docx_revision(session, doc_id=doc_id, op="accept", filter={"date": "2026-01-15"})
        assert result["accepted"] == 2
        assert result["remaining_by_author"] == {"Bob": 2}


class TestRunMergePostPass:
    def test_merge_respects_rpr_differences(self) -> None:
        session, doc_id = open_docx(
            document_xml(
                (
                    '<w:p><w:r><w:t xml:space="preserve">plain </w:t></w:r>'
                    f'<w:ins w:id="1" w:author="Alice" w:date="{DATE_ALICE}">'
                    "<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:ins>"
                    '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> tail</w:t></w:r></w:p>'
                ),
                SECT_PR,
            )
        )
        docx_revision(session, doc_id=doc_id, op="accept_all")
        data = part_bytes(session, doc_id)
        assert (
            b'<w:r><w:t xml:space="preserve">plain </w:t></w:r>'
            b"<w:r><w:rPr><w:b/></w:rPr><w:t>bold tail</w:t></w:r>"
        ) in data

    def test_merge_is_rsid_blind(self) -> None:
        session, doc_id = open_docx(
            document_xml(
                (
                    '<w:p><w:r w:rsidR="00AA0001"><w:t>x</w:t></w:r>'
                    f'<w:ins w:id="1" w:author="Alice" w:date="{DATE_ALICE}">'
                    "<w:r><w:t>z</w:t></w:r></w:ins>"
                    '<w:r w:rsidR="00BB0002"><w:t>y</w:t></w:r></w:p>'
                ),
                SECT_PR,
            )
        )
        docx_revision(session, doc_id=doc_id, op="accept_all")
        assert b'<w:r w:rsidR="00AA0001"><w:t>xzy</w:t></w:r>' in part_bytes(session, doc_id)

    def test_merge_only_touches_affected_paragraphs(self) -> None:
        # P1 has two mergeable runs but no revisions; resolving P2 must not touch P1.
        session, doc_id = open_docx(
            document_xml(
                "<w:p><w:r><w:t>split</w:t></w:r><w:r><w:t> runs</w:t></w:r></w:p>",
                (
                    f'<w:p><w:ins w:id="1" w:author="Alice" w:date="{DATE_ALICE}">'
                    "<w:r><w:t>added</w:t></w:r></w:ins></w:p>"
                ),
                SECT_PR,
            )
        )
        docx_revision(session, doc_id=doc_id, op="accept_all")
        data = part_bytes(session, doc_id)
        assert b"<w:p><w:r><w:t>split</w:t></w:r><w:r><w:t> runs</w:t></w:r></w:p>" in data
        assert b"<w:p><w:r><w:t>added</w:t></w:r></w:p>" in data


class TestTrackedEditInterplay:
    def test_tracked_replace_then_accept_author(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DOCXENGINE_FIXED_DATE", FIXED_DATE)
        session, doc_id = open_docx()
        docx_replace(
            session,
            doc_id=doc_id,
            anchor=A1,
            old="Revision History",
            new="Change Log",
            track_changes=True,
            author="Claude",
        )
        listed = docx_revision(session, doc_id=doc_id, op="list", filter={"author": "Claude"})
        assert [r["id"] for r in list(listed["revisions"])] == ["R5", "R6"]
        result = docx_revision(session, doc_id=doc_id, op="accept", filter={"author": "Claude"})
        assert result["accepted"] == 2
        assert result["remaining_by_author"] == {"Alice": 2, "Bob": 2}
        assert result["anchors"] == [paragraph_anchor(1, "Change Log")]
        content = str(docx_read(session, doc_id=doc_id)["content"])
        assert content.startswith(f"[{paragraph_anchor(1, 'Change Log')}] Change Log")
        assert "[del by Claude]" not in content


# ---------------------------------------------------------------------------
# Conformance corpus parity (pins the cases in conformance/cases/)
# ---------------------------------------------------------------------------

CORPUS = Path(__file__).resolve().parents[2] / "conformance" / "corpus"


@pytest.mark.skipif(not CORPUS.is_dir(), reason="conformance corpus not present")
class TestConformanceParity:
    def open_corpus(self, name: str) -> tuple[Session, str]:
        session = Session()
        result = docx_open(session, path=str(CORPUS / name / "input.docx"))
        return session, str(result["doc_id"])

    def test_revision_accept_author(self) -> None:
        session, doc_id = self.open_corpus("redlines")
        result = docx_revision(session, doc_id=doc_id, op="accept", filter={"author": "Alice"})
        assert result["accepted"] == 2
        data = part_bytes(session, doc_id)
        assert b"<w:ins" in data  # Bob's revisions preserved

    def test_revision_reject_all(self) -> None:
        session, doc_id = self.open_corpus("redlines")
        result = docx_revision(session, doc_id=doc_id, op="reject_all")
        assert result["rejected"] == 4
        data = part_bytes(session, doc_id)
        assert b"<w:ins" not in data and b"<w:del" not in data
