"""Comment tests: docx_comment (algorithms.md §18).

Covers add (validates a clean anchor + wires all five places: range start/end,
the reference run, the comments.xml definition, the CommentReference style),
reply/resolve via w15 commentsExtended, list (thread roots with replies and
resolved state), and delete (removes all five places, including replies). The
validator must stay green on every produced document.
"""

from __future__ import annotations

import base64

import pytest
from conftest import build_docx

from docxengine import (
    Session,
    ToolError,
    docx_comment,
    docx_open,
    docx_validate,
    paragraph_anchor,
)

A1 = "P1#515a"
A2 = paragraph_anchor(2, "The term is five (5) years from the Effective Date.")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCXENGINE_FIXED_DATE", "2026-06-11T00:00:00Z")
    monkeypatch.delenv("DOCXENGINE_AUTHOR", raising=False)


def open_docx(parts: dict[str, str] | None = None) -> tuple[Session, str]:
    session = Session()
    result = docx_open(session, bytes=base64.b64encode(build_docx(parts)).decode())
    return session, str(result["doc_id"])


def part(session: Session, doc_id: str, name: str) -> str:
    return session.get(doc_id).package.part(name).decode("utf-8")


def main_xml(session: Session, doc_id: str) -> str:
    package = session.get(doc_id).package
    return package.part(package.main_document_part()).decode("utf-8")


# ---------------------------------------------------------------------------
# add — five-place wiring (§18)
# ---------------------------------------------------------------------------


class TestAdd:
    def test_add_creates_comments_part_with_content_type_and_rel(self) -> None:
        session, doc_id = open_docx()
        docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="Mutual?", author="J.Doe")
        package = session.get(doc_id).package
        assert package.has_part("word/comments.xml")
        assert "word/comments.xml" in package.content_types().overrides
        rels = package.rels(package.main_document_part())
        rel_types = {r.rel_type.rsplit("/", 1)[-1] for r in rels}
        assert "comments" in rel_types

    def test_add_wires_all_five_places(self) -> None:
        session, doc_id = open_docx()
        res = docx_comment(
            session, doc_id=doc_id, op="add", anchor=A1, text="Should this be mutual?",
            author="Jane Q. Doe",
        )
        assert res["comment_id"] == "C0"
        assert res["anchor"] == A1
        md = main_xml(session, doc_id)
        # (1) range start, (2) range end, (3) the reference run.
        assert '<w:commentRangeStart w:id="0"/>' in md
        assert '<w:commentRangeEnd w:id="0"/>' in md
        assert '<w:rStyle w:val="CommentReference"/>' in md
        assert '<w:commentReference w:id="0"/>' in md
        # (4) the comments.xml definition, with author/date/initials.
        comments = part(session, doc_id, "word/comments.xml")
        assert 'w:id="0"' in comments
        assert 'w:author="Jane Q. Doe"' in comments
        assert 'w:date="2026-06-11T00:00:00Z"' in comments
        assert 'w:initials="JQD"' in comments
        assert "Should this be mutual?" in comments
        # (5) the ensured CommentReference style.
        styles = part(session, doc_id, "word/styles.xml")
        assert 'w:styleId="CommentReference"' in styles

    def test_add_range_start_after_ppr_before_runs(self) -> None:
        session, doc_id = open_docx()
        docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="x", author="A")
        md = main_xml(session, doc_id)
        # The range start follows the paragraph's w:pPr (so the comment brackets
        # only the runs), not the pPr itself.
        ppr_end = md.index("</w:pPr>") + len("</w:pPr>")
        assert md[ppr_end:].startswith('<w:commentRangeStart w:id="0"/>')

    def test_add_id_is_max_plus_one(self) -> None:
        session, doc_id = open_docx()
        first = docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="a", author="A")
        second = docx_comment(session, doc_id=doc_id, op="add", anchor=A2, text="b", author="B")
        assert first["comment_id"] == "C0"
        assert second["comment_id"] == "C1"

    def test_add_validates_stale_anchor(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_comment(session, doc_id=doc_id, op="add", anchor="P1#0000", text="x", author="A")
        assert err.value.code == "anchor_stale"

    def test_add_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="x", author="A")
        assert docx_validate(session, doc_id=doc_id)["valid"] is True

    def test_empty_author_yields_empty_initials(self) -> None:
        session, doc_id = open_docx()
        docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="x", author="")
        comments = part(session, doc_id, "word/comments.xml")
        assert 'w:initials=""' in comments


# ---------------------------------------------------------------------------
# list (§18)
# ---------------------------------------------------------------------------


class TestList:
    def test_list_returns_thread_root_with_anchor(self) -> None:
        session, doc_id = open_docx()
        docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="root", author="Jane")
        result = docx_comment(session, doc_id=doc_id, op="list")
        comments = result["comments"]
        assert isinstance(comments, list) and len(comments) == 1
        entry = comments[0]
        assert entry["id"] == "C0"
        assert entry["anchor"] == A1
        assert entry["author"] == "Jane"
        assert entry["text"] == "root"
        assert entry["resolved"] is False
        assert entry["replies"] == []

    def test_list_empty_when_no_comments(self) -> None:
        session, doc_id = open_docx()
        assert docx_comment(session, doc_id=doc_id, op="list") == {"comments": []}


# ---------------------------------------------------------------------------
# reply / resolve via w15 (§18)
# ---------------------------------------------------------------------------


class TestReplyResolve:
    def test_reply_appends_comment_and_w15_commentex(self) -> None:
        session, doc_id = open_docx()
        root = docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="root", author="A")
        reply = docx_comment(
            session, doc_id=doc_id, op="reply", comment_id=root["comment_id"], text="re", author="B"
        )
        assert reply["comment_id"] == "C1"
        package = session.get(doc_id).package
        assert package.has_part("word/commentsExtended.xml")
        ext = part(session, doc_id, "word/commentsExtended.xml")
        assert "w15:commentEx" in ext
        assert "w15:paraIdParent" in ext
        # The reply shows up nested under the root in list.
        entry = docx_comment(session, doc_id=doc_id, op="list")["comments"][0]  # type: ignore[index]
        assert entry["replies"] == [{"author": "B", "date": "2026-06-11T00:00:00Z", "text": "re"}]

    def test_reply_to_unknown_comment_is_not_found(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_comment(session, doc_id=doc_id, op="reply", comment_id="C9", text="x", author="A")
        assert err.value.code == "not_found"

    def test_resolve_sets_w15_done(self) -> None:
        session, doc_id = open_docx()
        root = docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="root", author="A")
        docx_comment(session, doc_id=doc_id, op="resolve", comment_id=root["comment_id"])
        ext = part(session, doc_id, "word/commentsExtended.xml")
        assert 'w15:done="1"' in ext
        entry = docx_comment(session, doc_id=doc_id, op="list")["comments"][0]  # type: ignore[index]
        assert entry["resolved"] is True

    def test_reply_and_resolve_keep_document_valid(self) -> None:
        session, doc_id = open_docx()
        root = docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="root", author="A")
        docx_comment(
            session, doc_id=doc_id, op="reply",
            comment_id=root["comment_id"], text="r", author="B",
        )
        docx_comment(session, doc_id=doc_id, op="resolve", comment_id=root["comment_id"])
        assert docx_validate(session, doc_id=doc_id)["valid"] is True


# ---------------------------------------------------------------------------
# delete (§18)
# ---------------------------------------------------------------------------


class TestDelete:
    def test_delete_removes_all_five_places(self) -> None:
        session, doc_id = open_docx()
        res = docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="x", author="A")
        docx_comment(session, doc_id=doc_id, op="delete", comment_id=res["comment_id"])
        md = main_xml(session, doc_id)
        assert "commentRangeStart" not in md
        assert "commentRangeEnd" not in md
        assert "commentReference" not in md
        comments = part(session, doc_id, "word/comments.xml")
        assert "<w:comment " not in comments
        assert docx_comment(session, doc_id=doc_id, op="list") == {"comments": []}

    def test_delete_removes_replies_too(self) -> None:
        session, doc_id = open_docx()
        root = docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="root", author="A")
        docx_comment(
            session, doc_id=doc_id, op="reply",
            comment_id=root["comment_id"], text="r", author="B",
        )
        docx_comment(session, doc_id=doc_id, op="delete", comment_id=root["comment_id"])
        comments = part(session, doc_id, "word/comments.xml")
        assert "<w:comment " not in comments
        ext = part(session, doc_id, "word/commentsExtended.xml")
        assert "w15:commentEx" not in ext

    def test_delete_unknown_is_not_found(self) -> None:
        session, doc_id = open_docx()
        with pytest.raises(ToolError) as err:
            docx_comment(session, doc_id=doc_id, op="delete", comment_id="C9")
        assert err.value.code == "not_found"

    def test_delete_keeps_document_valid(self) -> None:
        session, doc_id = open_docx()
        first = docx_comment(session, doc_id=doc_id, op="add", anchor=A1, text="a", author="A")
        docx_comment(session, doc_id=doc_id, op="add", anchor=A2, text="b", author="B")
        docx_comment(session, doc_id=doc_id, op="delete", comment_id=first["comment_id"])
        assert docx_validate(session, doc_id=doc_id)["valid"] is True
