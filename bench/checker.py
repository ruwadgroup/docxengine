#!/usr/bin/env python3
"""Assertion engine for the agent task benchmark (stdlib + docxengine).

Each task in ``bench/tasks/*.json`` carries a ``checks`` list of element-level
ground-truth assertions. After the runner drives the MCP server and the task
saves its output document, :func:`run_checks` evaluates every check against that
saved ``.docx`` and returns the failures (empty list == task passed).

Checks read the document through docxengine's ``Package`` API — the same OPC
model the engine writes — so a check sees exactly the bytes Word would open. No
third-party dependencies: XML is parsed with ``xml.etree.ElementTree``.

Supported check types
---------------------
- ``doc_text_contains {text}``       coalesced body text contains ``text``.
- ``doc_text_absent {text}``         coalesced body text does not contain ``text``.
- ``paragraph_text {ordinal, equals}``  1-based body paragraph N reads exactly ``equals``.
- ``paragraph_count {n}``            the body has exactly ``n`` paragraphs.
- ``revision_count {author?, type_?, n}``  count of w:ins/w:del matching the
  optional author and type (``ins``/``del``) equals ``n``.
- ``validate_clean``                 docx_validate on the reopened doc reports valid.
- ``outline_contains {level, text}`` a heading at ``level`` reads exactly ``text``.
- ``style_color {style, hex}``       style ``style`` in styles.xml resolves to run
  color ``hex`` (case-insensitive, ``#`` optional).
- ``comment_count {author?, n}``     comments matching the optional author == ``n``.
- ``comment_text_contains {text}``   some comment body contains ``text`` (case-insensitive).
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Any

from docxengine import Package, call

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


# ---------------------------------------------------------------------------
# Document readers (operate on the saved output package)
# ---------------------------------------------------------------------------


def _root(pkg: Package, part: str) -> ET.Element | None:
    if part not in pkg.part_names:
        return None
    return ET.fromstring(pkg.part(part))


def _body_paragraphs(pkg: Package) -> list[ET.Element]:
    root = _root(pkg, "word/document.xml")
    if root is None:
        return []
    body = root.find(_w("body"))
    if body is None:
        return []
    return [child for child in body if child.tag == _w("p")]


def _paragraph_text(paragraph: ET.Element) -> str:
    """As-if-accepted text: w:t survives, w:delText (deleted runs) is dropped."""
    parts: list[str] = []
    for elem in paragraph.iter():
        if elem.tag == _w("t") and elem.text:
            parts.append(elem.text)
    return "".join(parts)


def _body_text(pkg: Package) -> str:
    """All story text in document order, including paragraphs inside table cells."""
    root = _root(pkg, "word/document.xml")
    if root is None:
        return ""
    body = root.find(_w("body"))
    if body is None:
        return ""
    return "\n".join(_paragraph_text(p) for p in body.iter(_w("p")))


def _revisions(pkg: Package) -> list[tuple[str, str]]:
    """Return (type, author) for every w:ins / w:del in document.xml."""
    root = _root(pkg, "word/document.xml")
    out: list[tuple[str, str]] = []
    if root is None:
        return out
    for elem in root.iter():
        if elem.tag == _w("ins"):
            out.append(("ins", elem.get(_w("author"), "")))
        elif elem.tag == _w("del"):
            out.append(("del", elem.get(_w("author"), "")))
    return out


def _comments(pkg: Package) -> list[tuple[str, str]]:
    """Return (author, body_text) for every top-level comment in comments.xml."""
    root = _root(pkg, "word/comments.xml")
    out: list[tuple[str, str]] = []
    if root is None:
        return out
    for comment in root.findall(_w("comment")):
        author = comment.get(_w("author"), "")
        text = "".join(t.text or "" for t in comment.iter(_w("t")))
        out.append((author, text))
    return out


def _style_color(pkg: Package, style_id: str) -> str | None:
    """Resolved run color (w:rPr/w:color@w:val) of a paragraph/character style."""
    root = _root(pkg, "word/styles.xml")
    if root is None:
        return None
    for style in root.findall(_w("style")):
        if style.get(_w("styleId")) != style_id:
            continue
        rpr = style.find(_w("rPr"))
        if rpr is None:
            return None
        color = rpr.find(_w("color"))
        if color is None:
            return None
        return color.get(_w("val"))
    return None


# ---------------------------------------------------------------------------
# Check evaluation
# ---------------------------------------------------------------------------


def _norm_hex(value: str) -> str:
    return value.lstrip("#").upper()


def evaluate_check(check: dict[str, Any], pkg: Package, doc_id: str) -> str | None:
    """Return a failure reason, or None if the check passes."""
    kind = check.get("type")

    if kind == "doc_text_contains":
        text = check["text"]
        return None if text in _body_text(pkg) else f"text not found: {text!r}"

    if kind == "doc_text_absent":
        text = check["text"]
        return f"text should be absent: {text!r}" if text in _body_text(pkg) else None

    if kind == "paragraph_text":
        ordinal = check["ordinal"]
        paras = _body_paragraphs(pkg)
        if ordinal < 1 or ordinal > len(paras):
            return f"paragraph {ordinal} out of range (have {len(paras)})"
        actual = _paragraph_text(paras[ordinal - 1])
        expected = check["equals"]
        return None if actual == expected else f"P{ordinal}: expected {expected!r}, got {actual!r}"

    if kind == "paragraph_count":
        actual = len(_body_paragraphs(pkg))
        expected = check["n"]
        return None if actual == expected else f"expected {expected} paragraphs, got {actual}"

    if kind == "revision_count":
        want_author = check.get("author")
        want_type = check.get("type_")
        revs = _revisions(pkg)
        matched = [
            (t, a)
            for (t, a) in revs
            if (want_author is None or a == want_author)
            and (want_type is None or t == want_type)
        ]
        expected = check["n"]
        if len(matched) == expected:
            return None
        sel = []
        if want_author is not None:
            sel.append(f"author={want_author!r}")
        if want_type is not None:
            sel.append(f"type={want_type!r}")
        selector = f" ({', '.join(sel)})" if sel else ""
        return f"expected {expected} revision(s){selector}, got {len(matched)}"

    if kind == "validate_clean":
        verdict = call("docx_validate", {"doc_id": doc_id})
        if verdict.get("valid") is True:
            return None
        return f"validate not clean: {verdict.get('issues')}"

    if kind == "outline_contains":
        level = check["level"]
        text = check["text"]
        outline = call("docx_outline", {"doc_id": doc_id}).get("outline", [])
        for entry in outline:
            if entry.get("level") == level and entry.get("text") == text:
                return None
        return f"no level-{level} heading reads {text!r}"

    if kind == "style_color":
        style = check["style"]
        want = _norm_hex(check["hex"])
        actual = _style_color(pkg, style)
        if actual is None:
            return f"style {style!r} has no run color"
        return None if _norm_hex(actual) == want else f"style {style!r} color {actual!r} != {want!r}"

    if kind == "comment_count":
        want_author = check.get("author")
        comments = _comments(pkg)
        matched = [c for c in comments if want_author is None or c[0] == want_author]
        expected = check["n"]
        if len(matched) == expected:
            return None
        sel = f" (author={want_author!r})" if want_author is not None else ""
        return f"expected {expected} comment(s){sel}, got {len(matched)}"

    if kind == "comment_text_contains":
        needle = check["text"].lower()
        for _, body in _comments(pkg):
            if needle in body.lower():
                return None
        return f"no comment contains {check['text']!r}"

    return f"unknown check type {kind!r}"


def run_checks(checks: list[dict[str, Any]], output_path: str) -> list[str]:
    """Open ``output_path`` and evaluate every check; return the failure reasons."""
    pkg = Package.open(output_path)
    opened = call("docx_open", {"path": output_path})
    doc_id = opened["doc_id"]
    failures: list[str] = []
    for check in checks:
        reason = evaluate_check(check, pkg, doc_id)
        if reason is not None:
            failures.append(reason)
    return failures
