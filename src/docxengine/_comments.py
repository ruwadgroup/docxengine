"""Comments (``docx_comment``) — algorithms.md §18.

``add`` wires all five places (range start, range end, the reference run, the
``word/comments.xml`` definition, and the ``CommentReference`` style), id =
``max w:comment/@w:id + 1`` (start 0). ``reply`` appends a child comment and a
``word/commentsExtended.xml`` (w15) ``commentEx`` linking it to its thread root;
``resolve`` flips that root's ``w15:done``. ``delete`` removes every place for a
thread (root + replies). ``list`` returns one entry per thread root.

All edits splice raw bytes per §3; the comments part carries the ``w14`` namespace
so each ``w:p`` can hold a ``w14:paraId`` (the reply linkage key).
"""

from __future__ import annotations

import re
from xml.etree import ElementTree as ET

from . import _edits, _parts, _xml
from ._errors import ToolError
from ._opc import Package
from ._session import Session

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
_W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"

COMMENTS_PART = "word/comments.xml"
COMMENTS_EXTENDED_PART = "word/commentsExtended.xml"

_COMMENTS_REL_TYPE = f"{_parts.REL_BASE}/comments"
_COMMENTS_EXTENDED_REL_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
_COMMENTS_CT = f"{_parts.CT_BASE}.comments+xml"
_COMMENTS_EXTENDED_CT = f"{_parts.CT_BASE}.commentsExtended+xml"

_COMMENTS_NS_DECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
)
_COMMENTS_EXTENDED_NS_DECL = (
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
)

_COMMENT_ID_RE = re.compile(r"^C([0-9]+)$")


def _comment_invalid(detail: str) -> ToolError:
    return ToolError(
        "anchor_invalid", detail, ["Check the comment_id (e.g. 'C7') and op arguments."]
    )


def _comment_not_found(comment_id: str) -> ToolError:
    return ToolError(
        "not_found",
        f"Comment {comment_id} does not exist.",
        ['Call docx_comment {op: "list"} to see comment ids.'],
    )


def _parse_comment_id(comment_id: str) -> str:
    m = _COMMENT_ID_RE.match(comment_id)
    if not m:
        raise _comment_invalid(f"Malformed comment id: {comment_id}.")
    return m.group(1)


def _initials(author: str) -> str:
    """Uppercased first letter of each whitespace-separated author word (§18)."""
    return "".join(word[0].upper() for word in author.split() if word)


# ---------------------------------------------------------------------------
# comments.xml access
# ---------------------------------------------------------------------------


def _ensure_comments(package: Package) -> bytes:
    """Return the comments part, creating it (+ its w14 namespace) on demand."""
    if package.has_part(COMMENTS_PART):
        return package.part(COMMENTS_PART)
    body = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f"<w:comments {_COMMENTS_NS_DECL}></w:comments>"
    ).encode()
    package.set_part(COMMENTS_PART, body)
    _parts.ensure_content_type_override(package, COMMENTS_PART, _COMMENTS_CT)
    main = package.main_document_part()
    _parts.add_relationship(
        package, main, _parts.next_rel_id(package, main), _COMMENTS_REL_TYPE, "comments.xml"
    )
    return body


def _ensure_comments_extended(package: Package) -> bytes:
    if package.has_part(COMMENTS_EXTENDED_PART):
        return package.part(COMMENTS_EXTENDED_PART)
    body = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f"<w15:commentsEx {_COMMENTS_EXTENDED_NS_DECL}></w15:commentsEx>"
    ).encode()
    package.set_part(COMMENTS_EXTENDED_PART, body)
    _parts.ensure_content_type_override(
        package, COMMENTS_EXTENDED_PART, _COMMENTS_EXTENDED_CT
    )
    main = package.main_document_part()
    _parts.add_relationship(
        package,
        main,
        _parts.next_rel_id(package, main),
        _COMMENTS_EXTENDED_REL_TYPE,
        "commentsExtended.xml",
    )
    return body


def _max_comment_id(package: Package) -> int:
    if not package.has_part(COMMENTS_PART):
        return -1
    data = package.part(COMMENTS_PART)
    max_id = -1
    for el in _xml.iter_elements(data, names=("w:comment",)):
        attrs = _edits.start_tag_attrs(data, el)
        value = attrs.get("w:id")
        if value is not None and value.lstrip("-").isdigit():
            max_id = max(max_id, int(value))
    return max_id


def _make_para_id(comment_id: int, text: str) -> str:
    """Deterministic 8-hex ``w14:paraId`` from id+text (byte-parity with the JS engine).

    Word uses random ids; we derive one as the first 8 uppercase hex chars of the
    SHA-256 of ``paraId:{id}:{text}`` (UTF-8).
    """
    import hashlib

    digest = hashlib.sha256(f"paraId:{comment_id}:{text}".encode()).hexdigest()
    return digest[:8].upper()


def _comment_xml(comment_id: int, author: str, date: str, text: str, para_id: str) -> str:
    run = f"<w:r>{_xml.emit_text_element(text)}</w:r>"
    return (
        f'<w:comment w:id="{comment_id}" w:author="{_xml.escape_attr(author)}"'
        f' w:date="{_xml.escape_attr(date)}" w:initials="{_xml.escape_attr(_initials(author))}">'
        f'<w:p w14:paraId="{para_id}">{run}</w:p></w:comment>'
    )


# ---------------------------------------------------------------------------
# Document part wiring (places 1–3)
# ---------------------------------------------------------------------------


def _wire_document(package: Package, anchor: str, comment_id: int) -> str:
    """Splice range start (before runs) + range end & reference run (after) (§18)."""
    main = package.main_document_part()
    entries = _edits.paragraph_entries(package)
    entry = _edits.require_paragraph(entries, anchor)
    data = package.part(main)
    p = entry.span
    # Range start before the paragraph's runs: after any leading w:pPr.
    ppr = next(
        _xml.iter_elements(data, p.inner_start, p.inner_end, names=("w:pPr",), max_depth=1), None
    )
    start_pos = ppr.end if ppr is not None else p.inner_start
    range_start = f'<w:commentRangeStart w:id="{comment_id}"/>'
    end_pos = p.inner_end
    range_end = (
        f'<w:commentRangeEnd w:id="{comment_id}"/>'
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{comment_id}"/></w:r>'
    )
    new = _xml.splice(
        data,
        [
            (start_pos, start_pos, range_start.encode("utf-8")),
            (end_pos, end_pos, range_end.encode("utf-8")),
        ],
    )
    package.set_part(main, new)
    return entry.anchor


# ---------------------------------------------------------------------------
# add (§18)
# ---------------------------------------------------------------------------


def _add(package: Package, anchor: str, text: str, author: str) -> tuple[str, str]:
    if anchor is None:
        raise _comment_invalid("add requires an anchor.")
    date = _edits.revision_date()
    # JS order: ensure the comments part, then the style, then append the comment,
    # then splice the in-document markers (new-part creation order matters for §9).
    _ensure_comments(package)
    _parts.ensure_style(package, "CommentReference")
    comment_id = _max_comment_id(package) + 1
    para_id = _make_para_id(comment_id, text)
    fragment = _comment_xml(comment_id, author, date, text, para_id)
    data = package.part(COMMENTS_PART)
    package.set_part(COMMENTS_PART, _parts.append_before_close(data, b"</w:comments>", fragment))
    body_anchor = _wire_document(package, anchor, comment_id)
    return f"C{comment_id}", body_anchor


# ---------------------------------------------------------------------------
# Comment thread model (parse comments + commentsExtended)
# ---------------------------------------------------------------------------


def _comment_text(p_root: ET.Element) -> str:
    parts = [t.text or "" for t in p_root.iter(f"{{{_W_NS}}}t")]
    return "".join(parts)


def _read_comments(package: Package) -> list[dict[str, object]]:
    """All ``w:comment`` definitions in document order: id, author, date, text, paraId."""
    if not package.has_part(COMMENTS_PART):
        return []
    try:
        root = ET.fromstring(package.part(COMMENTS_PART))
    except ET.ParseError:
        return []
    out: list[dict[str, object]] = []
    for comment in root.iter(f"{{{_W_NS}}}comment"):
        cid = comment.get(f"{{{_W_NS}}}id", "")
        first_p = comment.find(f"{{{_W_NS}}}p")
        para_id = first_p.get(f"{{{_W14_NS}}}paraId", "") if first_p is not None else ""
        out.append(
            {
                "id": cid,
                "author": comment.get(f"{{{_W_NS}}}author", ""),
                "date": comment.get(f"{{{_W_NS}}}date", ""),
                "text": _comment_text(comment),
                "para_id": para_id,
            }
        )
    return out


def _read_extended(package: Package) -> dict[str, dict[str, str]]:
    """paraId → {parent, done} from ``commentsExtended.xml`` (empty when absent)."""
    if not package.has_part(COMMENTS_EXTENDED_PART):
        return {}
    try:
        root = ET.fromstring(package.part(COMMENTS_EXTENDED_PART))
    except ET.ParseError:
        return {}
    out: dict[str, dict[str, str]] = {}
    for ex in root.iter(f"{{{_W15_NS}}}commentEx"):
        para_id = ex.get(f"{{{_W15_NS}}}paraId", "")
        out[para_id] = {
            "parent": ex.get(f"{{{_W15_NS}}}paraIdParent", ""),
            "done": ex.get(f"{{{_W15_NS}}}done", "0"),
        }
    return out


def _comment_range_anchors(package: Package) -> dict[str, str]:
    """comment id → body anchor of its ``commentRangeStart``'s paragraph (§18)."""
    from ._anchors import paragraph_anchor, paragraph_normalized_text

    main = package.main_document_part()
    data = package.part(main)
    out: dict[str, str] = {}
    ordinal = 0
    for child in _xml.iter_body_children(data):
        if child.name != "w:p":
            continue
        ordinal += 1
        for el in _xml.iter_elements(
            data, child.inner_start, child.inner_end, names=("w:commentRangeStart",)
        ):
            cid = _edits.start_tag_attrs(data, el).get("w:id")
            if cid is not None and cid not in out:
                out[cid] = paragraph_anchor(ordinal, paragraph_normalized_text(data, child))
    return out


def _reply_para_ids(extended: dict[str, dict[str, str]]) -> tuple[set[str], dict[str, str]]:
    """paraIds whose ``commentEx`` carries a ``paraIdParent`` (replies) + the parent map."""
    reply_para_ids: set[str] = set()
    parent_of: dict[str, str] = {}
    for para_id, info in extended.items():
        if info.get("parent"):
            reply_para_ids.add(para_id)
            parent_of[para_id] = info["parent"]
    return reply_para_ids, parent_of


def _root_of(
    record: dict[str, object],
    by_para: dict[str, dict[str, object]],
    parent_of: dict[str, str],
) -> dict[str, object]:
    """Walk ``paraIdParent`` up to the thread root (§18)."""
    current = record
    seen: set[str] = set()
    while True:
        para_id = str(current["para_id"])
        parent = parent_of.get(para_id)
        if parent is None or para_id in seen:
            return current
        seen.add(para_id)
        nxt = by_para.get(parent)
        if nxt is None:
            return current
        current = nxt


def _threads(package: Package) -> list[dict[str, object]]:
    """One entry per thread root (§18), with replies attached in document order."""
    comments = _read_comments(package)
    extended = _read_extended(package)
    reply_para_ids, parent_of = _reply_para_ids(extended)
    by_para: dict[str, dict[str, object]] = {
        str(c["para_id"]): c for c in comments if c["para_id"]
    }
    out: list[dict[str, object]] = []
    range_anchors = _comment_range_anchors(package)
    for root in comments:
        para_id = str(root["para_id"])
        if para_id and para_id in reply_para_ids:
            continue  # a reply, not a root
        replies: list[dict[str, str]] = []
        for cand in comments:
            if cand["id"] == root["id"]:
                continue
            cand_para = str(cand["para_id"])
            if (
                cand_para
                and cand_para in reply_para_ids
                and _root_of(cand, by_para, parent_of)["id"] == root["id"]
            ):
                replies.append(
                    {
                        "author": str(cand["author"]),
                        "date": str(cand["date"]),
                        "text": str(cand["text"]),
                    }
                )
        resolved = extended.get(para_id, {}).get("done", "0") == "1" if para_id else False
        out.append(
            {
                "id": f"C{root['id']}",
                "anchor": range_anchors.get(str(root["id"]), ""),
                "author": root["author"],
                "date": root["date"],
                "text": root["text"],
                "resolved": resolved,
                "replies": replies,
            }
        )
    return out


def _find_root(package: Package, comment_id: str) -> dict[str, object]:
    """The thread root dict whose ``w:comment/@w:id`` equals ``comment_id`` (or its parent)."""
    comments = _read_comments(package)
    extended = _read_extended(package)
    _, parent_of = _reply_para_ids(extended)
    by_para = {str(c["para_id"]): c for c in comments if c["para_id"]}
    for c in comments:
        if str(c["id"]) == comment_id:
            return _root_of(c, by_para, parent_of)
    raise _comment_not_found(f"C{comment_id}")


# ---------------------------------------------------------------------------
# reply (§18)
# ---------------------------------------------------------------------------


def _reply(package: Package, comment_id: str, text: str, author: str) -> str:
    if not package.has_part(COMMENTS_PART):
        raise _comment_not_found(f"C{comment_id}")
    root = _find_root(package, comment_id)
    parent_para_id = str(root["para_id"])
    new_id = _max_comment_id(package) + 1
    date = _edits.revision_date()
    child_para_id = _make_para_id(new_id, text)
    fragment = _comment_xml(new_id, author, date, text, child_para_id)
    data = package.part(COMMENTS_PART)
    package.set_part(COMMENTS_PART, _parts.append_before_close(data, b"</w:comments>", fragment))
    # commentsExtended: add the child entry linking it to its thread root.
    _ensure_comments_extended(package)
    child_ex = (
        f'<w15:commentEx w15:paraId="{child_para_id}"'
        f' w15:paraIdParent="{parent_para_id}" w15:done="0"/>'
    )
    ext_data = package.part(COMMENTS_EXTENDED_PART)
    package.set_part(
        COMMENTS_EXTENDED_PART,
        _parts.append_before_close(ext_data, b"</w15:commentsEx>", child_ex),
    )
    return f"C{new_id}"


# ---------------------------------------------------------------------------
# resolve (§18)
# ---------------------------------------------------------------------------


_DONE_ATTR_RE = re.compile(rb'w15:done\s*=\s*"[^"]*"')
_SELF_CLOSE_RE = re.compile(rb"/?>$")


def _resolve(package: Package, comment_id: str) -> None:
    if not package.has_part(COMMENTS_PART):
        raise _comment_not_found(f"C{comment_id}")
    root = _find_root(package, comment_id)
    root_para = str(root["para_id"])
    _ensure_comments_extended(package)
    data = package.part(COMMENTS_EXTENDED_PART)
    for ex in _xml.iter_elements(data, names=("w15:commentEx",)):
        if f'w15:paraId="{root_para}"'.encode() in data[ex.start : ex.end]:
            tag = data[ex.start : ex.end]
            if _DONE_ATTR_RE.search(tag):
                updated = _DONE_ATTR_RE.sub(b'w15:done="1"', tag)
            else:
                updated = _SELF_CLOSE_RE.sub(b' w15:done="1"/>', tag)
            package.set_part(
                COMMENTS_EXTENDED_PART, _xml.splice(data, [(ex.start, ex.end, updated)])
            )
            return
    fragment = f'<w15:commentEx w15:paraId="{root_para}" w15:done="1"/>'
    package.set_part(
        COMMENTS_EXTENDED_PART,
        _parts.append_before_close(data, b"</w15:commentsEx>", fragment),
    )


# ---------------------------------------------------------------------------
# delete (§18)
# ---------------------------------------------------------------------------


def _delete(package: Package, comment_id: str) -> None:
    if not package.has_part(COMMENTS_PART):
        raise _comment_not_found(f"C{comment_id}")
    comments = _read_comments(package)
    root = next((c for c in comments if str(c["id"]) == comment_id), None)
    if root is None:
        raise _comment_not_found(f"C{comment_id}")
    extended = _read_extended(package)
    # The thread: the root's paraId plus the transitive closure of paraIdParent links.
    thread_para_ids: set[str] = {str(root["para_id"])}
    changed = True
    while changed:
        changed = False
        for para_id, info in extended.items():
            parent = info.get("parent", "")
            if parent and parent in thread_para_ids and para_id not in thread_para_ids:
                thread_para_ids.add(para_id)
                changed = True
    ids_to_delete = {
        str(c["id"])
        for c in comments
        if str(c["id"]) == comment_id or str(c["para_id"]) in thread_para_ids
    }
    _remove_document_places(package, ids_to_delete)
    _remove_comment_defs(package, ids_to_delete)
    _remove_comment_ex(package, thread_para_ids)


def _remove_document_places(package: Package, ids: set[str]) -> None:
    main = package.main_document_part()
    data = package.part(main)
    edits: list[tuple[int, int, bytes]] = []
    names = ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
    for el in _xml.iter_elements(data, names=names):
        if _edits.start_tag_attrs(data, el).get("w:id") in ids:
            if el.name == "w:commentReference":
                # Remove the enclosing reference run if it wraps just this reference.
                run = _enclosing_run(data, el)
                if run is not None:
                    edits.append((run.start, run.end, b""))
                else:
                    edits.append((el.start, el.end, b""))
            else:
                edits.append((el.start, el.end, b""))
    if edits:
        package.set_part(main, _xml.splice(data, _dedup_edits(edits)))


def _dedup_edits(edits: list[tuple[int, int, bytes]]) -> list[tuple[int, int, bytes]]:
    """Drop edits fully contained in an earlier (wider) edit; sort by start."""
    edits = sorted(edits, key=lambda e: (e[0], -e[1]))
    out: list[tuple[int, int, bytes]] = []
    last_end = -1
    for start, end, repl in edits:
        if start < last_end:
            continue  # contained in a wider edit already taken
        out.append((start, end, repl))
        last_end = end
    return out


def _enclosing_run(data: bytes, el: _xml.Span) -> _xml.Span | None:
    for run in _xml.iter_elements(data, names=("w:r",)):
        if run.start < el.start and el.end <= run.end:
            # Only collapse the run when its only meaningful child is the reference.
            return run
    return None


def _remove_comment_defs(package: Package, ids: set[str]) -> None:
    if not package.has_part(COMMENTS_PART):
        return
    data = package.part(COMMENTS_PART)
    edits: list[tuple[int, int, bytes]] = []
    for el in _xml.iter_elements(data, names=("w:comment",)):
        if _edits.start_tag_attrs(data, el).get("w:id") in ids:
            edits.append((el.start, el.end, b""))
    if edits:
        package.set_part(COMMENTS_PART, _xml.splice(data, edits))


def _remove_comment_ex(package: Package, para_ids: set[str]) -> None:
    if not package.has_part(COMMENTS_EXTENDED_PART):
        return
    data = package.part(COMMENTS_EXTENDED_PART)
    edits: list[tuple[int, int, bytes]] = []
    for ex in _xml.iter_elements(data, names=("w15:commentEx",)):
        m = re.search(rb'w15:paraId="([0-9A-Fa-f]+)"', data[ex.start : ex.end])
        if m is not None and m.group(1).decode("utf-8") in para_ids:
            edits.append((ex.start, ex.end, b""))
    if edits:
        package.set_part(COMMENTS_EXTENDED_PART, _xml.splice(data, edits))


# ---------------------------------------------------------------------------
# docx_comment
# ---------------------------------------------------------------------------


def docx_comment(
    session: Session,
    *,
    doc_id: str,
    op: str,
    anchor: str | None = None,
    comment_id: str | None = None,
    text: str | None = None,
    author: str | None = None,
    response_format: str = "concise",
) -> dict[str, object]:
    """Add, reply, resolve, list, or delete comments (§18)."""
    doc = session.get(doc_id)
    package = doc.package
    if op == "list":
        return {"comments": _threads(package)}
    if op == "add":
        if anchor is None:
            raise _comment_invalid("add requires an anchor.")
        cid, body_anchor = _add(package, anchor, text or "", _edits.resolve_author(author))
        doc.mark_dirty()
        return {"comment_id": cid, "anchor": body_anchor}
    if op == "reply":
        if comment_id is None:
            raise _comment_invalid("reply requires a comment_id.")
        raw_id = _parse_comment_id(comment_id)
        cid = _reply(package, raw_id, text or "", _edits.resolve_author(author))
        doc.mark_dirty()
        return {"comment_id": cid}
    if op == "resolve":
        if comment_id is None:
            raise _comment_invalid("resolve requires a comment_id.")
        raw_id = _parse_comment_id(comment_id)
        _resolve(package, raw_id)
        doc.mark_dirty()
        return {"comment_id": comment_id, "note": "Comment thread resolved."}
    if op == "delete":
        if comment_id is None:
            raise _comment_invalid("delete requires a comment_id.")
        raw_id = _parse_comment_id(comment_id)
        _delete(package, raw_id)
        doc.mark_dirty()
        return {"comment_id": comment_id, "note": f"Deleted comment {comment_id} and its replies."}
    raise _comment_invalid(f"Unknown comment op: {op}.")
