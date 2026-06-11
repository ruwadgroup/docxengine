"""Package validator and mechanical repair (algorithms.md §8/§8a).

``validate_package`` runs the five MVP checks (a–e) and returns the pinned
issue list — exact ordering, messages, and fix hints matter, because the
conformance harness deep-compares ``docx_validate`` results across the Python
and TypeScript implementations. ``repair_package`` applies the §8a fixes by
splicing raw part bytes (never re-serializing), then re-validates.
"""

from __future__ import annotations

import posixpath
import re
from dataclasses import dataclass

from . import _xml
from ._edits import REVISION_NAMES, next_revision_id, start_tag_attrs
from ._opc import CONTENT_TYPES_PART, Package, rels_part_for, resolve_rel_target

COMMENTS_PART = "word/comments.xml"
FOOTNOTES_PART = "word/footnotes.xml"

#: Relationship types (last path segment) Word consumes without an explicit
#: r:id reference in the document part — exempt from the unreferenced warning.
IMPLICIT_REL_TYPES = frozenset(
    {
        "styles",
        "settings",
        "webSettings",
        "fontTable",
        "numbering",
        "theme",
        "customXml",
        "comments",
        "commentsExtended",
        "footnotes",
        "endnotes",
        "glossaryDocument",
    }
)

#: §8a content types for repaired ``Default`` entries; anything else gets octet-stream.
DEFAULT_CONTENT_TYPES = {
    "rels": "application/vnd.openxmlformats-package.relationships+xml",
    "xml": "application/xml",
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
}
FALLBACK_CONTENT_TYPE = "application/octet-stream"

_R_ID_RE = re.compile(rb'\sr:id\s*=\s*(?:"([^"]*)"|\'([^\']*)\')')
_R_REF_RE = re.compile(rb'\sr:(?:id|embed|link)\s*=\s*(?:"([^"]*)"|\'([^\']*)\')')
_W_ID_VALUE_RE = re.compile(rb'(\sw:id\s*=\s*")([^"]*)(")')


@dataclass(frozen=True, slots=True)
class Issue:
    """One validator finding in the §8a pinned shape."""

    severity: str  # "error" | "warning"
    part: str
    message: str
    fix_hint: str

    def to_payload(self) -> dict[str, str]:
        return {
            "severity": self.severity,
            "part": self.part,
            "message": self.message,
            "fix_hint": self.fix_hint,
        }


# ---------------------------------------------------------------------------
# Scan helpers
# ---------------------------------------------------------------------------


def _content_parts(package: Package) -> list[str]:
    """Package entries subject to check a, in zip order."""
    return [
        name
        for name in package.part_names
        if not name.endswith("/") and name != CONTENT_TYPES_PART
    ]


def _extension(part_name: str) -> str:
    base = posixpath.basename(part_name)
    if "." not in base:
        return ""
    return base.rpartition(".")[2].lower()


def _source_part_for(rels_name: str) -> str | None:
    """Inverse of :func:`~docxengine._opc.rels_part_for` (``None`` = package root)."""
    directory, base = posixpath.split(rels_name)
    parent = posixpath.dirname(directory)  # strip the trailing "_rels"
    source = base.removesuffix(".rels")
    if not source:
        return None
    return f"{parent}/{source}" if parent else source


def _rels_parts(package: Package) -> list[str]:
    return [name for name in package.part_names if name.endswith(".rels")]


def _attr_values(data: bytes, pattern: re.Pattern[bytes]) -> list[str]:
    """Distinct matched attribute values in order of first occurrence."""
    out: list[str] = []
    for m in pattern.finditer(data):
        value = (m.group(1) if m.group(1) is not None else m.group(2)).decode("utf-8")
        if value not in out:
            out.append(value)
    return out


def _element_ids(data: bytes, names: tuple[str, ...]) -> list[tuple[_xml.Span, str]]:
    """``(span, w:id)`` for every named element, in document order."""
    found = [
        (el, start_tag_attrs(data, el).get("w:id", ""))
        for el in _xml.iter_elements(data, names=names)
    ]
    found.sort(key=lambda pair: pair[0].start)
    return found


def _story_refs_and_defs(
    package: Package, main: str, ref_name: str, defs_part: str, def_name: str
) -> tuple[list[str], list[tuple[str, str]]]:
    """Reference ids (distinct, first occurrence) and ``(id, w:type)`` definitions."""
    refs: list[str] = []
    for _, ref_id in _element_ids(package.part(main), (ref_name,)):
        if ref_id not in refs:
            refs.append(ref_id)
    defs: list[tuple[str, str]] = []
    if package.has_part(defs_part):
        data = package.part(defs_part)
        for el in sorted(_xml.iter_elements(data, names=(def_name,)), key=lambda s: s.start):
            attrs = start_tag_attrs(data, el)
            defs.append((attrs.get("w:id", ""), attrs.get("w:type", "")))
    return refs, defs


# ---------------------------------------------------------------------------
# Validation (§8/§8a)
# ---------------------------------------------------------------------------


def validate_package(package: Package) -> list[Issue]:
    """All §8 checks, issues in the pinned a → e order."""
    issues: list[Issue] = []
    main = package.main_document_part()
    main_data = package.part(main)
    main_rels_name = rels_part_for(main)

    # a — content-type coverage, package entries in zip order.
    ct = package.content_types()
    for name in _content_parts(package):
        if ct.content_type_of(name) is None:
            ext = _extension(name)
            issues.append(
                Issue(
                    "error",
                    name,
                    f"Part {name} is not covered by [Content_Types].xml "
                    f"(no Override, no Default for extension '{ext}').",
                    "docx_repair adds a content-type Default for the extension.",
                )
            )

    # b — r:id references of the document part resolve in its rels.
    rel_ids = {rel.rel_id for rel in package.rels(main)}
    for rid in _attr_values(main_data, _R_ID_RE):
        if rid not in rel_ids:
            issues.append(
                Issue(
                    "error",
                    main,
                    f"r:id {rid} is referenced in {main} but not defined in {main_rels_name}.",
                    "Add the missing relationship or remove the referencing element; "
                    "not auto-repairable.",
                )
            )

    # c — every non-External relationship target exists; then unreferenced warnings.
    for rels_name in _rels_parts(package):
        source = _source_part_for(rels_name)
        for rel in package.rels(source):
            if rel.is_external:
                continue
            target = resolve_rel_target(source, rel.target)
            if not package.has_part(target):
                issues.append(
                    Issue(
                        "error",
                        rels_name,
                        f"Relationship {rel.rel_id} targets missing part {target}.",
                        "docx_repair drops the orphaned relationship.",
                    )
                )
    referenced = set(_attr_values(main_data, _R_REF_RE))
    for rel in package.rels(main):
        short_type = rel.rel_type.rstrip("/").rpartition("/")[2]
        if short_type in IMPLICIT_REL_TYPES or rel.rel_id in referenced:
            continue
        issues.append(
            Issue(
                "warning",
                main_rels_name,
                f"Relationship {rel.rel_id} ({short_type}) is never referenced.",
                "Harmless; remove the unused relationship to tidy the package.",
            )
        )

    # d — w:ins/w:del id uniqueness (counted together), first-occurrence order.
    counts: dict[str, int] = {}
    for _, rev_id in _element_ids(main_data, REVISION_NAMES):
        counts[rev_id] = counts.get(rev_id, 0) + 1
    for rev_id, n in counts.items():
        if n > 1:
            issues.append(
                Issue(
                    "error",
                    main,
                    f"Duplicate revision id {rev_id} on {n} w:ins/w:del elements.",
                    "docx_repair renumbers the later duplicates.",
                )
            )

    # e — comment and footnote references resolve both directions.
    for noun, ref_name, defs_part, def_name in (
        ("Comment", "w:commentReference", COMMENTS_PART, "w:comment"),
        ("Footnote", "w:footnoteReference", FOOTNOTES_PART, "w:footnote"),
    ):
        refs, defs = _story_refs_and_defs(package, main, ref_name, defs_part, def_name)
        def_ids = {def_id for def_id, _ in defs}
        for ref_id in refs:
            if ref_id not in def_ids:
                issues.append(
                    Issue(
                        "error",
                        defs_part,
                        f"{noun} id={ref_id} referenced in body but missing.",
                        "docx_repair removes the orphaned reference.",
                    )
                )
        for def_id, def_type in defs:
            if def_id in refs or def_type in ("separator", "continuationSeparator"):
                continue
            issues.append(
                Issue(
                    "warning",
                    defs_part,
                    f"{noun} id={def_id} defined but never referenced.",
                    "Harmless; delete the unused definition to tidy the package.",
                )
            )

    return issues


def is_valid(issues: list[Issue]) -> bool:
    """§8a: valid iff no error-severity issue (warnings never block)."""
    return not any(issue.severity == "error" for issue in issues)


# ---------------------------------------------------------------------------
# Repair (§8/§8a)
# ---------------------------------------------------------------------------


def _drop_orphaned_relationships(package: Package, fixed: list[str]) -> None:
    for rels_name in _rels_parts(package):
        source = _source_part_for(rels_name)
        data = package.part(rels_name)
        edits: list[tuple[int, int, bytes]] = []
        for el in _xml.iter_elements(data, names=("Relationship",)):
            attrs = start_tag_attrs(data, el)
            if attrs.get("TargetMode", "Internal") == "External":
                continue
            target = resolve_rel_target(source, attrs.get("Target", ""))
            if not package.has_part(target):
                edits.append((el.start, el.end, b""))
                fixed.append(f"removed orphaned relationship {attrs.get('Id', '')} ({rels_name})")
        if edits:
            package.set_part(rels_name, _xml.splice(data, edits))


def _add_missing_content_type_defaults(package: Package, fixed: list[str]) -> None:
    ct = package.content_types()
    missing: list[str] = []
    for name in _content_parts(package):
        if ct.content_type_of(name) is not None:
            continue
        ext = _extension(name)
        if ext and ext not in missing:
            missing.append(ext)  # an extension-less part is not fixable here
    if not missing:
        return
    data = package.part(CONTENT_TYPES_PART)
    close = data.rfind(b"</Types>")
    if close < 0:
        return
    inserted = "".join(
        f'<Default Extension="{ext}" '
        f'ContentType="{DEFAULT_CONTENT_TYPES.get(ext, FALLBACK_CONTENT_TYPE)}"/>'
        for ext in missing
    )
    package.set_part(
        CONTENT_TYPES_PART, _xml.splice(data, [(close, close, inserted.encode("utf-8"))])
    )
    fixed.extend(f"added content-type Default for extension '{ext}'" for ext in missing)


def _renumber_duplicate_revision_ids(package: Package, fixed: list[str]) -> None:
    main = package.main_document_part()
    data = package.part(main)
    next_id = next_revision_id(data)  # max existing + 1 (§8a: later duplicates onward)
    seen: set[str] = set()
    edits: list[tuple[int, int, bytes]] = []
    for el, _ in _element_ids(data, REVISION_NAMES):
        tag_end = el.end if el.empty else el.inner_start
        m = _W_ID_VALUE_RE.search(data, el.start, tag_end)
        if m is None:
            continue
        value = m.group(2).decode("utf-8")
        if value not in seen:
            seen.add(value)
            continue
        edits.append((m.start(2), m.end(2), str(next_id).encode("utf-8")))
        fixed.append(f"renumbered duplicate revision id {value} -> {next_id}")
        next_id += 1
    if edits:
        package.set_part(main, _xml.splice(data, edits))


def _remove_orphaned_story_references(package: Package, fixed: list[str]) -> None:
    main = package.main_document_part()
    for noun, ref_name, defs_part, def_name, range_names in (
        (
            "comment",
            "w:commentReference",
            COMMENTS_PART,
            "w:comment",
            ("w:commentRangeStart", "w:commentRangeEnd"),
        ),
        ("footnote", "w:footnoteReference", FOOTNOTES_PART, "w:footnote", ()),
    ):
        refs, defs = _story_refs_and_defs(package, main, ref_name, defs_part, def_name)
        orphaned = {ref_id for ref_id in refs if ref_id not in {d for d, _ in defs}}
        if not orphaned:
            continue
        data = package.part(main)
        edits: list[tuple[int, int, bytes]] = []
        for el, el_id in _element_ids(data, (ref_name, *range_names)):
            if el_id not in orphaned:
                continue
            edits.append((el.start, el.end, b""))
            if el.name == ref_name:
                fixed.append(f"removed orphaned {noun} reference id={el_id}")
        if edits:
            package.set_part(main, _xml.splice(data, edits))


def repair_package(package: Package) -> tuple[list[str], list[str]]:
    """Apply the §8a fixes in order; returns ``(fixed, remaining)``.

    ``remaining`` is the message of every error-severity issue still present
    after re-validation. The caller reports both; fixes are applied either way.
    """
    fixed: list[str] = []
    _drop_orphaned_relationships(package, fixed)
    _add_missing_content_type_defaults(package, fixed)
    _renumber_duplicate_revision_ids(package, fixed)
    _remove_orphaned_story_references(package, fixed)
    remaining = [issue.message for issue in validate_package(package) if issue.severity == "error"]
    return fixed, remaining
