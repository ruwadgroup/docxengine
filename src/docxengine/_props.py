"""The closed formatting prop set and its pinned emission order (algorithms.md §16).

Both ``docx_style`` (style definitions) and ``docx_format`` (style-selector merges
and direct formatting) draw on the same canonical property names and the same
``w:rPr``/``w:pPr`` emission order. Shorthand wire names (``size``, ``spacing_after``)
map onto the canonical spec names (``size_pt``, ``spacing_after_pt``); ``spacing``
(a line multiplier) is out of the closed set the style/format writers honor (§13).
"""

from __future__ import annotations

from collections.abc import Mapping

from . import _xml

#: Wire shorthand → canonical spec property name (§13).
_ALIASES = {
    "size": "size_pt",
    "spacing_after": "spacing_after_pt",
    "spacing_before": "spacing_before_pt",
}

#: ``w:rPr`` child property names in their §16 emission order.
RPR_PROPS = ("bold", "italic", "underline", "color", "size_pt")
#: ``w:pPr`` child property names (after any ``w:pStyle``) in §16 emission order.
PPR_PROPS = ("alignment", "spacing_before_pt", "spacing_after_pt")

_JC = {"left": "left", "center": "center", "right": "right", "justify": "both", "both": "both"}


def canonical_props(props: Mapping[str, object]) -> dict[str, object]:
    """Resolve shorthand keys to their canonical names; ``spacing`` is dropped (§13)."""
    out: dict[str, object] = {}
    for key, value in props.items():
        if key == "spacing":  # a line multiplier, ignored by style/format writers
            continue
        out[_ALIASES.get(key, key)] = value
    return out


def _bool_toggle(tag: str, value: object) -> str:
    """A boolean toggle: ``<w:b/>`` for true, ``<w:b w:val="0"/>`` for false (§16)."""
    return f"<{tag}/>" if value else f'<{tag} w:val="0"/>'


def _rpr_child(name: str, value: object) -> str | None:
    if name == "bold":
        return _bool_toggle("w:b", value)
    if name == "italic":
        return _bool_toggle("w:i", value)
    if name == "underline":
        return '<w:u w:val="single"/>' if value else '<w:u w:val="none"/>'
    if name == "color":
        hex6 = str(value).lstrip("#").upper()
        return f'<w:color w:val="{_xml.escape_attr(hex6)}"/>'
    if name == "size_pt":
        return f'<w:sz w:val="{round(float(value) * 2)}"/>'  # type: ignore[arg-type]
    return None


def _ppr_child(name: str, value: object) -> str | None:
    if name == "alignment":
        jc = _JC.get(str(value))
        return f'<w:jc w:val="{jc}"/>' if jc is not None else None
    if name in ("spacing_before_pt", "spacing_after_pt"):
        attr = "w:before" if name == "spacing_before_pt" else "w:after"
        return f'<w:spacing {attr}="{round(float(value) * 20)}"/>'  # type: ignore[arg-type]
    return None


def rpr_children(props: Mapping[str, object]) -> list[tuple[str, str]]:
    """``(local-tag-name, xml)`` for each present ``w:rPr`` prop, in §16 order.

    The local tag name (``w:b``, ``w:sz``…) is the merge key for replacing a same-named
    child in an existing ``w:rPr`` (§16 style-selector merge).
    """
    out: list[tuple[str, str]] = []
    for name in RPR_PROPS:
        if name not in props:
            continue
        xml = _rpr_child(name, props[name])
        if xml is not None:
            out.append((_tag_name(xml), xml))
    return out


def ppr_children(props: Mapping[str, object]) -> list[tuple[str, str]]:
    """``(local-tag-name, xml)`` for each present ``w:pPr`` prop, in §16 order."""
    out: list[tuple[str, str]] = []
    for name in PPR_PROPS:
        if name not in props:
            continue
        xml = _ppr_child(name, props[name])
        if xml is not None:
            out.append((_tag_name(xml), xml))
    return out


def _tag_name(xml: str) -> str:
    """The element name of a freshly emitted child (``<w:sz w:val="22"/>`` → ``w:sz``)."""
    end = len(xml)
    for i, ch in enumerate(xml):
        if ch in " />" and i > 0:
            end = i
            break
    return xml[1:end]


def rpr_xml(props: Mapping[str, object]) -> str:
    """A complete ``w:rPr`` for the present run props (``""`` when none apply)."""
    children = "".join(xml for _, xml in rpr_children(props))
    return f"<w:rPr>{children}</w:rPr>" if children else ""


def ppr_children_xml(props: Mapping[str, object]) -> str:
    """The ``w:pPr`` child run for the present paragraph props (``""`` when none)."""
    return "".join(xml for _, xml in ppr_children(props))
