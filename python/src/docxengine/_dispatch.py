"""Tool dispatcher: route a spec tool name + args object to its implementation.

As of Phase 2 every one of the 24 tools defined in ``spec/tools/`` routes to a
handler — no path returns ``not_implemented`` for a defined tool. The
``not_implemented`` guard is retained for forward-compat: a future spec tool
shipped without a handler still fails cleanly instead of ``KeyError``.
Argument validation is minimal per the spec: required keys must be present
(``invalid_args``), and only schema-declared properties are forwarded.

Doc state is per process: :func:`call` uses one module-level
:class:`~docxengine._session.Session` unless the caller passes its own.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping

from . import (
    _comments,
    _convert,
    _create,
    _fields,
    _lists,
    _media,
    _render,
    _sections,
    _spec,
    _styles,
    _tables,
    _template,
    _tools_edit,
    _tools_lifecycle,
    _tools_read,
)
from ._errors import ToolError
from ._session import Session

_HANDLERS: dict[str, Callable[..., dict[str, object]]] = {
    "docx_open": _tools_read.docx_open,
    "docx_outline": _tools_read.docx_outline,
    "docx_read": _tools_read.docx_read,
    "docx_search": _tools_read.docx_search,
    "docx_replace": _tools_edit.docx_replace,
    "docx_edit_paragraph": _tools_edit.docx_edit_paragraph,
    "docx_insert": _tools_edit.docx_insert,
    "docx_delete": _tools_edit.docx_delete,
    "docx_revision": _tools_edit.docx_revision,
    "docx_validate": _tools_lifecycle.docx_validate,
    "docx_repair": _tools_lifecycle.docx_repair,
    "docx_save": _tools_lifecycle.docx_save,
    "docx_table": _tables.docx_table,
    "docx_style": _styles.docx_style,
    "docx_format": _styles.docx_format,
    "docx_list": _lists.docx_list,
    "docx_comment": _comments.docx_comment,
    "docx_section": _sections.docx_section,
    "docx_media": _media.docx_media,
    "docx_field": _fields.docx_field,
    "docx_template_fill": _template.docx_template_fill,
    "docx_create": _create.docx_create,
    "docx_convert": _convert.docx_convert,
    "docx_render_preview": _render.docx_render_preview,
}

MVP_TOOLS: frozenset[str] = frozenset(_HANDLERS)

_process_session: Session | None = None


def _default_session() -> Session:
    global _process_session
    if _process_session is None:
        _process_session = Session()
    return _process_session


def call(
    tool: str,
    args: Mapping[str, object] | None = None,
    *,
    session: Session | None = None,
) -> dict[str, object]:
    """Invoke ``tool`` with the schema-named ``args``; returns the result object.

    Failures raise :class:`~docxengine._errors.ToolError`; serialize
    ``exc.to_payload()`` at a process boundary.
    """
    if session is None:
        session = _default_session()
    schema = _spec.input_schema(tool)
    if schema is None:
        raise ToolError(
            "not_implemented",
            f"Tool {tool} is not defined in spec/tools/.",
            ["See docs/tools/index.md for the tool catalog."],
        )
    if tool not in _HANDLERS:  # forward-compat guard: unreached for the 24 Phase-2 tools
        raise ToolError("not_implemented", f"{tool} has no registered handler.", [])
    if args is None:
        args = {}
    if not isinstance(args, Mapping):
        raise ToolError(
            "invalid_args",
            f"{tool}: args must be a JSON object.",
            [f"Check the tool's input_schema in spec/tools/{tool}.json."],
        )
    required = schema.get("required", [])
    missing = [key for key in required if key not in args] if isinstance(required, list) else []
    if missing:
        raise ToolError(
            "invalid_args",
            f"{tool}: missing required argument(s): {', '.join(missing)}.",
            [f"Check the tool's input_schema in spec/tools/{tool}.json."],
        )
    properties = schema.get("properties", {})
    known = properties.keys() if isinstance(properties, dict) else ()
    kwargs = {key: value for key, value in args.items() if key in known}
    return _HANDLERS[tool](session, **kwargs)
