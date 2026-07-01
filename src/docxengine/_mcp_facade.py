"""MCP file facade: the path-based projection of the ``doc_id`` tool contract.

The wire contract (``spec/tools``, the CLI, the SDK ``call()`` surface) addresses
documents by an in-memory ``doc_id`` handle. The MCP server is a deployment where
a document *is* a file, so this module projects that contract onto a file-first
surface: every tool takes a ``path``; each call opens the file into an ephemeral
:class:`~docxengine._session.Session`, runs the underlying tool, and — when the
edit dirtied the document — validates and atomically saves it back to the same
path. ``doc_id`` never leaves the process, ``docx_save`` is folded into every
mutation, and ``docx_create`` writes its file immediately.

``mcp.py`` imports :func:`facade_tool_schemas` for ``tools/list`` and
:func:`call_path_tool` for ``tools/call``. The core (:func:`docxengine._dispatch.call`,
the handlers, ``spec/``) is untouched.
"""

from __future__ import annotations

import threading
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any, cast

from ._dispatch import call
from ._errors import ToolError
from ._paths import resolve_path
from ._session import Session
from ._spec import spec_tool_names, tool_schemas
from ._tools_lifecycle import docx_save

#: Spec tools not exposed over MCP (saving is automatic on every edit).
_DROPPED: frozenset[str] = frozenset({"docx_save"})

_PATH_DESC = (
    "Path to the .docx file. Relative paths resolve against the server's working "
    "directory (or DOCXENGINE_ROOT when set). Edits are validated and saved back "
    "to this file automatically — there is no separate save step."
)

#: Per-tool description rewrites where the spec text assumes the doc_id workflow.
_DESC_OVERRIDES: dict[str, str] = {
    "docx_open": (
        "Open a .docx file and get a human-readable summary of what it contains "
        "(paragraph count, whether it carries tracked changes or comments). Read-only "
        "— never modifies the file. You do not need to call this before other tools; "
        "every tool takes the file path directly."
    ),
    "docx_create": (
        "Create a new .docx file at `path` from Markdown (headings, lists, tables, "
        "emphasis) or a structured spec with explicit styles and sections. The file is "
        "written and validated immediately. Provide exactly one of content_md or spec."
    ),
    "docx_convert": (
        "Convert a .docx file to another format: md/html are produced in-engine "
        "(lossless for content; revisions and comments annotated inline) and returned "
        "as content; pdf/png go through the render adapter and are written to "
        "output_path, never returned as inline bytes."
    ),
    "docx_template_fill": (
        "Fill a mustache-style template .docx — placeholders, loops, conditions — with "
        "data and write the result to `path`. Placeholders fragmented across split runs "
        "are coalesced before matching; an empty 'unfilled' list is the success check."
    ),
    "docx_outline": (
        "Return the document's heading tree and table list with anchors, resolved "
        "through the style cascade. The cheap map of the document — call it first on an "
        "unfamiliar file, then use the anchors for targeted reads."
    ),
    "docx_validate": (
        "Check the document's package integrity (duplicate IDs, broken relationships, "
        "revision well-formedness) and report issues with fix hints. Every edit runs "
        "this gate for you and refuses to save a broken package; 'warning' issues never "
        "block."
    ),
}


def _path_prop(desc: str = _PATH_DESC) -> dict[str, str]:
    return {"type": "string", "description": desc}


# ---------------------------------------------------------------------------
# Schema transform (spec doc_id contract -> MCP path surface)
# ---------------------------------------------------------------------------


def _doc_id_to_path(input_schema: dict[str, Any]) -> None:
    """Generic rule: replace the ``doc_id`` property/required slot with ``path``."""
    props: dict[str, Any] = dict(input_schema.get("properties", {}))
    props.pop("doc_id", None)
    input_schema["properties"] = {"path": _path_prop(), **props}
    required = input_schema.get("required", [])
    required = required if isinstance(required, list) else []
    input_schema["required"] = ["path", *(r for r in required if r != "doc_id")]


def _transform_schema(name: str, schema: dict[str, Any]) -> None:
    if name in _DESC_OVERRIDES:
        schema["description"] = _DESC_OVERRIDES[name]
    input_schema = schema["input_schema"]
    props: dict[str, Any] = input_schema["properties"]

    if name == "docx_open":
        props.pop("bytes", None)
        props["path"] = _path_prop("Path to the .docx file to open. Read-only.")
        input_schema["required"] = ["path"]
    elif name == "docx_create":
        new_desc = "Path to write the new .docx file. Written and validated immediately."
        input_schema["properties"] = {"path": _path_prop(new_desc), **props}
        input_schema["required"] = ["path"]
    elif name == "docx_template_fill":
        rebuilt: dict[str, Any] = {}
        for key, value in props.items():
            rebuilt[key] = value
            if key == "template":
                rebuilt["path"] = _path_prop("Path to write the filled .docx file.")
        input_schema["properties"] = rebuilt
        input_schema["required"] = ["template", "path", "data"]
    elif name == "docx_convert":
        props.pop("doc_id", None)
        output = props.pop("path", None)
        if output is not None:
            output["description"] = (
                "Output path for file-producing targets. Required when to is 'pdf' or "
                "'png'; ignored for 'md' and 'html'."
            )
            props["output_path"] = output
        input_schema["properties"] = {"path": _path_prop("Path to the source .docx file."), **props}
        input_schema["required"] = ["path", "to"]
    elif name == "docx_media":
        props.pop("doc_id", None)
        output = props.pop("path", None)
        if output is not None:
            output["description"] = "Output file path for the extracted image (op 'extract')."
            props["output_path"] = output
        input_schema["properties"] = {"path": _path_prop("Path to the source .docx file."), **props}
        input_schema["required"] = ["path", "op"]
    else:
        _doc_id_to_path(input_schema)


def facade_tool_schemas() -> list[dict[str, Any]]:
    """The spec tool schemas projected onto the path surface (``docx_save`` dropped)."""
    out: list[dict[str, Any]] = []
    for schema in tool_schemas():  # fresh deep copies
        if schema["name"] in _DROPPED:
            continue
        _transform_schema(str(schema["name"]), schema)
        out.append(schema)
    return out


# ---------------------------------------------------------------------------
# Call path (open -> run -> validate -> save back)
# ---------------------------------------------------------------------------


class FacadeContext:
    """Process-wide facade state: the sandbox root and per-path write locks.

    One context is shared by the stdio loop and every HTTP session — the locks
    must be global so two sessions targeting the same file serialize correctly.
    """

    def __init__(self, root: Path | None) -> None:
        self.root = root
        self._locks: dict[str, threading.Lock] = {}
        self._meta = threading.Lock()

    def lock_for(self, resolved: Path) -> threading.Lock:
        key = str(resolved)
        with self._meta:
            lock = self._locks.get(key)
            if lock is None:
                lock = threading.Lock()
                self._locks[key] = lock
            return lock


def _facade_names() -> frozenset[str]:
    return spec_tool_names() - _DROPPED


def _strip_doc_id(result: dict[str, Any]) -> dict[str, Any]:
    result.pop("doc_id", None)
    return result


def _require(args: Mapping[str, Any], key: str, name: str) -> Any:
    if key not in args:
        raise ToolError(
            "invalid_args",
            f"{name}: missing required argument: {key}.",
            [f"Pass {key} (the .docx file path)."],
        )
    return args[key]


def _save_back(session: Session, doc_id: str, resolved: Path) -> int:
    """Validate + atomically write the dirty doc back; refusal leaves the file as-is."""
    try:
        saved = docx_save(session, doc_id=doc_id, path=str(resolved))
    except ToolError as exc:
        if exc.code == "validation_failed":
            exc.suggestions.append(
                "The file on disk was not modified. Run docx_repair on the same path, then retry."
            )
        raise
    return cast(int, saved["bytes"])


def _open_run_save(
    ctx: FacadeContext,
    name: str,
    user_path: object,
    core_args: dict[str, Any],
    *,
    echo_path: bool = True,
) -> dict[str, Any]:
    """Open ``user_path``, run ``name`` against it, and save back if it dirtied."""
    resolved = resolve_path(user_path, ctx.root)
    with ctx.lock_for(resolved):
        session = Session()
        doc = session.open_doc(resolved)  # open_failed propagates
        core = dict(call(name, {**core_args, "doc_id": doc.doc_id}, session=session))
        result = _strip_doc_id(core)
        if doc.dirty:
            result["bytes"] = _save_back(session, doc.doc_id, resolved)
            result["saved"] = True
    if echo_path:
        result.setdefault("path", user_path)
    return result


def _call_generic(name: str, args: dict[str, Any], ctx: FacadeContext) -> dict[str, Any]:
    user_path = _require(args, "path", name)
    core_args = {k: v for k, v in args.items() if k != "path"}
    return _open_run_save(ctx, name, user_path, core_args)


def _call_open(name: str, args: dict[str, Any], ctx: FacadeContext) -> dict[str, Any]:
    user_path = _require(args, "path", name)
    resolved = resolve_path(user_path, ctx.root)
    with ctx.lock_for(resolved):
        session = Session()
        result = _strip_doc_id(dict(call("docx_open", {"path": str(resolved)}, session=session)))
    result["path"] = user_path
    return result


def _call_create(name: str, args: dict[str, Any], ctx: FacadeContext) -> dict[str, Any]:
    user_path = _require(args, "path", name)
    resolved = resolve_path(user_path, ctx.root)
    core_args = {k: v for k, v in args.items() if k != "path"}
    with ctx.lock_for(resolved):
        session = Session()
        created = _strip_doc_id(dict(call("docx_create", core_args, session=session)))
        doc_id = next(iter(session.doc_ids()))
        created["bytes"] = _save_back(session, doc_id, resolved)
    created["saved"] = True
    created["path"] = user_path
    return created


def _call_template_fill(name: str, args: dict[str, Any], ctx: FacadeContext) -> dict[str, Any]:
    template = resolve_path(_require(args, "template", name), ctx.root)
    user_path = _require(args, "path", name)
    resolved = resolve_path(user_path, ctx.root)
    core_args: dict[str, Any] = {"template": str(template), "data": args.get("data")}
    for opt in ("syntax", "strict"):
        if opt in args:
            core_args[opt] = args[opt]
    with ctx.lock_for(resolved):
        session = Session()
        filled = _strip_doc_id(dict(call("docx_template_fill", core_args, session=session)))
        doc_id = next(iter(session.doc_ids()))
        filled["bytes"] = _save_back(session, doc_id, resolved)
    filled["saved"] = True
    filled["path"] = user_path
    return filled


def _call_convert(name: str, args: dict[str, Any], ctx: FacadeContext) -> dict[str, Any]:
    user_path = _require(args, "path", name)
    core_args: dict[str, Any] = {"to": args.get("to")}
    if "output_path" in args:
        core_args["path"] = str(resolve_path(args["output_path"], ctx.root))
    return _open_run_save(ctx, name, user_path, core_args, echo_path=False)


def _call_media(name: str, args: dict[str, Any], ctx: FacadeContext) -> dict[str, Any]:
    user_path = _require(args, "path", name)
    core_args = {k: v for k, v in args.items() if k not in ("path", "output_path", "image")}
    if "image" in args:
        core_args["image"] = str(resolve_path(args["image"], ctx.root))
    if "output_path" in args:
        core_args["path"] = str(resolve_path(args["output_path"], ctx.root))
    return _open_run_save(ctx, name, user_path, core_args, echo_path=False)


def _call_render_preview(name: str, args: dict[str, Any], ctx: FacadeContext) -> dict[str, Any]:
    user_path = _require(args, "path", name)
    core_args = {k: v for k, v in args.items() if k != "path"}
    result = _open_run_save(ctx, name, user_path, core_args)
    pages = result.get("pages")
    if isinstance(pages, list):
        for page in pages:
            image = page.get("image") if isinstance(page, dict) else None
            if isinstance(image, str) and "/preview/" in image:
                page["image"] = f"docx://{user_path}/preview/{image.split('/preview/', 1)[1]}"
    return result


_SPECIAL: dict[str, Callable[[str, dict[str, Any], FacadeContext], dict[str, Any]]] = {
    "docx_open": _call_open,
    "docx_create": _call_create,
    "docx_template_fill": _call_template_fill,
    "docx_convert": _call_convert,
    "docx_media": _call_media,
    "docx_render_preview": _call_render_preview,
}


def call_path_tool(name: str, args: Mapping[str, Any] | None, ctx: FacadeContext) -> dict[str, Any]:
    """Dispatch one path-addressed MCP tool call; raises :class:`ToolError` on failure."""
    if name in _DROPPED:
        raise ToolError(
            "invalid_args",
            "docx_save is not exposed over MCP — every edit is validated and saved to its "
            "file automatically.",
            ["Call the edit tool with the file path; the change is persisted on success."],
        )
    if name not in _facade_names():
        raise ToolError(
            "invalid_args",
            f"Unknown tool: {name}.",
            ["See tools/list for the available tools."],
        )
    handler = _SPECIAL.get(name, _call_generic)
    return handler(name, dict(args or {}), ctx)
