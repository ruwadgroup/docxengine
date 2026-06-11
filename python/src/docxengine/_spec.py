"""Packaged spec contracts: tool schemas and error codes.

``_specdata/`` holds verbatim copies of ``spec/tools/*.json`` and
``spec/errors.json`` (synced by ``python/scripts/sync_spec.py``) so the
contracts ship inside the wheel. Public accessors return fresh copies; the
private loaders are cached and must be treated as read-only.
"""

from __future__ import annotations

import copy
import json
from functools import cache
from importlib import resources


@cache
def _tools() -> tuple[dict[str, object], ...]:
    root = resources.files(__package__).joinpath("_specdata/tools")
    schemas = [
        json.loads(entry.read_text(encoding="utf-8"))
        for entry in sorted(root.iterdir(), key=lambda e: e.name)
        if entry.name.endswith(".json")
    ]
    return tuple(schemas)


@cache
def _errors() -> tuple[dict[str, object], ...]:
    raw = resources.files(__package__).joinpath("_specdata/errors.json").read_text(encoding="utf-8")
    return tuple(json.loads(raw)["errors"])


@cache
def _schemas_by_name() -> dict[str, dict[str, object]]:
    return {str(schema["name"]): schema for schema in _tools()}


def spec_tool_names() -> frozenset[str]:
    """Every tool name defined in the spec (implemented or not)."""
    return frozenset(_schemas_by_name())


def input_schema(tool: str) -> dict[str, object] | None:
    """The tool's raw ``input_schema`` (shared, read-only), or ``None`` if unknown."""
    schema = _schemas_by_name().get(tool)
    if schema is None:
        return None
    result = schema.get("input_schema")
    return result if isinstance(result, dict) else {}


def result_schema(tool: str) -> dict[str, object] | None:
    """The tool's raw ``result_schema`` (shared, read-only), or ``None`` if unknown."""
    schema = _schemas_by_name().get(tool)
    if schema is None:
        return None
    result = schema.get("result_schema")
    return result if isinstance(result, dict) else {}


def error_codes() -> frozenset[str]:
    """Every error code defined in ``spec/errors.json``."""
    return frozenset(str(entry["code"]) for entry in _errors())


def tool_schemas() -> list[dict[str, object]]:
    """The full spec tool schemas, alphabetical by name (fresh copies)."""
    return copy.deepcopy(list(_tools()))


def openai_tools() -> list[dict[str, object]]:
    """The tool schemas in the OpenAI function-calling shape."""
    return [
        {
            "type": "function",
            "function": {
                "name": schema["name"],
                "description": schema["description"],
                "parameters": copy.deepcopy(schema.get("input_schema", {})),
            },
        }
        for schema in _tools()
    ]


def anthropic_tools() -> list[dict[str, object]]:
    """The tool schemas in the Anthropic tool-use shape."""
    return [
        {
            "name": schema["name"],
            "description": schema["description"],
            "input_schema": copy.deepcopy(schema.get("input_schema", {})),
        }
        for schema in _tools()
    ]
