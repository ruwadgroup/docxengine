"""DocxEngine — AI-optimized DOCX manipulation engine (Python implementation).

Deterministic OOXML editing with tracked changes, hash-anchored addressing, and a
token-efficient agent view. The shared algorithm spec lives in ``spec/algorithms.md``.
"""

from ._anchors import (
    AnchorEntry,
    anchor_hash,
    build_anchor_index,
    normalized_text,
    paragraph_anchor,
    table_anchor,
)
from ._comments import docx_comment
from ._convert import docx_convert
from ._create import docx_create
from ._dispatch import call
from ._document import Document, Paragraph
from ._errors import ToolError
from ._fields import docx_field
from ._lists import docx_list
from ._media import docx_media
from ._opc import ContentTypes, Package, Relationship, rels_part_for, resolve_rel_target
from ._projector import (
    ProjectedBlock,
    project_body,
    project_outline,
    project_read,
    project_search,
)
from ._render import docx_render_preview
from ._sections import docx_section
from ._session import OpenDocument, Session
from ._spec import anthropic_tools, openai_tools, tool_schemas
from ._styles import docx_format, docx_style
from ._tables import docx_table
from ._template import docx_template_fill
from ._tools_edit import (
    docx_delete,
    docx_edit_paragraph,
    docx_insert,
    docx_replace,
    docx_revision,
)
from ._tools_lifecycle import docx_repair, docx_save, docx_validate
from ._tools_read import docx_open, docx_outline, docx_read, docx_search
from ._validate import Issue, repair_package, validate_package

__version__ = "0.0.0"

__all__ = [
    "AnchorEntry",
    "ContentTypes",
    "Document",
    "Issue",
    "OpenDocument",
    "Package",
    "Paragraph",
    "ProjectedBlock",
    "Relationship",
    "Session",
    "ToolError",
    "__version__",
    "anchor_hash",
    "anthropic_tools",
    "build_anchor_index",
    "call",
    "docx_comment",
    "docx_convert",
    "docx_create",
    "docx_delete",
    "docx_edit_paragraph",
    "docx_field",
    "docx_format",
    "docx_insert",
    "docx_list",
    "docx_media",
    "docx_open",
    "docx_outline",
    "docx_read",
    "docx_render_preview",
    "docx_repair",
    "docx_replace",
    "docx_revision",
    "docx_save",
    "docx_search",
    "docx_section",
    "docx_style",
    "docx_table",
    "docx_template_fill",
    "docx_validate",
    "normalized_text",
    "openai_tools",
    "paragraph_anchor",
    "project_body",
    "project_outline",
    "project_read",
    "project_search",
    "rels_part_for",
    "repair_package",
    "resolve_rel_target",
    "table_anchor",
    "tool_schemas",
    "validate_package",
]
