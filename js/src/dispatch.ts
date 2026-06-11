/**
 * Tool dispatcher: route a spec tool name + args object to its implementation.
 *
 * The MVP tools route to their handlers; every other tool defined in
 * `spec/tools/` returns `not_implemented` ("<tool> lands in Phase 2").
 * Argument validation is minimal per the spec: required keys must be present
 * (`invalid_args`), and only schema-declared properties are forwarded.
 *
 * Doc state is per process: `call()` uses one module-level Session; the CLI
 * and MCP layers pass their own via `dispatch()`. Errors are thrown as
 * ToolError in-language and serialize as `{"error", "message", "suggestions"}`
 * at the CLI/MCP boundary.
 */
import { ToolError } from "./errors.js";
import { Session } from "./session.js";
import { TOOL_SPECS } from "./specdata/index.js";
import { docxComment } from "./comments.js";
import { docxConvert } from "./convert.js";
import { docxCreate } from "./create.js";
import { docxField } from "./fields.js";
import { docxList } from "./lists.js";
import { docxMedia } from "./media.js";
import { docxRenderPreview } from "./toolsRender.js";
import { docxSection } from "./sections.js";
import { docxStyle, docxFormat } from "./styles.js";
import { docxTable } from "./tables.js";
import { docxTemplateFill } from "./template.js";
import {
  docxDelete,
  docxEditParagraph,
  docxInsert,
  docxReplace,
  docxRevision,
} from "./toolsEdit.js";
import { docxRepair, docxSave, docxValidate } from "./toolsLifecycle.js";
import { docxOpen, docxOutline, docxRead, docxSearch } from "./toolsRead.js";

type ToolFn = (session: Session, args: Record<string, unknown>) => unknown;

const HANDLERS = {
  docx_open: docxOpen,
  docx_outline: docxOutline,
  docx_read: docxRead,
  docx_search: docxSearch,
  docx_replace: docxReplace,
  docx_edit_paragraph: docxEditParagraph,
  docx_insert: docxInsert,
  docx_delete: docxDelete,
  docx_revision: docxRevision,
  docx_validate: docxValidate,
  docx_repair: docxRepair,
  docx_save: docxSave,
  // Phase 2 — stage 1: tables, styles, lists.
  docx_table: docxTable,
  docx_style: docxStyle,
  docx_format: docxFormat,
  docx_list: docxList,
  // Phase 2 — stage 2: comments, sections, media, fields.
  docx_comment: docxComment,
  docx_section: docxSection,
  docx_media: docxMedia,
  docx_field: docxField,
  // Phase 2 — stage 3: templates, create, convert, render adapter.
  docx_template_fill: docxTemplateFill,
  docx_create: docxCreate,
  docx_convert: docxConvert,
  docx_render_preview: docxRenderPreview,
} as unknown as Readonly<Record<string, ToolFn>>;

/** The Phase-1 tool names (everything else in spec/tools/ declines). */
export const MVP_TOOLS: ReadonlySet<string> = new Set(Object.keys(HANDLERS));

const INPUT_SCHEMAS = new Map<string, Record<string, unknown>>(
  TOOL_SPECS.map((t) => [t.name, t.input_schema]),
);

/**
 * Dispatch one tool call against an explicit session (the CLI and MCP layers
 * use this with their process-lifetime session). Throws ToolError.
 */
export function dispatch(session: Session, tool: string, args: unknown = {}): unknown {
  const schema = INPUT_SCHEMAS.get(tool);
  if (schema === undefined) {
    throw new ToolError("not_implemented", `Tool ${tool} is not defined in spec/tools/.`, [
      "See docs/tools/index.md for the tool catalog.",
    ]);
  }
  const fn = HANDLERS[tool];
  if (fn === undefined) {
    throw new ToolError("not_implemented", `${tool} lands in Phase 2 (see ROADMAP.md)`, []);
  }
  const provided = args ?? {};
  if (typeof provided !== "object" || Array.isArray(provided)) {
    throw new ToolError("invalid_args", `${tool}: args must be a JSON object.`, [
      `Check the tool's input_schema in spec/tools/${tool}.json.`,
    ]);
  }
  const argObj = provided as Record<string, unknown>;
  const required = schema["required"];
  const missing = Array.isArray(required)
    ? (required as string[]).filter((key) => !(key in argObj))
    : [];
  if (missing.length > 0) {
    throw new ToolError(
      "invalid_args",
      `${tool}: missing required argument(s): ${missing.join(", ")}.`,
      [`Check the tool's input_schema in spec/tools/${tool}.json.`],
    );
  }
  const properties = schema["properties"];
  const known =
    properties !== null && typeof properties === "object"
      ? new Set(Object.keys(properties))
      : new Set<string>();
  const filtered: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(argObj)) {
    if (known.has(key)) filtered[key] = value;
  }
  return fn(session, filtered);
}

/** The process-wide session backing `call()` and the native Document surface. */
export const defaultSession = new Session();

/**
 * The contract surface (docs/sdks/javascript.md): identical names, JSON
 * shapes, and errors to the Python package and the MCP server.
 */
export async function call(
  tool: string,
  args: Record<string, unknown> = {},
): Promise<Record<string, unknown>> {
  return dispatch(defaultSession, tool, args) as Record<string, unknown>;
}
