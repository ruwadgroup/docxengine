/**
 * Structured errors per spec/errors.json.
 *
 * In-language these are thrown as `ToolError`; at the CLI/MCP boundary they
 * serialize as `{"error": code, "message": str, "suggestions": [str]}`.
 */
import { ERROR_SPECS } from "./specdata/index.js";

/** Every error code defined in spec/errors.json (the closed set). */
export const ERROR_CODES: ReadonlySet<string> = new Set(ERROR_SPECS.map((e) => e.code));

export class ToolError extends Error {
  readonly code: string;
  readonly suggestions: readonly string[];

  constructor(code: string, message: string, suggestions: readonly string[] = []) {
    super(message);
    this.name = "ToolError";
    this.code = code;
    this.suggestions = suggestions;
  }

  toJSON(): { error: string; message: string; suggestions: string[] } {
    return { error: this.code, message: this.message, suggestions: [...this.suggestions] };
  }
}
