/**
 * Line-oriented JSON CLI per spec/algorithms.md §11 (the conformance driver).
 *
 * - One JSON request object per stdin line: {"tool": "docx_replace", "args": {…}}.
 * - Exactly one JSON object per stdout line, in request order: the tool's
 *   result, or {"error": code, "message": …, "suggestions": […]}.
 * - doc_ids persist for the process lifetime; blank lines are ignored;
 *   EOF on stdin → exit 0. stderr is free-form logging.
 *
 * Entry point: `node js/dist/cli.js`.
 */
import * as readline from "node:readline";

import { dispatch } from "./dispatch.js";
import { ToolError } from "./errors.js";
import { Session } from "./session.js";

const LINE_HINT = 'Send one {"tool": …, "args": {…}} object per line.';

/** Answer one request line with one response object. */
function respond(session: Session, line: string): Record<string, unknown> {
  let request: unknown;
  try {
    request = JSON.parse(line);
  } catch (e) {
    return new ToolError(
      "invalid_args",
      `Request line is not valid JSON: ${(e as Error).message}.`,
      [LINE_HINT],
    ).toJSON();
  }
  const tool = (request as { tool?: unknown } | null)?.tool;
  if (
    typeof request !== "object" ||
    request === null ||
    Array.isArray(request) ||
    typeof tool !== "string"
  ) {
    return new ToolError("invalid_args", "Request must be a JSON object with a string 'tool'.", [
      LINE_HINT,
    ]).toJSON();
  }
  try {
    return dispatch(session, tool, (request as { args?: unknown }).args) as Record<string, unknown>;
  } catch (e) {
    if (e instanceof ToolError) return e.toJSON();
    throw e; // engine bug: crash loudly rather than mask a parity break
  }
}

const session = new Session();
const rl = readline.createInterface({ input: process.stdin, terminal: false });
rl.on("line", (raw) => {
  const line = raw.trim();
  if (line === "") return;
  process.stdout.write(`${JSON.stringify(respond(session, line))}\n`);
});
rl.on("close", () => {
  process.exitCode = 0;
});
