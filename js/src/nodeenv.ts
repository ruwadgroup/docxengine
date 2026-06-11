/**
 * Lazy access to Node built-ins so the browser-safe paths — open from bytes,
 * read/search (docs/sdks/javascript.md §Browser) — never carry a top-level
 * `node:fs` import. Filesystem-touching code resolves the modules at call
 * time via `process.getBuiltinModule` (Node ≥22): bundlers see no static
 * `node:` specifier, and outside Node the call throws a plain Error that the
 * call sites wrap into the appropriate ToolError.
 */

interface BuiltinLoader {
  getBuiltinModule?: (id: string) => unknown;
}

function builtin<T>(id: string): T {
  const proc = (globalThis as { process?: BuiltinLoader }).process;
  const mod = proc?.getBuiltinModule?.(id);
  if (mod === undefined) {
    throw new Error(
      `${id} is unavailable in this runtime (Node ≥22 required); ` +
        "in the browser, open documents from bytes and skip path-based save",
    );
  }
  return mod as T;
}

/** `node:fs`, resolved at call time (throws outside Node). */
export function nodeFs(): typeof import("node:fs") {
  return builtin<typeof import("node:fs")>("node:fs");
}

/** `node:path`, resolved at call time (throws outside Node). */
export function nodePath(): typeof import("node:path") {
  return builtin<typeof import("node:path")>("node:path");
}

/** `node:child_process`, resolved at call time (throws outside Node). */
export function nodeChildProcess(): typeof import("node:child_process") {
  return builtin<typeof import("node:child_process")>("node:child_process");
}

/** `node:os`, resolved at call time (throws outside Node). */
export function nodeOs(): typeof import("node:os") {
  return builtin<typeof import("node:os")>("node:os");
}
