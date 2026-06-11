/**
 * Lifecycle tools: docx_validate, docx_repair, docx_save (algorithms.md §8/§9).
 *
 * `docx_save` is the always-on gate: it refuses to write any package carrying
 * an error-severity validation issue (`validation_failed`, suggesting
 * `docx_repair`). Warnings never block. Saving never closes the doc_id.
 */
import { ToolError } from "./errors.js";
import { nodeFs } from "./nodeenv.js";
import type { Session } from "./session.js";
import { type ValidationIssue, isValid, repairDoc, validateDoc } from "./validate.js";

type ResponseFormat = "concise" | "detailed";

// ---------------------------------------------------------------------------
// docx_validate
// ---------------------------------------------------------------------------

export interface DocxValidateArgs {
  doc_id: string;
  response_format?: ResponseFormat | undefined;
}

export interface DocxValidateResult {
  valid: boolean;
  issues: ValidationIssue[];
}

/** Run the §8 package checks; issues carry severity/part/message/fix_hint. */
export function docxValidate(session: Session, args: DocxValidateArgs): DocxValidateResult {
  const issues = validateDoc(session.get(args.doc_id));
  return { valid: isValid(issues), issues };
}

// ---------------------------------------------------------------------------
// docx_repair
// ---------------------------------------------------------------------------

export interface DocxRepairArgs {
  doc_id: string;
}

export interface DocxRepairResult {
  fixed: string[];
  remaining: string[];
}

/** Apply the §8a mechanical fixes; reports what was fixed and what remains. */
export function docxRepair(session: Session, args: DocxRepairArgs): DocxRepairResult {
  const doc = session.get(args.doc_id);
  const { fixed, remaining } = repairDoc(doc);
  if (fixed.length > 0) doc.markDirty();
  return { fixed, remaining };
}

// ---------------------------------------------------------------------------
// docx_save (§9)
// ---------------------------------------------------------------------------

export interface DocxSaveArgs {
  doc_id: string;
  path: string;
}

export interface DocxSaveResult {
  ok: boolean;
  validated: boolean;
  bytes: number;
}

/** Validate (§8), then write atomically (§9). Refuses on error-severity issues. */
export function docxSave(session: Session, args: DocxSaveArgs): DocxSaveResult {
  const doc = session.get(args.doc_id);
  const errors = validateDoc(doc).filter((issue) => issue.severity === "error");
  if (errors.length > 0) {
    throw new ToolError("validation_failed", "Package would trigger Word repair; save refused.", [
      "Run docx_repair, then re-validate.",
      ...errors.map((issue) => issue.message),
    ]);
  }
  doc.pkg.save(args.path); // §9 steps 2–4: normalized entries, atomic temp-file rename
  doc.markSaved();
  return { ok: true, validated: true, bytes: nodeFs().statSync(args.path).size };
}

/**
 * Validate (§8 gate) then return the document's .docx bytes — no filesystem.
 *
 * The bytes-out counterpart to {@link docxSave} for embedding contexts: the
 * browser, serverless, blob storage — anywhere persistence is the host's job.
 * Not a wire tool (returning base64 through an agent's context is token waste;
 * the host already holds the `Session`). Refuses an error-severity package
 * exactly as `docxSave` does.
 */
export function exportBytes(session: Session, docId: string): Uint8Array {
  const doc = session.get(docId);
  const errors = validateDoc(doc).filter((issue) => issue.severity === "error");
  if (errors.length > 0) {
    throw new ToolError("validation_failed", "Package would trigger Word repair; export refused.", [
      "Run docx_repair, then re-validate.",
      ...errors.map((issue) => issue.message),
    ]);
  }
  const bytes = doc.pkg.toBytes();
  doc.markSaved();
  return bytes;
}
