/**
 * Render-surface tool wrapper: `docx_render_preview` (algorithms.md §24).
 *
 * Resolves the `doc_id` against the session and delegates to the render adapter
 * (`render.ts`). Preview never errors when no renderer is installed — it returns
 * the structural fallback. The render adapter keeps all `node:child_process`
 * access behind lazy imports so this module stays browser-importable.
 */
import {
  renderPreview,
  type DocxRenderPreviewArgs,
  type DocxRenderPreviewResult,
} from "./render.js";
import type { Session } from "./session.js";

export type { DocxRenderPreviewArgs, DocxRenderPreviewResult };

/** Render preview pages (resource links) or the structural fallback (§24). */
export function docxRenderPreview(
  session: Session,
  args: DocxRenderPreviewArgs,
): DocxRenderPreviewResult {
  const doc = session.get(args.doc_id);
  return renderPreview(doc, args);
}
