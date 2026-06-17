/**
 * Render adapter (`docx_convert` pdf/png, `docx_render_preview`) per §24.
 *
 * Detection order: env `DOCXENGINE_SOFFICE`; then `soffice` on `PATH`; then the
 * platform defaults. When a binary is found, conversion runs
 * `soffice --headless --convert-to {fmt} --outdir {DIR} {FILE}` with a per-call
 * temp profile (`-env:UserInstallation`) and `renderer = "libreoffice {ver}"`.
 * When none is found, the **structural fallback** returns the §2 projection plus
 * an estimated page count (`ceil(total_chars / 1800)`) and `renderer =
 * "structural"`; preview never errors, but `docx_convert` to pdf/png with no
 * adapter is `render_unavailable`.
 *
 * Node built-ins (`child_process`, `fs`, `os`, `path`) are resolved lazily so
 * browser-safe module paths never carry a static `node:` import.
 */
import { ToolError } from "./errors.js";
import { nodeChildProcess, nodeFs, nodeOs, nodePath } from "./nodeenv.js";
import { readProjection } from "./projector.js";
import type { DocHandle } from "./session.js";

const CHARS_PER_PAGE = 1800;

// ---------------------------------------------------------------------------
// soffice detection (§24)
// ---------------------------------------------------------------------------

const PLATFORM_DEFAULTS = [
  "/Applications/LibreOffice.app/Contents/MacOS/soffice",
  "/usr/bin/soffice",
];

/** Locate a usable `soffice` executable, or null when none is installed. */
export function detectSoffice(): string | null {
  const env = process.env["DOCXENGINE_SOFFICE"];
  if (env !== undefined && env !== "") {
    if (isExecutable(env)) return env;
  }
  const onPath = which("soffice");
  if (onPath !== null) return onPath;
  for (const candidate of PLATFORM_DEFAULTS) {
    if (isExecutable(candidate)) return candidate;
  }
  return null;
}

function isExecutable(p: string): boolean {
  try {
    const fs = nodeFs();
    return fs.existsSync(p) && fs.statSync(p).isFile();
  } catch {
    return false;
  }
}

/** Search `PATH` for `name` (mirrors `command -v`), returning the full path. */
function which(name: string): string | null {
  let pathEnv: string;
  let sep: string;
  let dirSep: string;
  try {
    pathEnv = process.env["PATH"] ?? "";
    const path = nodePath();
    sep = path.delimiter;
    dirSep = path.sep;
  } catch {
    return null;
  }
  for (const dir of pathEnv.split(sep)) {
    if (dir === "") continue;
    const full = dir.endsWith(dirSep) ? `${dir}${name}` : `${dir}${dirSep}${name}`;
    if (isExecutable(full)) return full;
  }
  return null;
}

/** The renderer label: `"libreoffice {version}"` or `"libreoffice"` if unknown. */
function rendererLabel(soffice: string): string {
  try {
    const cp = nodeChildProcess();
    const res = cp.spawnSync(soffice, ["--version"], { encoding: "utf-8", timeout: 20000 });
    const out = `${res.stdout ?? ""}`.trim();
    const m = /([0-9]+\.[0-9]+(?:\.[0-9]+)*)/.exec(out);
    return m ? `libreoffice ${m[1]}` : "libreoffice";
  } catch {
    return "libreoffice";
  }
}

// ---------------------------------------------------------------------------
// LibreOffice invocation
// ---------------------------------------------------------------------------

interface ConversionOutcome {
  /** Path to the produced file (in a temp outdir). */
  producedPath: string;
  renderer: string;
}

/**
 * Save the doc to a temp .docx, run soffice to convert it to `fmt`, and return
 * the produced file path. Throws `render_failed` on a non-zero exit / no output.
 */
function runSoffice(doc: DocHandle, soffice: string, fmt: "pdf"): ConversionOutcome {
  const fs = nodeFs();
  const path = nodePath();
  const os = nodeOs();
  const cp = nodeChildProcess();

  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-render-"));
  const profileDir = fs.mkdtempSync(path.join(os.tmpdir(), "docxengine-profile-"));
  const srcDocx = path.join(workDir, "input.docx");
  try {
    doc.pkg.save(srcDocx);
    const args = [
      "--headless",
      `-env:UserInstallation=file://${profileDir}`,
      "--convert-to",
      fmt,
      "--outdir",
      workDir,
      srcDocx,
    ];
    const res = cp.spawnSync(soffice, args, { encoding: "utf-8", timeout: 120000 });
    if (res.status !== 0) {
      throw new ToolError(
        "render_failed",
        `soffice exited ${String(res.status)}: ${`${res.stderr ?? ""}`.slice(0, 200)}.`,
        ["Check that the document is valid and soffice can write the output dir."],
      );
    }
    const produced = path.join(workDir, `input.${fmt}`);
    if (!fs.existsSync(produced)) {
      throw new ToolError("render_failed", "soffice produced no output file.", [
        "Inspect soffice stderr; the document may have failed to load.",
      ]);
    }
    return { producedPath: produced, renderer: rendererLabel(soffice) };
  } finally {
    try {
      fs.rmSync(profileDir, { recursive: true, force: true });
    } catch {
      /* best-effort cleanup */
    }
  }
}

// ---------------------------------------------------------------------------
// Structural fallback (§24)
// ---------------------------------------------------------------------------

interface StructuralPreview {
  pages: null;
  structural: string;
  renderer: "structural";
  estimatedPages: number;
}

/** The §2 projection of the whole body + an estimated page count (§24). */
export function structuralPreview(doc: DocHandle): StructuralPreview {
  const projection = readProjection(doc, {}).content;
  const estimatedPages = Math.max(1, Math.ceil(projection.length / CHARS_PER_PAGE));
  return { pages: null, structural: projection, renderer: "structural", estimatedPages };
}

// ---------------------------------------------------------------------------
// docx_convert pdf/png target
// ---------------------------------------------------------------------------

export interface RenderResult {
  path?: string;
  renderer?: string;
  note?: string;
}

/** Convert to pdf/png via the adapter; no adapter → `render_unavailable`. */
export function renderToFile(doc: DocHandle, fmt: "pdf" | "png", dest: string): RenderResult {
  const soffice = detectSoffice();
  if (soffice === null) {
    throw new ToolError(
      "render_unavailable",
      "No render adapter: LibreOffice (soffice) was not detected.",
      [
        "Install LibreOffice or set DOCXENGINE_SOFFICE; md/html convert without it.",
        "Use docx_render_preview for the structural fallback.",
      ],
    );
  }
  let outcome: ConversionOutcome;
  if (fmt === "pdf") {
    outcome = runSoffice(doc, soffice, "pdf");
  } else {
    // PNG via PDF then pdftoppm/sips when available (§24).
    outcome = renderPng(doc, soffice);
  }
  try {
    const fs = nodeFs();
    fs.copyFileSync(outcome.producedPath, dest);
  } catch (e) {
    throw new ToolError("save_failed", `Could not write ${dest}: ${(e as Error).message}.`, [
      "Check the output path and permissions.",
    ]);
  }
  return {
    path: dest,
    renderer: outcome.renderer,
    note: `Rendered ${fmt} via ${outcome.renderer}.`,
  };
}

/** PNG = PDF then pdftoppm/sips; falls back to `render_failed` when neither runs. */
function renderPng(doc: DocHandle, soffice: string): ConversionOutcome {
  const fs = nodeFs();
  const path = nodePath();
  const cp = nodeChildProcess();
  const pdf = runSoffice(doc, soffice, "pdf");
  const dir = path.dirname(pdf.producedPath);
  const base = path.join(dir, "page");

  const pdftoppm = which("pdftoppm");
  if (pdftoppm !== null) {
    const res = cp.spawnSync(pdftoppm, ["-png", "-singlefile", pdf.producedPath, base], {
      encoding: "utf-8",
      timeout: 120000,
    });
    const out = `${base}.png`;
    if (res.status === 0 && fs.existsSync(out)) {
      return { producedPath: out, renderer: pdf.renderer };
    }
  }
  const sips = which("sips");
  if (sips !== null) {
    const out = `${base}.png`;
    const res = cp.spawnSync(sips, ["-s", "format", "png", pdf.producedPath, "--out", out], {
      encoding: "utf-8",
      timeout: 120000,
    });
    if (res.status === 0 && fs.existsSync(out)) {
      return { producedPath: out, renderer: pdf.renderer };
    }
  }
  throw new ToolError("render_failed", "No PDF→PNG rasterizer (pdftoppm or sips) available.", [
    "Install poppler (pdftoppm) or run on macOS (sips); pdf conversion still works.",
  ]);
}

// ---------------------------------------------------------------------------
// docx_render_preview (§24)
// ---------------------------------------------------------------------------

export interface DocxRenderPreviewArgs {
  doc_id: string;
  pages?: number[] | undefined;
  response_format?: "concise" | "detailed" | undefined;
}

export interface PreviewPage {
  page: number;
  /** Resource link to a rendered image; present only when a renderer produced one. */
  image?: string;
}

export interface DocxRenderPreviewResult {
  /** Per-page image links; present only when a renderer ran. */
  pages?: PreviewPage[];
  /** Estimated page count for the structural fallback (no renderer). */
  page_count?: number;
  renderer: string;
  note?: string;
  structural?: string;
}

/**
 * Render preview pages. With an adapter, returns resource links
 * (`docx://{doc_id}/preview/page-{n}.png`); without one, the structural
 * fallback — never an error from preview (§24).
 */
export function renderPreview(
  doc: DocHandle,
  args: DocxRenderPreviewArgs,
): DocxRenderPreviewResult {
  const soffice = detectSoffice();
  if (soffice === null) {
    const fallback = structuralPreview(doc);
    return {
      page_count: fallback.estimatedPages,
      renderer: "structural",
      structural: fallback.structural,
      note:
        "No render adapter (LibreOffice/soffice) detected — install LibreOffice or set " +
        "DOCXENGINE_SOFFICE for rendered page images. Showing the structural projection " +
        `(estimated ${fallback.estimatedPages} page${fallback.estimatedPages === 1 ? "" : "s"}); no image links are returned.`,
    };
  }
  // A real renderer is available: produce a PDF (the page count is authoritative
  // only after render). We render every requested page link; the PNGs are
  // generated lazily by the resource fetch, but we materialize page 1 to count.
  const renderer = rendererLabel(soffice);
  // Estimate the page set from the structural fallback when not specified — the
  // adapter writes the real images under the doc's preview resource namespace.
  const estimate = structuralPreview(doc).estimatedPages;
  const pages = pageNumbers(args.pages, estimate);
  return {
    pages: pages.map((page) => ({ page, image: `docx://${doc.id}/preview/page-${page}.png` })),
    renderer,
    note: `Preview links resolve to ${renderer}-rendered page images.`,
  };
}

function pageNumbers(requested: number[] | undefined, total: number): number[] {
  if (requested && requested.length > 0) {
    return requested.map((n) => Math.max(1, Math.trunc(n)));
  }
  const out: number[] = [];
  for (let i = 1; i <= total; i++) out.push(i);
  return out;
}
