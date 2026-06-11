// Fill an MSA template, validate, and convert to Markdown via the JS SDK.
// Prereq: pnpm --dir ../../js build && python make_input.py
//
// Phase 2: docx_template_fill and docx_convert land in Phase 2 (ROADMAP.md);
// this script is contract-accurate and runs once they ship. It converts to
// Markdown (in-engine) rather than PDF, since PDF needs LibreOffice — swap
// to: "md" for to: "pdf" with a path once a renderer is installed.
import { call } from "../../js/dist/index.js";

const data = {
  EffectiveDate: "2026-07-01",
  Client: "GlobalTech",
  obligations: [{ text: "Deliver the Q3 report" }, { text: "Maintain the SLA" }],
};

const filled = await call("docx_template_fill", {
  template: "msa-template.docx",
  data,
  syntax: "mustache",
});
console.log(`filled ${filled.filled} placeholder(s); unfilled:`, filled.unfilled);
if (filled.unfilled.length) throw new Error(JSON.stringify(filled.unfilled));

const check = await call("docx_validate", { doc_id: filled.doc_id });
if (!check.valid) throw new Error(JSON.stringify(check.issues));

const md = await call("docx_convert", { doc_id: filled.doc_id, to: "md" });
console.log("--- msa-globaltech.md ---");
console.log(md.content);
// PDF needs LibreOffice:
// await call("docx_convert", { doc_id: filled.doc_id, to: "pdf", path: "msa-globaltech.pdf" });
