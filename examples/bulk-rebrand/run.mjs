// Bulk rebrand via the JS SDK. Prereq: pnpm --dir ../../js build && python make_input.py
import { call } from "../../js/dist/index.js";

const doc = await call("docx_open", { path: "report.docx" });

const hits = await call("docx_search", { doc_id: doc.doc_id, query: "Acme Corp" });
console.log(`found ${hits.n_matches} occurrence(s)`);

const args = {
  doc_id: doc.doc_id,
  old: "Acme Corp",
  new: "GlobalTech Inc",
  all: true,
  track_changes: true,
  author: "Rebrand Bot",
};
const first = await call("docx_replace", args);
console.log(`replaced ${first.n_replaced}`);

const again = await call("docx_replace", args);
if (again.n_replaced !== 0) throw new Error("all:true must be idempotent");

await call("docx_save", { doc_id: doc.doc_id, path: "report-rebranded.docx" });
console.log("saved report-rebranded.docx");
