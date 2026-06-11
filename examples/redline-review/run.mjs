// Redline review via the JS SDK. Prereq: pnpm --dir ../../js build && python make_input.py
import { call } from "../../js/dist/index.js";

const doc = await call("docx_open", { path: "contract.docx" });
console.log(`opened ${doc.doc_id}: ${doc.summary}`);

const revs = await call("docx_revision", { doc_id: doc.doc_id, op: "list" });
console.log(
  "revisions:",
  revs.revisions.map((r) => [r.author, r.type]),
);

const result = await call("docx_revision", {
  doc_id: doc.doc_id,
  op: "accept",
  filter: { author: "Jane Doe" },
});
console.log(`accepted ${result.accepted}; remaining:`, result.remaining_by_author);

const check = await call("docx_validate", { doc_id: doc.doc_id });
if (!check.valid) throw new Error(JSON.stringify(check.issues));

await call("docx_save", { doc_id: doc.doc_id, path: "contract-reviewed.docx" });
console.log("saved contract-reviewed.docx");
