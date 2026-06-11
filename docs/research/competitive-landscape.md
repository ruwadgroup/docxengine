# Competitive landscape

| Tool                   | Edits existing docs w/ fidelity | Tracked changes | Comments       | Footnotes   | Agent-native (stable IDs, token-efficient view) | Verification              | Server-side / offline |
| ---------------------- | ------------------------------- | --------------- | -------------- | ----------- | ----------------------------------------------- | ------------------------- | --------------------- |
| python-docx            | partial                         | ✗               | recent/limited | recent      | ✗                                               | ✗                         | ✓                     |
| docx (dolanmiu, JS)    | generation-focused              | ✗               | ✗              | partial     | ✗                                               | ✗                         | ✓                     |
| docxtemplater          | template-only                   | ✗               | ✗              | paid module | ✗ (placeholders)                                | ✗                         | ✓                     |
| Pandoc                 | lossy round-trip                | read-only       | lossy          | partial     | ✗                                               | ✗                         | ✓                     |
| LibreOffice headless   | ✓ (heavy)                       | via UNO         | via UNO        | ✓           | ✗                                               | ✓ (slow)                  | ✓                     |
| Office.js              | ✓                               | ✓               | ✓              | ✓           | ✗                                               | host renders              | ✗ (host only)         |
| Anthropic docx skill   | ✓ (manual XML)                  | ✓               | ✓              | ✓           | partial (grep/line-nums)                        | ✓ (render)                | ✓                     |
| SecurityRonin/docx-mcp | ✓                               | ✓               | ✓              | ✓           | partial (para_id)                               | validate only             | ✓                     |
| pablospe/docx-editor   | ✓                               | ✓               | ✓              | partial     | ✓ (hash anchors)                                | ✗                         | ✓                     |
| **DocxEngine**         | **✓**                           | **✓**           | **✓**          | **✓**       | **✓ (hash anchors + MD view)**                  | **✓ (render + validate)** | **✓**                 |

## What makes this design different

DocxEngine is the only entry combining:

1. **Direct-XML fidelity** — tracked changes, comments, footnotes survive (proven approach: docx skill, docx-mcp);
2. **An agent-native view** — token-efficient Markdown projection with hash-validated anchors (proven addressing: docx-editor, safe-docx);
3. **Always-on validation + repair** — preventing Word "repair" rewrites (the docx-mcp lesson, made non-optional);
4. **A render-based self-verification loop** — the plan-edit-check pattern that wins in the PPTPilot evidence;
5. **Triple distribution from one conformance-tested contract** — MCP + idiomatic Python + idiomatic JS, with thin adapters.

## Notable per-tool gaps (receipts)

- **python-docx** — tracked-changes request open since **Dec 2016** (#340, also #1025); reading `.text` of a doc with unaccepted revisions returns wrong text (#566); comments only landed in 1.2.0. Forks (bayoo-docx, python-docx-oss) patch gaps piecemeal.
- **docx (dolanmiu)** — ~10.5M weekly downloads; excellent declarative _builder_, not a surgical editor of arbitrary existing files.
- **docxtemplater** — strong mustache merge; HTML/image/footnote modules are paid; no arbitrary structural edits.
- **Pandoc** — best-in-class DOCX↔Markdown extraction (`--track-changes=all`) but round-trips drop comments on tracked insertions (#9833) and can produce "repair"-prompting output.
- **LibreOffice headless** — high fidelity but each conversion boots much of the office runtime; font substitution shifts layout without metric-compatible fonts.
- **Office.js** — strong object model, but only inside a hosted Word add-in; no server-side/offline path.
- **Most other docx MCP servers** (GongRzhe/Office-Word-MCP-Server, MeterLong/MCP-Doc, aiexplorations/docx-mcp) wrap python-docx and inherit its gaps.
