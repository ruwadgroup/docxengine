<!-- Thanks for contributing to DocxEngine! -->

## What & why

<!-- What does this change and why? Link the issue: Closes #123 -->

## Type

- [ ] feat
- [ ] fix
- [ ] docs
- [ ] refactor / chore
- [ ] ci

## Area

- [ ] core (OPC / XML patcher / run coalescing)
- [ ] anchors (hash index / re-anchoring)
- [ ] projector (Markdown view / outline / search)
- [ ] edit (replace / insert / delete / format / tables)
- [ ] revisions & comments (tracked changes / w:ins / w:del)
- [ ] validate / render (validation gate / repair / preview)
- [ ] mcp (server / transports / resources)
- [ ] py / js / adapters
- [ ] spec / conformance / bench
- [ ] docs

## Testing

<!-- What did you test? Which corpus documents? Did you verify the result
opens in Word/LibreOffice without a "repair" prompt? -->

- Implementations touched: Python / TS / both
- Conformance cases added/updated:

## Checklist

- [ ] **Fidelity preserved** — tracked changes, comments, footnotes, and media survive the edit path; round-trip stays byte-stable modulo normalization.
- [ ] **Validation gate honored** — no path saves a document that fails `docx_validate`; zero Word "repair" prompts.
- [ ] **Anchors stay safe** — edits validate the content hash first and return fresh anchors.
- [ ] **Token economy** — no tool response can exceed ~25k tokens; raw OOXML is not exposed by default.
- [ ] **Parity** — behavior matches across Python and TS, with conformance cases proving it (or the divergence is documented and tracked).
- [ ] Contract changes in `spec/` are versioned with a changelog note.
- [ ] Docs updated if behavior or contract changed.
- [ ] Commits follow Conventional Commits.
