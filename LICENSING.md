# Licensing notes

DocxEngine's own code is licensed under [Apache-2.0](LICENSE). This document clarifies the licensing posture of things DocxEngine touches but does not ship.

## External tools (orchestrated, not bundled)

| Tool           | License     | How DocxEngine uses it                                                                                               |
| -------------- | ----------- | -------------------------------------------------------------------------------------------------------------------- |
| LibreOffice    | MPL-2.0     | Optional render adapter — invoked as an external headless process for PDF/PNG previews. Never linked, never bundled. |
| Microsoft Word | proprietary | Never invoked. Used only as the external ground truth for fidelity testing ("does Word open it without repair?").    |

Because these run as separate processes under their own licenses, they impose no licensing obligations on DocxEngine or its users beyond their own installation terms.

## Prior art (studied, not copied)

The design draws on published research and open projects — approaches were studied, **no code was copied**:

- **Anthropic docx skill** (anthropics/skills) — proprietary-licensed; its unpack→patch→validate→render approach informed the architecture. Do not vendor its code or reference files into this repository.
- **SecurityRonin/docx-mcp**, **pablospe/docx-editor**, **UseJunior/safe-docx** — open projects whose validation-gate and hash-anchor patterns this design builds on. Check their licenses before porting any code, and attribute in NOTICE if any is ever adapted.
- **SWE-agent** (Yang et al., NeurIPS 2024), Anthropic engineering posts, ECMA-376/ISO 29500 — cited in [docs/research/prior-art.md](docs/research/prior-art.md).

## Specifications

ECMA-376 (OOXML) is freely available from Ecma International; implementing it carries no licensing obligation. `[MS-DOCX]` extensions are documented under Microsoft's Open Specifications Promise.

## Contributions

By contributing you agree your contribution is licensed under Apache-2.0 (inbound = outbound). No CLA is required.
