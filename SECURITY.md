# Security Policy

## Reporting a vulnerability

**Do not open a public issue for security problems.** Report privately via [GitHub Security Advisories](https://github.com/ruwadgroup/docxengine/security/advisories/new). You'll get an acknowledgement within 72 hours and a remediation plan or assessment within 14 days. Coordinated disclosure is appreciated; we'll credit you unless you prefer otherwise.

## Supported versions

Pre-1.0: only the latest release receives security fixes.

## Threat model

DocxEngine parses untrusted documents and is driven by untrusted (LLM-generated) tool calls. Areas we treat as security surfaces:

- **Malicious archives** — zip bombs, path traversal via crafted part names (`../`), oversized parts, decompression-ratio abuse. The OPC layer canonicalizes part names (relationship targets cannot escape the package root) and enforces caps on part count, total and per-part uncompressed size, and compression ratio **before decompressing** — refused as `doc_too_large`. The caps are tunable via the `DOCXENGINE_MAX_*` environment variables (see [spec/algorithms.md](spec/algorithms.md) §27); defaults are generous for real documents.
- **Hostile XML** — XXE, entity-expansion (billion laughs), and external DTD fetching. The no-DOM scanners never expand entities beyond the five XML built-ins; the DOM-parsed parts (content-types, rels, styles, comments) are screened at the OPC chokepoint, which **rejects any `<!DOCTYPE`/`<!ENTITY` declaration** (`malicious_content`) before a parser sees it. There is no network access during parse, and XML nested past `DOCXENGINE_MAX_XML_DEPTH` is refused.
- **Tool-call injection** — path arguments from agent calls are confined to configured roots; `docx_open`/`docx_save` never follow arbitrary filesystem paths in server deployments.
- **Render adapter** — LibreOffice is invoked headless on untrusted input; deployments should sandbox it (container/seccomp) and treat it as the highest-risk component. The core never requires it. When no `soffice` is present, rendering **auto-fetches** an official LibreOffice build (on by default): the artifact and its checksum are pulled over **HTTPS from The Document Foundation** and the download is **SHA-256-verified against the publisher's `.sha256` sidecar** before it is extracted or executed. Non-HTTPS URLs are refused and the download is size-capped. Disable fetching with `DOCXENGINE_AUTO_FETCH_SOFFICE=0`, pin a mirror with `DOCXENGINE_SOFFICE_MIRROR`, or point `DOCXENGINE_SOFFICE` at a vetted binary. Auto-fetch env vars: `DOCXENGINE_AUTO_FETCH_SOFFICE`, `DOCXENGINE_SOFFICE_CACHE`, `DOCXENGINE_SOFFICE_VERSION`, `DOCXENGINE_SOFFICE_MIRROR`.
- **Secret leakage** — document content may be sensitive; the engine never logs document text at default log levels, and previews/conversions are written only to caller-specified locations.
- **Resource exhaustion** — per-document memory caps and response-size caps (the ~25k-token rule doubles as a DoS bound).

## Scope

In scope: the core engine, the MCP server, and the Python package in this repository. Out of scope: vulnerabilities in Word, LibreOffice, or downstream agent frameworks themselves (report upstream — but tell us if DocxEngine's defaults make exploitation easier).
