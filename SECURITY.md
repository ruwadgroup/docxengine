# Security Policy

## Reporting a vulnerability

**Do not open a public issue for security problems.** Report privately via [GitHub Security Advisories](https://github.com/ruwadgroup/docxengine/security/advisories/new). You'll get an acknowledgement within 72 hours and a remediation plan or assessment within 14 days. Coordinated disclosure is appreciated; we'll credit you unless you prefer otherwise.

## Supported versions

Pre-1.0: only the latest release receives security fixes.

## Threat model

DocxEngine parses untrusted documents and is driven by untrusted (LLM-generated) tool calls. Areas we treat as security surfaces:

- **Malicious archives** — zip bombs, path traversal via crafted part names (`../`), oversized parts, decompression-ratio abuse. The OPC layer enforces part-name canonicalization and size/ratio limits.
- **Hostile XML** — XXE, entity-expansion (billion laughs), and external DTD fetching. Parsers run with external entities and DTD resolution disabled; no network access during parse.
- **Tool-call injection** — path arguments from agent calls are confined to configured roots; `docx_open`/`docx_save` never follow arbitrary filesystem paths in server deployments.
- **Render adapter** — LibreOffice is invoked headless on untrusted input; deployments should sandbox it (container/seccomp) and treat it as the highest-risk component. The core never requires it.
- **Secret leakage** — document content may be sensitive; the engine never logs document text at default log levels, and previews/conversions are written only to caller-specified locations.
- **Resource exhaustion** — per-document memory caps and response-size caps (the ~25k-token rule doubles as a DoS bound).

## Scope

In scope: the core engine, MCP server, Python/JS packages, and conformance harness in this repository. Out of scope: vulnerabilities in Word, LibreOffice, or downstream agent frameworks themselves (report upstream — but tell us if DocxEngine's defaults make exploitation easier).
