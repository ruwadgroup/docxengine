"""Fill an MSA template, validate, and convert to Markdown via the docxengine package.

Prereq: pip install -e ../../python && python make_input.py

Phase 2: docx_template_fill and docx_convert land in Phase 2 (ROADMAP.md); this
script is contract-accurate and runs once they ship. It converts to Markdown
(in-engine, no external tooling) rather than PDF, since PDF needs LibreOffice —
swap `to="md"` for `to="pdf"` with a `path` once a renderer is installed.
"""

from docxengine import call

DATA = {
    "EffectiveDate": "2026-07-01",
    "Client": "GlobalTech",
    "obligations": [
        {"text": "Deliver the Q3 report"},
        {"text": "Maintain the SLA"},
    ],
}


def main() -> None:
    filled = call(
        "docx_template_fill",
        {"template": "msa-template.docx", "data": DATA, "syntax": "mustache"},
    )
    print(f"filled {filled['filled']} placeholder(s); unfilled: {filled['unfilled']}")
    assert not filled["unfilled"], filled["unfilled"]

    check = call("docx_validate", {"doc_id": filled["doc_id"]})
    assert check["valid"], check["issues"]

    md = call("docx_convert", {"doc_id": filled["doc_id"], "to": "md"})
    print("--- msa-globaltech.md ---")
    print(md["content"])
    # PDF needs LibreOffice:
    # call("docx_convert", {"doc_id": filled["doc_id"], "to": "pdf", "path": "msa-globaltech.pdf"})


if __name__ == "__main__":
    main()
