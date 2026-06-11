"""Bulk rebrand via the Python SDK. Prereq: pip install -e ../../python && python make_input.py"""

from docxengine import call


def main() -> None:
    doc = call("docx_open", {"path": "report.docx"})
    doc_id = doc["doc_id"]

    hits = call("docx_search", {"doc_id": doc_id, "query": "Acme Corp"})
    print(f"found {hits['n_matches']} occurrence(s)")

    first = call(
        "docx_replace",
        {
            "doc_id": doc_id,
            "old": "Acme Corp",
            "new": "GlobalTech Inc",
            "all": True,
            "track_changes": True,
            "author": "Rebrand Bot",
        },
    )
    print(f"replaced {first['n_replaced']}")

    again = call(
        "docx_replace",
        {
            "doc_id": doc_id,
            "old": "Acme Corp",
            "new": "GlobalTech Inc",
            "all": True,
            "track_changes": True,
            "author": "Rebrand Bot",
        },
    )
    assert again["n_replaced"] == 0, "all:true must be idempotent"

    call("docx_save", {"doc_id": doc_id, "path": "report-rebranded.docx"})
    print("saved report-rebranded.docx")


if __name__ == "__main__":
    main()
