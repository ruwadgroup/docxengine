"""Redline review via the docxengine package. Prereq: pip install -e ../../python && python make_input.py"""

from docxengine import call


def main() -> None:
    doc = call("docx_open", {"path": "contract.docx"})
    print(f"opened {doc['doc_id']}: {doc['summary']}")

    revs = call("docx_revision", {"doc_id": doc["doc_id"], "op": "list"})
    print(f"revisions: {[(r['author'], r['type']) for r in revs['revisions']]}")

    result = call(
        "docx_revision",
        {"doc_id": doc["doc_id"], "op": "accept", "filter": {"author": "Jane Doe"}},
    )
    print(f"accepted {result['accepted']}; remaining by author: {result['remaining_by_author']}")

    check = call("docx_validate", {"doc_id": doc["doc_id"]})
    assert check["valid"], check["issues"]

    call("docx_save", {"doc_id": doc["doc_id"], "path": "contract-reviewed.docx"})
    print("saved contract-reviewed.docx")


if __name__ == "__main__":
    main()
