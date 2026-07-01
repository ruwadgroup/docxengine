"""CLI contract tests (algorithms.md §11) — real subprocess over a stdin pipe."""

from __future__ import annotations

import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from conftest import build_docx


def run_cli(lines: list[str]) -> tuple[list[dict[str, object]], int]:
    """Pipe request lines into ``python -m docxengine.cli``; parsed responses + rc."""
    proc = subprocess.run(
        [sys.executable, "-m", "docxengine.cli"],
        input="".join(line + "\n" for line in lines),
        capture_output=True,
        text=True,
        timeout=60,
    )
    responses = [json.loads(line) for line in proc.stdout.splitlines() if line.strip()]
    return responses, proc.returncode


def request(tool: str, **args: object) -> str:
    return json.dumps({"tool": tool, "args": args})


@pytest.fixture
def docx_path(docx_bytes: bytes, tmp_path: Path) -> Path:
    path = tmp_path / "fixture.docx"
    path.write_bytes(docx_bytes)
    return path


class TestRoundTrip:
    def test_open_replace_save(self, docx_path: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        responses, rc = run_cli(
            [
                request("docx_open", path=str(docx_path)),
                request(
                    "docx_replace", doc_id="d1", old="five (5) years", new="three (3) years"
                ),
                request("docx_save", doc_id="d1", path=str(out)),
            ]
        )
        assert rc == 0
        assert len(responses) == 3  # exactly one response line per request, in order
        opened, replaced, saved = responses
        assert opened["doc_id"] == "d1"
        assert opened["n_paragraphs"] == 3
        assert replaced["n_replaced"] == 1
        assert isinstance(replaced["new_anchor"], str)
        assert saved == {"ok": True, "validated": True, "bytes": out.stat().st_size}
        with zipfile.ZipFile(out) as zf:
            assert b"three (3) years" in zf.read("word/document.xml")

    def test_doc_state_persists_for_process_lifetime(self, docx_path: Path) -> None:
        responses, rc = run_cli(
            [
                request("docx_open", path=str(docx_path)),
                request("docx_open", path=str(docx_path)),
                request("docx_read", doc_id="d2"),
            ]
        )
        assert rc == 0
        assert [r.get("doc_id") for r in responses[:2]] == ["d1", "d2"]
        assert "Master Services Agreement" in str(responses[2]["content"])

    def test_save_refusal_round_trips_as_error_object(
        self, docx_bytes: bytes, tmp_path: Path
    ) -> None:
        from conftest import DOCUMENT_XML, FIXTURE_PARTS

        parts = dict(FIXTURE_PARTS)
        parts["word/document.xml"] = DOCUMENT_XML.replace('w:ins w:id="2"', 'w:ins w:id="1"')
        corrupt = tmp_path / "corrupt.docx"
        corrupt.write_bytes(build_docx(parts))
        out = tmp_path / "out.docx"
        responses, rc = run_cli(
            [
                request("docx_open", path=str(corrupt)),
                request("docx_save", doc_id="d1", path=str(out)),
                request("docx_repair", doc_id="d1"),
                request("docx_save", doc_id="d1", path=str(out)),
            ]
        )
        assert rc == 0
        assert responses[1]["error"] == "validation_failed"
        assert not out.exists() or responses[3]["ok"] is True
        assert responses[2]["fixed"] == ["renumbered duplicate revision id 1 -> 2"]
        assert responses[2]["remaining"] == []
        assert responses[3] == {"ok": True, "validated": True, "bytes": out.stat().st_size}


class TestProtocol:
    def test_error_payload_shape(self) -> None:
        responses, rc = run_cli([request("docx_read", doc_id="d404")])
        assert rc == 0
        assert responses == [
            {
                "error": "doc_not_found",
                "message": "Unknown or expired doc_id: d404.",
                "suggestions": ["Call docx_open again."],
            }
        ]

    def test_every_spec_tool_is_implemented(self) -> None:
        # Phase 2 complete: a spec tool with missing args reports invalid_args, not
        # not_implemented. (docx_template_fill needs the 'template' argument.)
        responses, _ = run_cli([request("docx_template_fill", data={})])
        assert responses[0]["error"] == "invalid_args"
        assert "template" in str(responses[0]["message"])

    def test_unknown_tool_is_not_implemented(self) -> None:
        responses, _ = run_cli([request("docx_frobnicate")])
        assert responses[0]["error"] == "not_implemented"

    def test_malformed_json_line_does_not_kill_the_process(self, docx_path: Path) -> None:
        responses, rc = run_cli(
            ["{this is not json", request("docx_open", path=str(docx_path))]
        )
        assert rc == 0
        assert responses[0]["error"] == "invalid_args"
        assert responses[1]["doc_id"] == "d1"

    def test_non_object_request_is_invalid_args(self) -> None:
        responses, _ = run_cli(['["docx_open"]', '{"args": {}}'])
        assert [r["error"] for r in responses] == ["invalid_args", "invalid_args"]

    def test_missing_required_args_report_invalid_args(self) -> None:
        responses, _ = run_cli([request("docx_replace", doc_id="d1")])
        assert responses[0]["error"] == "invalid_args"
        assert "old" in str(responses[0]["message"])

    def test_blank_lines_are_skipped_and_eof_exits_zero(self, docx_path: Path) -> None:
        responses, rc = run_cli(["", "   ", request("docx_open", path=str(docx_path))])
        assert rc == 0
        assert len(responses) == 1
