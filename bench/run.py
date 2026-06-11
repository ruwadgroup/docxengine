#!/usr/bin/env python3
"""Agent task benchmark runner (stdlib + docxengine).

For each task in ``bench/tasks/*.json`` the runner:

1. generates the fixture it needs — the conformance corpus via
   ``conformance/harness/make_fixtures.py`` (reused, never reimplemented), plus
   the MSA template for the template task,
2. starts the MCP stdio server (``.venv/bin/docxengine-mcp``) as a subprocess
   with ``DOCXENGINE_FIXED_DATE=2026-01-01T00:00:00Z`` so tracked-change output
   is deterministic,
3. drives it over JSON-RPC 2.0 (newline-framed): ``initialize`` then the task's
   ``tools/call`` sequence — an implicit ``docx_open`` of the fixture first (for
   corpus tasks), then each scripted step with the live ``doc_id`` injected and
   the ``{out}`` / ``{template}`` placeholders substituted,
4. collects metrics: wall time per call, total calls, tool errors, and an
   approximate token cost (``len(json)/4`` over every request and response),
5. runs :mod:`bench.checker` against the saved output document,
6. prints a results table and writes ``bench/results.json``.

Exit status: 0 iff every selected task passes its checks (and produced output);
non-zero otherwise. ``--driver scripted`` is the default and only driver today —
the LLM driver (an agent actually choosing the calls) is documented in the
README but not yet implemented. ``--phase2`` includes the Phase 2 tasks (tables,
styles, comments, templates); by default only the MVP tasks run, so the
benchmark is green against the current engine.
"""

from __future__ import annotations

import argparse
import json
import os
import queue
import subprocess
import sys
import threading
import time
from collections import deque
from pathlib import Path
from typing import Any

BENCH_DIR = Path(__file__).resolve().parent
REPO_DIR = BENCH_DIR.parent
TASKS_DIR = BENCH_DIR / "tasks"
OUT_DIR = BENCH_DIR / ".out"
RESULTS_PATH = BENCH_DIR / "results.json"

CONFORMANCE_HARNESS = REPO_DIR / "conformance" / "harness"
CORPUS_DIR = REPO_DIR / "conformance" / "corpus"
TEMPLATE_EXAMPLE = REPO_DIR / "examples" / "template-to-pdf"
MCP_SERVER = REPO_DIR / ".venv" / "bin" / "docxengine-mcp"

sys.path.insert(0, str(BENCH_DIR))
sys.path.insert(0, str(CONFORMANCE_HARNESS))
import checker  # noqa: E402
import make_fixtures  # noqa: E402

FIXED_ENV = {"DOCXENGINE_FIXED_DATE": "2026-01-01T00:00:00Z"}
RESPONSE_TIMEOUT = 60.0


# ---------------------------------------------------------------------------
# MCP stdio client (JSON-RPC 2.0, newline-framed)
# ---------------------------------------------------------------------------


class MCPError(Exception):
    """A transport- or protocol-level failure talking to the server."""


class MCPClient:
    """Speaks the MCP stdio transport to one ``docxengine-mcp`` subprocess."""

    def __init__(self, env: dict[str, str]) -> None:
        try:
            self.proc = subprocess.Popen(
                [str(MCP_SERVER)],
                cwd=str(REPO_DIR),
                env=env,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
            )
        except OSError as exc:
            raise MCPError(f"could not start {MCP_SERVER}: {exc}") from exc
        self._lines: queue.Queue[str | None] = queue.Queue()
        self._stderr_tail: deque[str] = deque(maxlen=5)
        self._next_id = 0
        threading.Thread(target=self._read_loop, daemon=True).start()
        threading.Thread(target=self._stderr_loop, daemon=True).start()
        # Metrics accumulated across the session.
        self.req_chars = 0
        self.resp_chars = 0

    def _read_loop(self) -> None:
        assert self.proc.stdout is not None
        for line in self.proc.stdout:
            self._lines.put(line)
        self._lines.put(None)

    def _stderr_loop(self) -> None:
        assert self.proc.stderr is not None
        for line in self.proc.stderr:
            self._stderr_tail.append(line.rstrip("\n"))

    def _stderr_hint(self) -> str:
        tail = " | ".join(self._stderr_tail)
        return f" (stderr: {tail[:300]})" if tail else ""

    def _send(self, message: dict[str, Any]) -> dict[str, Any]:
        assert self.proc.stdin is not None
        payload = json.dumps(message, ensure_ascii=False)
        self.req_chars += len(payload)
        try:
            self.proc.stdin.write(payload + "\n")
            self.proc.stdin.flush()
        except (BrokenPipeError, OSError) as exc:
            raise MCPError(f"pipe broke sending {message.get('method')}: {exc}") from exc
        try:
            line = self._lines.get(timeout=RESPONSE_TIMEOUT)
        except queue.Empty:
            raise MCPError(f"timed out after {RESPONSE_TIMEOUT}s") from None
        if line is None:
            raise MCPError(f"server exited mid-conversation{self._stderr_hint()}")
        self.resp_chars += len(line.rstrip("\n"))
        try:
            response = json.loads(line)
        except json.JSONDecodeError as exc:
            raise MCPError(f"non-JSON from server: {line[:200]!r}") from exc
        return response

    def initialize(self) -> None:
        self._next_id += 1
        message = {
            "jsonrpc": "2.0",
            "id": self._next_id,
            "method": "initialize",
            "params": {
                "protocolVersion": "2025-03-26",
                "capabilities": {},
                "clientInfo": {"name": "docxengine-bench", "version": "0"},
            },
        }
        response = self._send(message)
        if "error" in response:
            raise MCPError(f"initialize failed: {response['error']}")

    def tools_call(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        """Invoke a tool; return the parsed result payload (engine error JSON if isError)."""
        self._next_id += 1
        message = {
            "jsonrpc": "2.0",
            "id": self._next_id,
            "method": "tools/call",
            "params": {"name": name, "arguments": arguments},
        }
        response = self._send(message)
        if "error" in response:  # JSON-RPC level error (bad request shape)
            raise MCPError(f"tools/call {name} JSON-RPC error: {response['error']}")
        result = response.get("result", {})
        content = result.get("content", [])
        text = content[0]["text"] if content else "{}"
        payload = json.loads(text)
        payload["__isError"] = bool(result.get("isError"))
        return payload

    def close(self) -> None:
        try:
            if self.proc.stdin is not None:
                self.proc.stdin.close()
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()


# ---------------------------------------------------------------------------
# Fixture generation
# ---------------------------------------------------------------------------

_corpus_built = False


def ensure_corpus() -> None:
    global _corpus_built
    if not _corpus_built:
        make_fixtures.build_corpus(CORPUS_DIR, quiet=True)
        _corpus_built = True


def ensure_template() -> Path:
    """Build the MSA template via the example's make_input.py; return its path."""
    template = TEMPLATE_EXAMPLE / "msa-template.docx"
    make_input = TEMPLATE_EXAMPLE / "make_input.py"
    if make_input.exists():
        subprocess.run(
            [str(REPO_DIR / ".venv" / "bin" / "python"), str(make_input)],
            cwd=str(TEMPLATE_EXAMPLE),
            check=True,
            capture_output=True,
        )
    return template


def fixture_path(name: str) -> Path:
    if name == "msa-template":
        return ensure_template()
    ensure_corpus()
    return CORPUS_DIR / name / "input.docx"


# ---------------------------------------------------------------------------
# Task execution
# ---------------------------------------------------------------------------


class TaskResult:
    def __init__(self, name: str) -> None:
        self.name = name
        self.ok = False
        self.calls = 0
        self.errors = 0
        self.repairs = 0
        self.wall_ms = 0.0
        self.call_ms: list[float] = []
        self.tokens = 0
        self.reasons: list[str] = []


def _substitute(value: Any, subs: dict[str, str]) -> Any:
    """Recursively replace {out}/{template} placeholders in script args."""
    if isinstance(value, str):
        for key, repl in subs.items():
            value = value.replace("{" + key + "}", repl)
        return value
    if isinstance(value, dict):
        return {k: _substitute(v, subs) for k, v in value.items()}
    if isinstance(value, list):
        return [_substitute(v, subs) for v in value]
    return value


def run_task(task: dict[str, Any], env: dict[str, str]) -> TaskResult:
    name = task["name"]
    result = TaskResult(name)
    fixture = fixture_path(task["fixture"])
    out_path = OUT_DIR / f"{name}.docx"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink()
    subs = {"out": str(out_path), "template": str(fixture)}

    script = task["script"]
    # Corpus tasks open the fixture implicitly; template tasks open it themselves.
    implicit_open = bool(script) and script[0]["tool"] != "docx_template_fill"

    try:
        client = MCPClient(env)
    except MCPError as exc:
        result.reasons.append(str(exc))
        return result

    started = time.perf_counter()
    doc_id: str | None = None
    try:
        client.initialize()

        if implicit_open:
            doc_id = _timed_call(
                client, result, "docx_open", {"path": str(fixture)}, doc_id=None
            ).get("doc_id")
            if doc_id is None:
                raise MCPError("implicit docx_open returned no doc_id")

        for step in script:
            tool = step["tool"]
            args = _substitute(dict(step.get("args", {})), subs)
            payload = _timed_call(client, result, tool, args, doc_id=doc_id)
            # Template fill mints a fresh doc_id the later steps must thread.
            if "doc_id" in payload and tool == "docx_template_fill":
                doc_id = payload["doc_id"]
    except MCPError as exc:
        result.reasons.append(str(exc))
    finally:
        client.close()

    result.wall_ms = (time.perf_counter() - started) * 1000.0
    result.tokens = (client.req_chars + client.resp_chars) // 4

    if result.reasons:
        return result
    if not out_path.exists():
        result.reasons.append("task produced no output document")
        return result

    # Element-level ground truth on the saved document.
    failures = checker.run_checks(task.get("checks", []), str(out_path))
    result.reasons.extend(failures)
    result.ok = not result.reasons
    return result


def _timed_call(
    client: MCPClient,
    result: TaskResult,
    tool: str,
    args: dict[str, Any],
    *,
    doc_id: str | None,
) -> dict[str, Any]:
    """Inject doc_id (when the tool takes one), time the call, fold in metrics."""
    call_args = dict(args)
    if doc_id is not None and tool != "docx_template_fill" and "doc_id" not in call_args:
        call_args["doc_id"] = doc_id
    start = time.perf_counter()
    payload = client.tools_call(tool, call_args)
    elapsed = (time.perf_counter() - start) * 1000.0
    result.calls += 1
    result.call_ms.append(elapsed)
    if payload.pop("__isError", False):
        result.errors += 1
        result.reasons.append(f"{tool} errored: {json.dumps(payload)[:200]}")
    # A Word-repair event surfaces as a validation_failed save or a repaired note.
    if tool == "docx_validate" and payload.get("valid") is False:
        result.repairs += 1
    return payload


# ---------------------------------------------------------------------------
# Loading + reporting
# ---------------------------------------------------------------------------


def load_tasks(only: str | None, include_phase2: bool) -> list[dict[str, Any]]:
    tasks: list[dict[str, Any]] = []
    for path in sorted(TASKS_DIR.glob("*.json")):
        task = json.loads(path.read_text(encoding="utf-8"))
        if only is not None:
            if task["name"] == only:
                tasks.append(task)
            continue
        if task.get("phase") == 2 and not include_phase2:
            continue
        tasks.append(task)
    if only is not None and not tasks:
        sys.exit(f"unknown task: {only}")
    return tasks


def print_table(results: list[TaskResult]) -> None:
    width = max((len(r.name) for r in results), default=4)
    header = (
        f"{'TASK'.ljust(width)}  {'RESULT':>6}  {'CALLS':>5}  "
        f"{'ERRORS':>6}  {'REPAIR':>6}  {'TOKENS':>6}  {'WALL_MS':>8}"
    )
    print(header)
    print("-" * len(header))
    for r in results:
        verdict = "pass" if r.ok else "FAIL"
        print(
            f"{r.name.ljust(width)}  {verdict:>6}  {r.calls:>5}  "
            f"{r.errors:>6}  {r.repairs:>6}  {r.tokens:>6}  {r.wall_ms:>8.1f}"
        )


def write_results(results: list[TaskResult], driver: str) -> None:
    payload = {
        "driver": driver,
        "fixed_date": FIXED_ENV["DOCXENGINE_FIXED_DATE"],
        "n_tasks": len(results),
        "n_passed": sum(1 for r in results if r.ok),
        "tasks": [
            {
                "name": r.name,
                "ok": r.ok,
                "calls": r.calls,
                "errors": r.errors,
                "word_repairs": r.repairs,
                "tokens": r.tokens,
                "wall_ms": round(r.wall_ms, 1),
                "call_ms": [round(ms, 1) for ms in r.call_ms],
                "failures": r.reasons,
            }
            for r in results
        ],
    }
    RESULTS_PATH.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="DocxEngine agent task benchmark")
    parser.add_argument("--task", help="run a single task by name")
    parser.add_argument(
        "--driver",
        choices=["scripted"],
        default="scripted",
        help="task driver (only 'scripted' is implemented; 'llm' is documented but not built)",
    )
    parser.add_argument(
        "--phase2",
        action="store_true",
        help="include Phase 2 tasks (tables, styles, comments, templates)",
    )
    opts = parser.parse_args(argv)

    if not MCP_SERVER.exists():
        sys.exit(f"MCP server not found at {MCP_SERVER} — install the package: pip install -e python")

    tasks = load_tasks(opts.task, opts.phase2)
    env = {**os.environ, **FIXED_ENV}

    results: list[TaskResult] = []
    for task in tasks:
        results.append(run_task(task, env))

    print_table(results)
    write_results(results, opts.driver)

    failures = [r for r in results if not r.ok]
    if failures:
        print()
        for r in failures:
            for reason in r.reasons:
                print(f"FAIL {r.name}: {reason}")
        print(f"\n{len(failures)} of {len(results)} task(s) failed")
        return 1
    print(f"\nall {len(results)} task(s) passed (driver={opts.driver})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
