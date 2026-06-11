#!/usr/bin/env python3
"""Cross-implementation conformance harness (stdlib only).

Usage:
    python conformance/harness/run.py [--impl py|js] [case-name ...]

For every case in conformance/cases/*.json the harness:

1. regenerates the corpus (make_fixtures.py — deterministic, cheap),
2. spawns each implementation CLI (one process per case per impl) and speaks
   the line-oriented JSON protocol of spec/algorithms.md §11:
   docx_open(corpus input) -> case tool (doc_id injected) -> docx_save when the
   tool mutates -> invariant probes,
3. checks `expect` (partial result match, or error code),
4. cross-compares the two implementations: the case-tool result objects must be
   deep-equal after masking volatile keys ("doc_id", "bytes"), and the saved
   output packages must be equal per the §10 normalization (decompressed bytes,
   canonical-XML fallback for .xml/.rels parts).

Environment: every CLI runs with DOCXENGINE_FIXED_DATE=2026-06-10T00:00:00Z and
DOCXENGINE_AUTHOR=Harness so tracked-change output is deterministic.

Expectation matching (`expect.result`):
- dict      -> subset: each expected key must exist in the actual and match;
- list      -> ["$contains", m1, m2, ...] means each matcher must match a
               *distinct* element (order-free, no length constraint);
               any other list is a same-length positional match;
- scalar    -> equality.
`expect.error` -> the response must be {"error": <code>, ...} with that code.

Invariants (`expect.invariants`):
- "roundtrip"           saved output ==(§10) corpus input;
- "no_word_repair"      docx_open the saved output, docx_validate it -> valid;
- "revisions_preserved" output word/document.xml still contains w:ins/w:del;
- "validate_clean"      docx_validate the live doc_id after the tool -> valid.

Exit status: 0 iff every case passes on every requested implementation and
(when both run) every parity comparison passes.
"""

from __future__ import annotations

import argparse
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import zipfile
from collections import deque
from pathlib import Path
from typing import Any
from xml.parsers import expat

HARNESS_DIR = Path(__file__).resolve().parent
CONFORMANCE_DIR = HARNESS_DIR.parent
REPO_DIR = CONFORMANCE_DIR.parent
CORPUS_DIR = CONFORMANCE_DIR / "corpus"
CASES_DIR = CONFORMANCE_DIR / "cases"
OUT_DIR = CONFORMANCE_DIR / ".out"

sys.path.insert(0, str(HARNESS_DIR))
import make_fixtures  # noqa: E402

FIXED_ENV = {
    "DOCXENGINE_FIXED_DATE": "2026-06-10T00:00:00Z",
    "DOCXENGINE_AUTHOR": "Harness",
}

IMPLS: dict[str, list[str]] = {
    "py": [str(REPO_DIR / ".venv" / "bin" / "python"), "-m", "docxengine.cli"],
    "js": ["node", str(REPO_DIR / "js" / "dist" / "cli.js")],
}

ALWAYS_MUTATING = {"docx_replace", "docx_edit_paragraph", "docx_insert", "docx_delete",
                   "docx_repair", "docx_format"}
# Phase 2 op-style tools mutate the live doc on every op except the read-only ones
# below; for these the harness saves the result and compares output packages too.
NEVER_MUTATING_OPS: dict[str, set[str]] = {
    "docx_style": {"list"},
    "docx_section": {"list"},
    "docx_comment": {"list"},
    "docx_media": {"extract"},
}
OP_MUTATING = {"docx_table", "docx_style", "docx_list", "docx_section",
               "docx_comment", "docx_media", "docx_field"}
# Volatile/non-load-bearing keys excluded from result-object parity: ids and
# byte counts vary by construction, and `note` is free-form human prose that the
# tools word differently per language (it carries no machine-checkable contract).
MASKED_KEYS = {"doc_id", "bytes", "note"}
RESPONSE_TIMEOUT = 60.0


def mutates(tool: str, args: dict[str, Any]) -> bool:
    if tool in ALWAYS_MUTATING:
        return True
    if tool == "docx_revision":
        return args.get("op") in {"accept", "reject", "accept_all", "reject_all"}
    if tool in OP_MUTATING:
        return args.get("op") not in NEVER_MUTATING_OPS.get(tool, set())
    return False


# ---------------------------------------------------------------------------
# CLI subprocess (line-oriented JSON protocol, §11)
# ---------------------------------------------------------------------------


class HarnessError(Exception):
    """A case-level failure with a human-readable reason."""


class CLIProc:
    def __init__(self, cmd: list[str], env: dict[str, str]) -> None:
        try:
            self.proc = subprocess.Popen(
                cmd,
                cwd=str(REPO_DIR),
                env=env,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
            )
        except OSError as exc:
            raise HarnessError(f"CLI not available ({cmd[0]}): {exc}") from exc
        self._lines: queue.Queue[str | None] = queue.Queue()
        self._stderr_tail: deque[str] = deque(maxlen=5)
        threading.Thread(target=self._read_loop, daemon=True).start()
        threading.Thread(target=self._stderr_loop, daemon=True).start()

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

    def call(self, tool: str, args: dict[str, Any]) -> dict[str, Any]:
        assert self.proc.stdin is not None
        request = json.dumps({"tool": tool, "args": args}, ensure_ascii=False)
        try:
            self.proc.stdin.write(request + "\n")
            self.proc.stdin.flush()
        except (BrokenPipeError, OSError) as exc:
            raise HarnessError(f"CLI pipe broke on {tool}: {exc}") from exc
        try:
            line = self._lines.get(timeout=RESPONSE_TIMEOUT)
        except queue.Empty:
            raise HarnessError(f"CLI timed out on {tool} ({RESPONSE_TIMEOUT}s)") from None
        if line is None:
            raise HarnessError(f"CLI exited mid-conversation on {tool}{self._stderr_hint()}")
        try:
            response = json.loads(line)
        except json.JSONDecodeError as exc:
            raise HarnessError(f"CLI wrote non-JSON for {tool}: {line[:200]!r}") from exc
        if not isinstance(response, dict):
            raise HarnessError(f"CLI wrote a non-object for {tool}: {line[:200]!r}")
        return response

    def close(self) -> None:
        try:
            if self.proc.stdin is not None:
                self.proc.stdin.close()
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()


# ---------------------------------------------------------------------------
# Partial expectation matching
# ---------------------------------------------------------------------------


def match_partial(expected: Any, actual: Any, path: str = "$") -> list[str]:
    if isinstance(expected, dict):
        if not isinstance(actual, dict):
            return [f"{path}: expected object, got {type(actual).__name__}"]
        errors: list[str] = []
        for key, exp_val in expected.items():
            if key not in actual:
                errors.append(f"{path}.{key}: missing")
            else:
                errors.extend(match_partial(exp_val, actual[key], f"{path}.{key}"))
        return errors
    if isinstance(expected, list):
        if expected and expected[0] == "$substr":
            # ["$substr", s1, s2, …]: each si must be a substring of the actual
            # string. Used for md/html conversion content where exact byte
            # equality is asserted by the parity comparison, not the golden.
            if not isinstance(actual, str):
                return [f"{path}: expected string for $substr, got {type(actual).__name__}"]
            return [
                f"{path}: substring {s!r} not found"
                for s in expected[1:]
                if s not in actual
            ]
        if not isinstance(actual, list):
            return [f"{path}: expected array, got {type(actual).__name__}"]
        if expected and expected[0] == "$contains":
            return _match_contains(expected[1:], actual, path)
        if len(expected) != len(actual):
            return [f"{path}: expected {len(expected)} items, got {len(actual)}"]
        errors = []
        for i, (exp_val, act_val) in enumerate(zip(expected, actual)):
            errors.extend(match_partial(exp_val, act_val, f"{path}[{i}]"))
        return errors
    if isinstance(expected, bool) or isinstance(actual, bool):
        return [] if expected is actual else [f"{path}: expected {expected!r}, got {actual!r}"]
    return [] if expected == actual else [f"{path}: expected {expected!r}, got {actual!r}"]


def _match_contains(matchers: list[Any], actual: list[Any], path: str) -> list[str]:
    """Each matcher must match a distinct element (backtracking assignment)."""

    def assign(i: int, used: frozenset[int]) -> bool:
        if i == len(matchers):
            return True
        for j, item in enumerate(actual):
            if j in used:
                continue
            if not match_partial(matchers[i], item, path):
                if assign(i + 1, used | {j}):
                    return True
        return False

    if assign(0, frozenset()):
        return []
    return [f"{path}: $contains not satisfied by {len(actual)} element(s)"]


def mask_volatile(obj: Any) -> Any:
    """Drop volatile keys (MASKED_KEYS) entirely so neither their value nor
    their *presence* affects result-object parity — e.g. an optional `note`
    emitted by only one implementation must not fail parity."""
    if isinstance(obj, dict):
        return {k: mask_volatile(v) for k, v in obj.items() if k not in MASKED_KEYS}
    if isinstance(obj, list):
        return [mask_volatile(v) for v in obj]
    return obj


# ---------------------------------------------------------------------------
# Package comparison (algorithms.md §10)
# ---------------------------------------------------------------------------


def canon_xml(data: bytes) -> str:
    """Canonical serialization: attributes sorted by name as written, no
    whitespace-only text between elements, comments/PIs dropped."""
    root_children: list[Any] = []
    stack: list[list[Any]] = [root_children]

    def start(name: str, attrs: list[str]) -> None:
        pairs = sorted(zip(attrs[0::2], attrs[1::2]))
        node = ("el", name, pairs, [])
        stack[-1].append(node)
        stack.append(node[3])

    def end(_name: str) -> None:
        stack.pop()

    def chardata(text: str) -> None:
        children = stack[-1]
        if children and children[-1][0] == "text":
            children[-1] = ("text", children[-1][1] + text)
        else:
            children.append(("text", text))

    parser = expat.ParserCreate()
    parser.ordered_attributes = True
    parser.buffer_text = True
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = chardata
    parser.Parse(data, True)

    def esc_text(s: str) -> str:
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def esc_attr(s: str) -> str:
        return esc_text(s).replace('"', "&quot;")

    def ser(node: Any, out: list[str]) -> None:
        _, name, pairs, children = node
        out.append(f"<{name}")
        for k, v in pairs:
            out.append(f' {k}="{esc_attr(v)}"')
        out.append(">")
        has_elem = any(c[0] == "el" for c in children)
        for child in children:
            if child[0] == "text":
                if has_elem and not child[1].strip():
                    continue  # whitespace-only text between elements
                out.append(esc_text(child[1]))
            else:
                ser(child, out)
        out.append(f"</{name}>")

    out: list[str] = []
    for node in root_children:
        if node[0] == "el":
            ser(node, out)
    return "".join(out)


def compare_packages(path_a: Path, path_b: Path, label_a: str, label_b: str) -> list[str]:
    diffs: list[str] = []
    with zipfile.ZipFile(path_a) as za, zipfile.ZipFile(path_b) as zb:
        names_a, names_b = za.namelist(), zb.namelist()
        if names_a != names_b:
            return [f"entry-name order differs: {label_a}={names_a} {label_b}={names_b}"]
        for name in names_a:
            da, db = za.read(name), zb.read(name)
            if da == db:
                continue
            if name.endswith((".xml", ".rels")):
                try:
                    ca, cb = canon_xml(da), canon_xml(db)
                except expat.ExpatError as exc:
                    diffs.append(f"{name}: unparseable XML during fallback ({exc})")
                    continue
                if ca == cb:
                    continue
                idx = next(
                    (i for i, (x, y) in enumerate(zip(ca, cb)) if x != y),
                    min(len(ca), len(cb)),
                )
                lo = max(0, idx - 60)
                diffs.append(
                    f"{name}: canonical XML differs at char {idx}: "
                    f"{label_a}=…{ca[lo:idx + 60]!r} {label_b}=…{cb[lo:idx + 60]!r}"
                )
            else:
                diffs.append(f"{name}: binary content differs")
    return diffs


# ---------------------------------------------------------------------------
# Case execution
# ---------------------------------------------------------------------------


def load_cases(only: list[str]) -> list[tuple[str, dict[str, Any]]]:
    cases = []
    for path in sorted(CASES_DIR.glob("*.json")):
        if only and path.stem not in only:
            continue
        cases.append((path.stem, json.loads(path.read_text(encoding="utf-8"))))
    if only:
        missing = set(only) - {name for name, _ in cases}
        if missing:
            sys.exit(f"unknown case(s): {', '.join(sorted(missing))}")
    return cases


class CaseRun:
    """Outcome of one case on one implementation."""

    def __init__(self) -> None:
        self.reasons: list[str] = []
        self.result: dict[str, Any] | None = None
        self.output: Path | None = None

    @property
    def ok(self) -> bool:
        return not self.reasons


def run_case(impl: str, name: str, case: dict[str, Any], env: dict[str, str]) -> CaseRun:
    run = CaseRun()
    tool = case["tool"]
    expect = case.get("expect", {})
    input_docx = CORPUS_DIR / case["doc"] / "input.docx"
    out_path = OUT_DIR / name / f"{impl}.docx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        proc = CLIProc(IMPLS[impl], env)
    except HarnessError as exc:
        run.reasons.append(str(exc))
        return run

    try:
        opened = proc.call("docx_open", {"path": str(input_docx)})
        if "error" in opened:
            raise HarnessError(f"docx_open failed: {opened}")
        doc_id = opened["doc_id"]

        # Optional `setup` pre-steps run against the same doc before the case
        # tool (e.g. docx_style define before docx_style apply). Each gets
        # doc_id injected; any error aborts the case.
        for i, step in enumerate(case.get("setup", [])):
            step_args = dict(step.get("args", {}))
            step_args["doc_id"] = doc_id
            step_resp = proc.call(step["tool"], step_args)
            if "error" in step_resp:
                raise HarnessError(
                    f"setup step {i} ({step['tool']}) failed: {json.dumps(step_resp)[:200]}"
                )

        args = dict(case.get("args", {}))
        args["doc_id"] = doc_id
        if tool == "docx_save":
            args["path"] = str(out_path)
        response = proc.call(tool, args)
        run.result = response

        if "error" in expect:
            if response.get("error") != expect["error"]:
                raise HarnessError(
                    f"expected error {expect['error']!r}, got {json.dumps(response)[:300]}"
                )
            return run  # error expected and matched; nothing to save/compare
        if "error" in response:
            raise HarnessError(f"tool errored: {json.dumps(response)[:300]}")
        run.reasons.extend(match_partial(expect.get("result", {}), response))

        if tool == "docx_save":
            run.output = out_path
        elif mutates(tool, args):
            saved = proc.call("docx_save", {"doc_id": doc_id, "path": str(out_path)})
            if "error" in saved:
                raise HarnessError(f"post-edit docx_save failed: {json.dumps(saved)[:300]}")
            run.output = out_path

        for invariant in expect.get("invariants", []):
            check_invariant(invariant, proc, doc_id, input_docx, run)

        # `output_contains`: {part_name: [substrings, …]} asserted against the
        # saved package (e.g. a docx_format style_selector edit must land in
        # word/styles.xml). Also accepts {part: {"absent": [substrings]}} to
        # assert the input had a marker the edit was supposed to introduce.
        oc = expect.get("output_contains")
        if oc:
            if run.output is None:
                run.reasons.append("output_contains: no output document was produced")
            else:
                check_output_contains(oc, run.output, input_docx, run)

        # `file_equals_part`: {"path_key": <result key naming a written file>,
        #  "part": "<fixture>/<part>"} — asserts the bytes a tool wrote to disk
        #  (e.g. docx_media extract) equal a corpus part. Resolves the path
        #  relative to REPO_DIR (the CLI's cwd).
        fe = expect.get("file_equals_part")
        if fe:
            check_file_equals_part(fe, run, impl)
    except HarnessError as exc:
        run.reasons.append(str(exc))
    finally:
        proc.close()
    return run


def check_file_equals_part(spec: dict[str, Any], run: CaseRun, impl: str) -> None:
    """Assert a file written by the tool equals a corpus part's bytes."""
    if run.result is None:
        run.reasons.append("file_equals_part: tool produced no result")
        return
    written = run.result.get(spec["path_key"])
    if not isinstance(written, str):
        run.reasons.append(f"file_equals_part: result has no string {spec['path_key']!r}")
        return
    written_path = Path(written)
    if not written_path.is_absolute():
        written_path = REPO_DIR / written_path
    if not written_path.is_file():
        run.reasons.append(f"file_equals_part: {written_path} was not written")
        return
    fixture, _, part = spec["part"].partition("/")
    expected = CORPUS_DIR / fixture / "input.docx"
    with zipfile.ZipFile(expected) as zf:
        want = zf.read(part)
    if written_path.read_bytes() != want:
        run.reasons.append(f"file_equals_part: {written_path} bytes != {spec['part']}")


def check_output_contains(
    spec: dict[str, Any], output: Path, input_docx: Path, run: CaseRun
) -> None:
    """Assert substrings in named parts of the saved package.

    spec maps a part name to either a list of required substrings, or a dict
    {"present": [...], "absent": [...], "absent_in_input": [...]}: `present`
    must be in the output part, `absent` must NOT be (e.g. a deleted marker),
    and `absent_in_input` proves the edit *introduced* a marker by requiring it
    absent from the input fixture.
    """
    with zipfile.ZipFile(output) as zf:
        out_names = set(zf.namelist())
        for part, want in spec.items():
            if part not in out_names:
                run.reasons.append(f"output_contains: output missing part {part}")
                continue
            data = zf.read(part).decode("utf-8")
            if isinstance(want, str):  # a bare string means "this substring must be present"
                want = [want]
            present = want if isinstance(want, list) else want.get("present", [])
            absent = [] if isinstance(want, list) else want.get("absent", [])
            absent_in_input = [] if isinstance(want, list) else want.get("absent_in_input", [])
            for sub in present:
                if sub not in data:
                    run.reasons.append(f"output_contains: {part} lacks {sub!r}")
            for sub in absent:
                if sub in data:
                    run.reasons.append(f"output_contains: {part} still contains {sub!r}")
            if absent_in_input:
                with zipfile.ZipFile(input_docx) as zin:
                    src = (
                        zin.read(part).decode("utf-8") if part in zin.namelist() else ""
                    )
                for sub in absent_in_input:
                    if sub in src:
                        run.reasons.append(
                            f"output_contains: {part} already had {sub!r} in the input "
                            "(edit did not introduce it)"
                        )


def check_invariant(
    invariant: str, proc: CLIProc, doc_id: str, input_docx: Path, run: CaseRun
) -> None:
    if invariant == "roundtrip":
        if run.output is None:
            run.reasons.append("roundtrip: no output document was produced")
            return
        diffs = compare_packages(input_docx, run.output, "input", "output")
        run.reasons.extend(f"roundtrip: {d}" for d in diffs)
    elif invariant == "no_word_repair":
        if run.output is None:
            run.reasons.append("no_word_repair: no output document was produced")
            return
        reopened = proc.call("docx_open", {"path": str(run.output)})
        if "error" in reopened:
            run.reasons.append(f"no_word_repair: reopen failed: {json.dumps(reopened)[:200]}")
            return
        verdict = proc.call("docx_validate", {"doc_id": reopened["doc_id"]})
        if verdict.get("valid") is not True:
            run.reasons.append(f"no_word_repair: output invalid: {json.dumps(verdict)[:300]}")
    elif invariant == "revisions_preserved":
        if run.output is None:
            run.reasons.append("revisions_preserved: no output document was produced")
            return
        with zipfile.ZipFile(run.output) as zf:
            doc = zf.read("word/document.xml")
        if b"<w:ins " not in doc and b"<w:del " not in doc:
            run.reasons.append("revisions_preserved: output has no w:ins/w:del left")
    elif invariant == "validate_clean":
        verdict = proc.call("docx_validate", {"doc_id": doc_id})
        if verdict.get("valid") is not True:
            run.reasons.append(f"validate_clean: still invalid: {json.dumps(verdict)[:300]}")
    else:
        run.reasons.append(f"unknown invariant {invariant!r}")


def check_parity(py: CaseRun, js: CaseRun) -> list[str]:
    reasons: list[str] = []
    if py.result is not None and js.result is not None:
        masked_py, masked_js = mask_volatile(py.result), mask_volatile(js.result)
        if masked_py != masked_js:
            reasons.append(
                "result objects differ: "
                f"py={json.dumps(masked_py, sort_keys=True)[:300]} "
                f"js={json.dumps(masked_js, sort_keys=True)[:300]}"
            )
    elif (py.result is None) != (js.result is None):
        reasons.append("one implementation produced a result and the other did not")
    if py.output is not None and js.output is not None:
        reasons.extend(compare_packages(py.output, js.output, "py", "js"))
    elif (py.output is None) != (js.output is None):
        reasons.append("one implementation produced an output document and the other did not")
    return reasons


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="DocxEngine conformance harness")
    parser.add_argument("--impl", choices=["py", "js"], help="run a single implementation")
    parser.add_argument("cases", nargs="*", help="case names to run (default: all)")
    opts = parser.parse_args(argv)
    impls = [opts.impl] if opts.impl else ["py", "js"]

    make_fixtures.build_corpus(CORPUS_DIR, quiet=True)
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    cases = load_cases(opts.cases)
    if not cases:
        sys.exit("no cases found in conformance/cases/")

    env = {**os.environ, **FIXED_ENV}

    rows: list[tuple[str, dict[str, str]]] = []
    failures: list[tuple[str, str, list[str]]] = []
    for name, case in cases:
        runs: dict[str, CaseRun] = {}
        cells: dict[str, str] = {}
        for impl in impls:
            runs[impl] = run_case(impl, name, case, env)
            cells[impl] = "pass" if runs[impl].ok else "FAIL"
            if not runs[impl].ok:
                failures.append((name, impl, runs[impl].reasons))
        if len(impls) == 2:
            if runs["py"].ok and runs["js"].ok:
                parity = check_parity(runs["py"], runs["js"])
                cells["parity"] = "pass" if not parity else "FAIL"
                if parity:
                    failures.append((name, "parity", parity))
            else:
                cells["parity"] = "-"
        rows.append((name, cells))

    columns = impls + (["parity"] if len(impls) == 2 else [])
    width = max(len(name) for name, _ in rows)
    header = "CASE".ljust(width) + "".join(f"  {c.upper():>6}" for c in columns)
    print(header)
    print("-" * len(header))
    for name, cells in rows:
        print(name.ljust(width) + "".join(f"  {cells.get(c, '-'):>6}" for c in columns))

    if failures:
        print()
        for name, side, reasons in failures:
            for reason in reasons:
                print(f"FAIL {name} [{side}]: {reason}")
        print(f"\n{len(failures)} failing case-side(s) of {len(rows)} cases")
        return 1
    print(f"\nall {len(rows)} cases passed on: {', '.join(columns)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
