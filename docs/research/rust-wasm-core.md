# Rust/WASM core unification — v2 evaluation

The [ROADMAP](../../ROADMAP.md) Phase 3 calls for a "v2 evaluation of Rust/WASM core unification," and the decision-thresholds table lists the trigger: _"Conformance-caught drift costs more than an FFI/WASM build, or edit throughput bottlenecks."_ This document is that evaluation. It is a decision record, not an advocacy piece — the question is whether to collapse the two hand-maintained engines into one Rust core, and the honest answer for v1.0-alpha is **defer**, with concrete triggers that would change the answer.

The [ARCHITECTURE](../../ARCHITECTURE.md#core-sharing-strategy) "Core-sharing strategy" table already recorded this as "Defer to v2." This doc revisits that call against the codebase as it actually stands and sharpens the triggers into something measurable.

## 1. Context — the problem a single core would solve

DocxEngine ships two full implementations of the same deterministic engine:

- **Python** (`python/src/docxengine/`) — zero runtime dependencies (`dependencies = []` in `python/pyproject.toml`; `mcp` is an optional extra). Pure stdlib: `zipfile`, the byte-splice scanner, the anchor index, the validator.
- **TypeScript** (`js/src/`) — one runtime dependency, `fflate`, for DEFLATE. Everything else is hand-written TS.

These are not a core plus a thin shell. They are parallel reimplementations of the same algorithms — the file lists mirror each other almost one-to-one (`_anchors.py`/`anchors.ts`, `_projector.py`/`projector.ts`, `_edits.py`/`edits.ts`, `_validate.py`/`validate.ts`, and so on). The split-run coalescer in [spec/algorithms.md §4](../../spec/algorithms.md) exists twice; the byte-splice strategy in §3 exists twice; the tracked-change writer exists twice.

The maintenance tax is real and recurring:

- **Every feature is written twice.** A new tool, a new validator check, a new projection rule lands as two diffs in two languages with two test suites (427 Python tests, 328 TS tests as of Phase 2).
- **Every bugfix is written twice**, and the second one is easy to forget.
- **Parity is debugged by hand.** When `conformance/harness/run.py` reports a byte difference, someone diffs two language implementations of the same algorithm to find which one is wrong. This is the costliest failure mode — it is the cost the trigger is measured against.

What keeps the two honest is the cross-implementation conformance suite (`conformance/`): 36 cases, each running the same input `.docx` and the same tool call through both engines, normalizing output (fixed ZIP order, zeroed timestamps, canonicalized XML in touched parts per §10), and failing on any byte difference or violated invariant. Parity is **Invariant 6** in ARCHITECTURE. The suite is the safety net that makes the dual-engine approach viable at all — and the thing a single core would make unnecessary.

## 2. What "core unification" means

One implementation of the deterministic core, written once (the natural candidate is Rust), compiled to two delivery targets, with the existing thin language faces kept on top:

- **(a) Native Python extension** — via PyO3 + maturin (a Python-aware binding + wheel builder), or a plain C ABI consumed through `ctypes`/`cffi`. The `Document`/`call()` surface in `python/src/docxengine/` becomes a shell over the compiled core.
- **(b) WASM module** — via wasm-bindgen / wasm-pack, for Node and the browser. The `@docxengine/core` surface in `js/src/` becomes a shell over the `.wasm` binary.

The faces stay as thin as ARCHITECTURE already demands ("translate formats only; behavior lives in the core" — Invariant 7). The difference from today is that "the core" is one artifact instead of two source trees kept in lockstep by tests.

This is the first row of the ARCHITECTURE core-sharing table ("Single Rust core → PyO3 + wasm-bindgen"), promoted from rejected-for-v1 to the thing under evaluation.

## 3. Options

### A. Status quo — two engines + conformance suite (current)

Two idiomatic implementations, parity enforced by `conformance/`.

- **Distribution:** the strongest property. Pure `pip install docxengine` with zero native code and zero third-party deps; pure `npm` with one small pure-JS dep. No compiler, no platform matrix, no `manylinux`/`musl` builds, no WASM bundle. Installs anywhere Python 3.12+ or Node runs, including locked-down environments.
- **Browser:** native. The TS engine runs in-browser today; `exportBytes` (`js/src/toolsLifecycle.ts:95`) returns a `Uint8Array` with no filesystem, mirroring Python's `export_bytes` (`python/src/docxengine/_tools_lifecycle.py:50`). The no-DOM raw-bytes splice design (§3) is plain string/byte work — no native dependency to port.
- **Contributors:** anyone who knows Python or TypeScript can contribute to their half.
- **CI:** two language toolchains, no cross-compilation.
- **Cost:** double implementation; drift risk, mitigated (not eliminated) by 36 conformance cases.

### B. Rust core + PyO3 + WASM (the unification candidate)

- **Distribution — the central cost.** The pure-Python install is **lost**: wheels must be built per platform/arch (manylinux x86_64/aarch64, macOS x86_64/arm64, Windows; musl for Alpine), with an sdist fallback that forces a Rust toolchain on the user. The npm package ships a `.wasm` blob (bundle size and per-environment loading become real concerns; bundlers and CDNs vary in WASM handling).
- **Browser:** still works — WASM runs in all modern browsers — but `exportBytes` now crosses the JS↔WASM boundary, and the whole document buffer must be marshalled in and out of linear memory.
- **FFI/WASM boundary cost vs. the splice design.** This is the design-specific catch. The engine holds **every package part as raw bytes** and edits by splicing byte ranges (§3); the session model keeps documents open server-side keyed by `doc_id`. If the core owns the bytes and the face only sends tool-call JSON across the boundary, marshalling is cheap and bounded. If buffers are copied across the boundary per operation, large documents pay a copy on every call — the exact throughput axis the trigger names. A unification that does not keep document state inside the core would trade one bottleneck for another.
- **Contributors:** narrows sharply. Rust + PyO3 + wasm-bindgen is a much smaller contributor pool than Python/TS, for a project whose adoption thesis (per ARCHITECTURE) is friction-free integration into private systems.
- **CI:** the heaviest option — a cross-compilation matrix (cibuildwheel-class) plus a WASM build, replacing two simple language jobs.
- **Benefit:** one implementation, no drift, conformance suite demoted to a regression check, and likely throughput gains on large documents (measurable via `bench/perf.py`).

### C. Rust core via C ABI / WASI

A C ABI consumed by `cffi` (Python) and a WASI build for JS hosts, instead of language-specific bindings.

- **Pro:** one stable ABI; in principle any language binds to it (relevant if a third binding is ever requested).
- **Con:** WASI in the **browser** is not a solved, universal story the way `wasm-bindgen` is; the hand-written marshalling that PyO3/wasm-bindgen generate becomes ours to maintain. Same per-platform binary-distribution burden as B, with a rougher browser path. Worse than B for DocxEngine's explicit browser requirement.

### D. Single TS core run under Python via a JS runtime — **rejected**

Compile/transpile one TS core and execute it under Python through an embedded JS runtime (Node subprocess, or a bundled engine).

- Forces a JavaScript runtime dependency into every Python install — exactly the "Node in a Python shop" failure the ARCHITECTURE table already rejected, and a direct violation of the zero-dependency pure-Python install that is a current selling point.
- Adds IPC/embedding latency and a heavyweight dependency to the most constrained deployment target. **Rejected.**

### E. Shared-spec codegen (adjacent, not exclusive)

Generate the mechanical, table-driven parts (error catalog, tool schemas, content-type tables) from one source into both languages. Already partially true: `spec/` is the contract, `errors.json` is shared data. This shrinks the duplicated surface without a binary toolchain and is **complementary to A** — worth pursuing regardless of the Rust decision, because it lowers the dual-engine tax that the Rust trigger is measured against. It does not unify the hard algorithmic core (splice, coalescing, tracked-change emission), which is where drift actually bites.

## 4. Cost / benefit vs. dual-engine + conformance

**Lost under unification (B/C):**

- The zero-dependency, zero-toolchain **pure-Python install** — a stated competitive property.
- The **pure-JS** simplicity (one tiny dep, trivial bundling, no `.wasm`).
- **Contributor accessibility** — Python/TS devs vs. a Rust+FFI+WASM skill set.
- A **simple CI** in exchange for a per-platform wheel matrix + WASM build, plus the binary-distribution support burden (musl, glibc versions, arch coverage, bundler quirks) that lands on users as install failures.

**Gained under unification:**

- **One implementation** — features and fixes written once.
- **No drift** — Invariant 6 becomes structural, not test-enforced.
- **Throughput** on large documents (Rust over interpreted Python/JS), to the extent `bench/perf.py` shows the engine — not LibreOffice rendering or model latency — is the bottleneck.

The trade is concrete: unification removes a real, recurring maintenance tax but spends DocxEngine's distribution advantage to do it. At today's surface size (~16 tools, complete and green in both languages) the dual-engine tax is being paid down, not accumulating — Phases 0–2 shipped full parity. The conformance suite is doing its job; drift is caught before release, not after.

## 5. Trigger thresholds

The ROADMAP row, sharpened into measurable triggers. Any one firing reopens this decision:

1. **Drift escapes the net.** A parity defect ships in **two consecutive releases**, or a single release ships with a Python/TS behavioral divergence the 36-case suite did not catch — i.e., the conformance net has holes and widening it is itself becoming the tax. This is the literal "drift costs more than an FFI/WASM build" trigger, made countable.
2. **Throughput dominates the agent loop.** `bench/perf.py` shows core edit/replace/save on documents above a stated size (e.g. **>5 MB / >20k paragraphs**) is the dominant per-turn cost — above tool-call and model latency, and above LibreOffice render time (the separate renderer threshold in ARCHITECTURE). `perf.py` measures exactly these numbers per operation (`open`, `search`, `replace_all`, `validate`, `to_bytes`) and writes `bench/perf-results.json`; this trigger is "the numbers crossed the line," not a guess.
3. **A third binding is requested.** A serious ask for a third language surface changes the arithmetic — at three reimplementations the one-core-many-bindings (C) calculus flips, where at two it does not.
4. **The dual-engine tax outpaces feature velocity.** Sustained periods where the second-language port is the critical-path blocker on shipping (qualitative, but visible in PR history).

Until a trigger fires, the cheaper move when the duplication tax stings is option E (codegen the table-driven surface), which buys relief without the binary toolchain.

## 6. Recommendation

**Defer Rust/WASM unification for v1.0-alpha. Keep the dual-engine + conformance approach.**

The motivation for unification — eliminating the double-implementation tax and drift risk — is real, but at the current surface it is a tax being serviced, not a debt compounding: Phases 0–2 shipped full parity, the 36-case conformance suite is catching divergence pre-release, and no stated trigger has fired. Against that, unification would spend DocxEngine's clearest adoption advantage (zero-dependency pure-`pip` / pure-`npm` installs, native browser execution of the splice design via `export_bytes`/`exportBytes`) on a heavier CI matrix, a per-platform binary-distribution burden, and a narrower contributor pool — for a deterministic core whose throughput has not been shown to bottleneck the agent loop. The correct posture is to keep `bench/perf.py` and the conformance suite running as the instruments that watch for triggers 1 and 2, pursue codegen (option E) opportunistically to shrink the duplicated surface, and revisit this decision the moment a measurable trigger fires — at which point option B (Rust core + PyO3 + wasm-bindgen) is the front-runner, provided it keeps document state inside the core so the FFI/WASM boundary carries tool-call JSON, not whole-document buffers.
