# DocxEngine task runner (boilerplate — targets activate as code lands).

.PHONY: help setup format format-check lint test test-py test-js conformance bench clean

help: ## Show this help.
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-14s\033[0m %s\n", $$1, $$2}'

setup: ## Install tooling (pnpm deps + husky hooks).
	pnpm install

format: ## Format docs, config, and source.
	pnpm format

format-check: ## Check formatting without writing.
	pnpm format:check

lint: format-check ## Lint everything that exists so far.
	@if [ -f python/pyproject.toml ] && command -v ruff >/dev/null; then ruff check python; fi

test: test-py test-js ## Run all test suites.

test-py: ## Run Python tests (no-op until code lands).
	@if [ -d python/tests ]; then cd python && pytest; else echo "python: no tests yet"; fi

test-js: ## Run JS tests (no-op until code lands).
	@if [ -d js/test ]; then cd js && pnpm test; else echo "js: no tests yet"; fi

conformance: ## Run the cross-implementation conformance harness.
	.venv/bin/python conformance/harness/run.py

bench: ## Run the agent task benchmark.
	.venv/bin/python bench/run.py

clean: ## Remove build artifacts.
	rm -rf dist coverage python/dist js/dist
