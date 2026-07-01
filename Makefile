# DocxEngine task runner.

.PHONY: help setup lint test bench perf fidelity clean

help: ## Show this help.
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-14s\033[0m %s\n", $$1, $$2}'

setup: ## Install the Python package with dev extras.
	pip install -e ".[dev]"

lint: ## Lint the Python package.
	ruff check src tests scripts

test: ## Run the Python test suite.
	pytest

bench: ## Run the agent task benchmark.
	.venv/bin/python bench/run.py

perf: ## Run the large-document performance benchmark.
	.venv/bin/python bench/perf.py

fidelity: ## Run the renderer fidelity harness (structural + LibreOffice when present).
	.venv/bin/python conformance/fidelity/run.py

clean: ## Remove build artifacts.
	rm -rf dist coverage
