# Repository Guidelines

## Project Structure & Module Organization
- `agent.py` is the CLI entry point and the main dev harness.
- `excel_agent/` holds the UI-agnostic core (config, events, orchestration).
- `tools/` contains Excel tool wrappers used by the agent.
- `libs/` includes shared libraries; `libs/deepagents` is a git submodule.
- `ui/` contains `cli/` and `web/` frontends.
- `tests/` is the pytest suite.
- `skills/` holds skill definitions; `docs/` contains supporting documentation.

## Build, Test, and Development Commands
- `uv sync` installs runtime and dev dependencies.
- `uv run python agent.py` starts the interactive CLI.
- `uv run python agent.py --web` launches the Web UI.
- `uv run python -m pytest` runs the full test suite.
- `uv run ruff check .` and `uv run ruff format .` lint and format the codebase.
- `git submodule update --init --recursive` fetches `libs/deepagents` if missing.

## Coding Style & Naming Conventions
- Python 3.12+, 4-space indentation, and standard PEP 8 layout.
- Use `snake_case` for functions/variables and `PascalCase` for classes.
- Prefer small, focused modules and keep Excel COM logic in `libs/excel_com/`.
- Formatting and linting are enforced by Ruff (use the commands above).

## Testing Guidelines
- Framework: pytest with conventions from `pyproject.toml`.
- Test files: `tests/test_*.py`, classes `Test*`, functions `test_*`.
- Integration tests should be marked with `@pytest.mark.integration`.

## Commit & Pull Request Guidelines
- Commit messages are short, imperative, and start with a verb (e.g., “Add …”, “Fix …”, “Refactor …”, “Update …”).
- PRs should include a concise summary, testing notes, and any Excel/Windows prerequisites.
- For UI changes, include a CLI log snippet or a Web UI screenshot when applicable.

## Configuration Notes
- This project requires Windows with Microsoft Excel installed.
- LLM provider credentials are expected via environment variables; document any new ones in `README.md`.
