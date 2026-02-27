# Repository Guidelines

## Project Structure & Module Organization
- `src/`: React + TypeScript frontend (Vite, Tailwind, shadcn/ui).
- `src-python/`: Python backend (pywebview window, FastAPI server, agent logic).
- `docs/`: project documentation and notes.
- `dist/`: production frontend build output.
- `test_excel_flow.py`: end-to-end streaming test script; writes artifacts to `test_output/`.

## Build, Test, and Development Commands
- `bun run dev:app` (or `bun run dev:app`): starts FastAPI + Vite and launches pywebview.
- `bun run dev`: frontend-only Vite dev server on `http://localhost:5173`.
- `bun run build`: TypeScript build + Vite production build into `dist/`.
- `bun run pywebview:prod`: run the app against the built `dist/`.
- `bun run package`: build a standalone executable via Nuitka.
- `bun run lint`: run ESLint across the repo.
- `uv sync`: install Python dependencies.
- `uv run ruff check src-python/` and `uv run ruff format src-python/`: lint/format Python.
- `uv run pyright src-python/`: type-check Python.

## Coding Style & Naming Conventions
- TypeScript: 2-space indentation; use ESLint defaults from `eslint.config.js`.
- Python: formatted with Ruff; keep modules and functions snake_case; classes in PascalCase.
- Paths and imports: use `@/` alias for `src/` in frontend.

## Testing Guidelines
- Primary E2E script: `test_excel_flow.py` (non-pytest). Typical flow:
  - Start backend: `EXCEL_AGENT_DEV=1 uv run uvicorn server:app --host 127.0.0.1 --port 8765`
  - Run tests: `uv run python -u test_excel_flow.py`
  - Single case: `uv run python -u test_excel_flow.py --case test_03_formula`
- Keep new test cases in the same file; follow the existing `test_##_name` pattern.

## Commit & Pull Request Guidelines
- Commit messages: history shows short, imperative summaries (often sentence case). Keep to one line and be specific.
- Pull requests: include a concise summary, list tests run (or “not run” with reason), and add UI screenshots/GIFs for frontend changes.

## Security & Configuration
- The backend uses an auth token (`EXCEL_AGENT_TOKEN`) and persists settings at `~/.excel_agent/settings.json`. Avoid committing secrets or local settings files.
