# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Development (starts Vite + pywebview together)
bun run dev:app

# Frontend only
bun run dev          # Vite dev server on port 5173
bun run lint         # ESLint

# Production
bun run build        # tsc + vite build
bun run pywebview:prod  # Run with built dist/

# Package to standalone executable (Nuitka)
bun run package

# Python deps managed via uv
uv sync

# Python linting/formatting
uv run ruff check src-python/
uv run ruff format src-python/  # excludes libs/deepagents (configured in pyproject.toml)

# Python type checking
uv run pyright src-python/
```

## Architecture

This is a **pywebview + React desktop app** — Python backend exposed to a React frontend via pywebview's JS bridge.

### Frontend (`src/`)

- `src/App.tsx` — root with ErrorBoundary
- `src/stores/chatStore.ts` — Zustand store; single source of truth for conversations, messages, streaming state
- `src/services/pywebview.ts` — HTTP client for the FastAPI backend; uses SSE for streaming, REST for settings/dialogs
- `src/components/chat/` — streaming chat UI (AssistantBubble, ToolGroupBlock, TaskPanel, ThinkingBlock)
- `src/components/layout/` — AppLayout with TitleBar, sidebars, MainContent, Footer

### Backend (`src-python/`)

- `src-python/main.py` — creates pywebview window, handles dev (connects to Vite URL) vs prod (loads `dist/`)
- `src-python/server.py` — FastAPI app on port 8765; token-based auth (`EXCEL_AGENT_TOKEN` env var or auto-generated); endpoints:
  - `GET /health` — no auth required
  - `GET/POST /api/settings` — provider/model/api_key persisted to `~/.excel_agent/settings.json`
  - `POST /api/stream` — SSE streaming endpoint that runs the agent and emits events
  - `POST /api/dialog/open`, `POST /api/dialog/save` — file dialogs via pywebview
- `src-python/agent/core.py` — `AgentCore`: initializes the DeepAgents LLM agent, streams queries via `astream_query()`, converts `MessageParser` events → `AgentEvent` types
- `src-python/agent/config.py` — `AgentConfig` dataclass; default model is `"zhipu:glm-4.7"` (format: `provider:model_name`)
- `src-python/agent/model_provider.py` — maps provider names to OpenAI-compatible API configs; supported: `zhipu`, `openai`, `deepseek`, `moonshot`
- `src-python/agent/events.py` — `AgentEvent` type definitions (EventType enum + event dataclasses)
- `src-python/tools/excel_tools/` — excel tool implementations (range, worksheet, chart, pivot, VBA, etc.)
- `src-python/libs/excel/` — utilities for working with Excel via `pywin32` (COM automation, cell references, A1/R1C1 notation)
- `src-python/libs/deepagents/` — DeepAgents framework (vendored, not a pip package); contains `graph.py`, backends, and middleware
- `src-python/libs/stream_msg_parser/` — streaming LLM message parser

### Dev orchestration

- `dev.py` — starts uvicorn (FastAPI) on port 8765 with `--reload`, starts Vite on port 5173, then launches pywebview pointing at `http://localhost:5173`
- `build.py` — builds React first, then compiles Python to a standalone exe via Nuitka (`--onefile`)

### Streaming event protocol

The frontend connects to `POST /api/stream` via SSE. The server yields `data: <json>\n\n` lines with a `type` field. Event types:

- `stream:content-token`, `stream:thinking-token`
- `stream:tool-group-start`, `stream:tool-call-update`, `stream:tool-group-done`
- `stream:tasks-init`, `stream:task-update`
- `stream:done` (payload may include `{error: string}` on failure)

The Zustand store in `chatStore.ts` consumes these events and drives all UI updates.

### UI

- shadcn/ui (New York style) + Tailwind CSS v4
- Add components: `bunx shadcn@latest add <component>` (MCP server configured in `.mcp.json`)
- Path alias: `@/` → `src/`
