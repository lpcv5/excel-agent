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

This is a **pywebview + React desktop app** — Python FastAPI backend served locally, consumed by a React frontend via HTTP/SSE.

### Frontend (`src/`)

- `src/App.tsx` — root with ErrorBoundary
- `src/router.tsx` — client-side routing
- `src/stores/` — Zustand stores: `chatStore.ts`, `projectStore.ts`, `settingsStore.ts`, `fileTreeStore.ts`
- `src/services/api.ts` — HTTP client for the FastAPI backend; uses SSE for streaming, REST for everything else
- `src/components/chat/` — streaming chat UI (AssistantBubble, ToolGroupBlock, TaskPanel, ThinkingBlock)
- `src/components/layout/` — AppLayout with TitleBar, sidebars, MainContent, Footer
- `src/components/project/` — project management UI
- `src/components/settings/` — settings panel
- `src/locales/` — i18n strings (`en-US.json`, `zh-CN.json`)

### Backend (`src-python/`)

- `src-python/main.py` — creates pywebview window; dev mode connects to Vite URL, prod loads `dist/`
- `src-python/server.py` — FastAPI app factory + lifespan; registers routers; token auth via `EXCEL_AGENT_TOKEN` env var or auto-generated
- `src-python/api/routers/` — route handlers:
  - `stream.py` — `POST /api/stream` SSE endpoint; runs the agent
  - `settings.py` — `GET/POST /api/settings`; persisted to `~/.excel_agent/settings.json`
  - `projects.py` — project CRUD endpoints
  - `files.py` — file browsing/dialog endpoints (`POST /api/dialog/open`, `POST /api/dialog/save`)
  - `ws_watcher.py` — WebSocket endpoint for file-system watch events
- `src-python/api/deps.py` — FastAPI dependency injection (auth, services)
- `src-python/api/errors.py` — global exception handlers
- `src-python/services/settings_service.py` — settings persistence
- `src-python/services/project_service.py` — project management logic
- `src-python/agent/core.py` — `AgentCore`: initializes the LangChain agent, streams queries via `astream_query()`, converts events → `AgentEvent` types
- `src-python/agent/config.py` — `AgentConfig` dataclass; default model `"zhipu:glm-4.7"` (format: `provider:model_name`); reads `AGENTS.md` and `skills/` from project root
- `src-python/agent/model_provider.py` — maps provider names to LangChain chat models; supported: `zhipu`, `openai`, `deepseek`, `moonshot`
- `src-python/agent/events.py` — `AgentEvent` type definitions (EventType enum + event dataclasses)
- `src-python/agent/context.py` — agent context/state passed through runs
- `src-python/agent/analysis_agent.py` — analysis sub-agent
- `src-python/tools/excel_tools/` — Excel tool implementations (range, worksheet, chart, pivot, VBA, etc.)
- `src-python/tools/base.py` — `ToolProvider` base class
- `src-python/tools/schema_tool.py` — schema introspection tool
- `src-python/tools/subagent.py` — sub-agent tool wrapper
- `src-python/libs/excel/` — utilities for Excel via `pywin32` (COM automation, cell references, A1/R1C1 notation)
- `src-python/libs/deepagents/` — DeepAgents framework (vendored); contains `graph.py`, backends, middleware
- `src-python/libs/stream_msg_parser/` — streaming LLM message parser

### Dev orchestration

- `dev.py` — starts uvicorn (FastAPI) on port 8765 with `--reload`, starts Vite on port 5173, then launches pywebview pointing at `http://localhost:5173`
- `build.py` — builds React first, then compiles Python to a standalone exe via Nuitka (`--onefile`)
- `AGENTS.md` — agent system prompt / instructions loaded at runtime

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
- i18n via `src/i18n.ts` + `src/locales/`
- Path alias: `@/` → `src/`
