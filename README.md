# Excel Agent

An AI-powered Excel automation desktop app. Chat with an LLM agent that reads, writes, and manipulates Excel files via natural language — built with **pywebview + React**.

## Tech Stack

- **Frontend**: React 19, TypeScript, Vite, Tailwind CSS v4, shadcn/ui
- **Backend**: Python, FastAPI, pywebview, pywin32 (Excel COM automation)
- **AI**: LangChain-based agent; supports zhipu, openai, deepseek, moonshot providers
- **Packaging**: Nuitka (compiles to standalone executable)

## Setup

```bash
# Install frontend dependencies
bun install

# Install Python dependencies (requires uv: https://docs.astral.sh/uv/)
uv sync
```

## Development

```bash
bun run dev:app
```

This starts the FastAPI backend (port 8765), Vite dev server (port 5173), and pywebview window together.

## Production

```bash
bun run build
bun run pywebview:prod
```

## Packaging (Nuitka)

```bash
bun run package
```

Produces a single standalone executable with the React frontend bundled inside. No Python or Node.js required on the target machine.
