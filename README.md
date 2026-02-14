# pywebview + React Desktop App Template

A desktop application template using **pywebview** for the native window and **React** for the UI. Inspired by Tauri's project layout with `src/` for the frontend and `src-python/` for the backend.

## Tech Stack

- **Frontend**: React 19, TypeScript, Vite, Tailwind CSS v4, shadcn/ui
- **Backend**: Python 3.9+, pywebview 5+
- **Packaging**: Nuitka (compiles to standalone executable)

## Setup

```bash
# Install frontend dependencies (npm, bun, yarn, or pnpm all work)
npm install   # or: bun install / yarn install / pnpm install

# Install Python dependencies (requires uv: https://docs.astral.sh/uv/)
uv sync
```

## Development

Single command (like `tauri dev`):

```bash
npm run dev:app
# or with a different package manager:
uv run python dev.py --pm bun
uv run python dev.py --pm yarn
uv run python dev.py --pm pnpm
```

## Production

```bash
npm run build
npm run pywebview:prod
```

## Packaging (Nuitka)

```bash
npm run package
```

Produces a single standalone executable with the React frontend bundled inside. No Python or Node.js required on the target machine.

## Adding New API Methods

1. Add a method to `src-python/api.py`:

```python
class Api:
    def greet(self, name: str) -> str:
        return f"Hello, {name}! This response came from Python."

    def add(self, a: float, b: float) -> float:
        return a + b
```

2. Update the TypeScript interface in `src/vite-env.d.ts`:

```typescript
interface PyWebViewApi {
  greet(name: string): Promise<string>;
  add(a: number, b: number): Promise<number>;
}
```

3. Call it from React using the bridge:

```typescript
const api = await getPyWebViewApi();
const result = await api.add(2, 3);
```
