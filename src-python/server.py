"""FastAPI app factory + lifespan. All routes live in api/routers/."""

import logging
import os
import secrets
import time
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware

from agent.logging_config import setup_app_logging
from api.errors import register_exception_handlers
from api.routers import files, projects, settings, stream, ws_watcher  # noqa: E501
from libs.excel_com.com_thread import shutdown_com_thread
from services.project_service import ProjectService
from services.settings_service import SettingsService

setup_app_logging()
logger = logging.getLogger("app.fastapi")

# ── token (module-level so main.py can import it before lifespan runs) ──────────
APP_TOKEN: str = os.environ.get("EXCEL_AGENT_TOKEN") or secrets.token_urlsafe(32)

# ── webview window reference (set by main.py after create_window) ───────────────
_webview_window = None


def set_webview_window(window) -> None:
    global _webview_window
    _webview_window = window
    try:
        app.state.webview_window = window
    except Exception:
        pass  # app.state not yet available before lifespan


@asynccontextmanager
async def lifespan(app: FastAPI):
    app.state.app_token = APP_TOKEN
    app.state.settings_service = SettingsService()
    app.state.project_service = ProjectService()
    app.state.webview_window = _webview_window
    yield
    shutdown_com_thread()


app = FastAPI(title="ExcelAgent API", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_methods=["*"],
    allow_headers=["*"],
)

register_exception_handlers(app)


@app.middleware("http")
async def log_requests(request: Request, call_next):
    start = time.perf_counter()
    response = await call_next(request)
    elapsed_ms = (time.perf_counter() - start) * 1000
    logger.info("%s %s %s %.1fms", request.method, request.url.path, response.status_code, elapsed_ms)
    return response


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/api/version")
async def get_version():
    return {"version": "0.0.1"}


app.include_router(stream.router)
app.include_router(settings.router)
app.include_router(projects.router)
app.include_router(files.router)
app.include_router(ws_watcher.router)
