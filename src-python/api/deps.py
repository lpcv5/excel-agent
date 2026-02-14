"""FastAPI dependency injection helpers."""

import os
import secrets
from typing import Optional

from fastapi import Depends, Header, HTTPException, Request

from api.errors import AppError, ErrorCode  # noqa: F401 (re-exported)
from services.project_service import ProjectService
from services.settings_service import SettingsService


def get_settings_service(request: Request) -> SettingsService:
    return request.app.state.settings_service


def get_project_service(request: Request) -> ProjectService:
    return request.app.state.project_service


def get_core(
    project_svc: ProjectService = Depends(get_project_service),
    settings_svc: SettingsService = Depends(get_settings_service),
):
    settings = settings_svc.load()
    return project_svc.get_core(settings)


async def verify_token(
    request: Request,
    x_app_token: Optional[str] = Header(default=None),
) -> None:
    if os.environ.get("EXCEL_AGENT_DEV") == "1":
        return
    app_token: str = request.app.state.app_token
    token = x_app_token or ""
    if not secrets.compare_digest(token, app_token):
        raise HTTPException(status_code=401, detail="Unauthorized")
