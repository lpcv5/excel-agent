"""Settings endpoints: GET/POST /api/settings."""

from fastapi import APIRouter, Depends

from agent.model_provider import PREDEFINED_PROVIDERS
from api.deps import get_project_service, get_settings_service, verify_token
from services.project_service import ProjectService
from services.settings_service import AppSettings, SettingsService

router = APIRouter(prefix="/api/settings", dependencies=[Depends(verify_token)])


@router.get("")
async def get_settings(svc: SettingsService = Depends(get_settings_service)):
    return svc.load_masked()


@router.post("")
async def save_settings(
    body: AppSettings,
    svc: SettingsService = Depends(get_settings_service),
    project_svc: ProjectService = Depends(get_project_service),
):
    svc.save(body)
    # Reset core so next request picks up new settings
    project_svc._reset_core()
    return {"ok": True}


@router.get("/providers")
async def get_providers():
    return [
        {"name": p.name, "default_model": p.default_model}
        for p in PREDEFINED_PROVIDERS.values()
    ]
