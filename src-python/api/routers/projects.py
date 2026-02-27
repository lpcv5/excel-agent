"""Projects endpoints: /api/projects/*"""

import dataclasses
from pathlib import Path

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from api.deps import get_project_service, get_settings_service, verify_token
from services.project_service import ProjectService
from services.settings_service import SettingsService

router = APIRouter(prefix="/api/projects", dependencies=[Depends(verify_token)])


class DataSourceBody(BaseModel):
    id: str
    type: str
    path: str
    name: str


class ProjectOptionsBody(BaseModel):
    data_cleaning_enabled: bool = True
    auto_save_memory: bool = True
    show_hidden_files: bool = False


class CreateProjectBody(BaseModel):
    project_path: str
    name: str
    data_sources: list[DataSourceBody] = []
    options: ProjectOptionsBody = ProjectOptionsBody()


class OpenProjectBody(BaseModel):
    project_path: str


@router.post("/create")
async def create_project_endpoint(
    body: CreateProjectBody,
    project_svc: ProjectService = Depends(get_project_service),
    settings_svc: SettingsService = Depends(get_settings_service),
):
    from agent.project import DataSource, ProjectOptions, create_project

    root = Path(body.project_path)
    if (root / ".excel-agent" / "project.json").exists():
        raise HTTPException(status_code=409, detail="Project already exists at this path")

    sources = [DataSource(id=s.id, type=s.type, path=s.path, name=s.name) for s in body.data_sources]
    opts = ProjectOptions(
        data_cleaning_enabled=body.options.data_cleaning_enabled,
        auto_save_memory=body.options.auto_save_memory,
        show_hidden_files=body.options.show_hidden_files,
    )
    cfg = create_project(root, body.name, sources, opts)
    project_svc.set_project(root)
    if sources:
        await project_svc.trigger_analysis(settings_svc.load())
    project_dict = {**dataclasses.asdict(cfg), "path": str(root)}
    project_dict["analysis_status"] = project_svc.get_analysis_status()["status"]
    return {"project": project_dict}


@router.post("/open")
async def open_project_endpoint(
    body: OpenProjectBody,
    project_svc: ProjectService = Depends(get_project_service),
    settings_svc: SettingsService = Depends(get_settings_service),
):
    from agent.project import add_to_recent, compute_data_sources_hash, load_project_config

    root = Path(body.project_path)
    if not root.exists() or not (root / ".excel-agent" / "project.json").exists():
        raise HTTPException(status_code=404, detail="Project not found at this path")

    cfg = load_project_config(root)
    add_to_recent(str(root), cfg.name)
    project_svc.set_project(root)

    # Auto-trigger if hash is stale or status is idle/error
    if cfg.data_sources:
        current_hash = compute_data_sources_hash(cfg.data_sources)
        if cfg.data_sources_hash != current_hash or cfg.analysis_status in ("idle", "error"):
            await project_svc.trigger_analysis(settings_svc.load())

    project_dict = {**dataclasses.asdict(cfg), "path": str(root)}
    project_dict["analysis_status"] = project_svc.get_analysis_status()["status"]
    return {"project": project_dict}


@router.get("/recent")
async def get_recent_projects():
    from agent.project import load_global_config

    cfg = load_global_config()
    valid = [r for r in cfg.recent_projects if Path(r.path).exists()]
    return {"recent_projects": [dataclasses.asdict(r) for r in valid]}


@router.get("/current")
async def get_current_project(project_svc: ProjectService = Depends(get_project_service)):
    root = project_svc.get_active_root()
    if root is None:
        return {"project": None}
    try:
        from agent.project import load_project_config

        cfg = load_project_config(root)
        return {"project": {**dataclasses.asdict(cfg), "path": str(root)}}
    except Exception:
        return {"project": None}


@router.post("/close")
async def close_project(project_svc: ProjectService = Depends(get_project_service)):
    project_svc.set_project(None)
    return {"ok": True}


class AddDataSourceBody(BaseModel):
    type: str
    path: str
    name: str


def _path_to_id(path: str) -> str:
    import hashlib
    return hashlib.sha1(path.encode()).hexdigest()[:8]


@router.post("/data-sources")
async def add_data_source(
    body: AddDataSourceBody,
    project_svc: ProjectService = Depends(get_project_service),
    settings_svc: SettingsService = Depends(get_settings_service),
):
    from agent.project import DataSource, load_project_config, save_project_config

    root = project_svc.get_active_root()
    if root is None:
        raise HTTPException(status_code=400, detail="No active project")

    cfg = load_project_config(root)
    if any(s.path == body.path for s in cfg.data_sources):
        raise HTTPException(status_code=409, detail="Data source already exists")

    source_id = _path_to_id(body.path)
    cfg.data_sources.append(DataSource(id=source_id, type=body.type, path=body.path, name=body.name))
    save_project_config(root, cfg)
    await project_svc.trigger_analysis(settings_svc.load())
    return {"project": {**__import__("dataclasses").asdict(cfg), "path": str(root)}}


@router.delete("/data-sources/{source_id}")
async def remove_data_source(
    source_id: str,
    project_svc: ProjectService = Depends(get_project_service),
    settings_svc: SettingsService = Depends(get_settings_service),
):
    from agent.project import load_project_config, save_project_config

    root = project_svc.get_active_root()
    if root is None:
        raise HTTPException(status_code=400, detail="No active project")

    cfg = load_project_config(root)
    cfg.data_sources = [s for s in cfg.data_sources if s.id != source_id]
    save_project_config(root, cfg)
    if cfg.data_sources:
        await project_svc.trigger_analysis(settings_svc.load())
    return {"project": {**__import__("dataclasses").asdict(cfg), "path": str(root)}}


@router.post("/analyze")
async def trigger_analysis(
    project_svc: ProjectService = Depends(get_project_service),
    settings_svc: SettingsService = Depends(get_settings_service),
):
    root = project_svc.get_active_root()
    if root is None:
        raise HTTPException(status_code=400, detail="No active project")
    await project_svc.trigger_analysis(settings_svc.load())
    return {"ok": True}


@router.get("/analysis-status")
async def get_analysis_status(
    project_svc: ProjectService = Depends(get_project_service),
):
    return project_svc.get_analysis_status()
