"""Project configuration and lifecycle management."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

GLOBAL_CONFIG_DIR = Path.home() / ".excel-agent"
GLOBAL_CONFIG_FILE = GLOBAL_CONFIG_DIR / "config.json"
PROJECT_DIR_NAME = ".excel-agent"
PROJECT_CONFIG_NAME = "project.json"
AGENT_MD_NAME = "Agent.md"
SKILLS_DIR_NAME = "skills"


@dataclass
class DataSource:
    id: str
    type: str  # "excel" | "csv"
    path: str
    name: str


@dataclass
class ProjectOptions:
    data_cleaning_enabled: bool = True
    auto_save_memory: bool = True
    show_hidden_files: bool = False


@dataclass
class ProjectConfig:
    name: str
    created_at: str
    modified_at: str
    data_sources: list[DataSource] = field(default_factory=list)
    options: ProjectOptions = field(default_factory=ProjectOptions)
    analysis_status: str = "idle"  # "idle"|"running"|"done"|"error"
    analysis_completed_at: Optional[str] = None
    data_sources_hash: Optional[str] = None


@dataclass
class RecentProject:
    path: str
    name: str
    last_opened: str


@dataclass
class GlobalConfig:
    recent_projects: list[RecentProject] = field(default_factory=list)


# ── serialization helpers ────────────────────────────────────────────────────


def compute_data_sources_hash(data_sources: list[DataSource]) -> str:
    import hashlib
    paths = sorted(s.path for s in data_sources)
    return hashlib.sha256("|".join(paths).encode()).hexdigest()[:16]


def _to_dict(obj) -> dict:
    return asdict(obj)


def _project_config_from_dict(d: dict) -> ProjectConfig:
    sources = [DataSource(**s) for s in d.get("data_sources", [])]
    opts_d = d.get("options", {})
    opts = ProjectOptions(
        **{k: v for k, v in opts_d.items() if k in ProjectOptions.__dataclass_fields__}
    )
    return ProjectConfig(
        name=d["name"],
        created_at=d["created_at"],
        modified_at=d["modified_at"],
        data_sources=sources,
        options=opts,
        analysis_status=d.get("analysis_status", "idle"),
        analysis_completed_at=d.get("analysis_completed_at"),
        data_sources_hash=d.get("data_sources_hash"),
    )


def _global_config_from_dict(d: dict) -> GlobalConfig:
    recents = [RecentProject(**r) for r in d.get("recent_projects", [])]
    return GlobalConfig(recent_projects=recents)


# ── global config ────────────────────────────────────────────────────────────


def load_global_config() -> GlobalConfig:
    try:
        if GLOBAL_CONFIG_FILE.exists():
            return _global_config_from_dict(
                json.loads(GLOBAL_CONFIG_FILE.read_text(encoding="utf-8"))
            )
    except Exception:
        pass
    return GlobalConfig()


def save_global_config(cfg: GlobalConfig) -> None:
    GLOBAL_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    GLOBAL_CONFIG_FILE.write_text(
        json.dumps(_to_dict(cfg), indent=2, ensure_ascii=False), encoding="utf-8"
    )


def add_to_recent(project_path: str, name: str) -> None:
    cfg = load_global_config()
    now = datetime.now(timezone.utc).isoformat()
    # upsert
    cfg.recent_projects = [r for r in cfg.recent_projects if r.path != project_path]
    cfg.recent_projects.insert(
        0, RecentProject(path=project_path, name=name, last_opened=now)
    )
    # cap at 10
    cfg.recent_projects = cfg.recent_projects[:10]
    save_global_config(cfg)


# ── project config ───────────────────────────────────────────────────────────


def _project_dir(project_root: Path) -> Path:
    return project_root / PROJECT_DIR_NAME


def load_project_config(project_root: Path) -> ProjectConfig:
    cfg_file = _project_dir(project_root) / PROJECT_CONFIG_NAME
    return _project_config_from_dict(json.loads(cfg_file.read_text(encoding="utf-8")))


def save_project_config(project_root: Path, cfg: ProjectConfig) -> None:
    cfg_file = _project_dir(project_root) / PROJECT_CONFIG_NAME
    cfg_file.write_text(
        json.dumps(_to_dict(cfg), indent=2, ensure_ascii=False), encoding="utf-8"
    )


def create_project(
    project_root: Path,
    name: str,
    data_sources: Optional[list[DataSource]] = None,
    options: Optional[ProjectOptions] = None,
) -> ProjectConfig:
    proj_dir = _project_dir(project_root)
    proj_dir.mkdir(parents=True, exist_ok=True)

    now = datetime.now(timezone.utc).isoformat()
    cfg = ProjectConfig(
        name=name,
        created_at=now,
        modified_at=now,
        data_sources=data_sources or [],
        options=options or ProjectOptions(),
    )
    save_project_config(project_root, cfg)

    agent_md = proj_dir / AGENT_MD_NAME
    if not agent_md.exists():
        agent_md.write_text(
            f"# {name}\n\nAgent memory and instructions for this project.\n",
            encoding="utf-8",
        )

    skills_dir = proj_dir / SKILLS_DIR_NAME
    skills_dir.mkdir(exist_ok=True)

    add_to_recent(str(project_root), name)
    return cfg


def get_project_agent_paths(project_root: Path) -> tuple[Path, Path]:
    """Returns (agent_md_path, skills_path)."""
    proj_dir = _project_dir(project_root)
    return proj_dir / AGENT_MD_NAME, proj_dir / SKILLS_DIR_NAME
