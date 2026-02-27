"""ProjectService: owns AgentCore lifecycle and stream guard."""

import asyncio
import logging
import re
import threading
from pathlib import Path
from typing import Optional

from api.errors import AppError, ErrorCode
from services.settings_service import AppSettings

logger = logging.getLogger("app.project")


def _resolve_model_entry(settings: AppSettings, model_id: str | None):
    """Find ModelEntry by ID, falling back to first model in list."""
    if model_id:
        for entry in settings.models:
            if entry.id == model_id:
                return entry
    return settings.models[0] if settings.models else None


def _update_agent_md(agent_md_path: Path, summary_content: str) -> None:
    """Replace ## Data Sources section in Agent.md with a brief summary only.

    Full schema details are stored in .excel-agent/schema/ and accessed
    via the read_datasource_schema tool on demand.
    """
    section_header = "## Data Sources"
    new_section = (
        f"{section_header}\n\n"
        f"{summary_content}\n\n"
        f"> Use `read_datasource_schema` tool to read full schema for any source.\n"
    )

    if agent_md_path.exists():
        text = agent_md_path.read_text(encoding="utf-8")
        pattern = r"## Data Sources\n.*?(?=\n## |\Z)"
        if re.search(pattern, text, re.DOTALL):
            text = re.sub(pattern, new_section.rstrip(), text, flags=re.DOTALL)
        else:
            text = text.rstrip() + "\n\n" + new_section
    else:
        text = new_section

    agent_md_path.write_text(text, encoding="utf-8")


class ProjectService:
    def __init__(self) -> None:
        self._core = None
        self._core_lock = threading.Lock()
        self._active_root: Optional[Path] = None
        self._project_lock = threading.Lock()
        self._stream_active = False
        self._analysis_status: str = "idle"
        self._analysis_lock = threading.Lock()
        self._analysis_task: Optional[asyncio.Task] = None

    def get_active_root(self) -> Optional[Path]:
        with self._project_lock:
            return self._active_root

    def set_project(self, root: Optional[Path]) -> None:
        if self._stream_active:
            raise AppError(
                ErrorCode.PROJECT_SWITCH_DURING_STREAM,
                "Cannot switch project while a stream is active",
                status=409,
            )
        with self._project_lock:
            self._active_root = root
        self._reset_core()

    def get_core(self, settings: AppSettings):
        """Lazy-init AgentCore with current settings and project root."""
        if self._core is not None:
            return self._core
        with self._core_lock:
            if self._core is not None:
                return self._core
            import os

            from agent.config import AgentConfig
            from agent.core import AgentCore
            from agent.model_provider import ModelProvider
            from tools.excel_tools import CompositeExcelToolProvider

            entry = _resolve_model_entry(settings, settings.main_model_id)
            if entry is None:
                raise AppError(ErrorCode.SETTINGS_INVALID, "No model configured", status=400)

            mp = ModelProvider.from_model_entry(entry)
            model_spec = str(mp)
            if entry.api_key:
                from agent.model_provider import PREDEFINED_PROVIDERS

                provider_cfg = PREDEFINED_PROVIDERS.get(entry.provider)
                if provider_cfg:
                    os.environ[provider_cfg.api_key_env] = entry.api_key

            project_root = self.get_active_root()
            if project_root:
                from agent.project import get_project_agent_paths

                agents_md, skills = get_project_agent_paths(project_root)
                config = AgentConfig(
                    model=model_spec,
                    working_dir=project_root,
                    agents_md_path=agents_md,
                    skills_path=skills,
                    tool_providers=[CompositeExcelToolProvider()],
                )
            else:
                config = AgentConfig(
                    model=model_spec,
                    tool_providers=[CompositeExcelToolProvider()],
                )
            self._core = AgentCore(config)
        return self._core

    def _reset_core(self) -> None:
        with self._core_lock:
            self._core = None

    def mark_stream_start(self) -> None:
        self._stream_active = True

    def mark_stream_end(self) -> None:
        self._stream_active = False

    def cancel_stream(self) -> None:
        if self._core is not None:
            self._core.cancel()

    # ── Analysis ──────────────────────────────────────────────────────────────

    def get_analysis_status(self) -> dict:
        with self._analysis_lock:
            root = self.get_active_root()
            return {"status": self._analysis_status, "project_root": str(root) if root else None}

    async def trigger_analysis(self, settings: AppSettings) -> None:
        """Fire-and-forget analysis task. No-op if already running."""
        with self._analysis_lock:
            if self._analysis_status == "running":
                return
            self._analysis_status = "running"

        root = self.get_active_root()
        if root is None:
            with self._analysis_lock:
                self._analysis_status = "idle"
            return

        self._analysis_task = asyncio.ensure_future(self._run_analysis_task(root, settings))

    async def _run_analysis_task(self, root: Path, settings: AppSettings) -> None:
        from datetime import datetime, timezone

        from agent.analysis_agent import run_analysis
        from agent.project import (
            AGENT_MD_NAME,
            PROJECT_DIR_NAME,
            compute_data_sources_hash,
            load_project_config,
            save_project_config,
        )

        try:
            cfg = load_project_config(root)
            if not cfg.data_sources:
                with self._analysis_lock:
                    self._analysis_status = "idle"
                return

            # Resolve model entry for analysis (fall back to main model)
            entry = _resolve_model_entry(settings, settings.analysis_model_id) or \
                    _resolve_model_entry(settings, settings.main_model_id)
            if entry is None:
                raise ValueError("No model configured for analysis")

            model_spec = f"{entry.provider}:{entry.model_name}"
            provider = entry.provider

            content = await run_analysis(
                data_sources=cfg.data_sources,
                model_spec=model_spec,
                api_key=entry.api_key,
                provider=provider,
                project_root=root,
            )

            agent_md_path = root / PROJECT_DIR_NAME / AGENT_MD_NAME
            _update_agent_md(agent_md_path, content)

            now = datetime.now(timezone.utc).isoformat()
            cfg.analysis_status = "done"
            cfg.analysis_completed_at = now
            cfg.data_sources_hash = compute_data_sources_hash(cfg.data_sources)
            save_project_config(root, cfg)

            with self._analysis_lock:
                self._analysis_status = "done"

            logger.info("Analysis complete for project: %s", root)

        except Exception as e:
            logger.error("Analysis failed: %s", e)
            with self._analysis_lock:
                self._analysis_status = "error"
            try:
                from agent.project import load_project_config, save_project_config
                cfg = load_project_config(root)
                cfg.analysis_status = "error"
                save_project_config(root, cfg)
            except Exception:
                pass
