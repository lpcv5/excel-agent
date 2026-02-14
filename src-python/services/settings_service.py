"""AppSettings model and SettingsService."""

import json
import logging
import uuid
from pathlib import Path
from typing import Optional

from pydantic import BaseModel

logger = logging.getLogger("app.settings")

_SETTINGS_FILE = Path.home() / ".excel_agent" / "settings.json"


class ModelEntry(BaseModel):
    id: str
    provider: str
    model_name: str
    api_key: str = ""
    base_url: Optional[str] = None
    display_name: str = ""


class AppSettings(BaseModel):
    models: list[ModelEntry] = []
    main_model_id: Optional[str] = None
    subagents_model_id: Optional[str] = None
    analysis_model_id: Optional[str] = None
    language: str = "zh-CN"


def _migrate_old_format(data: dict) -> AppSettings:
    """Convert old flat format {provider, model, api_key, analysis_model} to new structure."""
    entries: list[ModelEntry] = []
    main_id = str(uuid.uuid4())
    main_entry = ModelEntry(
        id=main_id,
        provider=data.get("provider", "zhipu"),
        model_name=data.get("model", "glm-4.7"),
        api_key=data.get("api_key", ""),
        display_name="",
    )
    entries.append(main_entry)

    analysis_model_id: Optional[str] = None
    analysis_model_str = data.get("analysis_model", "").strip()
    if analysis_model_str and analysis_model_str != f"{main_entry.provider}:{main_entry.model_name}":
        parts = analysis_model_str.split(":", 1)
        a_provider = parts[0] if len(parts) == 2 else main_entry.provider
        a_model = parts[1] if len(parts) == 2 else analysis_model_str
        a_id = str(uuid.uuid4())
        entries.append(ModelEntry(
            id=a_id,
            provider=a_provider,
            model_name=a_model,
            api_key=main_entry.api_key,
            display_name="",
        ))
        analysis_model_id = a_id

    return AppSettings(
        models=entries,
        main_model_id=main_id,
        analysis_model_id=analysis_model_id,
        language=data.get("language", "zh-CN"),
    )


class SettingsService:
    def load(self) -> AppSettings:
        try:
            if _SETTINGS_FILE.exists():
                data = json.loads(_SETTINGS_FILE.read_text())
                # Detect old format by presence of top-level "provider" key
                if "provider" in data:
                    logger.info("Migrating settings from old format")
                    settings = _migrate_old_format(data)
                    self.save(settings)
                    return settings
                return AppSettings(**data)
        except Exception as e:
            logger.warning("Failed to load settings: %s", e)
        return AppSettings()

    def save(self, settings: AppSettings) -> None:
        # Guard: if api_key == "********", restore real key from existing stored entry
        existing: Optional[AppSettings] = None
        try:
            if _SETTINGS_FILE.exists():
                raw = json.loads(_SETTINGS_FILE.read_text())
                if "provider" not in raw:
                    existing = AppSettings(**raw)
        except Exception:
            pass

        if existing:
            existing_by_id = {e.id: e for e in existing.models}
            for entry in settings.models:
                if entry.api_key == "********":
                    real = existing_by_id.get(entry.id)
                    if real:
                        entry.api_key = real.api_key

        _SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        _SETTINGS_FILE.write_text(json.dumps(settings.model_dump(), indent=2))

    def load_masked(self) -> dict:
        s = self.load()
        d = s.model_dump()
        for entry in d["models"]:
            entry["api_key"] = "********" if entry["api_key"] else ""
        return d
