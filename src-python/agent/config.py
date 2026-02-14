"""Agent configuration."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Literal, Optional
import uuid

from langchain_core.language_models import BaseChatModel

if TYPE_CHECKING:
    from tools.base import ToolProvider


@dataclass
class LoggingConfig:
    enabled: bool = True
    level: str = "DEBUG"
    output: Literal["console", "file", "both"] = "file"
    file_path: Optional[Path] = None
    log_full_prompt: bool = True
    log_token_usage: bool = True
    log_timing: bool = True


@dataclass
class AgentConfig:
    model: str | BaseChatModel = "zhipu:glm-4.7"
    working_dir: Path = field(default_factory=Path.cwd)
    agents_md_path: Optional[Path] = None
    skills_path: Optional[Path] = None
    thread_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    excel_visible: bool = False
    excel_display_alerts: bool = False
    excel_interactive: bool = True
    tool_providers: list["ToolProvider"] = field(default_factory=list)
    streaming_enabled: bool = True
    logging: LoggingConfig = field(default_factory=LoggingConfig)

    def __post_init__(self) -> None:
        if self.agents_md_path is None:
            # __file__ is src-python/agent/config.py
            # parent.parent.parent = project root
            self.agents_md_path = Path(__file__).parent.parent.parent / "AGENTS.md"
        if self.skills_path is None:
            self.skills_path = Path(__file__).parent.parent.parent / "skills"

    def get_model_instance(self) -> BaseChatModel:
        from agent.model_provider import create_model

        return create_model(self.model)
