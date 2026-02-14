"""Agent configuration.

This module defines the configuration for the Excel Agent,
including model settings, paths, and session options.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, Optional
import uuid


@dataclass
class LoggingConfig:
    """Configuration for LLM call logging.

    Attributes:
        enabled: Whether to enable LLM call logging
        level: Log level (DEBUG, INFO, WARNING, ERROR)
        output: Output destination (console, file, or both)
        file_path: Path to log file (used when output is file or both)
        log_full_prompt: Whether to log full prompt content
        log_token_usage: Whether to log token usage statistics
        log_timing: Whether to log call timing information
    """

    enabled: bool = False
    level: str = "DEBUG"
    output: Literal["console", "file", "both"] = "file"
    file_path: Optional[Path] = Path("agent.log")
    log_full_prompt: bool = True
    log_token_usage: bool = True
    log_timing: bool = True


@dataclass
class AgentConfig:
    """Configuration for Excel Agent.

    Attributes:
        model: LLM model to use (e.g., "openai:gpt-5-mini")
        working_dir: Working directory for file operations
        agents_md_path: Path to AGENTS.md memory file
        skills_path: Path to skills directory
        thread_id: Thread ID for conversation persistence
        excel_visible: Whether Excel window is visible
        excel_display_alerts: Whether Excel displays alerts
        streaming_enabled: Whether to use streaming output
        logging: LLM call logging configuration
    """

    # Model settings
    model: str = "openai:gpt-5-mini"

    # Paths
    working_dir: Path = field(default_factory=Path.cwd)
    agents_md_path: Optional[Path] = None
    skills_path: Optional[Path] = None

    # Session settings
    thread_id: str = field(default_factory=lambda: str(uuid.uuid4()))

    # Excel settings
    excel_visible: bool = False
    excel_display_alerts: bool = False

    # Streaming settings
    streaming_enabled: bool = True

    # Logging settings
    logging: LoggingConfig = field(default_factory=LoggingConfig)

    def __post_init__(self) -> None:
        """Set default paths based on working_dir if not specified."""
        if self.agents_md_path is None:
            # Default to AGENTS.md in the project root
            self.agents_md_path = Path(__file__).parent.parent / "AGENTS.md"
        if self.skills_path is None:
            # Default to skills directory in the project root
            self.skills_path = Path(__file__).parent.parent / "skills"

    @classmethod
    def from_cli_args(cls, args) -> "AgentConfig":
        """Create config from CLI arguments.

        Args:
            args: Parsed argparse namespace

        Returns:
            AgentConfig instance
        """
        log_level = getattr(args, "log_level", None)
        log_file = getattr(args, "log_file", None)

        logging_config = LoggingConfig(
            enabled=log_level is not None,
            level=log_level or "DEBUG",
            output="both" if log_file else "file",
            file_path=Path(log_file) if log_file else Path("agent.log"),
        )

        return cls(
            model=getattr(args, "model", "openai:gpt-5-mini"),
            working_dir=Path(args.dir) if getattr(args, "dir", None) else Path.cwd(),
            logging=logging_config,
        )
