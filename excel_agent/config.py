"""Agent configuration.

This module defines the configuration for the Excel Agent,
including model settings, paths, and session options.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import uuid


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
    """

    # Model settings
    model: str = "openai:gpt-4.1"

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
        return cls(
            model=getattr(args, "model", "openai:gpt-5-mini"),
            working_dir=Path(args.dir) if getattr(args, "dir", None) else Path.cwd(),
        )
