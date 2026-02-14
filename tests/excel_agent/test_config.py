"""Tests for excel_agent/config.py."""

import pytest
from pathlib import Path


class TestAgentConfig:
    """Tests for AgentConfig class."""

    def test_default_config(self):
        """Test default configuration values."""
        from excel_agent.config import AgentConfig

        config = AgentConfig()

        assert config.model == "openai:gpt-4.1"
        assert config.excel_visible is False
        assert config.streaming_enabled is True

    def test_custom_config(self):
        """Test custom configuration values."""
        from excel_agent.config import AgentConfig

        config = AgentConfig(
            model="anthropic:claude-3-opus",
            working_dir=Path("/custom/dir"),
        )

        assert config.model == "anthropic:claude-3-opus"
        assert config.working_dir == Path("/custom/dir")

    def test_thread_id_generated(self):
        """Test that thread_id is auto-generated."""
        from excel_agent.config import AgentConfig

        config = AgentConfig()

        assert config.thread_id is not None
        assert len(config.thread_id) > 0

    def test_custom_thread_id(self):
        """Test custom thread_id."""
        from excel_agent.config import AgentConfig

        config = AgentConfig(thread_id="my-custom-thread")

        assert config.thread_id == "my-custom-thread"
