"""Tests for excel_agent/core.py."""

from unittest.mock import MagicMock, patch
import pytest


class TestAgentCore:
    """Tests for AgentCore class."""

    def test_init_with_default_config(self):
        """Test AgentCore initialization with default config."""
        from excel_agent.core import AgentCore
        from excel_agent.config import AgentConfig

        core = AgentCore()

        assert core.config is not None
        assert isinstance(core.config, AgentConfig)

    def test_init_with_custom_config(self):
        """Test AgentCore initialization with custom config."""
        from excel_agent.core import AgentCore
        from excel_agent.config import AgentConfig

        config = AgentConfig(model="custom-model")
        core = AgentCore(config)

        assert core.config.model == "custom-model"

    def test_agent_lazy_loading(self):
        """Test that agent is lazy-loaded."""
        from excel_agent.core import AgentCore

        core = AgentCore()

        # Agent should not be created initially
        assert core._agent is None

        # Access agent property triggers creation
        with patch("excel_agent.core.AgentCore._create_agent") as mock_create:
            mock_create.return_value = MagicMock()
            _ = core.agent
            mock_create.assert_called_once()

    def test_invoke_returns_response(self):
        """Test AgentCore.invoke returns response."""
        from excel_agent.core import AgentCore

        core = AgentCore()
        core._agent = MagicMock()
        core._agent.invoke.return_value = {"output": "Test response"}

        result = core.invoke("test query")

        assert result == "Test response"

    def test_invoke_with_thread_id(self):
        """Test AgentCore.invoke with custom thread_id."""
        from excel_agent.core import AgentCore

        core = AgentCore()
        core._agent = MagicMock()
        core._agent.invoke.return_value = {"output": "Response"}

        core.invoke("test query", thread_id="custom-thread")

        core._agent.invoke.assert_called_once_with(
            {"messages": [("user", "test query")]},
            config={"configurable": {"thread_id": "custom-thread"}},
        )

    def test_stream_query_yields_events(self):
        """Test AgentCore.stream_query yields events."""
        from excel_agent.core import AgentCore
        from excel_agent.events import QueryStartEvent, TextEvent, QueryEndEvent

        core = AgentCore()
        core._agent = MagicMock()
        # Create a mock message chunk with text content
        mock_chunk = MagicMock()
        mock_chunk.content = [{"type": "text", "text": "Hello"}]
        core._agent.stream.return_value = iter([(mock_chunk, {})])

        events = list(core.stream_query("test query"))

        # Should yield QueryStartEvent, TextEvent, and QueryEndEvent
        assert len(events) == 3
        assert isinstance(events[0], QueryStartEvent)
        assert isinstance(events[1], TextEvent)
        assert isinstance(events[2], QueryEndEvent)

    def test_stream_query_handles_exception(self):
        """Test AgentCore.stream_query handles exceptions."""
        from excel_agent.core import AgentCore
        from excel_agent.events import ErrorEvent

        core = AgentCore()
        core._agent = MagicMock()
        core._agent.stream.side_effect = Exception("Test error")

        events = list(core.stream_query("test query"))

        # Should yield an ErrorEvent
        assert len(events) == 2  # QueryStartEvent + ErrorEvent
        assert isinstance(events[1], ErrorEvent)
        assert "Test error" in events[1].error_message

    def test_extract_response_with_output(self):
        """Test _extract_response with 'output' key."""
        from excel_agent.core import AgentCore

        core = AgentCore()
        result = {"output": "Direct output"}

        assert core._extract_response(result) == "Direct output"

    def test_extract_response_with_messages(self):
        """Test _extract_response with messages."""
        from excel_agent.core import AgentCore

        core = AgentCore()
        mock_message = MagicMock()
        mock_message.content = "Message content"
        result = {"messages": [mock_message]}

        assert core._extract_response(result) == "Message content"

    def test_extract_response_with_list_content(self):
        """Test _extract_response with list content."""
        from excel_agent.core import AgentCore

        core = AgentCore()
        mock_message = MagicMock()
        mock_message.content = [
            {"type": "text", "text": "Part 1"},
            {"type": "text", "text": "Part 2"},
        ]
        result = {"messages": [mock_message]}

        assert core._extract_response(result) == "Part 1Part 2"

    def test_extract_response_no_response(self):
        """Test _extract_response with empty result."""
        from excel_agent.core import AgentCore

        core = AgentCore()
        result = {}

        assert core._extract_response(result) == "No response"

    def test_new_session(self):
        """Test new_session generates new thread_id."""
        from excel_agent.core import AgentCore

        core = AgentCore()
        old_thread_id = core.config.thread_id

        new_id = core.new_session()

        assert new_id != old_thread_id
        assert core.config.thread_id == new_id

    def test_new_session_with_custom_id(self):
        """Test new_session with custom thread_id."""
        from excel_agent.core import AgentCore

        core = AgentCore()

        new_id = core.new_session("custom-thread-id")

        assert new_id == "custom-thread-id"
        assert core.config.thread_id == "custom-thread-id"
