"""Tests for ui/cli/runner.py."""

import asyncio
from unittest.mock import MagicMock, patch


class TestCLIRunner:
    """Tests for CLIRunner class."""

    @patch("builtins.input", side_effect=["quit"])
    @patch("builtins.print")
    def test_run_quit(self, mock_print, mock_input):
        """Test CLIRunner exits on quit command."""
        from ui.cli.runner import CLIRunner

        runner = CLIRunner()
        asyncio.run(runner.run())

        # Should print goodbye message
        assert any("Goodbye" in str(call) for call in mock_print.call_args_list)

    @patch("builtins.input", side_effect=["exit"])
    @patch("builtins.print")
    def test_run_exit(self, mock_print, mock_input):
        """Test CLIRunner exits on exit command."""
        from ui.cli.runner import CLIRunner

        runner = CLIRunner()
        asyncio.run(runner.run())

        assert any("Goodbye" in str(call) for call in mock_print.call_args_list)

    @patch("builtins.input", side_effect=["q"])
    @patch("builtins.print")
    def test_run_q(self, mock_print, mock_input):
        """Test CLIRunner exits on 'q' command."""
        from ui.cli.runner import CLIRunner

        runner = CLIRunner()
        asyncio.run(runner.run())

        assert any("Goodbye" in str(call) for call in mock_print.call_args_list)

    @patch("builtins.input", side_effect=["", "quit"])
    @patch("builtins.print")
    def test_run_empty_input(self, mock_print, mock_input):
        """Test CLIRunner handles empty input."""
        from ui.cli.runner import CLIRunner

        runner = CLIRunner()
        # Mock the core to prevent actual agent creation
        runner.core = MagicMock()
        asyncio.run(runner.run())

        # Empty input should be skipped, astream_query should not be called
        runner.core.astream_query.assert_not_called()

    @patch("builtins.input", side_effect=["hello", "quit"])
    @patch("builtins.print")
    def test_run_processes_query(self, mock_print, mock_input):
        """Test CLIRunner processes queries using astream_query."""
        from ui.cli.runner import CLIRunner

        runner = CLIRunner()
        runner.core = MagicMock()
        # astream_query() returns an async iterator of events
        runner.core.astream_query = MagicMock(return_value=iter([]))

        asyncio.run(runner.run())

        # Verify astream_query was called
        runner.core.astream_query.assert_called_once_with("hello")

    @patch("builtins.input", side_effect=KeyboardInterrupt())
    @patch("builtins.print")
    def test_run_keyboard_interrupt(self, mock_print, mock_input):
        """Test CLIRunner handles KeyboardInterrupt."""
        from ui.cli.runner import CLIRunner

        runner = CLIRunner()
        asyncio.run(runner.run())

        # Should exit gracefully
        assert any("Goodbye" in str(call) for call in mock_print.call_args_list)

    @patch("builtins.input", side_effect=["query", "quit"])
    @patch("builtins.print")
    def test_run_handles_exception(self, mock_print, mock_input):
        """Test CLIRunner handles exceptions during streaming."""
        from ui.cli.runner import CLIRunner

        runner = CLIRunner()
        runner.core = MagicMock()
        # Make astream_query raise an exception
        runner.core.astream_query.side_effect = Exception("Test error")

        asyncio.run(runner.run())

        # Should print error message
        assert any("Error" in str(call) for call in mock_print.call_args_list)


class TestCLIRunnerSingleQuery:
    """Tests for CLIRunner.run_single_query method."""

    @patch("builtins.print")
    def test_run_single_query_success(self, mock_print):
        """Test run_single_query with successful response."""
        from ui.cli.runner import CLIRunner

        runner = CLIRunner()
        runner.core = MagicMock()
        runner.core.invoke.return_value = "Result text"

        result = runner.run_single_query("test query")

        runner.core.invoke.assert_called_once_with("test query")
        mock_print.assert_called_with("Result text")
        assert result == "Result text"

    @patch("builtins.print")
    def test_run_single_query_with_config(self, mock_print):
        """Test run_single_query with custom config."""
        from ui.cli.runner import CLIRunner
        from excel_agent.config import AgentConfig

        config = AgentConfig(model="test-model")
        runner = CLIRunner(config)

        assert runner.config.model == "test-model"


class TestCLIRunnerRenderEvent:
    """Tests for CLIRunner._render_event method."""

    def test_render_thinking_event(self):
        """Test rendering thinking event."""
        from ui.cli.runner import CLIRunner
        from excel_agent.events import ThinkingEvent

        runner = CLIRunner()

        with patch("builtins.print") as mock_print:
            event = ThinkingEvent(content="thinking...")
            runner._render_event(event)

            # Should print thinking indicator
            assert any("Thinking" in str(call) for call in mock_print.call_args_list)

    def test_render_tool_call_start_event(self):
        """Test rendering tool call start event."""
        from ui.cli.runner import CLIRunner
        from excel_agent.events import ToolCallStartEvent

        runner = CLIRunner()

        with patch("builtins.print") as mock_print:
            event = ToolCallStartEvent(tool_name="excel_read_range")
            runner._render_event(event)

            # Should print tool call indicator
            assert any("Calling tool" in str(call) for call in mock_print.call_args_list)

    def test_render_text_event(self):
        """Test rendering text event."""
        from ui.cli.runner import CLIRunner
        from excel_agent.events import TextEvent

        runner = CLIRunner()

        with patch("builtins.print") as mock_print:
            event = TextEvent(content="Hello world")
            runner._render_event(event)

            # Should print text content
            assert any("Hello world" in str(call) for call in mock_print.call_args_list)

    def test_render_error_event(self):
        """Test rendering error event."""
        from ui.cli.runner import CLIRunner
        from excel_agent.events import ErrorEvent

        runner = CLIRunner()

        with patch("builtins.print") as mock_print:
            event = ErrorEvent(error_message="Test error")
            runner._render_event(event)

            # Should print error indicator
            assert any("Error" in str(call) for call in mock_print.call_args_list)
