"""Tests for agent.py."""

import sys
from unittest.mock import MagicMock, patch, PropertyMock

import pytest


class TestCreateExcelAgent:
    """Tests for create_excel_agent function."""

    def test_create_excel_agent_raises_when_deepagents_not_available(self):
        """Test create_excel_agent raises when deepagents not available."""
        # Import the module first
        import agent

        # Save the original value
        original_value = agent.DEEPAGENTS_AVAILABLE

        try:
            # Set the module-level variable
            agent.DEEPAGENTS_AVAILABLE = False

            with pytest.raises(RuntimeError, match="deepagents package is required"):
                agent.create_excel_agent()
        finally:
            # Restore the original value
            agent.DEEPAGENTS_AVAILABLE = original_value


class TestRunInteractive:
    """Tests for run_interactive function."""

    @patch("builtins.input", side_effect=["quit"])
    @patch("builtins.print")
    def test_run_interactive_quit(self, mock_print, mock_input):
        """Test run_interactive exits on quit command."""
        from agent import run_interactive

        mock_agent = MagicMock()
        run_interactive(mock_agent)

        # Should print goodbye message
        assert any("Goodbye" in str(call) for call in mock_print.call_args_list)

    @patch("builtins.input", side_effect=["exit"])
    @patch("builtins.print")
    def test_run_interactive_exit(self, mock_print, mock_input):
        """Test run_interactive exits on exit command."""
        from agent import run_interactive

        mock_agent = MagicMock()
        run_interactive(mock_agent)

        assert any("Goodbye" in str(call) for call in mock_print.call_args_list)

    @patch("builtins.input", side_effect=["", "quit"])
    @patch("builtins.print")
    def test_run_interactive_empty_input(self, mock_print, mock_input):
        """Test run_interactive handles empty input."""
        from agent import run_interactive

        mock_agent = MagicMock()
        run_interactive(mock_agent)

        # Empty input should be skipped, agent.invoke should not be called
        mock_agent.invoke.assert_not_called()

    @patch("builtins.input", side_effect=["hello", "quit"])
    @patch("builtins.print")
    def test_run_interactive_processes_query(self, mock_print, mock_input):
        """Test run_interactive processes queries."""
        from agent import run_interactive

        mock_agent = MagicMock()
        mock_agent.invoke.return_value = {"output": "Response text"}

        run_interactive(mock_agent)

        mock_agent.invoke.assert_called_once_with({"input": "hello"})

    @patch("builtins.input", side_effect=KeyboardInterrupt())
    @patch("builtins.print")
    def test_run_interactive_keyboard_interrupt(self, mock_print, mock_input):
        """Test run_interactive handles KeyboardInterrupt."""
        from agent import run_interactive

        mock_agent = MagicMock()
        run_interactive(mock_agent)

        # Should exit gracefully
        assert any("Goodbye" in str(call) for call in mock_print.call_args_list)

    @patch("builtins.input", side_effect=["query", "quit"])
    @patch("builtins.print")
    def test_run_interactive_handles_exception(self, mock_print, mock_input):
        """Test run_interactive handles exceptions."""
        from agent import run_interactive

        mock_agent = MagicMock()
        mock_agent.invoke.side_effect = Exception("Test error")

        run_interactive(mock_agent)

        # Should print error message
        assert any("Error" in str(call) for call in mock_print.call_args_list)


class TestRunSingleQuery:
    """Tests for run_single_query function."""

    @patch("builtins.print")
    def test_run_single_query_success(self, mock_print):
        """Test run_single_query with successful response."""
        from agent import run_single_query

        mock_agent = MagicMock()
        mock_agent.invoke.return_value = {"output": "Result text"}

        run_single_query(mock_agent, "test query")

        mock_agent.invoke.assert_called_once_with({"input": "test query"})
        mock_print.assert_called_with("Result text")

    @patch("builtins.print")
    def test_run_single_query_no_output(self, mock_print):
        """Test run_single_query with no output."""
        from agent import run_single_query

        mock_agent = MagicMock()
        mock_agent.invoke.return_value = {}

        run_single_query(mock_agent, "test query")

        mock_print.assert_called_with("No response")

    @patch("builtins.print")
    @patch("sys.exit")
    def test_run_single_query_exception(self, mock_exit, mock_print):
        """Test run_single_query handles exceptions."""
        from agent import run_single_query

        mock_agent = MagicMock()
        mock_agent.invoke.side_effect = Exception("Test error")

        run_single_query(mock_agent, "test query")

        mock_exit.assert_called_once_with(1)


class TestMain:
    """Tests for main function."""

    @patch("agent.DEEPAGENTS_AVAILABLE", True)
    @patch("agent.create_excel_agent")
    @patch("agent.run_interactive")
    @patch("sys.argv", ["agent.py"])
    def test_main_interactive_mode(self, mock_run_interactive, mock_create_agent):
        """Test main runs in interactive mode when no query provided."""
        mock_agent_instance = MagicMock()
        mock_create_agent.return_value = mock_agent_instance

        from agent import main
        main()

        mock_run_interactive.assert_called_once_with(mock_agent_instance)

    @patch("agent.DEEPAGENTS_AVAILABLE", True)
    @patch("agent.create_excel_agent")
    @patch("agent.run_single_query")
    @patch("sys.argv", ["agent.py", "test query"])
    def test_main_single_query(self, mock_run_single, mock_create_agent):
        """Test main runs single query when query provided."""
        mock_agent_instance = MagicMock()
        mock_create_agent.return_value = mock_agent_instance

        from agent import main
        main()

        mock_run_single.assert_called_once_with(mock_agent_instance, "test query")

    @patch("agent.DEEPAGENTS_AVAILABLE", True)
    @patch("agent.create_excel_agent")
    @patch("agent.run_single_query")
    @patch("sys.argv", ["agent.py", "--model", "anthropic:claude-3-opus", "query"])
    def test_main_custom_model(self, mock_run_single, mock_create_agent):
        """Test main with custom model."""
        mock_agent_instance = MagicMock()
        mock_create_agent.return_value = mock_agent_instance

        from agent import main
        main()

        mock_create_agent.assert_called_once()
        call_kwargs = mock_create_agent.call_args[1]
        assert call_kwargs["model"] == "anthropic:claude-3-opus"

    @patch("agent.DEEPAGENTS_AVAILABLE", True)
    @patch("agent.create_excel_agent")
    @patch("agent.run_single_query")
    @patch("sys.argv", ["agent.py", "--dir", "/custom/dir", "query"])
    def test_main_custom_working_dir(self, mock_run_single, mock_create_agent):
        """Test main with custom working directory."""
        mock_agent_instance = MagicMock()
        mock_create_agent.return_value = mock_agent_instance

        from agent import main
        main()

        call_kwargs = mock_create_agent.call_args[1]
        assert call_kwargs["working_dir"] == "/custom/dir"

    @patch("agent.EXCEL_TOOLS", [MagicMock(name="tool1"), MagicMock(name="tool2")])
    @patch("builtins.print")
    @patch("sys.argv", ["agent.py", "--list-tools"])
    def test_main_list_tools(self, mock_print):
        """Test main with --list-tools flag."""
        from agent import main

        main()

        # Should print tool names
        printed_calls = [str(call) for call in mock_print.call_args_list]
        assert any("Available Excel Tools" in str(call) for call in printed_calls)
