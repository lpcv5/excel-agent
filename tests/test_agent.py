"""Tests for agent.py entry point."""

from unittest.mock import MagicMock, patch
import pytest


class TestMain:
    """Tests for main function."""

    @patch("sys.argv", ["agent.py"])
    @patch("ui.cli.runner.CLIRunner")
    def test_main_cli_mode(self, mock_runner_class):
        """Test main runs in CLI mode by default."""
        mock_runner = MagicMock()
        mock_runner_class.return_value = mock_runner

        from agent import main

        main()

        mock_runner.run.assert_called_once()

    @patch("sys.argv", ["agent.py", "test query"])
    @patch("ui.cli.runner.CLIRunner")
    def test_main_cli_single_query(self, mock_runner_class):
        """Test main runs single query in CLI mode."""
        mock_runner = MagicMock()
        mock_runner_class.return_value = mock_runner

        from agent import main

        main()

        mock_runner.run_single_query.assert_called_once_with("test query")

    @patch("sys.argv", ["agent.py", "--web"])
    @patch("ui.web.server.run_server")
    def test_main_web_mode(self, mock_run_server):
        """Test main runs in Web UI mode with --web flag."""
        from agent import main

        main()

        mock_run_server.assert_called_once()

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

    @patch("sys.argv", ["agent.py", "--version"])
    def test_main_version(self, capsys):
        """Test main with --version flag."""
        from agent import main

        with pytest.raises(SystemExit):
            main()

        captured = capsys.readouterr()
        assert "0.2.0" in captured.out

    @patch("sys.argv", ["agent.py", "--model", "anthropic:claude-3-opus", "query"])
    @patch("ui.cli.runner.CLIRunner")
    def test_main_custom_model(self, mock_runner_class):
        """Test main with custom model."""
        mock_runner = MagicMock()
        mock_runner_class.return_value = mock_runner

        from agent import main

        main()

        # Verify the runner was created (config is passed)
        mock_runner_class.assert_called_once()
