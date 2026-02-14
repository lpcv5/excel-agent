"""
Excel Agent - Entry point with UI mode selection.

This module provides the main entry point for Excel Agent with
support for multiple UI modes (CLI by default, WebUI with --web flag).

The agent logic is separated into the excel_agent package, and UI
implementations are in the ui package.
"""

import argparse
from pathlib import Path

from excel_tools import EXCEL_TOOLS


def main():
    """Main entry point for the Excel Agent."""
    parser = argparse.ArgumentParser(
        description="Excel Agent - Process Excel files with natural language",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive CLI mode (default)
  uv run python agent.py

  # Single query
  uv run python agent.py "Read sales.xlsx and show the first 10 rows"

  # Web UI mode
  uv run python agent.py --web

  # With specific model
  uv run python agent.py --model openai:gpt-4 "Analyze data.xlsx"
        """,
    )

    # Positional argument for query
    parser.add_argument(
        "query",
        nargs="?",
        help="Natural language query. If not provided, enters interactive mode.",
    )

    # UI mode options
    parser.add_argument(
        "--web",
        action="store_true",
        help="Use Web UI mode instead of CLI (default is CLI)",
    )

    # Model options
    parser.add_argument(
        "--model",
        "-m",
        default="openai:gpt-5-mini",
        help="LLM model to use (default: openai:gpt-5-mini)",
    )

    # Directory option
    parser.add_argument(
        "--dir",
        "-d",
        default=None,
        help="Working directory (default: current directory)",
    )

    # Utility options
    parser.add_argument(
        "--list-tools",
        "-l",
        action="store_true",
        help="List available Excel tools and exit",
    )
    parser.add_argument(
        "--version",
        "-v",
        action="version",
        version="Excel Agent 0.2.0",
    )

    args = parser.parse_args()

    # Handle --list-tools
    if args.list_tools:
        print("Available Excel Tools:")
        print("=" * 40)
        for tool in EXCEL_TOOLS:
            print(f"  - {tool.name}: {tool.description[:60]}...")
        print()
        return

    # Build configuration
    from excel_agent.config import AgentConfig

    config = AgentConfig(
        model=args.model,
        working_dir=Path(args.dir) if args.dir else Path.cwd(),
    )

    # Select and run UI mode
    if args.web:
        # Web UI mode
        from ui.web.server import run_server

        run_server(config)
    else:
        # CLI mode (default)
        from ui.cli.runner import CLIRunner

        runner = CLIRunner(config)
        if args.query:
            runner.run_single_query(args.query)
        else:
            runner.run()


if __name__ == "__main__":
    main()
