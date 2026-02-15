"""
Excel Agent - Entry point with UI mode selection.

This module provides the main entry point for Excel Agent with
support for multiple UI modes (CLI by default, WebUI with --web flag).

The agent logic is separated into the excel_agent package, and UI
implementations are in the ui package.
"""

import argparse
import asyncio

from tools.excel_tool import EXCEL_TOOLS


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

  # With Zhipu GLM (default)
  uv run python agent.py --model zhipu:glm-5 "Analyze data.xlsx"

  # With OpenAI
  uv run python agent.py --model openai:gpt-4o "Analyze data.xlsx"

  # With DeepSeek
  uv run python agent.py --model deepseek:deepseek-chat "Analyze data.xlsx"

  # With Moonshot
  uv run python agent.py --model moonshot:moonshot-v1-8k "Analyze data.xlsx"

  # With custom provider (requires CUSTOM_API_KEY and CUSTOM_API_BASE env vars)
  uv run python agent.py --model custom:my-model "Analyze data.xlsx"

  # With LLM call logging (for debugging)
  uv run python agent.py --log-level DEBUG "Read data.xlsx"

  # With logging to file
  uv run python agent.py --log-level DEBUG --log-file logs/llm.log "Analyze data.xlsx"

Supported Providers:
  zhipu     - Zhipu GLM (glm-5, glm-4, etc.)
  openai    - OpenAI (gpt-4o, gpt-4o-mini, etc.)
  deepseek  - DeepSeek (deepseek-chat, deepseek-reasoner, etc.)
  moonshot  - Moonshot Kimi (moonshot-v1-8k, moonshot-v1-32k, etc.)
  custom    - Custom OpenAI-compatible API

Environment Variables:
  ZAI_API_KEY       - API key for Zhipu GLM (default provider)
  OPENAI_API_KEY    - API key for OpenAI
  DEEPSEEK_API_KEY  - API key for DeepSeek
  MOONSHOT_API_KEY  - API key for Moonshot
  CUSTOM_API_KEY    - API key for custom provider
  CUSTOM_API_BASE   - API base URL for custom provider
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
        default="zhipu:glm-5",
        help="LLM model to use, format: provider:model (default: zhipu:glm-5)",
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

    # Logging options
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default=None,
        help="LLM call logging level (default: no logging)",
    )
    parser.add_argument(
        "--log-file",
        type=str,
        default=None,
        help="Log file path (default: console only)",
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

    config = AgentConfig.from_cli_args(args)

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
            try:
                asyncio.run(runner.run())
            except KeyboardInterrupt:
                pass  # User exited via Ctrl+C, already handled in runner


if __name__ == "__main__":
    main()
