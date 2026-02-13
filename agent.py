"""
Excel Agent - A DeepAgents-based Excel processing agent.

This module provides a CLI interface for interacting with Excel files
using natural language commands. It supports data analysis, formula
generation, and report creation.
"""

import argparse
import sys
from pathlib import Path
from typing import Optional

from excel_tools import EXCEL_TOOLS

# Check if deepagents is available
try:
    from deepagents import create_deep_agent
    from deepagents.backends import FilesystemBackend

    DEEPAGENTS_AVAILABLE = True
except ImportError:
    DEEPAGENTS_AVAILABLE = False
    print("Warning: deepagents package not installed. Running in standalone mode.")


def create_excel_agent(
    model: str = "anthropic:claude-sonnet-4-20250514",
    working_dir: Optional[str] = None,
):
    """
    Create and configure the Excel Agent.

    Args:
        model: The LLM model to use (default: claude-sonnet-4)
        working_dir: Working directory for file operations

    Returns:
        Configured agent instance
    """
    if not DEEPAGENTS_AVAILABLE:
        raise RuntimeError(
            "deepagents package is required. Install it with: pip install deepagents"
        )

    # Determine working directory
    if working_dir is None:
        working_dir = Path.cwd()
    else:
        working_dir = Path(working_dir)

    # Create the agent
    agent = create_deep_agent(
        model=model,
        tools=EXCEL_TOOLS,
        system_prompt=None,  # Will use AGENTS.md
        memory=[str(Path(__file__).parent / "AGENTS.md")],
        skills=[str(Path(__file__).parent / "skills")],
        backend=FilesystemBackend(root_dir=str(working_dir)),
    )

    return agent


def run_interactive(agent):
    """Run the agent in interactive mode."""
    print("=" * 60)
    print("  Excel Agent - Interactive Mode")
    print("  Type your questions or commands. 'quit' to exit.")
    print("=" * 60)
    print()

    while True:
        try:
            user_input = input("You: ").strip()

            if not user_input:
                continue

            if user_input.lower() in ["quit", "exit", "q"]:
                print("\nGoodbye!")
                break

            # Run the agent
            result = agent.invoke({"input": user_input})

            # Print the response
            print(f"\nAgent: {result.get('output', 'No response')}\n")

        except KeyboardInterrupt:
            print("\n\nGoodbye!")
            break
        except Exception as e:
            print(f"\nError: {e}\n")


def run_single_query(agent, query: str):
    """Run a single query and print the result."""
    try:
        result = agent.invoke({"input": query})
        print(result.get("output", "No response"))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    """Main entry point for the Excel Agent CLI."""
    parser = argparse.ArgumentParser(
        description="Excel Agent - Process Excel files with natural language",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode
  uv run python agent.py

  # Single query
  uv run python agent.py "Read sales.xlsx and show the first 10 rows"

  # Analyze data
  uv run python agent.py "Analyze the data in report.xlsx and give me a summary"

  # Generate formula
  uv run python agent.py "Create a formula to calculate 10% bonus for sales over 5000"

  # With specific working directory
  uv run python agent.py --dir ./data "List all sheets in workbook.xlsx"
        """,
    )

    parser.add_argument(
        "query",
        nargs="?",
        help="Natural language query to execute. If not provided, enters interactive mode.",
    )
    parser.add_argument(
        "--model",
        "-m",
        default="anthropic:claude-sonnet-4-20250514",
        help="LLM model to use (default: claude-sonnet-4)",
    )
    parser.add_argument(
        "--dir",
        "-d",
        default=None,
        help="Working directory for file operations (default: current directory)",
    )
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
        version="Excel Agent 0.1.0",
    )

    args = parser.parse_args()

    # List tools mode
    if args.list_tools:
        print("Available Excel Tools:")
        print("=" * 40)
        for tool in EXCEL_TOOLS:
            print(f"  - {tool.name}: {tool.description[:60]}...")
        print()
        return

    # Check for deepagents availability
    if not DEEPAGENTS_AVAILABLE:
        print(
            "Error: The 'deepagents' package is required but not installed.",
            file=sys.stderr,
        )
        print("\nInstall it with:", file=sys.stderr)
        print("  pip install deepagents", file=sys.stderr)
        print("\nOr with uv:", file=sys.stderr)
        print("  uv pip install deepagents", file=sys.stderr)
        sys.exit(1)

    # Create the agent
    try:
        agent = create_excel_agent(
            model=args.model,
            working_dir=args.dir,
        )
    except Exception as e:
        print(f"Error creating agent: {e}", file=sys.stderr)
        sys.exit(1)

    # Run in appropriate mode
    if args.query:
        run_single_query(agent, args.query)
    else:
        run_interactive(agent)


if __name__ == "__main__":
    main()
