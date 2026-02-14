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

from deepagents import create_deep_agent
from deepagents.backends import FilesystemBackend
from langgraph.checkpoint.memory import MemorySaver


def create_excel_agent(
    model: str = "openai:gpt-5-mini",
):
    """
    Create and configure the Excel Agent using DeepAgents harness.

    Args:
        model: The LLM model to use (default: openai:gpt-5-mini)
        working_dir: Working directory for file operations

    Returns:
        Configured deep agent instance (CompiledStateGraph)
    """
    work_dir = Path.cwd()

    # Get paths for memory and skills
    base_path = Path(__file__).parent
    agents_md_path = base_path / "AGENTS.md"
    skills_path = base_path / "skills"

    # Configure memory (AGENTS.md)
    memory_paths = []
    if agents_md_path.exists():
        memory_paths.append(str(agents_md_path))

    # Configure skills directory
    skills_paths = []
    if skills_path.exists():
        skills_paths.append(str(skills_path))

    # Create the filesystem backend for file operations
    backend = FilesystemBackend(root_dir=work_dir)

    # Create checkpointer for conversation state persistence
    checkpointer = MemorySaver()

    # Create the deep agent with the DeepAgents harness
    agent = create_deep_agent(
        model=model,
        tools=EXCEL_TOOLS,
        memory=memory_paths if memory_paths else None,
        skills=skills_paths if skills_paths else None,
        backend=backend,
        checkpointer=checkpointer,
    )

    return agent


def extract_response(result: dict) -> str:
    """Extract response text from agent result."""
    if "output" in result:
        return result["output"]
    elif "messages" in result and result["messages"]:
        last_message = result["messages"][-1]
        content = last_message.content
        if isinstance(content, str):
            return content
        elif isinstance(content, list):
            text_parts = []
            for item in content:
                if isinstance(item, dict) and item.get("type") == "text":
                    text_parts.append(item.get("text", ""))
                elif isinstance(item, str):
                    text_parts.append(item)
            return "".join(text_parts)
        return str(content)
    return "No response"


def run_interactive(agent):
    """Run the agent in interactive mode."""
    print("=" * 60)
    print("  Excel Agent - Interactive Mode")
    print("  Type your questions or commands. 'quit' to exit.")
    print("=" * 60)
    print()

    # Use a thread_id for conversation persistence
    thread_id = "interactive-session"

    while True:
        try:
            user_input = input("You: ").strip()

            if not user_input:
                continue

            if user_input.lower() in ["quit", "exit", "q"]:
                print("\nGoodbye!")
                break

            # Run the agent with configurable thread_id for state persistence
            result = agent.invoke(
                {"messages": [("user", user_input)]},
                config={"configurable": {"thread_id": thread_id}},
            )

            # Print the response
            print(f"\nAgent: {extract_response(result)}\n")

        except KeyboardInterrupt:
            print("\n\nGoodbye!")
            break
        except Exception as e:
            print(f"\nError: {e}\n")


def run_single_query(agent, query: str):
    """Run a single query and print the result."""
    try:
        result = agent.invoke(
            {"messages": [("user", query)]},
            config={"configurable": {"thread_id": "single-query"}},
        )
        print(extract_response(result))
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
        default="openai:gpt-5-mini",
        help="LLM model to use (default: openai:gpt-5-mini)",
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

    # Create the agent
    try:
        agent = create_excel_agent(
            model=args.model,
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
