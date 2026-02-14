"""
Excel Agent - A DeepAgents-based Excel processing agent.

This module provides a CLI interface for interacting with Excel files
using natural language commands. It supports data analysis, formula
generation, and report creation.
"""

import argparse
import sys
from pathlib import Path

from excel_tools import EXCEL_TOOLS

from deepagents import create_deep_agent
from deepagents.backends import FilesystemBackend
from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph.state import CompiledStateGraph


def create_excel_agent(
    model: str = "openai:gpt-5-mini",
) -> CompiledStateGraph:
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
    # Enable virtual_mode so that "/" refers to work_dir, not the drive root
    backend = FilesystemBackend(root_dir=work_dir, virtual_mode=True)

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


def run_interactive(agent: CompiledStateGraph):
    """Run the agent in interactive mode with streaming output."""
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

            # Track state for formatting
            current_tool_name = None
            current_tool_args = ""
            last_content_type = None

            # Stream LLM tokens in real-time using messages mode
            # The messages mode returns (message_chunk, metadata) tuples directly
            for message_chunk, _metadata in agent.stream(
                {"messages": [("user", user_input)]},
                config={"configurable": {"thread_id": thread_id}},
                stream_mode="messages",
            ):
                # Get content from message_chunk
                # message_chunk can be AIMessageChunk with .content attribute
                # or the content itself in some cases
                content = getattr(message_chunk, "content", None)
                if content is None and isinstance(message_chunk, (list, str)):
                    content = message_chunk

                if not content:
                    continue

                # Handle list content (OpenAI Responses API style with type fields)
                if isinstance(content, list):
                    for item in content:
                        if not isinstance(item, dict):
                            # Handle string items in list
                            if isinstance(item, str):
                                if last_content_type != "text":
                                    print()  # New line before text
                                    last_content_type = "text"
                                print(item, end="", flush=True)
                            continue

                        item_type = item.get("type", "")

                        # Handle thinking/reasoning content
                        if item_type == "thinking":
                            thinking_text = item.get("thinking", "")
                            if thinking_text:
                                if last_content_type != "thinking":
                                    print("\n💭 Thinking:", flush=True)
                                    last_content_type = "thinking"
                                print(thinking_text, end="", flush=True)

                        # Handle tool calls
                        elif item_type == "function_call":
                            tool_name = item.get("name", "unknown")
                            # Only print if it's a new tool call
                            if current_tool_name != tool_name:
                                # Print previous tool's args if any
                                if current_tool_name and current_tool_args:
                                    print(f"   Args: {current_tool_args}", flush=True)
                                current_tool_name = tool_name
                                current_tool_args = ""
                                print(f"\n🔧 Calling tool: {tool_name}", flush=True)
                                last_content_type = "tool_call"

                        elif item_type == "function_call_arguments":
                            # Accumulate tool arguments
                            args_chunk = item.get("arguments", "")
                            if args_chunk:
                                current_tool_args += args_chunk

                        # Handle text output
                        elif item_type == "text":
                            text = item.get("text", "")
                            if text:
                                # Print accumulated tool args before text output
                                if current_tool_name and current_tool_args:
                                    print(f"   Args: {current_tool_args}", flush=True)
                                    current_tool_args = ""
                                if last_content_type not in ("text", None):
                                    print()  # New line before text output
                                last_content_type = "text"
                                print(text, end="", flush=True)

                        # Handle refusal (content filter)
                        elif item_type == "refusal":
                            refusal = item.get("refusal", "")
                            if refusal:
                                print(f"\n⚠️ Refusal: {refusal}", flush=True)

                        # Handle unknown types - print raw for debugging
                        elif item_type:
                            # Skip empty type items
                            pass

                # Handle string content (simple case - standard LangChain format)
                elif isinstance(content, str):
                    if content.strip():
                        # Print accumulated tool args before text output
                        if current_tool_name and current_tool_args:
                            print(f"   Args: {current_tool_args}", flush=True)
                            current_tool_args = ""
                        if last_content_type not in ("text", None):
                            print()  # New line before text
                        last_content_type = "text"
                        print(content, end="", flush=True)

            # Print any remaining tool args
            if current_tool_name and current_tool_args:
                print(f"   Args: {current_tool_args}", flush=True)

            print("\n")  # New line after streaming completes

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
