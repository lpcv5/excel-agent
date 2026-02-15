"""CLI implementation using the core layer.

This module provides the command-line interface for Excel Agent,
maintaining backward compatibility with the original agent.py.
"""

import asyncio
import sys
from typing import Optional

from excel_agent.config import AgentConfig
from excel_agent.events import (
    AgentEvent,
    ErrorEvent,
    EventType,
    RefusalEvent,
    TextEvent,
    ThinkingEvent,
    ToolCallArgsEvent,
    ToolCallStartEvent,
    ToolResultEvent,
)
from ui.base import BaseUI
from tools.excel_tool import cleanup_excel_resources


class CLIRunner(BaseUI):
    """Command-line interface runner.

    This class implements the CLI interface using the core layer.
    It maintains backward compatibility with the original agent.py.
    """

    def __init__(self, config: Optional[AgentConfig] = None):
        """Initialize the CLI runner.

        Args:
            config: Agent configuration. Creates default if not provided.
        """
        super().__init__(config)
        self._last_event_type: Optional[EventType] = None
        self._current_tool_name: Optional[str] = None
        self._current_tool_args: str = ""

    async def run(self) -> None:
        """Run interactive CLI mode."""
        print("=" * 60)
        print("  Excel Agent - Interactive Mode")
        print("  Type your questions or commands. 'quit' to exit.")
        print("=" * 60)
        print()

        try:
            while True:
                try:
                    user_input = input("You: ").strip()

                    if not user_input:
                        continue

                    if user_input.lower() in ["quit", "exit", "q"]:
                        print("\nGoodbye!")
                        break

                    # Process streaming events
                    async for event in self.core.astream_query(user_input):
                        self._render_event(event)

                    # Final newline after streaming completes
                    print("\n")

                    # Reset state for next query
                    self._reset_state()

                except KeyboardInterrupt:
                    print("\n\nGoodbye!")
                    break
                except asyncio.CancelledError:
                    print("\n\nGoodbye!")
                    break
                except Exception as e:
                    print(f"\nError: {e}\n")
        finally:
            self._cleanup_excel()

    def run_single_query(self, query: str) -> str:
        """Run a single query and print the result.

        Args:
            query: User's natural language query

        Returns:
            The agent's response
        """
        try:
            result = self.core.invoke(query)
            print(result)
            return result
        finally:
            self._cleanup_excel()

    def _render_event(self, event: AgentEvent) -> None:
        """Render an event to the console.

        Args:
            event: The event to render
        """
        # Force flush stdout before each output to ensure real-time streaming
        sys.stdout.flush()

        if isinstance(event, ThinkingEvent):
            if self._last_event_type != EventType.THINKING:
                print("\n💭 Thinking:", end="", flush=True)
            print(event.content, end="", flush=True)
            self._last_event_type = EventType.THINKING

        elif isinstance(event, ToolCallStartEvent):
            # Print previous tool's args if any
            if self._current_tool_name and self._current_tool_args:
                print(f"   Args: {self._current_tool_args}", flush=True)

            self._current_tool_name = event.tool_name
            # Show tool args if available (but not empty {} or empty string)
            initial_args = event.tool_args or ""
            # Skip if it's just empty braces
            if initial_args in ("", "{}"):
                self._current_tool_args = ""
            else:
                self._current_tool_args = initial_args
            print(f"\n🔧 Calling tool: {event.tool_name}", flush=True)
            # Don't print empty args
            self._last_event_type = EventType.TOOL_CALL_START

        elif isinstance(event, ToolCallArgsEvent):
            self._current_tool_args += event.content

        elif isinstance(event, TextEvent):
            # Print accumulated tool args before text
            if self._current_tool_name and self._current_tool_args:
                print(f"   Args: {self._current_tool_args}", flush=True)
                self._current_tool_args = ""

            if self._last_event_type not in (EventType.TEXT, None):
                print()  # New line before text

            print(event.content, end="", flush=True)
            self._last_event_type = EventType.TEXT

        elif isinstance(event, RefusalEvent):
            print(f"\n⚠️ Refusal: {event.content}", flush=True)

        elif isinstance(event, ToolResultEvent):
            # Print any accumulated tool args before showing result
            if self._current_tool_name and self._current_tool_args:
                print(f"   Args: {self._current_tool_args}", flush=True)
                self._current_tool_args = ""
            # Show tool result
            result_preview = event.content[:200] + "..." if len(event.content) > 200 else event.content
            print(f"\n✅ Tool result: {result_preview}", flush=True)
            self._last_event_type = EventType.TOOL_RESULT

        elif isinstance(event, ErrorEvent):
            print(f"\n⚠️ Error: {event.error_message}", flush=True)

    def _reset_state(self) -> None:
        """Reset state for next query."""
        self._last_event_type = None
        self._current_tool_name = None
        self._current_tool_args = ""

    @staticmethod
    def _cleanup_excel() -> None:
        """Best-effort cleanup of Excel COM resources on exit."""
        try:
            cleanup_excel_resources(force=False)
        except Exception:
            pass
