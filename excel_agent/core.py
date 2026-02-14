"""Core agent interface - UI-agnostic business logic.

This module provides the AgentCore class, which encapsulates all
agent logic independent of any UI implementation. It provides:
- Agent creation and configuration
- Query execution (streaming and non-streaming)
- Event-based output for UI consumption
"""

import logging
from collections.abc import AsyncGenerator
from typing import Any, Optional

from langgraph.graph.state import CompiledStateGraph
from libs.stream_msg_parser import MessageParser
from libs.stream_msg_parser.events import (
    ContentEvent as ParserContentEvent,
    ErrorEvent as ParserErrorEvent,
    ToolCallArgsEvent as ParserToolCallArgsEvent,
    ToolCallStartEvent as ParserToolCallStartEvent,
    ToolCallEndEvent,
    CompleteEvent,
    StreamEvent,
)

from excel_agent.config import AgentConfig
from excel_agent.events import (
    AgentEvent,
    ErrorEvent,
    QueryEndEvent,
    QueryStartEvent,
    TextEvent,
    ThinkingEvent,
    ToolCallArgsEvent,
    ToolCallStartEvent,
    ToolResultEvent,
)


class AgentCore:
    """UI-agnostic Excel Agent core.

    This class provides a clean interface for:
    - Creating and managing the DeepAgent
    - Executing queries (single or streaming)
    - Emitting events for UI consumption

    Usage:
        config = AgentConfig(model="openai:gpt-5-mini")
        core = AgentCore(config)

        # Streaming mode (async)
        async for event in core.astream_query("Read sales.xlsx"):
            handle_event(event)

        # Single query mode
        result = core.invoke("Read sales.xlsx")
    """

    def __init__(self, config: Optional[AgentConfig] = None):
        """Initialize the agent core.

        Args:
            config: Agent configuration. Uses defaults if not provided.
        """
        self.config = config or AgentConfig()
        self._agent: Optional[CompiledStateGraph] = None
        self._parser = MessageParser(track_tool_lifecycle=True)
        self._tool_args_buffer: dict[str, str] = {}  # Buffer for accumulating tool args
        self._logger: Optional[logging.Logger] = self._setup_logger()

    def _setup_logger(self) -> Optional[logging.Logger]:
        """Set up logger for LLM call logging.

        Returns:
            Configured logger instance or None if logging is disabled
        """
        if not self.config.logging.enabled:
            return None

        from excel_agent.logging_config import setup_logging

        return setup_logging(self.config.logging)

    def _log_stream_event(self, event: StreamEvent, elapsed_ms: Optional[float] = None) -> None:
        """Log a stream event.

        Args:
            event: The stream event to log
            elapsed_ms: Optional elapsed time in milliseconds
        """
        if self._logger is None:
            return

        import json
        from datetime import datetime

        log_data: dict[str, Any] = {
            "event": event.__class__.__name__,
            "timestamp": datetime.now().isoformat(),
        }

        if elapsed_ms is not None:
            log_data["elapsed_ms"] = round(elapsed_ms, 2)

        # Extract event-specific data
        if isinstance(event, ParserContentEvent):
            log_data["node"] = getattr(event, "node", None)
            log_data["content"] = event.content
        elif isinstance(event, ParserToolCallStartEvent):
            log_data["tool_name"] = event.name
            log_data["args"] = event.args
        elif isinstance(event, ToolCallEndEvent):
            log_data["tool_call_id"] = getattr(event, "id", None)
            log_data["tool_name"] = getattr(event, "name", None)
            log_data["status"] = getattr(event, "status", None)
            log_data["duration_ms"] = getattr(event, "duration_ms", None)
            # Truncate result for readability
            result = getattr(event, "result", None)
            if result:
                result_str = str(result)
                log_data["result"] = result_str[:500] + "..." if len(result_str) > 500 else result_str
            if hasattr(event, "error_message") and event.error_message:
                log_data["error_message"] = event.error_message
        elif isinstance(event, ParserErrorEvent):
            log_data["error"] = event.error

        self._logger.debug(
            f"Stream Event:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

    @property
    def agent(self) -> CompiledStateGraph:
        """Lazy-load the DeepAgent."""
        if self._agent is None:
            self._agent = self._create_agent()
        return self._agent

    def _create_agent(self) -> CompiledStateGraph:
        """Create the DeepAgent with Excel tools."""
        from tools.excel_tool import EXCEL_TOOLS

        from deepagents import create_deep_agent
        from deepagents.backends import FilesystemBackend
        from langgraph.checkpoint.memory import MemorySaver

        memory_paths = (
            [str(self.config.agents_md_path)]
            if self.config.agents_md_path and self.config.agents_md_path.exists()
            else None
        )

        skills_paths = (
            [str(self.config.skills_path)]
            if self.config.skills_path and self.config.skills_path.exists()
            else None
        )

        working_dir_str = str(self.config.working_dir.resolve())
        system_prompt = f"""## Working Directory

Your working directory is: `{working_dir_str}`

All file operations are relative to this directory. When the user asks about files
or wants to work with Excel files, use paths relative to this directory or absolute paths.
"""

        return create_deep_agent(
            model=self.config.model,
            tools=EXCEL_TOOLS,
            system_prompt=system_prompt,
            memory=memory_paths,
            skills=skills_paths,
            backend=FilesystemBackend(root_dir=self.config.working_dir),
            checkpointer=MemorySaver(),
        )

    async def astream_query(
        self,
        query: str,
        thread_id: Optional[str] = None,
    ) -> AsyncGenerator[AgentEvent, None]:
        """Execute a query and yield events asynchronously.

        This is the primary method for UI consumption. It uses
        langgraph-stream-parser to normalize streaming output.

        Args:
            query: User's natural language query
            thread_id: Optional thread ID for conversation continuity

        Yields:
            AgentEvent instances representing streaming output
        """
        import time

        thread = thread_id or self.config.thread_id
        start_time = time.perf_counter()

        # Yield query start event
        yield QueryStartEvent(content=query)

        try:
            # Use aiter to convert async stream for MessageParser
            stream = self.agent.astream(
                {"messages": [("user", query)]},
                config={"configurable": {"thread_id": thread}},
                stream_mode="messages",
            )

            # Use async parse
            async for parser_event in self._parser.aparse(stream):
                # Log the stream event
                elapsed = (time.perf_counter() - start_time) * 1000
                self._log_stream_event(parser_event, elapsed)

                # Convert parser events to our event types
                event = self._convert_event(parser_event)
                if event:
                    yield event

            # Yield query end event
            yield QueryEndEvent()

        except Exception as e:
            yield ErrorEvent(error_message=str(e))

    def _convert_event(self, parser_event: StreamEvent) -> Optional[AgentEvent]:
        """Convert langgraph-stream-parser event to our event type."""
        if isinstance(parser_event, ParserContentEvent):
            content = parser_event.content

            # Strip reasoning summary prefix if present
            # Format: {'id': 'rs_...', 'summary': [], 'type': 'reasoning'} actual text
            if isinstance(content, str) and content.startswith("{'id':"):
                import re
                # Check if this is a reasoning block
                if "'type': 'reasoning'" in content or '"type": "reasoning"' in content:
                    match = re.match(r"^\{[^}]*\}\s*(.*)$", content)
                    if match:
                        extracted = match.group(1)
                        # If there's actual content after the reasoning block, use it
                        if extracted and extracted.strip():
                            content = extracted
                        else:
                            return None  # Skip reasoning-only content
                    else:
                        return None  # Can't parse, skip

            # Skip content that is just tool call arguments dict
            # Format: {'arguments': '...', 'call_id': '...', 'name': '...'}
            if isinstance(content, str) and content.startswith("{'arguments':"):
                return None

            # Skip empty content (but preserve whitespace/newlines)
            if not content:
                return None

            # Check for thinking/reasoning content
            if hasattr(parser_event, "node") and parser_event.node == "thinking":
                return ThinkingEvent(content=content)
            return TextEvent(content=content)

        elif isinstance(parser_event, ParserToolCallStartEvent):
            import json as json_mod
            return ToolCallStartEvent(
                tool_name=parser_event.name,
                tool_args=json_mod.dumps(parser_event.args)
            )

        elif isinstance(parser_event, ParserToolCallArgsEvent):
            # Tool call args chunk - emit with the partial args
            import json as json_mod
            # Accumulate the args
            existing_args = getattr(self, '_tool_args_buffer', {})
            tool_id = parser_event.id
            if tool_id:
                existing_args[tool_id] = existing_args.get(tool_id, '') + parser_event.args
                self._tool_args_buffer = existing_args
            return ToolCallArgsEvent(
                tool_name=parser_event.name,
                content=parser_event.args,
            )

        elif isinstance(parser_event, ToolCallEndEvent):
            # Tool completed - emit result event
            # Check for error status
            if parser_event.status == "error":
                return ErrorEvent(
                    error_message=parser_event.error_message or "Tool call failed"
                )
            return ToolResultEvent(
                tool_name=parser_event.name,
                content=str(parser_event.result)[:1000] if parser_event.result else "",
                data={"status": parser_event.status, "duration_ms": parser_event.duration_ms}
            )

        elif isinstance(parser_event, ParserErrorEvent):
            return ErrorEvent(error_message=parser_event.error)

        elif isinstance(parser_event, CompleteEvent):
            # Stream completed - handled by QueryEndEvent
            return None

        return None

    def invoke(
        self,
        query: str,
        thread_id: Optional[str] = None,
    ) -> str:
        """Execute a single query and return the result.

        This is a convenience method for non-streaming use cases.

        Args:
            query: User's natural language query
            thread_id: Optional thread ID for conversation continuity

        Returns:
            The agent's response as a string
        """
        thread = thread_id or self.config.thread_id

        result = self.agent.invoke(
            {"messages": [("user", query)]},
            config={"configurable": {"thread_id": thread}},
        )

        return self._extract_response(result)

    def _extract_response(self, result: dict) -> str:
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

    def get_excel_status(self) -> dict:
        """Get current Excel application status.

        Returns:
            Dict with excel_running, workbook_count, open_workbooks
        """
        import json

        from tools.excel_tool import excel_status

        result = excel_status.invoke({})
        return json.loads(result)

    def new_session(self, thread_id: Optional[str] = None) -> str:
        """Start a new conversation session.

        Args:
            thread_id: New thread ID for the session. Generated if not provided.

        Returns:
            The new thread ID
        """
        import uuid

        self.config.thread_id = thread_id or str(uuid.uuid4())
        return self.config.thread_id
