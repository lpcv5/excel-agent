"""Core agent interface - UI-agnostic business logic.

This module provides the AgentCore class, which encapsulates all
agent logic independent of any UI implementation. It provides:
- Agent creation and configuration
- Query execution (streaming and non-streaming)
- Event-based output for UI consumption
"""

from collections.abc import Generator
from typing import Optional

from langgraph.graph.state import CompiledStateGraph

from excel_agent.config import AgentConfig
from excel_agent.events import (
    AgentEvent,
    ErrorEvent,
    EventType,
    QueryEndEvent,
    QueryStartEvent,
    RefusalEvent,
    TextEvent,
    ThinkingEvent,
    ToolCallArgsEvent,
    ToolCallStartEvent,
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

        # Streaming mode
        for event in core.stream_query("Read sales.xlsx"):
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

    @property
    def agent(self) -> CompiledStateGraph:
        """Lazy-load the DeepAgent."""
        if self._agent is None:
            self._agent = self._create_agent()
        return self._agent

    def _create_agent(self) -> CompiledStateGraph:
        """Create the DeepAgent with Excel tools."""
        from excel_tools import EXCEL_TOOLS

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

    def stream_query(
        self,
        query: str,
        thread_id: Optional[str] = None,
    ) -> Generator[AgentEvent, None, None]:
        """Execute a query and yield events.

        This is the primary method for UI consumption. It yields
        typed events that can be rendered by any UI implementation.

        Args:
            query: User's natural language query
            thread_id: Optional thread ID for conversation continuity

        Yields:
            AgentEvent instances representing streaming output
        """
        thread = thread_id or self.config.thread_id

        # Yield query start event
        yield QueryStartEvent(content=query)

        # Track state for event generation
        current_tool_name: Optional[str] = None
        last_event_type: Optional[EventType] = None

        try:
            for message_chunk, _metadata in self.agent.stream(
                {"messages": [("user", query)]},
                config={"configurable": {"thread_id": thread}},
                stream_mode="messages",
            ):
                content = getattr(message_chunk, "content", None)
                if content is None and isinstance(message_chunk, (list, str)):
                    content = message_chunk

                if not content:
                    continue

                # Process content and yield appropriate events
                for event in self._process_content(
                    content,
                    current_tool_name,
                    last_event_type,
                ):
                    # Update tracking state
                    if isinstance(event, ToolCallStartEvent):
                        current_tool_name = event.tool_name
                    elif event.type == EventType.TEXT:
                        current_tool_name = None

                    last_event_type = event.type
                    yield event

            # Yield query end event
            yield QueryEndEvent()

        except Exception as e:
            yield ErrorEvent(error_message=str(e))

    def _process_content(
        self,
        content,
        current_tool_name: Optional[str],
        last_event_type: Optional[EventType],
    ) -> Generator[AgentEvent, None, None]:
        """Process message content and yield events."""
        if isinstance(content, list):
            # OpenAI Responses API style with type fields
            for item in content:
                if not isinstance(item, dict):
                    if isinstance(item, str):
                        yield TextEvent(content=item)
                    continue

                item_type = item.get("type", "")

                if item_type == "thinking":
                    thinking_text = item.get("thinking", "")
                    if thinking_text:
                        yield ThinkingEvent(content=thinking_text)

                elif item_type == "function_call":
                    tool_name = item.get("name", "unknown")
                    yield ToolCallStartEvent(tool_name=tool_name)

                elif item_type == "function_call_arguments":
                    args_chunk = item.get("arguments", "")
                    if args_chunk and current_tool_name:
                        yield ToolCallArgsEvent(
                            tool_name=current_tool_name,
                            content=args_chunk,
                        )

                elif item_type == "text":
                    text = item.get("text", "")
                    if text:
                        yield TextEvent(content=text)

                elif item_type == "refusal":
                    refusal = item.get("refusal", "")
                    yield RefusalEvent(content=refusal)

        elif isinstance(content, str) and content.strip():
            yield TextEvent(content=content)

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

        from excel_tools import excel_status

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
