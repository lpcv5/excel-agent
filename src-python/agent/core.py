"""Core agent interface - UI-agnostic business logic."""

import logging
from collections.abc import AsyncGenerator
from pathlib import Path
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

from agent.config import AgentConfig
from agent.events import (
    AgentEvent,
    ErrorEvent,
    QueryEndEvent,
    QueryStartEvent,
    TextEvent,
    ThinkingEvent,
    ToolCallArgsEvent,
    ToolCallStartEvent,
    ToolResultEvent,
    TodoUpdateEvent,
)


class AgentCore:
    def __init__(self, config: Optional[AgentConfig] = None):
        self.config = config or AgentConfig()
        self._agent: Optional[CompiledStateGraph] = None
        self._parser = MessageParser(track_tool_lifecycle=True)
        self._tool_args_buffer: dict[str, str] = {}
        self._todo_args_cache: dict[str, str] = {}
        self._logger: Optional[logging.Logger] = self._setup_logger()
        self._cancelled: bool = False

    def _setup_logger(self) -> Optional[logging.Logger]:
        if not self.config.logging.enabled:
            return None
        from agent.context import get_app_context

        ctx = get_app_context()
        if ctx.initialized and ctx.logger:
            return ctx.logger
        from agent.logging_config import setup_logging

        return setup_logging(self.config.logging)

    @property
    def agent(self) -> CompiledStateGraph:
        if self._agent is None:
            self._agent = self._create_agent()
        return self._agent

    def _create_agent(self) -> CompiledStateGraph:
        from libs.deepagents import create_deep_agent
        from libs.deepagents.backends import FilesystemBackend
        from langgraph.checkpoint.memory import MemorySaver

        tools = [t for p in self.config.tool_providers for t in p.get_tools()]

        # Inject schema tool if project root is set
        if self.config.working_dir and self.config.working_dir != Path.cwd():
            from tools.schema_tool import SchemaToolProvider

            tools += SchemaToolProvider(self.config.working_dir).get_tools()

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

Use absolute Windows paths under this directory when working with files.

## Data Source Schema

This project may have pre-analyzed data source schemas available.
- Call `read_datasource_schema()` (no args) to get a summary of all data sources before starting any data task.
- Call `read_datasource_schema(source_name="filename.xlsx")` to get the full schema for a specific source.
- Always check the schema before reading or writing data — it tells you sheet names, column names, and data types.
"""
        model = self.config.get_model_instance()

        if self._logger:
            from agent.callbacks.llm_logger import LLMLoggingCallbackHandler

            model = model.with_config(
                callbacks=[
                    LLMLoggingCallbackHandler(
                        self._logger,
                        log_full_prompt=self.config.logging.log_full_prompt,
                        log_token_usage=self.config.logging.log_token_usage,
                        log_timing=self.config.logging.log_timing,
                    )
                ]
            )

        return create_deep_agent(
            model=model,  # type: ignore[arg-type]
            tools=tools,
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

        self._cancelled = False
        thread = thread_id or self.config.thread_id
        yield QueryStartEvent(content=query)
        try:
            stream = self.agent.astream(
                {"messages": [("user", query)]},
                config={"configurable": {"thread_id": thread}},
                stream_mode="messages",
            )
            async for parser_event in self._parser.aparse(stream):  # type: ignore[arg-type]
                if self._cancelled:
                    yield QueryEndEvent()
                    return
                for event in self._convert_event(parser_event):
                    yield event
            yield QueryEndEvent()
        except Exception as e:
            if self._logger:
                self._logger.exception("astream_query error")
            yield ErrorEvent(error_message=str(e))

    def cancel(self) -> None:
        """Signal the running stream to stop after the current event."""
        self._cancelled = True

    def _extract_todos_from_args(self, args_text: str) -> Optional[list[dict]]:
        if not args_text:
            return None
        try:
            import json as json_mod

            parsed = json_mod.loads(args_text)
        except Exception as e:
            if self._logger:
                self._logger.debug("Failed to parse todos args: %s", e)
            return None
        if isinstance(parsed, dict):
            todos = parsed.get("todos")
            if isinstance(todos, list):
                return todos
        return None

    def _convert_event(self, parser_event: StreamEvent) -> list[AgentEvent]:
        events: list[AgentEvent] = []

        if isinstance(parser_event, ParserContentEvent):
            content = parser_event.content
            if isinstance(content, str) and content.startswith("{'id':"):
                import re

                if "'type': 'reasoning'" in content or '"type": "reasoning"' in content:
                    match = re.match(r"^\{[^}]*\}\s*(.*)$", content)
                    if match:
                        extracted = match.group(1)
                        if extracted and extracted.strip():
                            content = extracted
                        else:
                            return events
                    else:
                        return events
            if isinstance(content, str) and content.startswith("{'arguments':"):
                return events
            if not content:
                return events
            if hasattr(parser_event, "node") and parser_event.node == "thinking":
                events.append(ThinkingEvent(content=content))
            else:
                events.append(TextEvent(content=content))
            return events

        elif isinstance(parser_event, ParserToolCallStartEvent):
            import json as json_mod

            tool_args = ""
            if (
                parser_event.args
                and isinstance(parser_event.args, dict)
                and len(parser_event.args) > 0
            ):
                try:
                    tool_args = json_mod.dumps(parser_event.args)
                except Exception as e:
                    if self._logger:
                        self._logger.debug("Failed to serialize tool args: %s", e)
                    tool_args = ""
            events.append(
                ToolCallStartEvent(
                    tool_name=parser_event.name,
                    tool_args=tool_args,
                    tool_call_id=getattr(parser_event, "id", None),
                )
            )
            if parser_event.name == "write_todos" and isinstance(
                parser_event.args, dict
            ):
                todos = parser_event.args.get("todos")
                if isinstance(todos, list):
                    events.append(TodoUpdateEvent(todos=todos))
            return events

        elif isinstance(parser_event, ParserToolCallArgsEvent):
            existing_args = getattr(self, "_tool_args_buffer", {})
            tool_id = parser_event.id
            if tool_id:
                existing_args[tool_id] = (
                    existing_args.get(tool_id, "") + parser_event.args
                )
                self._tool_args_buffer = existing_args
            events.append(
                ToolCallArgsEvent(
                    tool_name=parser_event.name,
                    content=parser_event.args,
                    tool_call_id=getattr(parser_event, "id", None),
                )
            )
            if parser_event.name == "write_todos":
                buffer_key = tool_id or f"name:{parser_event.name}"
                combined_args = (
                    existing_args.get(tool_id, "")
                    if tool_id
                    else (parser_event.args or "")
                )
                todos = self._extract_todos_from_args(combined_args)
                if todos is not None:
                    import json as json_mod_inner

                    todos_key = json_mod_inner.dumps(todos, sort_keys=True)
                    if self._todo_args_cache.get(buffer_key) != todos_key:
                        self._todo_args_cache[buffer_key] = todos_key
                        events.append(TodoUpdateEvent(todos=todos))
            return events

        elif isinstance(parser_event, ToolCallEndEvent):
            content = ""
            if parser_event.status == "error":
                content = parser_event.error_message or "Tool call failed"
            else:
                if parser_event.result:
                    result_str = str(parser_event.result)
                    if len(result_str) > 8000:
                        content = (
                            result_str[:8000]
                            + f" [truncated: {len(result_str)} total chars]"
                        )
                    else:
                        content = result_str
                else:
                    content = ""
            events.append(
                ToolResultEvent(
                    tool_name=parser_event.name,
                    content=content,
                    tool_call_id=getattr(parser_event, "id", None),
                    data={
                        "status": parser_event.status,
                        "duration_ms": parser_event.duration_ms,
                    },
                )
            )
            if parser_event.name == "write_todos":
                tool_id = getattr(parser_event, "id", None)
                buffer_key = tool_id or f"name:{parser_event.name}"
                combined_args = (
                    self._tool_args_buffer.get(tool_id, "") if tool_id else ""
                )
                todos = self._extract_todos_from_args(combined_args)
                if todos is not None:
                    import json as json_mod_inner

                    todos_key = json_mod_inner.dumps(todos, sort_keys=True)
                    if self._todo_args_cache.get(buffer_key) != todos_key:
                        self._todo_args_cache[buffer_key] = todos_key
                        events.append(TodoUpdateEvent(todos=todos))
                if tool_id:
                    self._tool_args_buffer.pop(tool_id, None)
                self._todo_args_cache.pop(buffer_key, None)
            return events

        elif isinstance(parser_event, ParserErrorEvent):
            events.append(ErrorEvent(error_message=parser_event.error))
            return events

        elif isinstance(parser_event, CompleteEvent):
            return events

        return events

    def new_session(self, thread_id: Optional[str] = None) -> str:
        import uuid

        self.config.thread_id = thread_id or str(uuid.uuid4())
        return self.config.thread_id
