"""Message parser for LangGraph streaming with stream_mode=messages.

This module provides a parser for handling LangGraph AIMessage streaming
output. Unlike langgraph-stream-parser which uses stream_mode="updates",
this parser works with stream_mode="messages" for token-level streaming.
"""

import time
from typing import Any, AsyncGenerator, Optional

from .events import (
    CompleteEvent,
    ContentEvent,
    ErrorEvent,
    StreamEvent,
    ToolCallArgsEvent,
    ToolCallEndEvent,
    ToolCallStartEvent,
)


class MessageParser:
    """Parser for LangGraph message stream mode.

    This parser handles AIMessage objects from langgraph streaming
    with stream_mode="messages", providing token-level content events
    and tool call lifecycle tracking.

    Usage:
        parser = MessageParser(track_tool_lifecycle=True)
        async for event in parser.aparse(stream):
            handle_event(event)
    """

    def __init__(self, track_tool_lifecycle: bool = True):
        """Initialize the parser.

        Args:
            track_tool_lifecycle: Whether to track tool call start/end events
        """
        self._track_tool_lifecycle = track_tool_lifecycle
        self._tool_start_times: dict[str, float] = {}
        self._seen_tool_calls: set[str] = set()  # Track tool calls we've already emitted
        self._current_tool_call_id: str = ""  # Track current tool call being streamed

    async def aparse(
        self, stream: AsyncGenerator[tuple[Any, dict[str, Any]], None]
    ) -> AsyncGenerator[StreamEvent, None]:
        """Parse an async stream of LangGraph messages.

        Args:
            stream: Async generator yielding (message, metadata) tuples
                   where metadata contains 'langgraph_node' key

        Yields:
            StreamEvent instances representing the parsed stream
        """
        try:
            async for message, metadata in stream:
                try:
                    # Extract node name from metadata
                    node_name = metadata.get("langgraph_node", "chatbot")

                    # Handle different message types
                    events = self._parse_message(node_name, message)
                    for event in events:
                        if event:
                            yield event
                except Exception as e:
                    yield ErrorEvent(error=f"Error parsing message: {e}")

            # Signal completion
            yield CompleteEvent()

        except Exception as e:
            yield ErrorEvent(error=f"Stream error: {e}")

    def _parse_message(
        self, node_name: str, message: Any
    ) -> list[Optional[StreamEvent]]:
        """Parse a single message from the stream.

        Args:
            node_name: The node that produced this message
            message: The message object (typically AIMessage)

        Returns:
            List of parsed events
        """
        events: list[Optional[StreamEvent]] = []

        # Import here to avoid circular imports
        from langchain_core.messages import AIMessage
        from langchain_core.messages import ToolMessage

        if isinstance(message, AIMessage):
            # Handle AIMessage - extract content and tool calls
            events.extend(self._parse_ai_message(node_name, message))

        elif isinstance(message, ToolMessage):
            # Handle ToolMessage - tool result
            if self._track_tool_lifecycle:
                events.append(self._parse_tool_result(message))

        return events

    def _parse_ai_message(
        self, node_name: str, message: Any
    ) -> list[Optional[StreamEvent]]:
        """Parse an AIMessage for content and tool calls.

        Args:
            node_name: The node that produced this message
            message: The AIMessage object

        Returns:
            List of parsed events
        """
        events: list[Optional[StreamEvent]] = []

        # Handle content (text) - preserve whitespace including newlines
        content = message.content
        if content:
            if isinstance(content, str):
                # Don't skip whitespace - emit all content including newlines
                if content:
                    events.append(ContentEvent(content=content, node=node_name))
            elif isinstance(content, list):
                # Handle list content (e.g., multiple content blocks)
                for item in content:
                    if isinstance(item, dict):
                        item_type = item.get("type", "")
                        if item_type == "text":
                            text = item.get("text", "")
                            if text:  # Emit all text including whitespace
                                events.append(ContentEvent(content=text, node=node_name))
                        elif item_type == "reasoning":
                            reasoning = item.get("reasoning", "")
                            if reasoning:
                                events.append(ContentEvent(content=reasoning, node="thinking"))

        # Handle reasoning content (newer langchain versions)
        if hasattr(message, "reasoning") and message.reasoning:
            reasoning = message.reasoning
            if reasoning:
                events.append(ContentEvent(content=reasoning, node="thinking"))

        # Handle tool call chunks (streaming arguments)
        if self._track_tool_lifecycle and hasattr(message, "tool_call_chunks") and message.tool_call_chunks:
            for chunk in message.tool_call_chunks:
                events.extend(self._parse_tool_call_chunk(chunk))

        # Handle tool calls (complete tool calls)
        if self._track_tool_lifecycle and hasattr(message, "tool_calls") and message.tool_calls:
            for tool_call in message.tool_calls:
                # Only emit if we haven't seen this tool call ID before
                tool_id = tool_call.get("id", "")
                if tool_id and tool_id not in self._seen_tool_calls:
                    self._seen_tool_calls.add(tool_id)
                    events.append(self._parse_tool_call(tool_call))

        # Handle invalid tool calls (errors)
        if self._track_tool_lifecycle and hasattr(message, "invalid_tool_calls") and message.invalid_tool_calls:
            for invalid_call in message.invalid_tool_calls:
                error = invalid_call.get("error")
                if error:
                    events.append(ErrorEvent(error=str(error)))

        return events

    def _parse_tool_call_chunk(self, chunk: dict[str, Any]) -> list[Optional[StreamEvent]]:
        """Parse a tool call chunk for streaming args.

        Args:
            chunk: The tool call chunk dict

        Returns:
            List of parsed events
        """
        events: list[Optional[StreamEvent]] = []

        tool_id = chunk.get("id", "") or self._current_tool_call_id
        tool_name = chunk.get("name")
        args_chunk = chunk.get("args", "")

        # If we have a tool name, emit ToolCallStartEvent
        if tool_name and tool_id:
            if tool_id not in self._seen_tool_calls:
                self._seen_tool_calls.add(tool_id)
                self._current_tool_call_id = tool_id
                self._tool_start_times[tool_id] = time.perf_counter()
                events.append(ToolCallStartEvent(
                    id=tool_id,
                    name=tool_name,
                    args={},  # Args will come in chunks
                ))

        # If we have args chunk, emit ToolCallArgsEvent
        if args_chunk and (tool_id or self._current_tool_call_id):
            # Use current tool call id if chunk doesn't have one
            effective_tool_id = tool_id or self._current_tool_call_id
            events.append(ToolCallArgsEvent(
                id=effective_tool_id,
                name=tool_name or "",  # Use empty string if name not in this chunk
                args=args_chunk,
            ))

        return events

    def _parse_tool_call(self, tool_call: dict[str, Any]) -> ToolCallStartEvent:
        """Parse a tool call from AIMessage.

        Args:
            tool_call: The tool call dict from AIMessage

        Returns:
            ToolCallStartEvent
        """
        tool_id = tool_call.get("id", "")
        tool_name = tool_call.get("name", "")
        tool_args = tool_call.get("args", {})

        # Track start time for duration calculation
        if tool_id:
            self._tool_start_times[tool_id] = time.perf_counter()

        return ToolCallStartEvent(
            id=tool_id,
            name=tool_name,
            args=tool_args,
        )

    def _parse_tool_result(self, message: Any) -> ToolCallEndEvent:
        """Parse a ToolMessage as a tool result.

        Args:
            message: The ToolMessage object

        Returns:
            ToolCallEndEvent
        """
        tool_call_id = message.tool_call_id
        tool_name = message.name if hasattr(message, "name") else ""
        content = message.content

        # Calculate duration
        duration_ms: Optional[float] = None
        if tool_call_id and tool_call_id in self._tool_start_times:
            start_time = self._tool_start_times.pop(tool_call_id)
            duration_ms = (time.perf_counter() - start_time) * 1000

        return ToolCallEndEvent(
            id=tool_call_id,
            name=tool_name,
            result=content,
            status="success",
            duration_ms=duration_ms,
        )


__all__ = ["MessageParser"]
