"""Event types for stream-msg-parser.

This module defines event classes for parsing LangGraph AIMessage
streaming output in "messages" stream mode.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Optional


@dataclass
class StreamEvent:
    """Base event class for stream parser."""

    timestamp: datetime = field(default_factory=datetime.now)


@dataclass
class ContentEvent(StreamEvent):
    """Content (text or reasoning) event from AIMessage.

    Attributes:
        content: The text or reasoning content
        node: Optional node name (e.g., "chatbot" or "thinking")
    """

    content: str = ""
    node: Optional[str] = None


@dataclass
class ToolCallStartEvent(StreamEvent):
    """Tool call started event.

    Attributes:
        id: Unique identifier for the tool call
        name: Name of the tool being called
        args: Arguments for the tool call (may be partial during streaming)
    """

    id: str = ""
    name: str = ""
    args: dict[str, Any] = field(default_factory=dict)


@dataclass
class ToolCallArgsEvent(StreamEvent):
    """Tool call arguments chunk event (for streaming args).

    Attributes:
        id: Unique identifier for the tool call
        name: Name of the tool being called
        args: Partial arguments being streamed
    """

    id: str = ""
    name: str = ""
    args: str = ""


@dataclass
class ToolCallEndEvent(StreamEvent):
    """Tool call completed event.

    Attributes:
        id: Unique identifier for the tool call
        name: Name of the tool that was called
        result: The result from the tool execution
        status: Status of the tool execution (success/error)
        error_message: Error message if status is error
        duration_ms: Duration of the tool execution in milliseconds
    """

    id: str = ""
    name: str = ""
    result: Any = None
    status: str = "success"
    error_message: Optional[str] = None
    duration_ms: Optional[float] = None


@dataclass
class ErrorEvent(StreamEvent):
    """Error event.

    Attributes:
        error: The error message
    """

    error: str = ""


@dataclass
class CompleteEvent(StreamEvent):
    """Stream completed event."""

    pass


__all__ = [
    "StreamEvent",
    "ContentEvent",
    "ToolCallStartEvent",
    "ToolCallEndEvent",
    "ErrorEvent",
    "CompleteEvent",
]
