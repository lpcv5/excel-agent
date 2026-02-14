"""Event types for streaming agent output.

This module defines the event system that decouples streaming output
from UI-specific rendering. All UI implementations (CLI, WebUI)
consume these events.
"""

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Any, Optional


class EventType(Enum):
    """Types of events emitted during agent execution."""

    # Lifecycle events
    SESSION_START = "session_start"
    SESSION_END = "session_end"
    QUERY_START = "query_start"
    QUERY_END = "query_end"
    ERROR = "error"

    # Streaming content events
    THINKING = "thinking"  # LLM reasoning
    TEXT = "text"  # Regular text output
    REFUSAL = "refusal"  # Content filter trigger

    # Tool events
    TOOL_CALL_START = "tool_call_start"
    TOOL_CALL_ARGS = "tool_call_args"
    TOOL_CALL_END = "tool_call_end"
    TOOL_RESULT = "tool_result"

    # Status events
    STATUS_UPDATE = "status_update"  # Excel status, workbook info


@dataclass
class AgentEvent:
    """Base event emitted during agent execution.

    Attributes:
        type: The type of this event
        timestamp: When this event occurred
        data: Additional event data
        content: Text content (for thinking, text events)
        tool_name: Tool name (for tool events)
        tool_args: Tool arguments (for tool events)
        error_message: Error message (for error events)
    """

    type: EventType
    timestamp: datetime = field(default_factory=datetime.now)
    data: dict[str, Any] = field(default_factory=dict)

    content: Optional[str] = None
    tool_name: Optional[str] = None
    tool_args: Optional[str] = None
    error_message: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        """Convert event to dictionary for serialization."""
        return {
            "type": self.type.value,
            "timestamp": self.timestamp.isoformat(),
            "content": self.content,
            "tool_name": self.tool_name,
            "tool_args": self.tool_args,
            "error_message": self.error_message,
            "data": self.data,
        }


@dataclass
class ThinkingEvent(AgentEvent):
    """LLM reasoning/thinking event."""

    type: EventType = EventType.THINKING
    content: str = ""


@dataclass
class TextEvent(AgentEvent):
    """Regular text output event."""

    type: EventType = EventType.TEXT
    content: str = ""


@dataclass
class RefusalEvent(AgentEvent):
    """Content filter refusal event."""

    type: EventType = EventType.REFUSAL
    content: str = ""


@dataclass
class ToolCallStartEvent(AgentEvent):
    """Tool call started event."""

    type: EventType = EventType.TOOL_CALL_START
    tool_name: str = ""


@dataclass
class ToolCallArgsEvent(AgentEvent):
    """Tool call arguments chunk event."""

    type: EventType = EventType.TOOL_CALL_ARGS
    tool_name: str = ""
    content: str = ""  # Chunk of arguments


@dataclass
class ToolResultEvent(AgentEvent):
    """Tool execution result event."""

    type: EventType = EventType.TOOL_RESULT
    tool_name: str = ""
    content: str = ""  # Result content
    data: dict[str, Any] = field(default_factory=dict)


@dataclass
class ErrorEvent(AgentEvent):
    """Error event."""

    type: EventType = EventType.ERROR
    error_message: str = ""


@dataclass
class QueryStartEvent(AgentEvent):
    """Query started event."""

    type: EventType = EventType.QUERY_START
    content: str = ""  # The query text


@dataclass
class QueryEndEvent(AgentEvent):
    """Query completed event."""

    type: EventType = EventType.QUERY_END
