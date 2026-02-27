"""Event types for streaming agent output."""

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Any, Optional

from libs.stream_msg_parser.events import (
    ContentEvent,
    ToolCallEndEvent,
    CompleteEvent,
)

__all__ = [
    "ContentEvent",
    "ToolCallStartEvent",
    "ToolCallEndEvent",
    "CompleteEvent",
    "EventType",
    "AgentEvent",
    "ThinkingEvent",
    "TextEvent",
    "RefusalEvent",
    "ToolCallArgsEvent",
    "ToolResultEvent",
    "ErrorEvent",
    "QueryStartEvent",
    "QueryEndEvent",
    "TodoUpdateEvent",
]


class EventType(Enum):
    QUERY_START = "query_start"
    QUERY_END = "query_end"
    ERROR = "error"
    THINKING = "thinking"
    TEXT = "text"
    REFUSAL = "refusal"
    TOOL_CALL_START = "tool_call_start"
    TOOL_CALL_ARGS = "tool_call_args"
    TOOL_CALL_END = "tool_call_end"
    TOOL_RESULT = "tool_result"
    TODO_UPDATE = "todo_update"


@dataclass
class AgentEvent:
    type: EventType
    timestamp: datetime = field(default_factory=datetime.now)
    data: dict[str, Any] = field(default_factory=dict)
    content: Optional[str] = None
    tool_name: Optional[str] = None
    tool_args: Optional[str] = None
    tool_call_id: Optional[str] = None
    error_message: Optional[str] = None
    todos: Optional[list[dict]] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": self.type.value,
            "timestamp": self.timestamp.isoformat(),
            "content": self.content,
            "tool_name": self.tool_name,
            "tool_args": self.tool_args,
            "tool_call_id": self.tool_call_id,
            "error_message": self.error_message,
            "data": self.data,
            "todos": self.todos,
        }


@dataclass
class ThinkingEvent(AgentEvent):
    type: EventType = EventType.THINKING
    content: str = ""  # type: ignore[assignment]


@dataclass
class TextEvent(AgentEvent):
    type: EventType = EventType.TEXT
    content: str = ""  # type: ignore[assignment]


@dataclass
class RefusalEvent(AgentEvent):
    type: EventType = EventType.REFUSAL
    content: str = ""  # type: ignore[assignment]


@dataclass
class ToolCallStartEvent(AgentEvent):
    type: EventType = EventType.TOOL_CALL_START
    tool_name: str = ""  # type: ignore[assignment]


@dataclass
class ToolCallArgsEvent(AgentEvent):
    type: EventType = EventType.TOOL_CALL_ARGS
    tool_name: str = ""  # type: ignore[assignment]
    content: str = ""  # type: ignore[assignment]


@dataclass
class ToolResultEvent(AgentEvent):
    type: EventType = EventType.TOOL_RESULT
    tool_name: str = ""  # type: ignore[assignment]
    content: str = ""  # type: ignore[assignment]
    data: dict[str, Any] = field(default_factory=dict)


@dataclass
class ErrorEvent(AgentEvent):
    type: EventType = EventType.ERROR
    error_message: str = ""  # type: ignore[assignment]


@dataclass
class QueryStartEvent(AgentEvent):
    type: EventType = EventType.QUERY_START
    content: str = ""  # type: ignore[assignment]


@dataclass
class QueryEndEvent(AgentEvent):
    type: EventType = EventType.QUERY_END


@dataclass
class TodoUpdateEvent(AgentEvent):
    type: EventType = EventType.TODO_UPDATE
    todos: list[dict] = field(default_factory=list)  # type: ignore[assignment]
