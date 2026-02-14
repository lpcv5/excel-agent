"""Excel Agent Core Package.

This package provides the UI-agnostic core logic for the Excel Agent.
It includes:
- AgentCore: Main agent interface
- AgentConfig: Configuration management
- Events: Streaming output events
- SessionManager: Conversation session management
"""

from excel_agent.config import AgentConfig
from excel_agent.core import AgentCore
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
    ToolResultEvent,
)
from excel_agent.session import Message, Session, SessionManager

__all__ = [
    # Core
    "AgentCore",
    "AgentConfig",
    # Events
    "AgentEvent",
    "EventType",
    "ThinkingEvent",
    "TextEvent",
    "RefusalEvent",
    "ToolCallStartEvent",
    "ToolCallArgsEvent",
    "ToolResultEvent",
    "ErrorEvent",
    "QueryStartEvent",
    "QueryEndEvent",
    # Session
    "Message",
    "Session",
    "SessionManager",
]
