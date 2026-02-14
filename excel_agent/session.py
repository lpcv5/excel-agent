"""Session management for Excel Agent.

This module provides session management for tracking conversations.
Note: The actual conversation state is managed by LangGraph's checkpointer.
This is a lightweight layer for UI purposes.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional
import uuid


@dataclass
class Message:
    """A message in the conversation.

    Attributes:
        role: "user" or "assistant"
        content: The message text
        timestamp: When the message was created
    """

    role: str
    content: str
    timestamp: datetime = field(default_factory=datetime.now)


@dataclass
class Session:
    """A conversation session.

    Attributes:
        id: Unique session identifier
        created_at: When the session was created
        messages: List of messages in the session
    """

    id: str
    created_at: datetime = field(default_factory=datetime.now)
    messages: list[Message] = field(default_factory=list)

    def add_message(self, role: str, content: str) -> Message:
        """Add a message to the session.

        Args:
            role: "user" or "assistant"
            content: The message text

        Returns:
            The created Message instance
        """
        msg = Message(role=role, content=content)
        self.messages.append(msg)
        return msg

    def get_last_message(self) -> Optional[Message]:
        """Get the last message in the session.

        Returns:
            The last Message or None if empty
        """
        return self.messages[-1] if self.messages else None

    def clear(self) -> None:
        """Clear all messages in the session."""
        self.messages.clear()


class SessionManager:
    """Manages conversation sessions.

    This is a lightweight session manager for UI purposes.
    The actual conversation state is managed by LangGraph's checkpointer.
    """

    def __init__(self, thread_id: Optional[str] = None):
        """Initialize the session manager.

        Args:
            thread_id: Optional thread ID. Generated if not provided.
        """
        self.thread_id = thread_id or str(uuid.uuid4())
        self._session = Session(id=self.thread_id)

    @property
    def session(self) -> Session:
        """Get the current session."""
        return self._session

    def new_session(self, thread_id: Optional[str] = None) -> Session:
        """Create a new session.

        Args:
            thread_id: Optional new thread ID. Generated if not provided.

        Returns:
            The new Session instance
        """
        self.thread_id = thread_id or str(uuid.uuid4())
        self._session = Session(id=self.thread_id)
        return self._session

    def add_user_message(self, content: str) -> Message:
        """Add a user message to the session.

        Args:
            content: The message text

        Returns:
            The created Message instance
        """
        return self._session.add_message("user", content)

    def add_assistant_message(self, content: str) -> Message:
        """Add an assistant message to the session.

        Args:
            content: The message text

        Returns:
            The created Message instance
        """
        return self._session.add_message("assistant", content)
