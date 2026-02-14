"""Abstract base class for UI implementations.

This module defines the contract that all UI implementations
(CLI, WebUI) must follow.
"""

from abc import ABC, abstractmethod
from typing import Optional

from excel_agent.config import AgentConfig
from excel_agent.core import AgentCore


class BaseUI(ABC):
    """Abstract base class for UI implementations.

    All UI implementations (CLI, WebUI) must inherit from this class.
    It provides common functionality and defines the interface contract.
    """

    def __init__(self, config: Optional[AgentConfig] = None):
        """Initialize the UI.

        Args:
            config: Agent configuration. Creates default if not provided.
        """
        self.config = config or AgentConfig()
        self.core = AgentCore(self.config)

    @abstractmethod
    def run(self) -> None:
        """Run the UI main loop.

        This method should block until the UI is closed.
        """
        pass

    @abstractmethod
    def run_single_query(self, query: str) -> str:
        """Run a single query and return the result.

        Args:
            query: User's natural language query

        Returns:
            The agent's response
        """
        pass

    def on_event(self, event) -> None:
        """Handle an agent event.

        This method can be overridden to customize event handling.
        Default implementation does nothing.

        Args:
            event: AgentEvent instance
        """
        pass

    def on_error(self, error: str) -> None:
        """Handle an error.

        Args:
            error: Error message
        """
        pass

    def new_session(self) -> str:
        """Start a new conversation session.

        Returns:
            The new thread ID
        """
        return self.core.new_session()
