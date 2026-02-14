"""WebUI server implementation (placeholder).

This module will contain the FastAPI-based web server with
WebSocket support for real-time streaming.

TODO: Implement full WebUI server with:
- FastAPI application
- WebSocket endpoints for streaming
- REST API for session management
- Static file serving for frontend
"""

from excel_agent.config import AgentConfig


def run_server(config: AgentConfig) -> None:
    """Run the WebUI server.

    Args:
        config: Agent configuration

    Raises:
        NotImplementedError: WebUI is not yet implemented
    """
    raise NotImplementedError(
        "WebUI mode is not yet implemented. "
        "Please use CLI mode (default) for now."
    )
