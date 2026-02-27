"""WebUI implementation package using pywebview.

This package provides the desktop WebUI for Excel Agent using
pywebview for native window rendering and communication between
Python and JavaScript.

Structure:
- server.py: pywebview window creation and lifecycle management
- api.py: Python API class exposed to frontend via pywebview.api
- frontend/: React + Vite frontend source code
- static/: Built frontend files (served in production mode)

Communication Flow:
1. Frontend calls window.pywebview.api.methodName()
2. Python API methods execute and return results
3. For streaming, Python uses window.evaluate_js() to push events
4. Frontend registers window.__onAgentEvent callback for events

Development Mode:
- If Vite dev server is running (localhost:5173), it's used for hot reload
- Otherwise, built static files are served from static/

Usage:
    from excel_agent.config import AgentConfig
    from ui.web.server import run_server

    config = AgentConfig()
    run_server(config)
"""

from ui.web.server import run_server, run_server_async

__all__ = ["run_server", "run_server_async"]
