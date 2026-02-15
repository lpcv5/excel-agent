"""WebUI server implementation using pywebview.

This module provides the desktop WebUI for Excel Agent using
pywebview for native window rendering and communication between
Python and JavaScript.

Usage:
    from excel_agent.config import AgentConfig
    from ui.web.server import run_server

    config = AgentConfig(model="openai:gpt-5-mini")
    run_server(config)

Development mode:
    - If Vite dev server is running (localhost:5173), use it for hot reload
    - Otherwise, serve built static files from ui/web/static/
"""

import asyncio
import threading
from pathlib import Path

import webview

from excel_agent.config import AgentConfig
from ui.web.api import ExcelAgentAPI


# Constants
VITE_DEV_PORT = 5173
STATIC_DIR = Path(__file__).parent / "static"
WINDOW_TITLE = "Excel Agent"
WINDOW_SIZE = (1200, 800)


def is_dev_server_running() -> bool:
    """Check if Vite dev server is running.

    Returns:
        True if localhost:5173 is accessible
    """
    import socket

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(1)
    try:
        result = sock.connect_ex(("localhost", VITE_DEV_PORT))
        return result == 0
    finally:
        sock.close()


def run_server(config: AgentConfig) -> None:
    """Run the WebUI server.

    This function blocks until the window is closed.

    Args:
        config: Agent configuration
    """
    # Create the API instance
    api = ExcelAgentAPI(config)

    # Determine URL to serve
    if is_dev_server_running():
        # Use Vite dev server for development
        url = f"http://localhost:{VITE_DEV_PORT}"
        print(f"[WebUI] Using Vite dev server: {url}")
    else:
        # Serve built static files
        if not STATIC_DIR.exists():
            raise RuntimeError(
                f"Static directory not found: {STATIC_DIR}\n"
                "Please build the frontend first:\n"
                "  cd ui/web/frontend && npm run build"
            )
        url = str(STATIC_DIR / "index.html")
        print(f"[WebUI] Serving static files from: {STATIC_DIR}")

    # Start asyncio event loop in a background thread
    loop = asyncio.new_event_loop()
    api.set_event_loop(loop)

    def run_event_loop():
        asyncio.set_event_loop(loop)
        loop.run_forever()

    loop_thread = threading.Thread(target=run_event_loop, daemon=True)
    loop_thread.start()

    # Create the window
    window = webview.create_window(
        title=WINDOW_TITLE,
        url=url,
        js_api=api,
        width=WINDOW_SIZE[0],
        height=WINDOW_SIZE[1],
        min_size=(800, 600),
    )

    # Set window reference in API
    api.set_window(window)

    # Start the webview (this blocks)
    try:
        webview.start(debug=False)
    finally:
        # Clean up Excel COM resources before exit
        # This is the most reliable place for cleanup as it runs on the main thread
        try:
            from tools.excel_tool import cleanup_excel_resources
            result = cleanup_excel_resources(force=False)
            if not result.get("success"):
                from libs.excel_com.manager import force_cleanup_excel_processes
                force_cleanup_excel_processes()
        except Exception:
            pass  # Ignore cleanup errors

        # Clean up the event loop
        loop.call_soon_threadsafe(loop.stop)


def run_server_async(config: AgentConfig) -> threading.Thread:
    """Run the WebUI server in a background thread.

    This is useful for testing or when you need non-blocking behavior.

    Args:
        config: Agent configuration

    Returns:
        The thread running the server
    """
    thread = threading.Thread(target=run_server, args=(config,), daemon=True)
    thread.start()
    return thread
