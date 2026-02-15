"""Python API exposed to the frontend via pywebview.

This module provides the ExcelAgentAPI class that bridges
the frontend JavaScript calls to the Python agent core.
"""

import asyncio
import json
from typing import Optional

import webview

from excel_agent.config import AgentConfig
from excel_agent.core import AgentCore
from excel_agent.events import AgentEvent


class ExcelAgentAPI:
    """API class exposed to the frontend via pywebview.api.

    This class provides synchronous methods that can be called
    from JavaScript through window.pywebview.api.

    For streaming events, this class uses window.evaluate_js()
    to push events to the frontend by calling window.__onAgentEvent().

    Example frontend usage:
        // Send a query
        const result = await window.pywebview.api.sendQuery('Read data.xlsx');

        // Listen for events
        window.__onAgentEvent = (event) => console.log(event);
    """

    def __init__(self, config: Optional[AgentConfig] = None):
        """Initialize the API.

        Args:
            config: Agent configuration. Uses defaults if not provided.
        """
        self._config = config or AgentConfig()
        self._core = AgentCore(self._config)
        self._window: Optional[webview.Window] = None
        self._loop: Optional[asyncio.AbstractEventLoop] = None
        self._is_streaming = False

    def set_window(self, window: webview.Window) -> None:
        """Set the pywebview window reference.

        Args:
            window: The pywebview Window instance
        """
        self._window = window

    def set_event_loop(self, loop: asyncio.AbstractEventLoop) -> None:
        """Set the asyncio event loop.

        Args:
            loop: The asyncio event loop running in a background thread
        """
        self._loop = loop

    def sendQuery(self, query: str) -> dict:
        """Send a query to the agent and start streaming events.

        This method starts an async task that streams agent events
        and pushes them to the frontend via window.evaluate_js().

        Args:
            query: User's natural language query

        Returns:
            Dict with success status and optional error message
        """
        if self._is_streaming:
            return {"success": False, "error": "A query is already in progress"}

        if not self._loop:
            return {"success": False, "error": "Event loop not initialized"}

        self._is_streaming = True

        # Schedule the async streaming task on the event loop
        asyncio.run_coroutine_threadsafe(
            self._stream_query(query), self._loop
        )

        # We return immediately; events are pushed via callbacks
        return {"success": True}

    async def _stream_query(self, query: str) -> None:
        """Stream query events and push them to the frontend.

        Args:
            query: User's natural language query
        """
        try:
            async for event in self._core.astream_query(query):
                self._emit_event(event)
        except Exception as e:
            # Emit error event
            from excel_agent.events import ErrorEvent
            error_event = ErrorEvent(error_message=str(e))
            self._emit_event(error_event)
        finally:
            self._is_streaming = False

    def _emit_event(self, event: AgentEvent) -> None:
        """Emit an event to the frontend.

        Args:
            event: The event to emit
        """
        if not self._window:
            return

        event_data = event.to_dict()
        event_json = json.dumps(event_data, ensure_ascii=False)

        # Use evaluate_js to call the frontend callback
        # The frontend should define window.__onAgentEvent
        js_code = f"""
        if (window.__onAgentEvent) {{
            window.__onAgentEvent({event_json});
        }}
        """

        try:
            self._window.evaluate_js(js_code)
        except Exception:
            # Ignore JS evaluation errors (window might be closing)
            pass

    def newSession(self) -> dict:
        """Start a new conversation session.

        Returns:
            Dict with success status and new thread_id
        """
        thread_id = self._core.new_session()
        return {"success": True, "thread_id": thread_id}

    def getExcelStatus(self) -> dict:
        """Get current Excel application status.

        Returns:
            Dict with excel_running, workbook_count, open_workbooks
        """
        try:
            status = self._core.get_excel_status()
            return {"success": True, **status}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def getConfig(self) -> dict:
        """Get current agent configuration.

        Returns:
            Dict with configuration details
        """
        return {
            "success": True,
            "model": self._config.model,
            "working_dir": str(self._config.working_dir),
            "thread_id": self._config.thread_id,
            "streaming_enabled": self._config.streaming_enabled,
        }

    def isStreaming(self) -> dict:
        """Check if a query is currently streaming.

        Returns:
            Dict with is_streaming boolean
        """
        return {"success": True, "is_streaming": self._is_streaming}

    def stopStreaming(self) -> dict:
        """Request to stop the current streaming query.

        Returns:
            Dict with success status
        """
        # Note: This is a request; the streaming will stop
        # at the next opportunity. Full cancellation would
        # require more complex implementation.
        if self._is_streaming:
            self._is_streaming = False
            return {"success": True, "message": "Stop requested"}
        return {"success": True, "message": "No query in progress"}

    # =========================================================================
    # Excel Cleanup Methods
    # =========================================================================

    def getUnsavedWorkbooks(self) -> dict:
        """Get list of workbooks with unsaved changes.

        Returns:
            Dict with unsaved_workbooks list
        """
        try:
            from tools.excel_tool import get_excel_manager
            manager = get_excel_manager()
            unsaved = manager.get_unsaved_workbooks()
            return {"success": True, "unsaved_workbooks": unsaved}
        except Exception as e:
            return {"success": False, "error": str(e), "unsaved_workbooks": []}

    def saveAllWorkbooks(self) -> dict:
        """Save all open workbooks.

        Returns:
            Dict with saved count and any errors
        """
        try:
            from tools.excel_tool import get_excel_manager
            manager = get_excel_manager()
            result = manager.save_all_workbooks()
            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def closeAllWorkbooks(self, save: bool = True) -> dict:
        """Close all open workbooks.

        Args:
            save: Whether to save changes before closing

        Returns:
            Dict with closed count and any errors
        """
        try:
            from tools.excel_tool import get_excel_manager
            manager = get_excel_manager()
            result = manager.close_all_workbooks(save=save)
            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def quitExcel(self) -> dict:
        """Quit Excel application completely.

        Returns:
            Dict with success status
        """
        try:
            from tools.excel_tool import get_excel_manager
            manager = get_excel_manager()
            success = manager.quit_excel()
            return {"success": success}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def prepareClose(self) -> dict:
        """Prepare for application close.

        This checks for unsaved workbooks and returns info needed
        for the close confirmation dialog.

        Returns:
            Dict with has_unsaved, unsaved_workbooks, can_close
        """
        try:
            from tools.excel_tool import get_excel_manager
            manager = get_excel_manager()

            unsaved = manager.get_unsaved_workbooks()
            has_unsaved = len(unsaved) > 0

            return {
                "success": True,
                "has_unsaved": has_unsaved,
                "unsaved_workbooks": unsaved,
                "can_close": True,
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "has_unsaved": False,
                "unsaved_workbooks": [],
                "can_close": True,
            }

    def forceClose(self, save: bool = False) -> dict:
        """Force close Excel, optionally saving workbooks.

        Args:
            save: Whether to save workbooks before closing

        Returns:
            Dict with success status
        """
        try:
            from tools.excel_tool import get_excel_manager, cleanup_excel_resources

            # First try to save workbooks if requested
            if save:
                try:
                    manager = get_excel_manager()
                    manager.save_all_workbooks()
                except Exception:
                    pass  # Ignore save errors during close

            # Perform comprehensive cleanup of all COM resources
            result = cleanup_excel_resources(force=True)

            return {"success": result["success"], "errors": result.get("errors")}
        except Exception as e:
            return {"success": False, "error": str(e)}
