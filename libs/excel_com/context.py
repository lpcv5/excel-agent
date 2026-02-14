"""
Excel Context Manager - Preserves user's Excel state during operations.

This module provides context managers that save and restore user's Excel
view state (active sheet, selection, scroll position, etc.) to ensure
MCP operations don't disrupt the user's workflow.
"""

from contextlib import contextmanager
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from win32com.client import Application  # type: ignore


class UserStatePreserver:
    """Preserves and restores user's Excel state during operations.

    This class saves the user's current Excel view state before an operation
    and restores it after the operation completes, ensuring minimal disruption
    to the user's workflow.

    State that is preserved:
    - Active workbook and worksheet
    - Current selection (selected cells/range)
    - Scroll position (active window's scroll area)
    - Active cell
    """

    def __init__(self, app: "Application"):
        """Initialize the state preserver.

        Args:
            app: Excel Application COM object
        """
        self._app = app
        self._saved_active_workbook: Optional[object] = None
        self._saved_active_worksheet: Optional[object] = None
        self._saved_selection: Optional[object] = None
        self._saved_active_cell: Optional[object] = None
        self._saved_window_scroll_row: Optional[int] = None
        self._saved_window_scroll_column: Optional[int] = None

    def save_state(self) -> None:
        """Save the current Excel view state."""
        try:
            # Save active workbook and worksheet
            self._saved_active_workbook = self._app.ActiveWorkbook
            if self._saved_active_workbook:
                self._saved_active_worksheet = self._app.ActiveSheet

            # Save current selection
            try:
                self._saved_selection = self._app.Selection
            except Exception:
                self._saved_selection = None

            # Save active cell
            try:
                self._saved_active_cell = self._app.ActiveCell
            except Exception:
                self._saved_active_cell = None

            # Save scroll position of active window
            try:
                if self._app.ActiveWindow:
                    self._saved_window_scroll_row = self._app.ActiveWindow.ScrollRow
                    self._saved_window_scroll_column = self._app.ActiveWindow.ScrollColumn
            except Exception:
                pass
        except Exception:
            # If saving state fails, we'll just skip restoration
            pass

    def restore_state(self) -> None:
        """Restore the previously saved Excel view state."""
        try:
            # Try to restore the active workbook and worksheet
            if self._saved_active_workbook:
                try:
                    # Check if the workbook is still open
                    _ = self._saved_active_workbook.Name  # Test if still valid
                    self._saved_active_workbook.Activate()

                    if self._saved_active_worksheet:
                        try:
                            # Check if worksheet is still valid
                            _ = self._saved_active_worksheet.Name
                            self._saved_active_worksheet.Activate()
                        except Exception:
                            pass  # Worksheet may have been deleted
                except Exception:
                    pass  # Workbook may have been closed

            # Restore scroll position
            try:
                if self._app.ActiveWindow and self._saved_window_scroll_row is not None:
                    self._app.ActiveWindow.ScrollRow = self._saved_window_scroll_row
                if self._app.ActiveWindow and self._saved_window_scroll_column is not None:
                    self._app.ActiveWindow.ScrollColumn = self._saved_window_scroll_column
            except Exception:
                pass

            # Note: We don't restore selection/active cell as it may be
            # disruptive if the user has started working in a different area
            # The scroll position restoration is usually sufficient
        except Exception:
            # If restoration fails, the Excel state will be as-is
            pass


@contextmanager
def preserve_user_state(app: "Application"):
    """Context manager that preserves user's Excel state during operations.

    Usage:
        with preserve_user_state(excel_app):
            # Perform Excel operations
            # User's view state will be restored when exiting the context

    Args:
        app: Excel Application COM object

    Yields:
        None
    """
    preserver = UserStatePreserver(app)
    preserver.save_state()

    try:
        yield
    finally:
        preserver.restore_state()
