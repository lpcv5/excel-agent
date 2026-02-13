"""
Excel Application Manager - Core Excel lifecycle and workbook management.

This module provides the ExcelAppManager class which handles creating,
managing, and cleaning up the Excel COM application instance.
"""

from pathlib import Path
from typing import TYPE_CHECKING, Optional

import win32com.client

if TYPE_CHECKING:
    from win32com.client import Application  # type: ignore
    from win32com.client import Workbook  # type: ignore


class ExcelAppManager:
    """Manages the Excel application instance and workbook operations.

    This class provides a high-level interface for working with Excel via COM.
    It handles the application lifecycle, workbook management, and provides
    helper methods for common operations.

    The manager can work in two modes:
    1. Attach to user's existing Excel instance (seamless mode)
    2. Create independent Excel instance (background mode)

    Workbooks opened by the user are tracked and will NOT be closed on shutdown.
    Only workbooks opened by this manager will be cleaned up.
    """

    def __init__(self, visible: bool = False, display_alerts: bool = False, attach_to_existing: bool = True):
        """Initialize the Excel application manager.

        Args:
            visible: Whether to show the Excel window (default: False for background operation)
            display_alerts: Whether to display Excel alert dialogs (default: False)
            attach_to_existing: Whether to attach to user's existing Excel instance (default: True)
        """
        self._app: Optional["Application"] = None
        self._visible = visible
        self._display_alerts = display_alerts
        self._attach_to_existing = attach_to_existing
        self._workbooks: dict[str, object] = {}  # Track open workbooks by path
        self._workbook_owned: dict[str, bool] = {}  # Track whether we own the workbook (can close it)

    @property
    def app(self) -> "Application":
        """Get the Excel application COM object.

        Returns:
            The Excel Application COM object

        Raises:
            RuntimeError: If Excel application is not initialized
        """
        if self._app is None:
            raise RuntimeError("Excel application is not initialized. Call start() first.")
        return self._app

    # =============================================================================
    # Lifecycle Management
    # =============================================================================

    def start(self) -> None:
        """Start the Excel application.

        Creates a new Excel COM application instance or attaches to an existing one.
        - If attach_to_existing=True and user has Excel open, attaches to that instance
        - Otherwise, creates a new independent Excel instance

        Should be called before any other operations.
        """
        if self._app is not None:
            return  # Already started

        if self._attach_to_existing:
            try:
                # Try to attach to existing Excel instance
                # This will connect to user's Excel if already open
                self._app = win32com.client.Dispatch("Excel.Application")
                # Check if we successfully connected to a running instance
                # by trying to access the Workbooks collection
                _ = self._app.Workbooks.Count
                self._attached_to_existing = True
            except Exception:
                # No existing Excel instance, create a new one
                self._app = win32com.client.DispatchEx("Excel.Application")
                self._attached_to_existing = False
        else:
            # Create a separate Excel instance
            self._app = win32com.client.DispatchEx("Excel.Application")
            self._attached_to_existing = False

        self._app.Visible = self._visible
        self._app.DisplayAlerts = self._display_alerts

        # Track all workbooks that are already open in the Excel instance
        # We should NOT close these on shutdown
        if self._app:
            try:
                for i in range(1, self._app.Workbooks.Count + 1):
                    try:
                        wb = self._app.Workbooks(i)
                        if wb.Path and wb.Name:
                            # Workbook has been saved, get full path
                            full_path = wb.FullName
                            self._workbooks[full_path] = wb
                            self._workbook_owned[full_path] = False  # We don't own it
                    except Exception:
                        # Unsaved workbook or error accessing it
                        pass
            except Exception:
                pass

    def stop(self) -> None:
        """Stop the Excel application and clean up resources.

        Only closes workbooks that were opened by this manager.
        Workbooks opened by the user are NOT closed.

        If we attached to an existing Excel instance, we only release our reference.
        If we created our own Excel instance, we quit it after closing our workbooks.

        Should be called when shutting down the server.
        """
        if self._app is None:
            return

        # Get a reference to the app before cleanup
        app_ref = self._app
        attached_to_existing = getattr(self, '_attached_to_existing', False)

        try:
            # Only close workbooks we own (opened by this manager)
            for workbook_path in list(self._workbooks.keys()):
                if self._workbook_owned.get(workbook_path, True):
                    try:
                        self.close_workbook(workbook_path, save=False)
                    except Exception:
                        pass  # Ignore errors during workbook cleanup
                else:
                    # Just remove from tracking, don't close user's workbook
                    self._workbooks.pop(workbook_path, None)
                    self._workbook_owned.pop(workbook_path, None)

            # Clear our tracking dicts
            self._workbooks.clear()
            self._workbook_owned.clear()
        except Exception:
            pass  # Ignore errors during cleanup
        finally:
            # Release our reference to the Excel app
            if not attached_to_existing:
                # Only quit if we created our own instance
                try:
                    app_ref.Quit()
                except Exception:
                    pass  # May have already exited

            self._app = None
            app_ref = None

            # Force garbage collection to release COM references
            import gc
            gc.collect()

    # =============================================================================
    # Workbook Operations
    # =============================================================================

    def open_workbook(self, filepath: str, read_only: bool = False) -> object:
        """Open an Excel workbook.

        If the workbook is already open in the Excel instance (e.g., by the user),
        it will be reused and NOT marked as owned by this manager.

        Args:
            filepath: Path to the Excel file
            read_only: Whether to open in read-only mode

        Returns:
            The Workbook COM object

        Raises:
            RuntimeError: If Excel application is not initialized
            FileNotFoundError: If the file doesn't exist
            Exception: For other COM-related errors
        """
        if self._app is None:
            raise RuntimeError("Excel application is not initialized. Call start() first.")

        path = Path(filepath).resolve()

        if not path.exists():
            raise FileNotFoundError(f"Excel file not found: {filepath}")

        str_path = str(path)

        # Return cached workbook if already tracked
        if str_path in self._workbooks:
            return self._workbooks[str_path]

        # Check if workbook is already open in the Excel instance
        workbook = self._find_open_workbook(str_path)
        if workbook:
            # Workbook is already open (likely by user), track it but don't own it
            self._workbooks[str_path] = workbook
            self._workbook_owned[str_path] = False
            return workbook

        try:
            # Open the workbook - we own this one
            workbook = self._app.Workbooks.Open(str_path, ReadOnly=read_only)
            self._workbooks[str_path] = workbook
            self._workbook_owned[str_path] = True  # We own it, can close it
            return workbook
        except Exception as e:
            raise Exception(f"Failed to open workbook '{filepath}': {e}") from e

    def close_workbook(self, filepath: str, save: bool = True) -> None:
        """Close an open workbook.

        Note: If the workbook was opened by the user (not by this manager),
        it will NOT be closed. It will only be removed from tracking.

        Args:
            filepath: Path to the workbook (must match the path used to open it)
            save: Whether to save changes before closing (only applies if we own the workbook)

        Raises:
            RuntimeError: If Excel application is not initialized
            ValueError: If the workbook is not open
            Exception: For other COM-related errors
        """
        if self._app is None:
            raise RuntimeError("Excel application is not initialized. Call start() first.")

        # Find workbook by path (case-insensitive matching)
        str_path = self._find_workbook_path(filepath)

        if str_path is None:
            raise ValueError(f"Workbook is not open: {filepath}")

        try:
            workbook = self._workbooks.pop(str_path)
            we_own_it = self._workbook_owned.pop(str_path, True)

            if we_own_it:
                # Only close workbooks we opened
                try:
                    workbook.Close(SaveChanges=save)
                except Exception:
                    # Excel may have already exited
                    pass
            else:
                # Workbook was opened by user, just save if requested and remove from tracking
                if save:
                    try:
                        workbook.Save()
                    except Exception:
                        pass
        except Exception as e:
            raise Exception(f"Failed to close workbook '{filepath}': {e}") from e

    def get_workbook(self, filepath: str) -> object:
        """Get an open workbook by path.

        Args:
            filepath: Path to the workbook

        Returns:
            The Workbook COM object

        Raises:
            RuntimeError: If Excel application is not initialized
            ValueError: If the workbook is not open
        """
        if self._app is None:
            raise RuntimeError("Excel application is not initialized. Call start() first.")

        str_path = self._find_workbook_path(filepath)

        if str_path is None:
            raise ValueError(f"Workbook is not open: {filepath}")

        return self._workbooks[str_path]

    def _find_workbook_path(self, filepath: str) -> str | None:
        """Find workbook path by case-insensitive matching.

        Args:
            filepath: Path to search for

        Returns:
            The actual path key if found, None otherwise
        """
        filepath_lower = filepath.lower()
        for open_path in self._workbooks.keys():
            if open_path.lower() == filepath_lower:
                return open_path
        return None

    def _find_open_workbook(self, filepath: str) -> object | None:
        """Check if a workbook is already open in the Excel instance.

        Args:
            filepath: Full path to the workbook

        Returns:
            The Workbook COM object if found, None otherwise
        """
        if self._app is None:
            return None

        try:
            filepath_lower = filepath.lower()
            for i in range(1, self._app.Workbooks.Count + 1):
                try:
                    wb = self._app.Workbooks(i)
                    if wb.FullName:
                        if wb.FullName.lower() == filepath_lower:
                            return wb
                except Exception:
                    continue
        except Exception:
            pass

        return None

    def get_active_workbook(self) -> Optional[object]:
        """Get the currently active workbook.

        Returns:
            The active Workbook COM object, or None if no workbook is open
        """
        if self._app is None:
            return None

        try:
            return self._app.ActiveWorkbook
        except Exception:
            return None

    def save_workbook(self, workbook: object, filepath: Optional[str] = None) -> None:
        """Save a workbook.

        Args:
            workbook: The Workbook COM object
            filepath: Optional path to save as (if different from original)

        Raises:
            Exception: If save operation fails
        """
        try:
            if filepath:
                workbook.SaveAs(filepath)
            else:
                workbook.Save()
        except Exception as e:
            raise Exception(f"Failed to save workbook: {e}") from e

    # =============================================================================
    # Worksheet Operations
    # =============================================================================

    def list_worksheets(self, workbook: object) -> list[str]:
        """List all worksheets in a workbook.

        Args:
            workbook: The Workbook COM object

        Returns:
            List of worksheet names
        """
        worksheets = []
        count = workbook.Worksheets.Count
        for i in range(1, count + 1):
            worksheets.append(workbook.Worksheets(i).Name)
        return worksheets

    def get_worksheet(self, workbook: object, name: str) -> object:
        """Get a worksheet by name.

        Args:
            workbook: The Workbook COM object
            name: Worksheet name

        Returns:
            The Worksheet COM object

        Raises:
            ValueError: If the worksheet doesn't exist
        """
        try:
            return workbook.Worksheets(name)
        except Exception as e:
            raise ValueError(f"Worksheet '{name}' not found in workbook") from e

    def get_active_worksheet(self, workbook: object) -> object:
        """Get the active worksheet in a workbook.

        Args:
            workbook: The Workbook COM object

        Returns:
            The active Worksheet COM object
        """
        return workbook.ActiveSheet

    # =============================================================================
    # Range Operations
    # =============================================================================

    def read_range(self, worksheet: object, range_address: str) -> list[list]:
        """Read data from a cell range.

        Args:
            worksheet: The Worksheet COM object
            range_address: Excel range address (e.g., "A1", "A1:Z100")

        Returns:
            2D list of cell values
        """
        range_obj = worksheet.Range(range_address)
        values = range_obj.Value

        # Convert COM variant to Python list
        if values is None:
            return []

        # Single cell
        if not isinstance(values, (list, tuple)):
            return [[self._normalize_value(values)]]

        # 2D array from Excel comes as ((row1_col1, row1_col2, ...), (row2_col1, ...))
        return [
            [self._normalize_value(cell) for cell in row]
            if isinstance(row, (list, tuple))
            else [self._normalize_value(row)]
            for row in values
        ]

    @staticmethod
    def _normalize_value(value):
        """Normalize a value from Excel COM.

        Converts float to int if the float represents a whole number.
        This is because Excel stores all numbers as floats.

        Args:
            value: Value from Excel COM

        Returns:
            Normalized value
        """
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value

    def write_range(self, worksheet: object, range_address: str, data: list[list]) -> None:
        """Write data to a cell range.

        Args:
            worksheet: The Worksheet COM object
            range_address: Excel range address (e.g., "A1", "A1:Z100")
            data: 2D list of values to write

        Raises:
            ValueError: If data is empty or not 2D
        """
        if not data or not data[0]:
            raise ValueError("Cannot write empty data")

        # Get the starting cell and its address
        start_cell = worksheet.Range(range_address)
        addr = start_cell.GetAddress()

        # Parse address like "$A$1" to extract column letter and row number
        import re
        match = re.match(r'\$([A-Z]+)\$(\d+)', addr)
        if match:
            # Calculate data dimensions
            rows = len(data)
            cols = len(data[0])

            # Parse starting position
            start_col_letter = match.group(1)
            start_row = int(match.group(2))

            # Convert column letter to number
            start_col = sum((ord(c) - ord('A') + 1) * (26 ** i)
                          for i, c in enumerate(reversed(start_col_letter)))

            # Calculate ending position
            end_col = start_col + cols - 1
            end_row = start_row + rows - 1
            end_col_letter = self._column_number_to_letter(end_col)

            # Create range from start to end cell
            end_cell_address = f"${end_col_letter}${end_row}"
            range_obj = worksheet.Range(range_address, end_cell_address)
        else:
            # Fallback: use the starting cell only
            range_obj = start_cell

        # Write data to the range
        range_obj.Value = data

    @staticmethod
    def _column_number_to_letter(col_num: int) -> str:
        """Convert column number to Excel column letter.

        Args:
            col_num: Column number (1=A, 2=B, ..., 27=AA, etc.)

        Returns:
            Excel column letter(s)
        """
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result

    # =============================================================================
    # State Query
    # =============================================================================

    def is_running(self) -> bool:
        """Check if Excel application is running.

        Returns:
            True if Excel is initialized and running
        """
        return self._app is not None

    def get_open_workbooks(self) -> list[str]:
        """Get list of open workbook paths.

        Returns:
            List of workbook filepaths currently open
        """
        return list(self._workbooks.keys())

    def is_workbook_owned(self, filepath: str) -> bool:
        """Check if a workbook is owned by this manager (can be closed).

        Args:
            filepath: Path to the workbook

        Returns:
            True if we own the workbook (can close it), False if user opened it
        """
        str_path = self._find_workbook_path(filepath)
        if str_path is None:
            return False
        return self._workbook_owned.get(str_path, True)

    def get_workbook_count(self) -> int:
        """Get number of open workbooks.

        Returns:
            Number of workbooks currently open
        """
        return len(self._workbooks)
