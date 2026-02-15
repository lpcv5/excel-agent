"""
Excel Application Manager - Core Excel lifecycle and workbook management.

This module provides the ExcelAppManager class which handles creating,
managing, and cleaning up the Excel COM application instance.
"""

import atexit
import gc
import os
import subprocess
import threading
import weakref
from pathlib import Path
from typing import TYPE_CHECKING, Optional

import pythoncom
import win32com.client

if TYPE_CHECKING:
    from win32com.client import Application  # type: ignore


# Global state for tracking Excel processes created by this application
_excel_pids: set[int] = set()
_excel_pids_lock = threading.Lock()
_cleanup_registered = False
_manager_registry: weakref.WeakSet = weakref.WeakSet()


def _register_cleanup():
    """Register the cleanup function with atexit."""
    global _cleanup_registered
    if not _cleanup_registered:
        atexit.register(force_cleanup_excel_processes)
        _cleanup_registered = True


def _track_excel_pid(pid: int) -> None:
    """Track an Excel process ID for cleanup."""
    with _excel_pids_lock:
        _excel_pids.add(pid)


def force_cleanup_excel_processes() -> None:
    """Force terminate all Excel processes created by this application.

    This is the ultimate cleanup method that ensures Excel processes
    are terminated even if COM cleanup fails due to threading issues.

    This function does NOT rely on COM objects and can be called from any thread.
    """
    # First, try graceful COM cleanup (may fail due to threading)
    try:
        cleanup_all_managers()
    except Exception:
        pass

    # Force multiple rounds of GC to release COM references
    for _ in range(5):
        gc.collect()

    # Get all Excel PIDs to terminate
    pids_to_terminate = set()
    with _excel_pids_lock:
        pids_to_terminate = _excel_pids.copy()
        _excel_pids.clear()

    # Terminate tracked Excel processes
    for pid in pids_to_terminate:
        try:
            _terminate_process_tree(pid)
        except Exception:
            pass

    # Always try to find and terminate any Excel processes that might have been
    # created but not tracked (e.g., from attach_to_existing mode or thread issues)
    try:
        _terminate_all_child_excel_processes()
    except Exception:
        pass


def _terminate_all_child_excel_processes() -> None:
    """Find and terminate ALL Excel processes spawned by this Python process."""
    try:
        import psutil
        current_process = psutil.Process()

        # Find all child processes
        try:
            children = current_process.children(recursive=True)
        except psutil.Error:
            children = []

        for child in children:
            try:
                if child.name().upper() == 'EXCEL.EXE':
                    child.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        # Also check any remaining Excel processes
        for proc in psutil.process_iter(['pid', 'name', 'ppid']):
            try:
                if proc.info['name'] and proc.info['name'].upper() == 'EXCEL.EXE':
                    # Check if parent is our process
                    if proc.info['ppid'] == os.getpid():
                        proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

    except ImportError:
        # psutil not available, fall back to tasklist/taskkill
        _terminate_orphaned_excel_processes()


def _terminate_process_tree(pid: int) -> None:
    """Terminate a process and all its children."""
    try:
        # Use taskkill to force terminate the process tree
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            capture_output=True,
            timeout=10
        )
    except Exception:
        pass


def _terminate_orphaned_excel_processes() -> None:
    """Find and terminate Excel processes that might have been created by this app."""
    try:
        # Get current process ID to avoid killing parent Excel if any
        current_pid = os.getpid()

        # Use WMIC to find Excel processes
        result = subprocess.run(
            ["wmic", "process", "where", "name='EXCEL.EXE'", "get", "processid,parentprocessid"],
            capture_output=True,
            text=True,
            timeout=10
        )

        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            for line in lines[1:]:  # Skip header
                parts = line.strip().split()
                if len(parts) >= 2:
                    try:
                        pid = int(parts[0])
                        parent_pid = int(parts[1])
                        # If this Excel was spawned by our process, terminate it
                        if parent_pid == current_pid:
                            _terminate_process_tree(pid)
                    except (ValueError, IndexError):
                        continue
    except Exception:
        pass


def release_com_object(obj) -> None:
    """Release a COM object reference."""
    if obj is None:
        return
    try:
        win32com.client.pythoncom.ReleaseObject(obj)
    except Exception:
        pass


def cleanup_all_managers() -> None:
    """Clean up all manager instances across all threads."""
    for manager in list(_manager_registry):
        try:
            manager.stop()
        except Exception:
            pass

    for _ in range(3):
        gc.collect()


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
        self._attached_to_existing = False  # Track if we attached to existing Excel
        self._com_initialized = False  # Track if we initialized COM

        # Register this manager for global cleanup
        _manager_registry.add(self)

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

        # Register atexit cleanup handler
        _register_cleanup()

        # Initialize COM for this thread (required for multi-threaded environments)
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
        except Exception:
            pass  # Already initialized or not needed

        # Get PIDs before creating Excel to detect new process
        pids_before = self._get_excel_pids()

        if self._attach_to_existing:
            try:
                # Try to attach to existing Excel instance
                # This will connect to user's Excel if already open
                self._app = win32com.client.Dispatch("Excel.Application")
                # Check if we successfully connected to a running instance
                # by trying to access the Workbooks collection
                _ = self._app.Workbooks.Count  # type: ignore
                self._attached_to_existing = True
            except Exception:
                # No existing Excel instance, create a new one
                self._app = win32com.client.DispatchEx("Excel.Application")
                self._attached_to_existing = False
        else:
            # Create a separate Excel instance
            self._app = win32com.client.DispatchEx("Excel.Application")
            self._attached_to_existing = False

        self._app.Visible = self._visible  # type: ignore
        self._app.DisplayAlerts = self._display_alerts  # type: ignore

        # Track the Excel PID if we created a new instance
        if not self._attached_to_existing:
            pids_after = self._get_excel_pids()
            new_pids = pids_after - pids_before
            for pid in new_pids:
                _track_excel_pid(pid)

        # Track all workbooks that are already open in the Excel instance
        # We should NOT close these on shutdown
        if self._app:
            try:
                for i in range(1, self._app.Workbooks.Count + 1):  # type: ignore
                    try:
                        wb = self._app.Workbooks(i)  # type: ignore
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

    def is_app_alive(self) -> bool:
        """Check if the Excel COM application object is still alive.

        Returns:
            True if the COM app is responsive, False otherwise.
        """
        if self._app is None:
            return False
        try:
            _ = self._app.Workbooks.Count  # type: ignore
            return True
        except Exception:
            # COM object likely disconnected; clear reference
            self._app = None
            self._workbooks.clear()
            self._workbook_owned.clear()
            return False

    def _get_excel_pids(self) -> set[int]:
        """Get all Excel process IDs."""
        pids = set()
        try:
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                for line in result.stdout.strip().split('\n'):
                    if 'EXCEL.EXE' in line:
                        parts = line.split(',')
                        if len(parts) >= 2:
                            try:
                                pid = int(parts[1].strip('"'))
                                pids.add(pid)
                            except ValueError:
                                continue
        except Exception:
            pass
        return pids

    def stop(self, force_quit: bool = False) -> None:
        """Stop the Excel application and clean up resources.

        Args:
            force_quit: If True, always quit Excel even if attached to existing instance

        Note:
            If we attached to an existing Excel instance, we only release our reference
            unless force_quit is True.
            If we created our own Excel instance, we quit it after closing our workbooks.
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
            # Always quit if force_quit, or if we created our own instance
            should_quit = force_quit or not attached_to_existing
            if should_quit:
                try:
                    # Close any remaining workbooks first
                    while app_ref.Workbooks.Count > 0:  # type: ignore
                        try:
                            app_ref.Workbooks(1).Close(SaveChanges=False)  # type: ignore
                        except Exception:
                            break
                    # Now quit
                    app_ref.Quit()  # type: ignore
                except Exception:
                    # Log but don't raise - we still need to clean up
                    pass

            # Clear the app reference
            self._app = None

            # Release COM object reference
            try:
                release_com_object(app_ref)
            except Exception:
                pass
            finally:
                app_ref = None

            # Force multiple rounds of garbage collection
            for _ in range(3):
                gc.collect()

            # Uninitialize COM if we initialized it
            if self._com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                finally:
                    self._com_initialized = False

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

        # Return cached workbook if already tracked and still alive
        if str_path in self._workbooks:
            workbook = self._workbooks[str_path]
            try:
                _ = workbook.Name  # type: ignore
                _ = workbook.FullName  # type: ignore
                return workbook
            except Exception:
                # Cached workbook is stale/disconnected; remove and reopen
                self._workbooks.pop(str_path, None)
                self._workbook_owned.pop(str_path, None)

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
            # Try alternative path formats for Chinese/multibyte filenames
            # Some Windows configurations need different encoding approaches
            import os

            try:
                # Try using the path as-is but ensure it's a string
                workbook = self._app.Workbooks.Open(str(str_path), ReadOnly=read_only)
                self._workbooks[str_path] = workbook
                self._workbook_owned[str_path] = True
                return workbook
            except Exception:
                try:
                    # Try with os.path for absolute path normalization
                    abs_path = os.path.abspath(str_path)
                    workbook = self._app.Workbooks.Open(abs_path, ReadOnly=read_only)
                    self._workbooks[str_path] = workbook
                    self._workbook_owned[str_path] = True
                    return workbook
                except Exception:
                    # All attempts failed, raise original error with details
                    raise Exception(f"Failed to open workbook '{filepath}' (tried: '{str_path}', '{os.path.abspath(str_path)}'): {e}") from e

    def close_workbook(self, filepath: str, save: bool = True, force: bool = False) -> None:
        """Close an open workbook.

        Args:
            filepath: Path to the workbook (must match the path used to open it)
            save: Whether to save changes before closing
            force: If True, close even if we don't own the workbook (default: False)

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

            # Always close the workbook when force=True or we own it
            # The "ownership" concept was causing issues where workbooks wouldn't close
            if force or we_own_it:
                try:
                    workbook.Close(SaveChanges=save)  # type: ignore
                except Exception:
                    # Excel may have already exited
                    pass
            else:
                # Not owned and not forced - just save if requested
                if save:
                    try:
                        workbook.Save()  # type: ignore
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
        # Normalize the path to handle different separators (forward vs backslash)
        try:
            normalized = str(Path(filepath).resolve()).lower()
        except Exception:
            normalized = filepath.lower()

        for open_path in self._workbooks.keys():
            if open_path.lower() == normalized:
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

    def create_workbook(self, filepath: Optional[str] = None) -> object:
        """Create a new Excel workbook.

        Creates a new blank workbook. If a filepath is provided, the workbook
        will be saved to that location.

        Args:
            filepath: Optional path to save the workbook to

        Returns:
            The Workbook COM object

        Raises:
            RuntimeError: If Excel application is not initialized
            Exception: For other COM-related errors
        """
        if self._app is None:
            raise RuntimeError("Excel application is not initialized. Call start() first.")

        try:
            # Create a new blank workbook
            workbook = self._app.Workbooks.Add()

            # If filepath provided, save the workbook
            if filepath:
                save_path = str(Path(filepath).resolve())
                workbook.SaveAs(save_path)
                self._workbooks[save_path] = workbook
                self._workbook_owned[save_path] = True
            else:
                # Track by name for unsaved workbooks
                workbook_name = workbook.Name
                self._workbooks[workbook_name] = workbook
                self._workbook_owned[workbook_name] = True

            return workbook
        except Exception as e:
            raise Exception(f"Failed to create workbook: {e}") from e

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
        Also converts datetime objects to ISO format strings for JSON serialization.

        Args:
            value: Value from Excel COM

        Returns:
            Normalized value
        """
        import datetime

        # Handle datetime objects (check by class name for pywintypes compatibility)
        value_type = type(value).__name__
        if isinstance(value, (datetime.datetime, datetime.date)):
            return value.isoformat()
        elif value_type in ('time', 'datetime', 'date') or hasattr(value, 'isoformat'):
            # Handle pywintypes datetime or other datetime-like objects
            try:
                return value.isoformat()
            except AttributeError:
                return str(value)

        # Convert float to int if whole number
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

    # =============================================================================
    # Cleanup and Save Operations
    # =============================================================================

    def get_unsaved_workbooks(self) -> list[dict]:
        """Get list of workbooks with unsaved changes.

        Returns:
            List of dicts with workbook info: {'name': str, 'path': str, 'saved': bool}
        """
        unsaved = []
        if self._app is None:
            return unsaved

        try:
            for i in range(1, self._app.Workbooks.Count + 1):
                try:
                    wb = self._app.Workbooks(i)
                    # Check if workbook has unsaved changes
                    # Saved property is True if no changes since last save
                    if not wb.Saved:
                        unsaved.append({
                            'name': wb.Name,
                            'path': wb.FullName if wb.Path else None,
                            'saved': wb.Saved,
                        })
                except Exception:
                    continue
        except Exception:
            pass

        return unsaved

    def save_all_workbooks(self) -> dict:
        """Save all open workbooks.

        Returns:
            Dict with 'saved' count and any 'errors'
        """
        saved_count = 0
        errors = []

        if self._app is None:
            return {'saved': 0, 'errors': []}

        try:
            for i in range(1, self._app.Workbooks.Count + 1):
                try:
                    wb = self._app.Workbooks(i)
                    if not wb.Saved:
                        if wb.Path:
                            # Has been saved before, just save
                            wb.Save()
                            saved_count += 1
                        else:
                            # New workbook, never saved - skip (would need a path)
                            errors.append(f"'{wb.Name}' has never been saved (no path)")
                except Exception as e:
                    errors.append(f"Failed to save workbook: {e}")
        except Exception as e:
            errors.append(f"Error accessing workbooks: {e}")

        return {'saved': saved_count, 'errors': errors}

    def close_all_workbooks(self, save: bool = True) -> dict:
        """Close all workbooks.

        Args:
            save: Whether to save changes before closing

        Returns:
            Dict with 'closed' count and any 'errors'
        """
        closed_count = 0
        errors = []

        if self._app is None:
            return {'closed': 0, 'errors': []}

        try:
            # Close workbooks from back to front to avoid index issues
            while self._app.Workbooks.Count > 0:
                try:
                    wb = self._app.Workbooks(self._app.Workbooks.Count)
                    name = wb.Name
                    try:
                        wb.Close(SaveChanges=save)
                        closed_count += 1
                    except Exception as e:
                        errors.append(f"Failed to close '{name}': {e}")
                except Exception:
                    break
        except Exception as e:
            errors.append(f"Error closing workbooks: {e}")

        # Clear tracking
        self._workbooks.clear()
        self._workbook_owned.clear()

        return {'closed': closed_count, 'errors': errors}

    def quit_excel(self) -> bool:
        """Quit Excel application completely.

        Returns:
            True if Excel was quit successfully
        """
        if self._app is None:
            return True

        app_ref = self._app

        try:
            app_ref.Quit()  # type: ignore
        except Exception:
            pass  # May have already exited

        # Clear reference
        self._app = None

        # Release COM object
        try:
            release_com_object(app_ref)
        except Exception:
            pass
        finally:
            app_ref = None

        # Force garbage collection
        for _ in range(3):
            gc.collect()

        # Uninitialize COM if we initialized it
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            finally:
                self._com_initialized = False

        return True
