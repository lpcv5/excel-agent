"""
Workbook Operations - Common Excel workbook operations.

This module provides high-level functions for common workbook operations
that will be exposed as Agent tools. These operations wrap the ExcelAppManager
to provide a clean interface for automation.

Operations include:
- Opening/closing workbooks
- Reading/writing ranges
- Saving workbooks
- Managing worksheets

IMPORTANT: Operations that modify user data (write, format, etc.) use
the preserve_user_state context manager to minimize disruption to the user's
Excel view state.
"""

from typing import Optional

from .manager import ExcelAppManager
from .context import preserve_user_state


# =============================================================================
# Workbook Lifecycle Operations
# =============================================================================

def open_workbook(
    manager: ExcelAppManager,
    filepath: str,
    read_only: bool = False
) -> tuple[object, list[str]]:
    """Open an Excel workbook and return workbook info.

    Args:
        manager: ExcelAppManager instance
        filepath: Path to the Excel file
        read_only: Whether to open in read-only mode

    Returns:
        Tuple of (workbook COM object, list of worksheet names)

    Raises:
        FileNotFoundError: If file doesn't exist
        Exception: For COM errors
    """
    workbook = manager.open_workbook(filepath, read_only)
    worksheets = manager.list_worksheets(workbook)
    return (workbook, worksheets)


def close_workbook(
    manager: ExcelAppManager,
    filepath: str,
    save: bool = True
) -> None:
    """Close an open workbook.

    Args:
        manager: ExcelAppManager instance
        filepath: Path to the workbook
        save: Whether to save changes before closing
    """
    manager.close_workbook(filepath, save)


def save_workbook(
    manager: ExcelAppManager,
    workbook: object,
    filepath: Optional[str] = None
) -> None:
    """Save a workbook.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        filepath: Optional path to save as
    """
    manager.save_workbook(workbook, filepath)


# =============================================================================
# Worksheet Operations
# =============================================================================

def add_worksheet(
    manager: ExcelAppManager,
    workbook: object,
    name: str,
    after: Optional[str] = None
) -> object:
    """Add a new worksheet to a workbook.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        name: Name for the new worksheet
        after: Name of worksheet to insert after (None = at end)

    Returns:
        The new Worksheet COM object
    """
    if after:
        after_sheet = manager.get_worksheet(workbook, after)
        worksheet = workbook.Worksheets.Add(After=after_sheet)
    else:
        worksheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    worksheet.Name = name
    return worksheet


def delete_worksheet(
    manager: ExcelAppManager,
    workbook: object,
    name: str
) -> None:
    """Delete a worksheet from a workbook.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        name: Name of worksheet to delete
    """
    worksheet = manager.get_worksheet(workbook, name)
    worksheet.Delete()


def rename_worksheet(
    manager: ExcelAppManager,
    workbook: object,
    old_name: str,
    new_name: str
) -> object:
    """Rename a worksheet.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        old_name: Current worksheet name
        new_name: New worksheet name

    Returns:
        The renamed Worksheet COM object
    """
    worksheet = manager.get_worksheet(workbook, old_name)
    worksheet.Name = new_name
    return worksheet


def copy_worksheet(
    manager: ExcelAppManager,
    workbook: object,
    name: str,
    new_name: Optional[str] = None
) -> object:
    """Copy a worksheet.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        name: Name of worksheet to copy
        new_name: Name for the copy (None = auto-name)

    Returns:
        The copied Worksheet COM object
    """
    worksheet = manager.get_worksheet(workbook, name)
    new_sheet = worksheet.Copy(After=worksheet)
    if new_name:
        workbook.ActiveSheet.Name = new_name
    return workbook.ActiveSheet


def get_used_range(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str
) -> tuple[str, int, int]:
    """Get the used range of a worksheet.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet

    Returns:
        Tuple containing:
            - Range address (str)
            - Row count (int)
            - Column count (int)
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    used_range = worksheet.UsedRange
    return (
        used_range.Address,
        used_range.Rows.Count,
        used_range.Columns.Count,
    )


# =============================================================================
# Range Operations
# =============================================================================

def read_range(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str
) -> list[list]:
    """Read data from a range.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range address (e.g., "A1", "A1:Z100")

    Returns:
        2D list of cell values
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    return manager.read_range(worksheet, range_address)


def write_range(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    data: list[list]
) -> None:
    """Write data to a range.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range address (e.g., "A1", "A1:Z100")
        data: 2D list of values to write
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        manager.write_range(worksheet, range_address, data)
