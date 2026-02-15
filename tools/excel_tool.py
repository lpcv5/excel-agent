"""
Excel tools for DeepAgents using Windows COM interface.

This module provides tools for reading, writing, and manipulating Excel files
via the Windows COM interface. Requires Microsoft Excel and Windows platform.

Low-level COM operations are in libs.excel_com package.
"""

import gc
import json
import threading
from contextlib import contextmanager
from pathlib import Path
from typing import Optional, Iterator, Any

from langchain_core.tools import tool

from libs.excel_com.manager import (
    ExcelAppManager,
    cleanup_all_managers,
    force_cleanup_excel_processes,
)
from libs.excel_com import workbook_ops, formatting_ops, formula_ops


# =============================================================================
# Singleton Excel Manager + Operation Lock
# =============================================================================

_excel_manager: ExcelAppManager | None = None
_excel_lock = threading.RLock()


def get_excel_manager() -> ExcelAppManager:
    """Get or create the singleton Excel manager instance (lazy-loaded)."""
    global _excel_manager
    with _excel_lock:
        if _excel_manager is None:
            _excel_manager = ExcelAppManager(visible=False, display_alerts=False)
        manager = _excel_manager

        if not manager.is_running():
            manager.start()
        elif not manager.is_app_alive():
            try:
                manager.stop(force_quit=False)
            except Exception:
                pass
            manager.start()

        return manager


def cleanup_excel_resources(force: bool = True) -> dict:
    """Clean up all Excel COM resources across all threads.

    This function should be called when the application is closing
    to ensure all Excel COM objects are properly released.

    Args:
        force: If True, force-terminate any remaining Excel processes
            created by this application.

    Returns:
        Dict with cleanup status
    """
    errors = []

    # First try graceful COM cleanup
    try:
        cleanup_all_managers()
    except Exception as e:
        errors.append(f"COM cleanup error: {str(e)}")

    # Clear singleton manager reference
    global _excel_manager
    _excel_manager = None

    # Force multiple rounds of garbage collection
    for _ in range(5):
        gc.collect()

    if force:
        # Force terminate any remaining Excel processes
        try:
            force_cleanup_excel_processes()
        except Exception as e:
            errors.append(f"Force cleanup error: {str(e)}")

    return {
        "success": len(errors) == 0,
        "errors": errors if errors else None,
        "forced": force,
    }


def normalize_filepath(filepath: str) -> str:
    """Normalize a filepath to absolute path for consistent matching.

    Args:
        filepath: The filepath to normalize

    Returns:
        Normalized absolute path as string
    """
    try:
        return str(Path(filepath).resolve())
    except Exception:
        return filepath


def _ok(summary: str, data: Optional[dict[str, Any]] = None) -> str:
    payload = {
        "success": True,
        "summary": summary,
        "data": data or {},
    }
    return json.dumps(payload, indent=2, ensure_ascii=False)


def _err(summary: str, error: str, data: Optional[dict[str, Any]] = None) -> str:
    payload = {
        "success": False,
        "summary": summary,
        "error": error,
    }
    if data is not None:
        payload["data"] = data
    return json.dumps(payload, indent=2, ensure_ascii=False)


@contextmanager
def workbook_context(
    filepath: str,
    read_only: bool,
    save_on_close: bool,
) -> Iterator[tuple[ExcelAppManager, object, str]]:
    """Open a workbook for a single operation and ensure it is closed."""
    manager = get_excel_manager()
    normalized_path = normalize_filepath(filepath)
    with _excel_lock:
        workbook = manager.open_workbook(normalized_path, read_only=read_only)
        try:
            yield manager, workbook, normalized_path
        finally:
            try:
                if manager.is_workbook_owned(normalized_path):
                    manager.close_workbook(normalized_path, save=save_on_close, force=True)
                else:
                    manager.close_workbook(normalized_path, save=save_on_close, force=False)
            except Exception:
                pass


# =============================================================================
# Status Tools (1)
# =============================================================================

@tool
def excel_status() -> str:
    """Get the current status of the Excel application.

    Returns information about whether Excel is running and which workbooks
    are currently open.

    Returns:
        JSON string with Excel status information.
    """
    try:
        if _excel_manager is None:
            return _ok(
                "Excel not initialized.",
                {"excel_running": False, "workbook_count": 0, "open_workbooks": []},
            )

        manager = _excel_manager
        is_running = manager.is_running() and manager.is_app_alive()
        result = {
            "excel_running": is_running,
            "workbook_count": manager.get_workbook_count() if is_running else 0,
            "open_workbooks": manager.get_open_workbooks() if is_running else [],
        }

        summary = "Excel is running." if is_running else "Excel is not running."
        return _ok(summary, result)

    except Exception as e:
        return _err("Failed to get Excel status.", str(e))


# =============================================================================
# Deprecated Workbook Tools (2)
# =============================================================================

@tool
def excel_open_workbook(filepath: str, read_only: bool = False) -> str:
    """Open an Excel workbook file.

    Opens the specified Excel file and returns information about the workbook
    including available worksheets.

    Args:
        filepath: Path to the Excel file (.xlsx, .xls, .xlsm)
        read_only: Open in read-only mode to prevent accidental changes (default: False)

    Returns:
        JSON string with workbook name and list of worksheet names.
    """
    return _err(
        "excel_open_workbook is no longer supported.",
        "Workbooks are opened and closed automatically per tool call.",
    )


# =============================================================================
# Workbook Tools (3)
# =============================================================================

@tool
def excel_create_workbook(filepath: str, sheet_names: Optional[str] = None) -> str:
    """Create a new Excel workbook.

    Creates a new blank workbook and optionally saves it to the specified path.
    The workbook will be opened automatically after creation.

    Args:
        filepath: Path where the workbook will be saved (.xlsx)
        sheet_names: Optional JSON string array of sheet names (e.g., '["Sheet1", "Sheet2"]')

    Returns:
        JSON string with workbook information.
    """
    try:
        manager = get_excel_manager()
        normalized_path = normalize_filepath(filepath)

        sheet_names_list = None
        if sheet_names:
            try:
                sheet_names_list = json.loads(sheet_names)
            except json.JSONDecodeError:
                return _err("Invalid sheet_names.", "sheet_names must be a valid JSON array")

        with _excel_lock:
            workbook, worksheets = workbook_ops.create_workbook(
                manager, normalized_path, sheet_names_list
            )
            workbook_name = workbook.Name
            try:
                manager.close_workbook(normalized_path, save=True, force=True)
            except Exception:
                pass

        return _ok(
            f"Workbook created: {workbook_name}",
            {
                "workbook_name": workbook_name,
                "workbook_path": normalized_path,
                "worksheets": worksheets,
            },
        )

    except Exception as e:
        return _err("Failed to create workbook.", str(e))


@tool
def excel_list_worksheets(filepath: str) -> str:
    """List all worksheets in an Excel workbook.

    Args:
        filepath: Path to the Excel file

    Returns:
        JSON string with list of worksheet names.
    """
    try:
        with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
            worksheets = manager.list_worksheets(workbook)
        return _ok(
            "Worksheets listed.",
            {"workbook_path": path, "worksheets": worksheets, "count": len(worksheets)},
        )
    except Exception as e:
        return _err("Failed to list worksheets.", str(e))


# =============================================================================
# Range Tools (2)
# =============================================================================

@tool
def excel_read_range(
    filepath: str,
    worksheet_name: str,
    range_address: str
) -> str:
    """Read data from a specified range in an Excel worksheet.

    Args:
        filepath: Path to the Excel file (opened automatically)
        worksheet_name: Name of the worksheet to read from
        range_address: Excel range address (e.g., "A1", "A1:D10", "B:B")

    Returns:
        JSON string with the data from the specified range.
    """
    try:
        with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
            data = workbook_ops.read_range(manager, workbook, worksheet_name, range_address)

        return _ok(
            "Range read successfully.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "range": range_address,
                "rows": len(data),
                "columns": len(data[0]) if data else 0,
                "data": data,
            },
        )
    except Exception as e:
        return _err("Failed to read range.", str(e))


@tool
def excel_write_range(
    filepath: str,
    worksheet_name: str,
    range_address: str,
    data: str
) -> str:
    """Write data to a specified range in an Excel worksheet.

    The data size must match or fit within the specified range.
    Data is preserved in the user's current view state.

    Args:
        filepath: Path to the Excel file (opened automatically)
        worksheet_name: Name of the worksheet to write to
        range_address: Starting cell address (e.g., "A1") or range (e.g., "A1:D10")
        data: JSON string representing 2D array of data to write

    Returns:
        JSON string with operation result.
    """
    try:
        parsed_data = json.loads(data)
        if not isinstance(parsed_data, list) or not all(isinstance(row, list) for row in parsed_data):
            return _err("Invalid data.", "Data must be a 2D array (list of lists)")

        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            workbook_ops.write_range(manager, workbook, worksheet_name, range_address, parsed_data)

        return _ok(
            "Range written successfully.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "range": range_address,
                "rows_written": len(parsed_data),
                "columns_written": len(parsed_data[0]) if parsed_data else 0,
            },
        )
    except json.JSONDecodeError as e:
        return _err("Invalid JSON data.", str(e))
    except Exception as e:
        return _err("Failed to write range.", str(e))


@tool
def excel_save_workbook(filepath: str, save_as: Optional[str] = None) -> str:
    """Save an open Excel workbook.

    Args:
        filepath: Path of the workbook to save
        save_as: Optional new filepath to save as (Save As functionality)

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=False) as (manager, workbook, path):
            manager.save_workbook(workbook, save_as)

        return _ok(
            "Workbook saved.",
            {"workbook_path": path, "saved_as": save_as if save_as else path},
        )
    except Exception as e:
        return _err("Failed to save workbook.", str(e))


@tool
def excel_close_workbook(filepath: str, save: bool = True) -> str:
    """Close an open Excel workbook.

    Args:
        filepath: Path of the workbook to close
        save: Whether to save changes before closing (default: True)

    Returns:
        JSON string with operation result.
    """
    return _err(
        "excel_close_workbook is no longer supported.",
        "Workbooks are opened and closed automatically per tool call.",
    )


# =============================================================================
# Worksheet Tools (5)
# =============================================================================

@tool
def excel_add_worksheet(
    filepath: str,
    worksheet_name: str,
    after: Optional[str] = None
) -> str:
    """Add a new worksheet to an Excel workbook.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name for the new worksheet
        after: Name of worksheet to insert after (optional, defaults to end)

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            workbook_ops.add_worksheet(manager, workbook, worksheet_name, after)

        return _ok(
            "Worksheet added.",
            {"workbook_path": path, "worksheet_name": worksheet_name, "inserted_after": after},
        )
    except Exception as e:
        return _err("Failed to add worksheet.", str(e))


@tool
def excel_delete_worksheet(filepath: str, worksheet_name: str) -> str:
    """Delete a worksheet from an Excel workbook.

    WARNING: This operation cannot be undone. The worksheet and all its data
    will be permanently deleted.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet to delete

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            workbook_ops.delete_worksheet(manager, workbook, worksheet_name)

        return _ok(
            "Worksheet deleted.",
            {"workbook_path": path, "deleted_worksheet": worksheet_name},
        )
    except Exception as e:
        return _err("Failed to delete worksheet.", str(e))


@tool
def excel_rename_worksheet(
    filepath: str,
    old_name: str,
    new_name: str
) -> str:
    """Rename a worksheet in an Excel workbook.

    Args:
        filepath: Path of the open workbook
        old_name: Current name of the worksheet
        new_name: New name for the worksheet

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            workbook_ops.rename_worksheet(manager, workbook, old_name, new_name)

        return _ok(
            "Worksheet renamed.",
            {"workbook_path": path, "old_name": old_name, "new_name": new_name},
        )
    except Exception as e:
        return _err("Failed to rename worksheet.", str(e))


@tool
def excel_copy_worksheet(
    filepath: str,
    worksheet_name: str,
    new_name: Optional[str] = None
) -> str:
    """Copy a worksheet within the same workbook.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet to copy
        new_name: Name for the copied worksheet (optional, auto-generated if not provided)

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            new_sheet = workbook_ops.copy_worksheet(manager, workbook, worksheet_name, new_name)

        return _ok(
            "Worksheet copied.",
            {
                "workbook_path": path,
                "source_worksheet": worksheet_name,
                "new_worksheet_name": new_sheet.Name,
            },
        )
    except Exception as e:
        return _err("Failed to copy worksheet.", str(e))


@tool
def excel_get_used_range(filepath: str, worksheet_name: str) -> str:
    """Get the used range of a worksheet.

    Returns the address and dimensions of the range that contains data.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet

    Returns:
        JSON string with range address and dimensions.
    """
    try:
        with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
            address, rows, cols = workbook_ops.get_used_range(manager, workbook, worksheet_name)

        return _ok(
            "Used range retrieved.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "used_range": address,
                "rows": rows,
                "columns": cols,
            },
        )
    except Exception as e:
        return _err("Failed to get used range.", str(e))


# =============================================================================
# Format Tools (4)
# =============================================================================

@tool
def excel_set_font_format(
    filepath: str,
    worksheet_name: str,
    range_address: str,
    font_name: Optional[str] = None,
    size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None
) -> str:
    """Set font formatting for a range of cells.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        range_address: Range address (e.g., "A1:D10")
        font_name: Font name (e.g., "Arial", "Calibri")
        size: Font size in points
        bold: Whether text is bold
        italic: Whether text is italic
        underline: Whether text is underlined
        color: RGB hex color (e.g., "FF0000" for red)

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            formatting_ops.set_font_format(
                manager, workbook, worksheet_name, range_address,
                font_name, size, bold, italic, underline, color
            )

        return _ok(
            "Font format applied.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "range": range_address,
                "formatting_applied": {
                    k: v for k, v in {
                        "font_name": font_name,
                        "size": size,
                        "bold": bold,
                        "italic": italic,
                        "underline": underline,
                        "color": color
                    }.items() if v is not None
                },
            },
        )
    except Exception as e:
        return _err("Failed to set font format.", str(e))


@tool
def excel_set_cell_format(
    filepath: str,
    worksheet_name: str,
    range_address: str,
    horizontal_alignment: Optional[str] = None,
    vertical_alignment: Optional[str] = None,
    wrap_text: Optional[bool] = None,
    number_format: Optional[str] = None
) -> str:
    """Set cell formatting for a range.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        range_address: Range address (e.g., "A1:D10")
        horizontal_alignment: Horizontal alignment ("left", "center", "right", "general")
        vertical_alignment: Vertical alignment ("top", "center", "bottom", "justify")
        wrap_text: Whether to wrap text in cells
        number_format: Excel number format code (e.g., "#,##0.00", "0%")

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            formatting_ops.set_cell_format(
                manager, workbook, worksheet_name, range_address,
                horizontal_alignment, vertical_alignment, wrap_text, number_format
            )

        return _ok(
            "Cell format applied.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "range": range_address,
                "formatting_applied": {
                    k: v for k, v in {
                        "horizontal_alignment": horizontal_alignment,
                        "vertical_alignment": vertical_alignment,
                        "wrap_text": wrap_text,
                        "number_format": number_format
                    }.items() if v is not None
                },
            },
        )
    except Exception as e:
        return _err("Failed to set cell format.", str(e))


@tool
def excel_set_border_format(
    filepath: str,
    worksheet_name: str,
    range_address: str,
    edge: str,
    style: Optional[str] = None,
    weight: Optional[int] = None,
    color: Optional[str] = None
) -> str:
    """Set border formatting for a range.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        range_address: Range address (e.g., "A1:D10")
        edge: Which edge(s) to format ("left", "right", "top", "bottom", "all")
        style: Border style ("continuous", "dash", "dot", "double")
        weight: Border weight (1-4, where 1=hairline, 2=thin, 3=medium, 4=thick)
        color: RGB hex color (e.g., "000000" for black)

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            formatting_ops.set_border_format(
                manager, workbook, worksheet_name, range_address,
                edge, style, weight, color
            )

        return _ok(
            "Border format applied.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "range": range_address,
                "border_applied": {
                    "edge": edge,
                    "style": style,
                    "weight": weight,
                    "color": color
                },
            },
        )
    except Exception as e:
        return _err("Failed to set border format.", str(e))


@tool
def excel_set_background_color(
    filepath: str,
    worksheet_name: str,
    range_address: str,
    color: str
) -> str:
    """Set background color for a range of cells.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        range_address: Range address (e.g., "A1:D10")
        color: RGB hex color (e.g., "FFFF00" for yellow, "FF0000" for red)

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            formatting_ops.set_background_color(
                manager, workbook, worksheet_name, range_address, color
            )

        return _ok(
            "Background color applied.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "range": range_address,
                "background_color": color,
            },
        )
    except Exception as e:
        return _err("Failed to set background color.", str(e))


# =============================================================================
# Formula Tools (2)
# =============================================================================

@tool
def excel_set_formula(
    filepath: str,
    worksheet_name: str,
    range_address: str,
    formula: str
) -> str:
    """Set a formula in a cell or range.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        range_address: Cell address (e.g., "A1") or range
        formula: Excel formula (must start with "=")

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            formula_ops.set_formula(
                manager, workbook, worksheet_name, range_address, formula
            )

        return _ok(
            "Formula set.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "cell": range_address,
                "formula": formula,
            },
        )
    except Exception as e:
        return _err("Failed to set formula.", str(e))


@tool
def excel_get_formula(
    filepath: str,
    worksheet_name: str,
    range_address: str
) -> str:
    """Get the formula from a cell.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        range_address: Cell address (e.g., "A1")

    Returns:
        JSON string with the formula or a message if the cell contains a value.
    """
    try:
        with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
            formula = formula_ops.get_formula(
                manager, workbook, worksheet_name, range_address
            )

        is_formula = formula.startswith("=") if formula else False
        return _ok(
            "Formula retrieved.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "cell": range_address,
                "formula": formula,
                "is_formula": is_formula,
            },
        )
    except Exception as e:
        return _err("Failed to get formula.", str(e))


# =============================================================================
# Column/Row Tools (3)
# =============================================================================

@tool
def excel_auto_fit_columns(
    filepath: str,
    worksheet_name: str,
    range_address: Optional[str] = None
) -> str:
    """Auto-fit column widths to fit content.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        range_address: Optional column range (e.g., "A:D"). If not provided, fits all used columns.

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            formatting_ops.auto_fit_columns(
                manager, workbook, worksheet_name, range_address
            )

        return _ok(
            "Columns auto-fitted.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "columns_auto_fitted": range_address if range_address else "all used columns",
            },
        )
    except Exception as e:
        return _err("Failed to auto-fit columns.", str(e))


@tool
def excel_set_column_width(
    filepath: str,
    worksheet_name: str,
    columns: str,
    width: float
) -> str:
    """Set column width for specific columns.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        columns: Column specification (e.g., "A", "A:C", "A,D,F")
        width: Width in characters (standard width is about 8.43)

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            formatting_ops.set_column_width(
                manager, workbook, worksheet_name, columns, width
            )

        return _ok(
            "Column width set.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "columns": columns,
                "width": width,
            },
        )
    except Exception as e:
        return _err("Failed to set column width.", str(e))


@tool
def excel_set_row_height(
    filepath: str,
    worksheet_name: str,
    rows: str,
    height: float
) -> str:
    """Set row height for specific rows.

    Args:
        filepath: Path of the open workbook
        worksheet_name: Name of the worksheet
        rows: Row specification (e.g., "1", "1:5", "1,3,5")
        height: Height in points (standard height is about 15)

    Returns:
        JSON string with operation result.
    """
    try:
        with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
            formatting_ops.set_row_height(
                manager, workbook, worksheet_name, rows, height
            )

        return _ok(
            "Row height set.",
            {
                "workbook_path": path,
                "worksheet": worksheet_name,
                "rows": rows,
                "height": height,
            },
        )
    except Exception as e:
        return _err("Failed to set row height.", str(e))


# =============================================================================
# Export all tools
# =============================================================================

EXCEL_TOOLS = [
    # Status
    excel_status,
    # Workbook
    excel_create_workbook,
    excel_list_worksheets,
    excel_save_workbook,
    # Range
    excel_read_range,
    excel_write_range,
    # Worksheet
    excel_add_worksheet,
    excel_delete_worksheet,
    excel_rename_worksheet,
    excel_copy_worksheet,
    excel_get_used_range,
    # Format
    excel_set_font_format,
    excel_set_cell_format,
    excel_set_border_format,
    excel_set_background_color,
    # Formula
    excel_set_formula,
    excel_get_formula,
    # Column/Row
    excel_auto_fit_columns,
    excel_set_column_width,
    excel_set_row_height,
]
