"""
Excel tools for DeepAgents using Windows COM interface.

CRUD-based API design:
- excel_open: Open/create workbook (lifecycle management)
- excel_save: Save workbook
- excel_info: Get workbook metadata
- excel_create: Create worksheet, rows, columns
- excel_read: Read range data, format, formula
- excel_update: Update range data, format, formula, worksheet properties, structure
- excel_delete: Delete worksheet, range content, rows, columns, format
"""

import gc
import json
import threading
from contextlib import contextmanager
from pathlib import Path
from typing import Optional, Iterator, Any, Literal

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
    """Clean up all Excel COM resources across all threads."""
    errors = []

    try:
        cleanup_all_managers()
    except Exception as e:
        errors.append(f"COM cleanup error: {str(e)}")

    global _excel_manager
    _excel_manager = None

    for _ in range(5):
        gc.collect()

    if force:
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
    """Normalize a filepath to absolute path for consistent matching."""
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
# Type Definitions
# =============================================================================

TargetCreate = Literal["workbook", "worksheet", "rows", "columns"]
TargetRead = Literal["worksheets", "range", "format", "formula", "used_range"]
TargetUpdate = Literal["range", "format", "formula", "worksheet", "structure"]
TargetDelete = Literal["worksheet", "range", "rows", "columns", "format"]


# =============================================================================
# Workbook Lifecycle Tools (3)
# =============================================================================

@tool
def excel_open(filepath: str, create_if_missing: bool = False, sheet_names: Optional[str] = None) -> str:
    """Open an Excel workbook, optionally creating it if it doesn't exist.

    This tool verifies the workbook can be accessed and returns its metadata.
    Workbooks are automatically managed (opened/closed) per operation.

    Args:
        filepath: Path to the Excel file (.xlsx, .xls, .xlsm)
        create_if_missing: If True, create a new workbook if file doesn't exist
        sheet_names: JSON array of sheet names for new workbook (e.g., '["Sheet1", "Data"]')

    Returns:
        JSON with workbook info including worksheets list.

    Examples:
        - Open existing: excel_open("data.xlsx")
        - Create new: excel_open("new.xlsx", create_if_missing=True)
        - Create with sheets: excel_open("report.xlsx", create_if_missing=True, sheet_names='["Summary", "Details"]')
    """
    try:
        normalized_path = normalize_filepath(filepath)
        path_exists = Path(normalized_path).exists()

        if not path_exists and not create_if_missing:
            return _err(
                "Workbook not found.",
                f"File does not exist: {normalized_path}. Set create_if_missing=True to create it."
            )

        manager = get_excel_manager()

        if not path_exists:
            # Create new workbook
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
                    "created": True,
                },
            )
        else:
            # Open existing workbook
            with workbook_context(filepath, read_only=True, save_on_close=False) as (mgr, workbook, path):
                worksheets = mgr.list_worksheets(workbook)
                workbook_name = workbook.Name

            return _ok(
                f"Workbook opened: {workbook_name}",
                {
                    "workbook_name": workbook_name,
                    "workbook_path": normalized_path,
                    "worksheets": worksheets,
                    "created": False,
                },
            )

    except Exception as e:
        return _err("Failed to open workbook.", str(e))


@tool
def excel_save(filepath: str, save_as: Optional[str] = None) -> str:
    """Save an Excel workbook.

    Args:
        filepath: Path of the workbook to save
        save_as: Optional new filepath for "Save As" functionality

    Returns:
        JSON with save operation result.

    Examples:
        - Save: excel_save("data.xlsx")
        - Save As: excel_save("data.xlsx", save_as="data_backup.xlsx")
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
def excel_info() -> str:
    """Get the current status of Excel application and open workbooks.

    Returns information about whether Excel is running and provides
    debugging information about the COM connection state.

    Returns:
        JSON with Excel status information.
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
# CRUD: Create (1)
# =============================================================================

@tool
def excel_create(filepath: str, target: str, params: str) -> str:
    """Create a new resource in an Excel workbook.

    Args:
        filepath: Path to the Excel file
        target: Type of resource to create ("worksheet", "rows", "columns")
        params: JSON string with creation parameters

    Returns:
        JSON with creation result.

    Target-specific params:
        worksheet:
            - name (str): Name for the new worksheet
            - after (str, optional): Insert after this worksheet
        rows:
            - worksheet (str): Worksheet name
            - index (int): Row number to insert before (1-based)
            - count (int, optional): Number of rows to insert (default: 1)
        columns:
            - worksheet (str): Worksheet name
            - index (str): Column letter to insert before (e.g., "A", "C")
            - count (int, optional): Number of columns to insert (default: 1)

    Examples:
        - Add worksheet: excel_create("data.xlsx", "worksheet", '{"name": "Summary"}')
        - Insert rows: excel_create("data.xlsx", "rows", '{"worksheet": "Sheet1", "index": 5, "count": 3}')
        - Insert columns: excel_create("data.xlsx", "columns", '{"worksheet": "Sheet1", "index": "C", "count": 2}')
    """
    try:
        p = json.loads(params)
    except json.JSONDecodeError as e:
        return _err("Invalid params.", f"params must be valid JSON: {e}")

    try:
        if target == "worksheet":
            name = p.get("name")
            if not name:
                return _err("Missing parameter.", "worksheet creation requires 'name' parameter")
            after = p.get("after")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                workbook_ops.add_worksheet(manager, workbook, name, after)

            return _ok(
                f"Worksheet '{name}' created.",
                {"workbook_path": path, "worksheet_name": name, "inserted_after": after},
            )

        elif target == "rows":
            worksheet = p.get("worksheet")
            index = p.get("index")
            count = p.get("count", 1)

            if not worksheet or index is None:
                return _err("Missing parameter.", "rows creation requires 'worksheet' and 'index' parameters")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                workbook_ops.insert_rows(manager, workbook, worksheet, index, count)

            return _ok(
                f"Inserted {count} row(s) at row {index}.",
                {"workbook_path": path, "worksheet": worksheet, "index": index, "count": count},
            )

        elif target == "columns":
            worksheet = p.get("worksheet")
            index = p.get("index")
            count = p.get("count", 1)

            if not worksheet or not index:
                return _err("Missing parameter.", "columns creation requires 'worksheet' and 'index' parameters")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                workbook_ops.insert_columns(manager, workbook, worksheet, index, count)

            return _ok(
                f"Inserted {count} column(s) at column {index}.",
                {"workbook_path": path, "worksheet": worksheet, "index": index, "count": count},
            )

        else:
            return _err(
                "Invalid target.",
                f"target must be one of: worksheet, rows, columns. Got: {target}"
            )

    except Exception as e:
        return _err(f"Failed to create {target}.", str(e))


# =============================================================================
# CRUD: Read (1)
# =============================================================================

@tool
def excel_read(filepath: str, target: str, params: str) -> str:
    """Read data from an Excel workbook.

    Args:
        filepath: Path to the Excel file
        target: Type of data to read ("worksheets", "range", "format", "formula", "used_range")
        params: JSON string with read parameters

    Returns:
        JSON with the requested data.

    Target-specific params:
        worksheets:
            (no params required) - Returns list of all worksheet names
        range:
            - worksheet (str): Worksheet name
            - range (str): Range address (e.g., "A1:D10")
        format:
            - worksheet (str): Worksheet name
            - range (str): Range address
        formula:
            - worksheet (str): Worksheet name
            - range (str): Cell address (e.g., "A1")
        used_range:
            - worksheet (str): Worksheet name

    Examples:
        - List worksheets: excel_read("data.xlsx", "worksheets", '{}')
        - Read range: excel_read("data.xlsx", "range", '{"worksheet": "Sheet1", "range": "A1:D10"}')
        - Get formula: excel_read("data.xlsx", "formula", '{"worksheet": "Sheet1", "range": "A1"}')
        - Get used range: excel_read("data.xlsx", "used_range", '{"worksheet": "Sheet1"}')
    """
    try:
        p = json.loads(params)
    except json.JSONDecodeError as e:
        return _err("Invalid params.", f"params must be valid JSON: {e}")

    try:
        if target == "worksheets":
            with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
                worksheets = manager.list_worksheets(workbook)

            return _ok(
                "Worksheets listed.",
                {"workbook_path": path, "worksheets": worksheets, "count": len(worksheets)},
            )

        elif target == "range":
            worksheet = p.get("worksheet")
            range_address = p.get("range")

            if not worksheet or not range_address:
                return _err("Missing parameter.", "range read requires 'worksheet' and 'range' parameters")

            with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
                data = workbook_ops.read_range(manager, workbook, worksheet, range_address)

            return _ok(
                "Range read successfully.",
                {
                    "workbook_path": path,
                    "worksheet": worksheet,
                    "range": range_address,
                    "rows": len(data),
                    "columns": len(data[0]) if data else 0,
                    "data": data,
                },
            )

        elif target == "formula":
            worksheet = p.get("worksheet")
            range_address = p.get("range")

            if not worksheet or not range_address:
                return _err("Missing parameter.", "formula read requires 'worksheet' and 'range' parameters")

            with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
                formula = formula_ops.get_formula(manager, workbook, worksheet, range_address)

            is_formula = formula.startswith("=") if formula else False
            return _ok(
                "Formula retrieved.",
                {
                    "workbook_path": path,
                    "worksheet": worksheet,
                    "range": range_address,
                    "formula": formula,
                    "is_formula": is_formula,
                },
            )

        elif target == "used_range":
            worksheet = p.get("worksheet")

            if not worksheet:
                return _err("Missing parameter.", "used_range read requires 'worksheet' parameter")

            with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
                address, rows, cols = workbook_ops.get_used_range(manager, workbook, worksheet)

            return _ok(
                "Used range retrieved.",
                {
                    "workbook_path": path,
                    "worksheet": worksheet,
                    "used_range": address,
                    "rows": rows,
                    "columns": cols,
                },
            )

        elif target == "format":
            worksheet = p.get("worksheet")
            range_address = p.get("range")

            if not worksheet or not range_address:
                return _err("Missing parameter.", "format read requires 'worksheet' and 'range' parameters")

            with workbook_context(filepath, read_only=True, save_on_close=False) as (manager, workbook, path):
                format_info = formatting_ops.get_format(manager, workbook, worksheet, range_address)

            return _ok(
                "Format retrieved.",
                {
                    "workbook_path": path,
                    "worksheet": worksheet,
                    "range": range_address,
                    "format": format_info,
                },
            )

        else:
            return _err(
                "Invalid target.",
                f"target must be one of: worksheets, range, format, formula, used_range. Got: {target}"
            )

    except Exception as e:
        return _err(f"Failed to read {target}.", str(e))


# =============================================================================
# CRUD: Update (1)
# =============================================================================

@tool
def excel_update(filepath: str, target: str, params: str) -> str:
    """Update data or properties in an Excel workbook.

    Args:
        filepath: Path to the Excel file
        target: Type of resource to update ("range", "format", "formula", "worksheet", "structure")
        params: JSON string with update parameters

    Returns:
        JSON with update result.

    Target-specific params:
        range:
            - worksheet (str): Worksheet name
            - range (str): Starting cell (e.g., "A1") or range
            - data (list[list]): 2D array of data to write
        format:
            - worksheet (str): Worksheet name
            - ranges (list): List of format specifications, each containing:
                - range (str): Range address
                - font (dict, optional): {name, size, bold, italic, underline, color}
                - fill (dict, optional): {color}
                - border (dict, optional): {edge, style, weight, color}
                - alignment (dict, optional): {horizontal, vertical, wrap_text}
                - number_format (str, optional): Excel format code
        formula:
            - worksheet (str): Worksheet name
            - range (str): Cell address
            - formula (str): Excel formula (must start with "=")
        worksheet:
            - name (str): Current worksheet name
            - new_name (str, optional): Rename to this name
            - copy_to (str, optional): Copy worksheet with this new name
        structure:
            - worksheet (str): Worksheet name
            - columns (dict, optional): Column widths {"A": 20, "B:D": "auto"}
            - rows (dict, optional): Row heights {"1": 30, "2:5": 20}

    Examples:
        - Write data: excel_update("data.xlsx", "range", '{"worksheet": "Sheet1", "range": "A1", "data": [[1,2],[3,4]]}')
        - Set format: excel_update("data.xlsx", "format", '{"worksheet": "Sheet1", "ranges": [{"range": "A1:B2", "font": {"bold": true}, "fill": {"color": "FFFF00"}}]}')
        - Set formula: excel_update("data.xlsx", "formula", '{"worksheet": "Sheet1", "range": "C1", "formula": "=A1+B1"}')
        - Rename sheet: excel_update("data.xlsx", "worksheet", '{"name": "Sheet1", "new_name": "Data"}')
        - Auto-fit columns: excel_update("data.xlsx", "structure", '{"worksheet": "Sheet1", "columns": {"A:D": "auto"}}')
    """
    try:
        p = json.loads(params)
    except json.JSONDecodeError as e:
        return _err("Invalid params.", f"params must be valid JSON: {e}")

    try:
        if target == "range":
            worksheet = p.get("worksheet")
            range_address = p.get("range")
            data = p.get("data")

            if not worksheet or not range_address or data is None:
                return _err("Missing parameter.", "range update requires 'worksheet', 'range', and 'data' parameters")

            if not isinstance(data, list) or not all(isinstance(row, list) for row in data):
                return _err("Invalid data.", "data must be a 2D array (list of lists)")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                workbook_ops.write_range(manager, workbook, worksheet, range_address, data)

            return _ok(
                "Range updated successfully.",
                {
                    "workbook_path": path,
                    "worksheet": worksheet,
                    "range": range_address,
                    "rows_written": len(data),
                    "columns_written": len(data[0]) if data else 0,
                },
            )

        elif target == "format":
            worksheet = p.get("worksheet")
            ranges = p.get("ranges", [])

            if not worksheet:
                return _err("Missing parameter.", "format update requires 'worksheet' parameter")

            if not ranges:
                return _err("Missing parameter.", "format update requires 'ranges' parameter with at least one format specification")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                for fmt_spec in ranges:
                    range_addr = fmt_spec.get("range")
                    if not range_addr:
                        continue

                    # Apply font formatting
                    font = fmt_spec.get("font", {})
                    if font:
                        formatting_ops.set_font_format(
                            manager, workbook, worksheet, range_addr,
                            font.get("name"), font.get("size"), font.get("bold"),
                            font.get("italic"), font.get("underline"), font.get("color")
                        )

                    # Apply fill/background
                    fill = fmt_spec.get("fill", {})
                    if fill.get("color"):
                        formatting_ops.set_background_color(
                            manager, workbook, worksheet, range_addr, fill["color"]
                        )

                    # Apply border
                    border = fmt_spec.get("border", {})
                    if border:
                        formatting_ops.set_border_format(
                            manager, workbook, worksheet, range_addr,
                            border.get("edge", "all"), border.get("style"),
                            border.get("weight"), border.get("color")
                        )

                    # Apply alignment and number format
                    alignment = fmt_spec.get("alignment", {})
                    number_format = fmt_spec.get("number_format")
                    if alignment or number_format:
                        formatting_ops.set_cell_format(
                            manager, workbook, worksheet, range_addr,
                            alignment.get("horizontal"), alignment.get("vertical"),
                            alignment.get("wrap_text"), number_format
                        )

            return _ok(
                "Format applied successfully.",
                {
                    "workbook_path": path,
                    "worksheet": worksheet,
                    "ranges_formatted": len(ranges),
                },
            )

        elif target == "formula":
            worksheet = p.get("worksheet")
            range_address = p.get("range")
            formula = p.get("formula")

            if not worksheet or not range_address or not formula:
                return _err("Missing parameter.", "formula update requires 'worksheet', 'range', and 'formula' parameters")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                formula_ops.set_formula(manager, workbook, worksheet, range_address, formula)

            return _ok(
                "Formula set successfully.",
                {
                    "workbook_path": path,
                    "worksheet": worksheet,
                    "range": range_address,
                    "formula": formula,
                },
            )

        elif target == "worksheet":
            name = p.get("name")
            new_name = p.get("new_name")
            copy_to = p.get("copy_to")

            if not name:
                return _err("Missing parameter.", "worksheet update requires 'name' parameter")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                if copy_to:
                    new_sheet = workbook_ops.copy_worksheet(manager, workbook, name, copy_to)
                    return _ok(
                        f"Worksheet '{name}' copied to '{new_sheet.Name}'.",
                        {
                            "workbook_path": path,
                            "source_worksheet": name,
                            "new_worksheet": new_sheet.Name,
                        },
                    )
                elif new_name:
                    workbook_ops.rename_worksheet(manager, workbook, name, new_name)
                    return _ok(
                        f"Worksheet renamed from '{name}' to '{new_name}'.",
                        {
                            "workbook_path": path,
                            "old_name": name,
                            "new_name": new_name,
                        },
                    )
                else:
                    return _err("Missing parameter.", "worksheet update requires 'new_name' or 'copy_to' parameter")

        elif target == "structure":
            worksheet = p.get("worksheet")
            columns = p.get("columns", {})
            rows = p.get("rows", {})

            if not worksheet:
                return _err("Missing parameter.", "structure update requires 'worksheet' parameter")

            if not columns and not rows:
                return _err("Missing parameter.", "structure update requires 'columns' or 'rows' parameter")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                # Process columns
                for col_spec, width in columns.items():
                    if width == "auto":
                        formatting_ops.auto_fit_columns(manager, workbook, worksheet, f"{col_spec}:{col_spec}" if ":" not in col_spec else col_spec)
                    else:
                        formatting_ops.set_column_width(manager, workbook, worksheet, col_spec, float(width))

                # Process rows
                for row_spec, height in rows.items():
                    formatting_ops.set_row_height(manager, workbook, worksheet, row_spec, float(height))

            return _ok(
                "Structure updated successfully.",
                {
                    "workbook_path": path,
                    "worksheet": worksheet,
                    "columns_updated": list(columns.keys()) if columns else [],
                    "rows_updated": list(rows.keys()) if rows else [],
                },
            )

        else:
            return _err(
                "Invalid target.",
                f"target must be one of: range, format, formula, worksheet, structure. Got: {target}"
            )

    except Exception as e:
        return _err(f"Failed to update {target}.", str(e))


# =============================================================================
# CRUD: Delete (1)
# =============================================================================

@tool
def excel_delete(filepath: str, target: str, params: str) -> str:
    """Delete resources or clear content in an Excel workbook.

    Args:
        filepath: Path to the Excel file
        target: Type of resource to delete ("worksheet", "range", "rows", "columns", "format")
        params: JSON string with deletion parameters

    Returns:
        JSON with deletion result.

    Target-specific params:
        worksheet:
            - name (str): Name of the worksheet to delete
        range:
            - worksheet (str): Worksheet name
            - range (str): Range to clear (e.g., "A1:D10")
        rows:
            - worksheet (str): Worksheet name
            - index (int): Starting row number (1-based)
            - count (int, optional): Number of rows to delete (default: 1)
        columns:
            - worksheet (str): Worksheet name
            - index (str): Starting column letter (e.g., "A", "C")
            - count (int, optional): Number of columns to delete (default: 1)
        format:
            - worksheet (str): Worksheet name
            - range (str): Range to clear formatting from

    Examples:
        - Delete worksheet: excel_delete("data.xlsx", "worksheet", '{"name": "Sheet2"}')
        - Clear range: excel_delete("data.xlsx", "range", '{"worksheet": "Sheet1", "range": "A1:D10"}')
        - Delete rows: excel_delete("data.xlsx", "rows", '{"worksheet": "Sheet1", "index": 5, "count": 3}')
        - Clear format: excel_delete("data.xlsx", "format", '{"worksheet": "Sheet1", "range": "A1:D10"}')
    """
    try:
        p = json.loads(params)
    except json.JSONDecodeError as e:
        return _err("Invalid params.", f"params must be valid JSON: {e}")

    try:
        if target == "worksheet":
            name = p.get("name")

            if not name:
                return _err("Missing parameter.", "worksheet deletion requires 'name' parameter")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                workbook_ops.delete_worksheet(manager, workbook, name)

            return _ok(
                f"Worksheet '{name}' deleted.",
                {"workbook_path": path, "deleted_worksheet": name},
            )

        elif target == "range":
            worksheet = p.get("worksheet")
            range_address = p.get("range")

            if not worksheet or not range_address:
                return _err("Missing parameter.", "range deletion requires 'worksheet' and 'range' parameters")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                workbook_ops.clear_range(manager, workbook, worksheet, range_address)

            return _ok(
                f"Range '{range_address}' cleared.",
                {"workbook_path": path, "worksheet": worksheet, "cleared_range": range_address},
            )

        elif target == "rows":
            worksheet = p.get("worksheet")
            index = p.get("index")
            count = p.get("count", 1)

            if not worksheet or index is None:
                return _err("Missing parameter.", "rows deletion requires 'worksheet' and 'index' parameters")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                workbook_ops.delete_rows(manager, workbook, worksheet, index, count)

            return _ok(
                f"Deleted {count} row(s) starting at row {index}.",
                {"workbook_path": path, "worksheet": worksheet, "start_index": index, "count": count},
            )

        elif target == "columns":
            worksheet = p.get("worksheet")
            index = p.get("index")
            count = p.get("count", 1)

            if not worksheet or not index:
                return _err("Missing parameter.", "columns deletion requires 'worksheet' and 'index' parameters")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                workbook_ops.delete_columns(manager, workbook, worksheet, index, count)

            return _ok(
                f"Deleted {count} column(s) starting at column {index}.",
                {"workbook_path": path, "worksheet": worksheet, "start_index": index, "count": count},
            )

        elif target == "format":
            worksheet = p.get("worksheet")
            range_address = p.get("range")

            if not worksheet or not range_address:
                return _err("Missing parameter.", "format deletion requires 'worksheet' and 'range' parameters")

            with workbook_context(filepath, read_only=False, save_on_close=True) as (manager, workbook, path):
                formatting_ops.clear_format(manager, workbook, worksheet, range_address)

            return _ok(
                f"Format cleared from range '{range_address}'.",
                {"workbook_path": path, "worksheet": worksheet, "cleared_range": range_address},
            )

        else:
            return _err(
                "Invalid target.",
                f"target must be one of: worksheet, range, rows, columns, format. Got: {target}"
            )

    except Exception as e:
        return _err(f"Failed to delete {target}.", str(e))


# =============================================================================
# Export all tools
# =============================================================================

EXCEL_TOOLS = [
    # Workbook lifecycle
    excel_open,       # Open/create workbook
    excel_save,       # Save workbook
    excel_info,       # Get Excel status
    # CRUD operations
    excel_create,     # Create: worksheet, rows, columns
    excel_read,       # Read: worksheets, range, format, formula, used_range
    excel_update,     # Update: range, format, formula, worksheet, structure
    excel_delete,     # Delete: worksheet, range, rows, columns, format
]
