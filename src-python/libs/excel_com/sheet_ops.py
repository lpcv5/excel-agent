"""Worksheet CRUD operations for Excel COM."""

import logging
from typing import Any

from .errors import SheetNotFoundError
from .utils import com_retry, pump_messages

logger = logging.getLogger("app.excel")


def list_sheets(wb: Any) -> list[str]:
    """Return all sheet names in the workbook."""
    return [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]


@com_retry(max_retries=2, delay=0.5)
def add_sheet(wb: Any, name: str, position: int | None = None) -> str:
    """Add a new worksheet. Returns the actual sheet name."""
    if position is not None:
        after = wb.Sheets(min(position, wb.Sheets.Count))
        ws = wb.Sheets.Add(After=after)
    else:
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    pump_messages()
    logger.info("Added sheet: %s", name)
    return ws.Name


@com_retry(max_retries=2, delay=0.5)
def delete_sheet(app: Any, wb: Any, name: str) -> None:
    """Delete a worksheet by name."""
    try:
        ws = wb.Sheets(name)
    except Exception as e:
        raise SheetNotFoundError(f"工作表 '{name}' 不存在", e)
    prev_alerts = app.DisplayAlerts
    app.DisplayAlerts = False
    try:
        ws.Delete()
        pump_messages()
    finally:
        app.DisplayAlerts = prev_alerts
    logger.info("Deleted sheet: %s", name)


@com_retry(max_retries=2, delay=0.5)
def rename_sheet(wb: Any, old_name: str, new_name: str) -> None:
    """Rename a worksheet."""
    try:
        ws = wb.Sheets(old_name)
    except Exception as e:
        raise SheetNotFoundError(f"工作表 '{old_name}' 不存在", e)
    ws.Name = new_name
    pump_messages()
    logger.info("Renamed sheet: %s -> %s", old_name, new_name)


@com_retry(max_retries=2, delay=0.5)
def copy_sheet(wb: Any, source_name: str, new_name: str) -> str:
    """Copy a worksheet. Returns the new sheet name."""
    try:
        ws = wb.Sheets(source_name)
    except Exception as e:
        raise SheetNotFoundError(f"工作表 '{source_name}' 不存在", e)
    ws.Copy(After=wb.Sheets(wb.Sheets.Count))
    new_ws = wb.Sheets(wb.Sheets.Count)
    new_ws.Name = new_name
    pump_messages()
    logger.info("Copied sheet: %s -> %s", source_name, new_name)
    return new_ws.Name
