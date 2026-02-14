"""Table (ListObject) operations for Excel COM."""

import logging
import time
from typing import Any

import pythoncom

from .constants import xlSrcRange, xlYes
from .errors import TableError
from .utils import com_retry, pump_messages

logger = logging.getLogger("app.excel")


@com_retry(max_retries=2, delay=0.5)
def create_table(
    sheet: Any, range_address: str, name: str, has_headers: bool = True
) -> Any:
    """Create a ListObject (table) from a range."""
    try:
        src = sheet.Range(range_address)
        tbl = sheet.ListObjects.Add(
            SourceType=xlSrcRange,
            Source=src,
            XlListObjectHasHeaders=xlYes if has_headers else 2,
        )
        tbl.Name = name
        pump_messages()
        pythoncom.PumpWaitingMessages()
        logger.info("Created table '%s' at %s", name, range_address)
        return tbl
    except Exception as e:
        raise TableError(f"创建表格 '{name}' 失败: {e}", e)


def list_tables(sheet: Any) -> list[dict[str, str]]:
    """List all tables on a sheet."""
    result = []
    for i in range(1, sheet.ListObjects.Count + 1):
        tbl = sheet.ListObjects(i)
        result.append(
            {
                "name": tbl.Name,
                "range": tbl.Range.Address,
                "rows": tbl.ListRows.Count,
                "columns": tbl.ListColumns.Count,
            }
        )
    return result


@com_retry(max_retries=2, delay=0.5)
def set_table_style(sheet: Any, table_name: str, style_name: str) -> None:
    """Set a table style."""
    tbl = _get_table(sheet, table_name)
    tbl.TableStyle = style_name
    pump_messages()


@com_retry(max_retries=2, delay=0.5)
def add_table_column(
    sheet: Any,
    table_name: str,
    column_name: str,
    formula: str | None = None,
    position: int | None = None,
) -> None:
    """Add a column to a table, optionally with a formula."""
    tbl = _get_table(sheet, table_name)
    try:
        if position is not None:
            col = tbl.ListColumns.Add(Position=position)
        else:
            col = tbl.ListColumns.Add()
        col.Name = column_name
        pump_messages()
        pythoncom.PumpWaitingMessages()
        time.sleep(0.3)

        if formula and tbl.ListRows.Count > 0:
            # Re-fetch column reference after pump
            col = tbl.ListColumns(column_name)
            try:
                col.DataBodyRange.Formula = formula
            except Exception:
                # Fallback: set cell by cell
                for cell in col.DataBodyRange:
                    cell.Formula = formula
            pump_messages()
        logger.info("Added column '%s' to table '%s'", column_name, table_name)
    except Exception as e:
        raise TableError(f"添加列 '{column_name}' 到表格 '{table_name}' 失败: {e}", e)


@com_retry(max_retries=2, delay=0.5)
def delete_table(sheet: Any, table_name: str) -> None:
    """Delete (unlist) a table, keeping the data."""
    tbl = _get_table(sheet, table_name)
    tbl.Unlist()
    pump_messages()
    logger.info("Deleted table: %s", table_name)


def _get_table(sheet: Any, table_name: str) -> Any:
    """Find a table by name on a sheet."""
    for i in range(1, sheet.ListObjects.Count + 1):
        tbl = sheet.ListObjects(i)
        if tbl.Name == table_name:
            return tbl
    available = [
        sheet.ListObjects(i).Name for i in range(1, sheet.ListObjects.Count + 1)
    ]
    raise TableError(f"表格 '{table_name}' 不存在，可用表格: {available}")
