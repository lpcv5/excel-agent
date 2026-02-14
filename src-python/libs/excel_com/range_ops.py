"""Range read/write operations for Excel COM."""

import logging
from datetime import datetime
from typing import Any

import pywintypes

from .errors import RangeError
from .utils import com_retry, pump_messages

logger = logging.getLogger("app.excel")

MAX_RETURN_ROWS = 100


def _convert_value(val: Any) -> Any:
    """Convert COM values to Python-friendly types."""
    if val is None:
        return None
    if isinstance(val, pywintypes.TimeType):
        return datetime.fromtimestamp(val.timestamp()).strftime("%Y-%m-%d")
    if isinstance(val, float) and val == int(val):
        return int(val)
    return val


def _values_to_list(raw: Any) -> list[list[Any]]:
    """Convert COM Range.Value result to a 2D list."""
    if raw is None:
        return [[None]]
    if not isinstance(raw, tuple):
        return [[_convert_value(raw)]]
    result = []
    for row in raw:
        if isinstance(row, tuple):
            result.append([_convert_value(c) for c in row])
        else:
            result.append([_convert_value(row)])
    return result


@com_retry(max_retries=2, delay=0.5)
def read_range(sheet: Any, range_address: str) -> list[list[Any]]:
    """Read a range and return as 2D list."""
    try:
        rng = sheet.Range(range_address)
    except Exception as e:
        raise RangeError(f"无效的区域地址 '{range_address}': {e}", e)
    raw = rng.Value
    return _values_to_list(raw)


def read_used_range(sheet: Any) -> tuple[str, list[list[Any]], int, int]:
    """Read the used range. Returns (address, data, total_rows, total_cols)."""
    ur = sheet.UsedRange
    if ur is None:
        return ("A1", [[]], 0, 0)
    addr = ur.Address
    total_rows = ur.Rows.Count
    total_cols = ur.Columns.Count
    raw = ur.Value
    data = _values_to_list(raw)
    return (addr, data, total_rows, total_cols)


@com_retry(max_retries=2, delay=0.5)
def write_range(sheet: Any, start_cell: str, values: list[list[Any]]) -> str:
    """Write a 2D list to a range starting at start_cell. Returns the written range address."""
    if not values or not values[0]:
        raise RangeError("写入数据不能为空")
    rows = len(values)
    cols = len(values[0])
    try:
        start = sheet.Range(start_cell)
        end = start.Offset(rows, cols)
        target = sheet.Range(start, end.Offset(-1, -1) if rows > 0 else start)
        # Build proper range from start to end
        r1 = start.Row
        c1 = start.Column
        target = sheet.Range(
            sheet.Cells(r1, c1),
            sheet.Cells(r1 + rows - 1, c1 + cols - 1),
        )
        target.Value = values
        pump_messages()
        return target.Address
    except Exception as e:
        raise RangeError(f"写入区域 '{start_cell}' 失败: {e}", e)


@com_retry(max_retries=2, delay=0.5)
def auto_fit_columns(sheet: Any, range_address: str | None = None) -> None:
    """Auto-fit column widths."""
    if range_address:
        sheet.Range(range_address).Columns.AutoFit()
    else:
        sheet.UsedRange.Columns.AutoFit()
    pump_messages()


@com_retry(max_retries=2, delay=0.5)
def set_number_format(sheet: Any, range_address: str, fmt: str) -> None:
    """Set number format on a range."""
    sheet.Range(range_address).NumberFormat = fmt
    pump_messages()
