"""Formula read/write operations for Excel COM."""

import logging
from typing import Any

from .errors import FormulaError
from .utils import com_retry, pump_messages

logger = logging.getLogger("app.excel")


@com_retry(max_retries=2, delay=0.5)
def get_formula(sheet: Any, range_address: str) -> list[list[str]]:
    """Read formulas from a range. Returns 2D list of formula strings."""
    rng = sheet.Range(range_address)
    raw = rng.Formula
    if raw is None:
        return [[""]]
    if not isinstance(raw, tuple):
        return [[str(raw) if raw else ""]]
    result = []
    for row in raw:
        if isinstance(row, tuple):
            result.append([str(c) if c else "" for c in row])
        else:
            result.append([str(row) if row else ""])
    return result


@com_retry(max_retries=2, delay=0.5)
def set_formula(sheet: Any, range_address: str, formula: str) -> None:
    """Set a formula on a range."""
    try:
        rng = sheet.Range(range_address)
        rng.Formula = formula
        pump_messages()
    except Exception as e:
        err_msg = str(e)
        if "0x800a03ec" in err_msg.lower():
            raise FormulaError(f"公式语法错误 '{formula}' 在 {range_address}: {e}", e)
        raise FormulaError(f"设置公式失败 '{formula}' 在 {range_address}: {e}", e)
