"""
Formula Operations - Excel formula operations.

This module provides functions for working with Excel formulas,
including setting and getting cell formulas.

IMPORTANT: set_formula uses the preserve_user_state context manager
to minimize disruption to the user's Excel view state.
"""

from .manager import ExcelAppManager
from .context import preserve_user_state


def set_formula(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    formula: str
) -> None:
    """Set a formula for a cell or range.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Cell or range address
        formula: Excel formula (must start with "=")
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        range_obj = worksheet.Range(range_address)
        range_obj.Formula = formula


def get_formula(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str
) -> str:
    """Get the formula for a cell.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Cell address

    Returns:
        The formula string (including "="), or empty string if no formula
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)
    formula = range_obj.Formula

    # Handle single cell vs range
    if isinstance(formula, (list, tuple)):
        formula = formula[0][0] if formula and formula[0] else ""

    return formula or ""
