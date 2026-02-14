"""
Formatting Operations - Excel cell and range formatting operations.

This module provides functions for formatting Excel cells and ranges,
including fonts, borders, alignment, and colors.

These operations will be exposed as Agent tools.

IMPORTANT: All formatting operations use the preserve_user_state context
manager to minimize disruption to the user's Excel view state.
"""

from typing import Optional

from .manager import ExcelAppManager
from .context import preserve_user_state
from .constants import (
    XL_UNDERLINE_STYLE_SINGLE,
    XL_UNDERLINE_STYLE_NONE,
    HORIZONTAL_ALIGNMENT_MAP,
    VERTICAL_ALIGNMENT_MAP,
    BORDER_STYLE_MAP,
    BORDER_EDGE_MAP,
)


# =============================================================================
# Font Formatting
# =============================================================================

def set_font_format(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    font_name: Optional[str] = None,
    size: Optional[int] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None
) -> None:
    """Set font formatting for a range.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range address
        font_name: Font name (e.g., "Arial", "Calibri")
        size: Font size in points
        bold: Whether text is bold
        italic: Whether text is italic
        underline: Whether text is underlined
        color: RGB hex color (e.g., "FF0000" for red)
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        range_obj = worksheet.Range(range_address)
        font = range_obj.Font

        if font_name is not None:
            font.Name = font_name
        if size is not None:
            font.Size = size
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if underline is not None:
            font.Underline = XL_UNDERLINE_STYLE_SINGLE if underline else XL_UNDERLINE_STYLE_NONE
        if color is not None:
            font.Color = int(color, 16)


# =============================================================================
# Cell Formatting
# =============================================================================

def set_cell_format(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    horizontal_alignment: Optional[str] = None,
    vertical_alignment: Optional[str] = None,
    wrap_text: Optional[bool] = None,
    number_format: Optional[str] = None
) -> None:
    """Set cell formatting for a range.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range address
        horizontal_alignment: "left", "center", "right", or "general"
        vertical_alignment: "top", "center", "bottom", or "justify"
        wrap_text: Whether to wrap text
        number_format: Excel number format code
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        range_obj = worksheet.Range(range_address)

        if horizontal_alignment is not None:
            range_obj.HorizontalAlignment = HORIZONTAL_ALIGNMENT_MAP[horizontal_alignment.lower()]

        if vertical_alignment is not None:
            range_obj.VerticalAlignment = VERTICAL_ALIGNMENT_MAP[vertical_alignment.lower()]

        if wrap_text is not None:
            range_obj.WrapText = wrap_text

        if number_format is not None:
            range_obj.NumberFormat = number_format


# =============================================================================
# Border Formatting
# =============================================================================

def set_border_format(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    edge: str,
    style: Optional[str] = None,
    weight: Optional[int] = None,
    color: Optional[str] = None
) -> None:
    """Set border formatting for a range.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range address
        edge: "left", "right", "top", "bottom", or "all"
        style: Border style ("continuous", "dash", "dot", "double")
        weight: Border weight (1-4)
        color: RGB hex color
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        range_obj = worksheet.Range(range_address)
        borders = range_obj.Borders

        if edge.lower() == "all":
            # Set all edges
            if style is not None:
                borders.LineStyle = BORDER_STYLE_MAP.get(style, 1)
            if weight is not None:
                borders.Weight = weight
            if color is not None:
                borders.Color = int(color, 16)
        else:
            # Set specific edge
            edge_const = BORDER_EDGE_MAP[edge.lower()]
            border = borders.Item(edge_const)
            if style is not None:
                border.LineStyle = BORDER_STYLE_MAP[style]
            if weight is not None:
                border.Weight = weight
            if color is not None:
                border.Color = int(color, 16)


# =============================================================================
# Background Color
# =============================================================================

def set_background_color(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    color: str
) -> None:
    """Set background color for a range.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range address
        color: RGB hex color (e.g., "FFFF00" for yellow)
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        range_obj = worksheet.Range(range_address)
        range_obj.Interior.Color = int(color, 16)


# =============================================================================
# Column/Row Operations
# =============================================================================

def auto_fit_columns(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: Optional[str] = None
) -> None:
    """Auto-fit column widths.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Optional column range (e.g., "A:D")
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)

        if range_address:
            range_obj = worksheet.Range(range_address)
            range_obj.Columns.AutoFit()
        else:
            worksheet.UsedRange.Columns.AutoFit()


def set_column_width(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    columns: str,
    width: float
) -> None:
    """Set column width.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        columns: Column specification (e.g., "A", "A:C")
        width: Width in points
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        range_obj = worksheet.Range(columns)
        range_obj.ColumnWidth = width


def set_row_height(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    rows: str,
    height: float
) -> None:
    """Set row height.

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        rows: Row specification (e.g., "1", "1:5")
        height: Height in points
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        range_obj = worksheet.Range(rows)
        range_obj.RowHeight = height
