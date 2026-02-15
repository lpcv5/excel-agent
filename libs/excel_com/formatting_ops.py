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


# =============================================================================
# Reverse Mappings for get_format
# =============================================================================

# Reverse lookup dictionaries (COM value -> string)
HORIZONTAL_ALIGNMENT_REVERSE = {
    -4131: "left",    # XL_LEFT
    -4108: "center",  # XL_CENTER
    -4152: "right",   # XL_RIGHT
    1: "general",     # XL_GENERAL
}

VERTICAL_ALIGNMENT_REVERSE = {
    -4160: "top",      # XL_TOP
    -4107: "bottom",   # XL_BOTTOM
    -4130: "justify",  # XL_JUSTIFY
    -4108: "center",   # XL_CENTER_V
}

BORDER_STYLE_REVERSE = {
    1: "continuous",     # XL_CONTINUOUS
    -4115: "dash",       # XL_DASH
    -4118: "dot",        # XL_DOT
    -4119: "double",     # XL_DOUBLE
    -4142: "none",       # XL_NONE
    -4135: "slant_dash_dot",  # XL_SLANT_DASH_DOT
}

BORDER_WEIGHT_REVERSE = {
    1: "hairline",  # XL_HAIRLINE
    2: "thin",      # XL_THIN
    -4138: "medium",  # XL_MEDIUM
    4: "thick",     # XL_THICK
}

UNDERLINE_STYLE_REVERSE = {
    2: True,     # XL_UNDERLINE_STYLE_SINGLE
    -4142: False,  # XL_UNDERLINE_STYLE_NONE
}


# =============================================================================
# Format Query Operations
# =============================================================================

def get_format(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str
) -> dict:
    """Get formatting information for a range.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range address

    Returns:
        Dictionary containing formatting information:
            - font: {name, size, bold, italic, underline, color}
            - fill: {color}
            - alignment: {horizontal, vertical, wrap_text}
            - border: {style, weight, color}
            - number_format: str
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)

    result = {}

    # Font info
    font = range_obj.Font
    font_info = {
        "name": font.Name,
        "size": font.Size,
        "bold": font.Bold,
        "italic": font.Italic,
        "underline": UNDERLINE_STYLE_REVERSE.get(font.Underline, False),
    }
    # Convert color from int to hex (if not automatic)
    try:
        if font.Color != -16776961:  # xlColorIndexAutomatic
            font_info["color"] = f"{int(font.Color) & 0xFFFFFF:06X}"
    except Exception:
        pass
    result["font"] = font_info

    # Fill/background info
    interior = range_obj.Interior
    try:
        if interior.ColorIndex != -4142:  # xlNone
            fill_info = {"color": f"{int(interior.Color) & 0xFFFFFF:06X}"}
            result["fill"] = fill_info
    except Exception:
        pass

    # Alignment info
    alignment_info = {}
    h_align = range_obj.HorizontalAlignment
    if h_align in HORIZONTAL_ALIGNMENT_REVERSE:
        alignment_info["horizontal"] = HORIZONTAL_ALIGNMENT_REVERSE[h_align]

    v_align = range_obj.VerticalAlignment
    if v_align in VERTICAL_ALIGNMENT_REVERSE:
        alignment_info["vertical"] = VERTICAL_ALIGNMENT_REVERSE[v_align]

    alignment_info["wrap_text"] = range_obj.WrapText
    result["alignment"] = alignment_info

    # Border info (check all edges)
    borders = range_obj.Borders
    try:
        # Get the overall border style if all edges are the same
        line_style = borders.LineStyle
        if line_style != -4142:  # xlNone
            border_info = {
                "style": BORDER_STYLE_REVERSE.get(line_style, "continuous"),
            }
            weight = borders.Weight
            if weight in BORDER_WEIGHT_REVERSE:
                border_info["weight"] = BORDER_WEIGHT_REVERSE[weight]
            try:
                border_info["color"] = f"{int(borders.Color) & 0xFFFFFF:06X}"
            except Exception:
                pass
            result["border"] = border_info
    except Exception:
        pass

    # Number format
    number_format = range_obj.NumberFormat
    if number_format and number_format != "General":
        result["number_format"] = number_format

    return result


def clear_format(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str
) -> None:
    """Clear all formatting from a range (keeps contents).

    This operation preserves the user's Excel view state (active sheet,
    selection, scroll position) to minimize disruption.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range address to clear formatting from
    """
    with preserve_user_state(manager.app):
        worksheet = manager.get_worksheet(workbook, worksheet_name)
        worksheet.Range(range_address).ClearFormats()
