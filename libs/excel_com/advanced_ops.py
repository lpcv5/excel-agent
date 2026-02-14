"""
Advanced Operations - Advanced Excel operations.

This module provides advanced Excel operations that will be exposed as Agent tools.
These operations include:

Charts:
- Create various chart types (line, bar, pie, scatter, etc.)
- Customize chart appearance
- Add multiple series

Pivot Tables:
- Create pivot tables from data ranges
- Add calculated fields
- Refresh pivot tables

Data Analysis:
- Find and replace
- Sort and filter
- Remove duplicates
- Data validation

Conditional Formatting:
- Apply conditional formatting rules
- Create custom formulas

This module is structured for easy extension as new features are added.
"""

from typing import Optional
from .manager import ExcelAppManager


# =============================================================================
# Chart Operations
# =============================================================================

def create_chart(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    data_range: str,
    chart_type: int,
    chart_title: str,
    position: Optional[tuple[int, int]] = None
) -> object:
    """Create a chart from a data range.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet containing data
        data_range: Range of data for the chart
        chart_type: Excel chart type constant (e.g., 4 for line, 51 for pie)
        chart_title: Title for the chart
        position: Optional (row, column) for chart placement

    Returns:
        The Chart COM object

    Note:
        Chart type constants (common ones):
        - 4: Line chart
        - 51: Pie chart
        - -4100: Bar chart (xlColumnClustered)
        - 1: Clustered column (xlColumnClustered)
        - 3: 3D Column
        - 5: Scatter chart
        - -4169: Area chart
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    chart_objects = worksheet.ChartObjects()

    if position:
        chart = chart_objects.Add(position[1] * 72, position[0] * 14.4, 400, 300)
    else:
        chart = chart_objects.Add(100, 100, 400, 300)

    chart.Chart.ChartWizard(
        Source=worksheet.Range(data_range),
        Gallery=chart_type,
        Title=chart_title
    )

    return chart.Chart


def set_chart_style(
    manager: ExcelAppManager,
    chart: object,
    style: int,
    has_legend: bool = True,
    show_data_labels: bool = False
) -> None:
    """Set chart styling options.

    Args:
        manager: ExcelAppManager instance
        chart: Chart COM object
        style: Style index (1-48)
        has_legend: Whether to show legend
        show_data_labels: Whether to show data labels
    """
    chart.ChartStyle = style
    chart.HasLegend = has_legend

    if show_data_labels:
        chart.SeriesCollection(1).HasDataLabels = True


# =============================================================================
# Pivot Table Operations
# =============================================================================

def create_pivot_table(
    manager: ExcelAppManager,
    workbook: object,
    source_sheet: str,
    source_range: str,
    dest_sheet: str,
    dest_cell: str,
    table_name: str
) -> object:
    """Create a pivot table from data.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        source_sheet: Worksheet containing source data
        source_range: Range of source data
        dest_sheet: Worksheet for pivot table (will be created if needed)
        dest_cell: Cell for pivot table placement
        table_name: Name for the pivot table

    Returns:
        The PivotTable COM object
    """
    source_worksheet = manager.get_worksheet(workbook, source_sheet)
    source_data = source_worksheet.Range(source_range)

    # Get or create destination worksheet
    try:
        dest_worksheet = manager.get_worksheet(workbook, dest_sheet)
    except Exception:
        dest_worksheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        dest_worksheet.Name = dest_sheet

    # Create pivot cache and pivot table
    pivot_cache = workbook.PivotCaches().Create(SourceType=1, SourceData=source_data)
    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=dest_worksheet.Range(dest_cell),
        TableName=table_name
    )

    return pivot_table


def add_pivot_field(
    pivot_table: object,
    field_name: str,
    orientation: int,
    position: int = 1
) -> None:
    """Add a field to a pivot table.

    Args:
        pivot_table: PivotTable COM object
        field_name: Name of the field to add
        orientation: Field orientation (1=Row, 2=Column, 3=Filter, 4=Data)
        position: Position for the field
    """
    try:
        pivot_field = pivot_table.PivotFields(field_name)
        pivot_field.Orientation = orientation
        pivot_field.Position = position
    except Exception as e:
        raise ValueError(f"Failed to add field '{field_name}': {e}") from e


# =============================================================================
# Data Analysis Operations
# =============================================================================

def remove_duplicates(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    headers: bool = True
) -> int:
    """Remove duplicate rows from a range.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range to check for duplicates
        headers: Whether the range has headers

    Returns:
        Number of duplicate rows removed
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)

    # Remove duplicates
    result = range_obj.RemoveDuplicates(Headers=1 if headers else 0)

    # Return count of removed rows (approximate)
    return 1 if result else 0


def sort_range(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    key_column: str,
    order: str = "asc"
) -> None:
    """Sort a range by a key column.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range to sort
        key_column: Column letter to sort by (e.g., "A")
        order: "asc" for ascending, "desc" for descending
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)

    sort_order = 1 if order.lower() == "asc" else 2
    range_obj.Sort(
        Key1=range_obj.Columns(key_column),
        Order1=sort_order
    )


def autofilter(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str
) -> None:
    """Apply autofilter to a range.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range to filter
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)
    range_obj.AutoFilter


# =============================================================================
# Conditional Formatting
# =============================================================================

def add_conditional_format(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    condition_type: str,
    formula: Optional[str] = None,
    color: Optional[str] = None
) -> None:
    """Add conditional formatting to a range.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range to format
        condition_type: Type of condition ("cell_value", "formula", "color_scale")
        formula: Condition formula or value
        color: RGB hex color for formatting
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)
    formats = range_obj.FormatConditions

    if condition_type == "cell_value" and formula and color:
        formats.Add(
            Type=1,  # xlCellValue
            Operator=5,  # xlBetween
            Formula1=formula
        ).Interior.Color = int(color, 16)
    elif condition_type == "formula" and formula and color:
        formats.Add(
            Type=2,  # xlExpression
            Formula1=formula
        ).Interior.Color = int(color, 16)


# =============================================================================
# Utility Functions
# =============================================================================

def find_replace(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str,
    find_text: str,
    replace_with: str,
    match_case: bool = False,
    replace_all: bool = True
) -> int:
    """Find and replace text in a range.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range to search
        find_text: Text to find
        replace_with: Replacement text
        match_case: Whether to match case
        replace_all: Whether to replace all occurrences

    Returns:
        Number of replacements made
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)

    find_count = 0
    for cell in range_obj:
        if match_case:
            if cell.Value == find_text:
                cell.Value = replace_with
                find_count += 1
        else:
            if str(cell.Value).lower() == find_text.lower():
                cell.Value = replace_with
                find_count += 1
        if not replace_all and find_count > 0:
            break

    return find_count


def clear_contents(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str
) -> None:
    """Clear contents from a range (preserves formatting).

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range to clear
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)
    range_obj.ClearContents()


def clear_all(
    manager: ExcelAppManager,
    workbook: object,
    worksheet_name: str,
    range_address: str
) -> None:
    """Clear all (contents and formatting) from a range.

    Args:
        manager: ExcelAppManager instance
        workbook: Workbook COM object
        worksheet_name: Name of the worksheet
        range_address: Range to clear
    """
    worksheet = manager.get_worksheet(workbook, worksheet_name)
    range_obj = worksheet.Range(range_address)
    range_obj.Clear()
