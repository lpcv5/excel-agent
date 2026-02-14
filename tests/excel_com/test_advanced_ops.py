"""Tests for excel_com/advanced_ops.py."""

from unittest.mock import MagicMock

import pytest

from libs.excel_com.manager import ExcelAppManager
from libs.excel_com import advanced_ops


class TestCreateChart:
    """Tests for create_chart function."""

    def test_create_chart(self):
        """Test create_chart creates a chart."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_chart_objects = MagicMock()
        mock_chart = MagicMock()
        mock_chart_objects.Add.return_value = mock_chart

        mock_worksheet.ChartObjects.return_value = mock_chart_objects
        manager.get_worksheet.return_value = mock_worksheet

        result = advanced_ops.create_chart(
            manager, mock_workbook, "Sheet1", "A1:D10",
            chart_type=4,  # Line chart
            chart_title="Test Chart"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        mock_chart_objects.Add.assert_called_once()
        mock_chart.Chart.ChartWizard.assert_called_once()
        assert result == mock_chart.Chart

    def test_create_chart_with_position(self):
        """Test create_chart with custom position."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_chart_objects = MagicMock()
        mock_chart = MagicMock()
        mock_chart_objects.Add.return_value = mock_chart

        mock_worksheet.ChartObjects.return_value = mock_chart_objects
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.create_chart(
            manager, mock_workbook, "Sheet1", "A1:D10",
            chart_type=51,  # Pie chart
            chart_title="Test Chart",
            position=(2, 3)  # row=2, col=3
        )

        # Position affects the Add call parameters
        mock_chart_objects.Add.assert_called_once()


class TestSetChartStyle:
    """Tests for set_chart_style function."""

    def test_set_chart_style(self):
        """Test set_chart_style sets style options."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_chart = MagicMock()
        mock_series = MagicMock()
        mock_series.HasDataLabels = False
        mock_chart.SeriesCollection.return_value = mock_series

        advanced_ops.set_chart_style(
            manager, mock_chart, style=5, has_legend=True, show_data_labels=False
        )

        assert mock_chart.ChartStyle == 5
        assert mock_chart.HasLegend is True

    def test_set_chart_style_with_data_labels(self):
        """Test set_chart_style with data labels."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_chart = MagicMock()
        mock_series = MagicMock()
        mock_chart.SeriesCollection.return_value = mock_series

        advanced_ops.set_chart_style(
            manager, mock_chart, style=10, has_legend=False, show_data_labels=True
        )

        assert mock_chart.HasLegend is False
        assert mock_series.HasDataLabels is True


class TestCreatePivotTable:
    """Tests for create_pivot_table function."""

    def test_create_pivot_table(self):
        """Test create_pivot_table creates a pivot table."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_source_sheet = MagicMock()
        mock_dest_sheet = MagicMock()
        mock_source_range = MagicMock()
        mock_pivot_cache = MagicMock()
        mock_pivot_table = MagicMock()

        mock_source_sheet.Range.return_value = mock_source_range
        mock_dest_sheet.Range.return_value = MagicMock()

        manager.get_worksheet.side_effect = [mock_source_sheet, mock_dest_sheet]

        mock_pivot_caches = MagicMock()
        mock_pivot_caches.Create.return_value = mock_pivot_cache
        mock_workbook.PivotCaches.return_value = mock_pivot_caches

        mock_pivot_cache.CreatePivotTable.return_value = mock_pivot_table

        result = advanced_ops.create_pivot_table(
            manager, mock_workbook,
            source_sheet="Data", source_range="A1:D100",
            dest_sheet="Pivot", dest_cell="A1",
            table_name="PivotTable1"
        )

        assert result == mock_pivot_table

    def test_create_pivot_table_creates_dest_sheet(self):
        """Test create_pivot_table creates destination sheet if needed."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_source_sheet = MagicMock()
        mock_new_sheet = MagicMock()

        mock_source_range = MagicMock()
        mock_source_sheet.Range.return_value = mock_source_range

        # First call succeeds, second call fails (dest sheet doesn't exist)
        manager.get_worksheet.side_effect = [
            mock_source_sheet,
            Exception("Sheet not found")
        ]

        mock_workbook.Worksheets.Count = 1
        mock_workbook.Worksheets.Add.return_value = mock_new_sheet

        mock_pivot_cache = MagicMock()
        mock_pivot_table = MagicMock()
        mock_pivot_cache.CreatePivotTable.return_value = mock_pivot_table

        mock_pivot_caches = MagicMock()
        mock_pivot_caches.Create.return_value = mock_pivot_cache
        mock_workbook.PivotCaches.return_value = mock_pivot_caches

        advanced_ops.create_pivot_table(
            manager, mock_workbook,
            source_sheet="Data", source_range="A1:D100",
            dest_sheet="NewPivot", dest_cell="A1",
            table_name="PivotTable1"
        )

        mock_workbook.Worksheets.Add.assert_called_once()


class TestAddPivotField:
    """Tests for add_pivot_field function."""

    def test_add_pivot_field(self):
        """Test add_pivot_field adds field to pivot table."""
        mock_pivot_table = MagicMock()
        mock_field = MagicMock()
        mock_pivot_table.PivotFields.return_value = mock_field

        advanced_ops.add_pivot_field(
            mock_pivot_table,
            field_name="Category",
            orientation=1,  # Row
            position=1
        )

        mock_pivot_table.PivotFields.assert_called_once_with("Category")
        assert mock_field.Orientation == 1
        assert mock_field.Position == 1

    def test_add_pivot_field_raises_on_error(self):
        """Test add_pivot_field raises on invalid field."""
        mock_pivot_table = MagicMock()
        mock_pivot_table.PivotFields.side_effect = Exception("Field not found")

        with pytest.raises(ValueError, match="Failed to add field"):
            advanced_ops.add_pivot_field(
                mock_pivot_table,
                field_name="NonExistent",
                orientation=1
            )


class TestRemoveDuplicates:
    """Tests for remove_duplicates function."""

    def test_remove_duplicates(self):
        """Test remove_duplicates removes duplicate rows."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.RemoveDuplicates.return_value = True

        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        result = advanced_ops.remove_duplicates(
            manager, mock_workbook, "Sheet1", "A1:D100", headers=True
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        mock_worksheet.Range.assert_called_once_with("A1:D100")
        mock_range.RemoveDuplicates.assert_called_once_with(Headers=1)
        assert result == 1

    def test_remove_duplicates_without_headers(self):
        """Test remove_duplicates without headers."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.RemoveDuplicates.return_value = True

        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.remove_duplicates(
            manager, mock_workbook, "Sheet1", "A1:D100", headers=False
        )

        mock_range.RemoveDuplicates.assert_called_once_with(Headers=0)


class TestSortRange:
    """Tests for sort_range function."""

    def test_sort_range_ascending(self):
        """Test sort_range with ascending order."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_columns = MagicMock()
        mock_range.Columns.return_value = mock_columns

        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.sort_range(
            manager, mock_workbook, "Sheet1", "A1:D100", key_column="A", order="asc"
        )

        mock_range.Sort.assert_called_once()
        # Order1=1 is ascending

    def test_sort_range_descending(self):
        """Test sort_range with descending order."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_columns = MagicMock()
        mock_range.Columns.return_value = mock_columns

        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.sort_range(
            manager, mock_workbook, "Sheet1", "A1:D100", key_column="B", order="desc"
        )

        mock_range.Sort.assert_called_once()
        # Order1=2 is descending


class TestAutofilter:
    """Tests for autofilter function."""

    def test_autofilter(self):
        """Test autofilter applies filter to range."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()

        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.autofilter(
            manager, mock_workbook, "Sheet1", "A1:D100"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        mock_worksheet.Range.assert_called_once_with("A1:D100")


class TestAddConditionalFormat:
    """Tests for add_conditional_format function."""

    def test_add_conditional_format_cell_value(self):
        """Test add_conditional_format with cell_value type."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_format_conditions = MagicMock()
        mock_condition = MagicMock()

        mock_format_conditions.Add.return_value = mock_condition
        mock_range.FormatConditions = mock_format_conditions
        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.add_conditional_format(
            manager, mock_workbook, "Sheet1", "A1:A10",
            condition_type="cell_value",
            formula="100",
            color="FF0000"
        )

        mock_format_conditions.Add.assert_called_once()

    def test_add_conditional_format_formula(self):
        """Test add_conditional_format with formula type."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_format_conditions = MagicMock()
        mock_condition = MagicMock()

        mock_format_conditions.Add.return_value = mock_condition
        mock_range.FormatConditions = mock_format_conditions
        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.add_conditional_format(
            manager, mock_workbook, "Sheet1", "A1:A10",
            condition_type="formula",
            formula="=A1>100",
            color="00FF00"
        )

        mock_format_conditions.Add.assert_called_once()


class TestFindReplace:
    """Tests for find_replace function."""

    def test_find_replace_match_case(self):
        """Test find_replace with case matching."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()

        # Create mock cells
        mock_cell1 = MagicMock()
        mock_cell1.Value = "Hello"
        mock_cell2 = MagicMock()
        mock_cell2.Value = "hello"
        mock_cell3 = MagicMock()
        mock_cell3.Value = "Hello"

        mock_range.__iter__ = MagicMock(return_value=iter([mock_cell1, mock_cell2, mock_cell3]))
        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.find_replace(
            manager, mock_workbook, "Sheet1", "A1:A3",
            find_text="Hello",
            replace_with="Hi",
            match_case=True,
            replace_all=True
        )

        # Only exact case matches should be replaced
        assert mock_cell1.Value == "Hi"
        assert mock_cell2.Value == "hello"  # Unchanged
        assert mock_cell3.Value == "Hi"

    def test_find_replace_ignore_case(self):
        """Test find_replace ignoring case."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()

        mock_cell1 = MagicMock()
        mock_cell1.Value = "Hello"
        mock_cell2 = MagicMock()
        mock_cell2.Value = "hello"

        mock_range.__iter__ = MagicMock(return_value=iter([mock_cell1, mock_cell2]))
        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.find_replace(
            manager, mock_workbook, "Sheet1", "A1:A2",
            find_text="hello",
            replace_with="Hi",
            match_case=False,
            replace_all=True
        )

        assert mock_cell1.Value == "Hi"
        assert mock_cell2.Value == "Hi"

    def test_find_replace_replace_first_only(self):
        """Test find_replace replacing only first match."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()

        mock_cell1 = MagicMock()
        mock_cell1.Value = "Hello"
        mock_cell2 = MagicMock()
        mock_cell2.Value = "Hello"

        mock_range.__iter__ = MagicMock(return_value=iter([mock_cell1, mock_cell2]))
        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.find_replace(
            manager, mock_workbook, "Sheet1", "A1:A2",
            find_text="Hello",
            replace_with="Hi",
            match_case=True,
            replace_all=False
        )

        assert mock_cell1.Value == "Hi"
        # Second cell might or might not be replaced depending on iteration order


class TestClearContents:
    """Tests for clear_contents function."""

    def test_clear_contents(self):
        """Test clear_contents clears range contents."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()

        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.clear_contents(
            manager, mock_workbook, "Sheet1", "A1:D10"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        mock_worksheet.Range.assert_called_once_with("A1:D10")
        mock_range.ClearContents.assert_called_once()


class TestClearAll:
    """Tests for clear_all function."""

    def test_clear_all(self):
        """Test clear_all clears range contents and formatting."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()

        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        advanced_ops.clear_all(
            manager, mock_workbook, "Sheet1", "A1:D10"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        mock_worksheet.Range.assert_called_once_with("A1:D10")
        mock_range.Clear.assert_called_once()
