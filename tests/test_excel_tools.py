"""Tests for excel_tools.py."""

import json
from unittest.mock import MagicMock, patch

import pytest


class TestGetExcelManager:
    """Tests for get_excel_manager function."""

    def test_get_excel_manager_singleton(self):
        """Test get_excel_manager returns singleton instance."""
        # Clear the global instance first
        import excel_tools
        excel_tools._excel_manager = None

        from excel_tools import get_excel_manager

        manager1 = get_excel_manager()
        manager2 = get_excel_manager()

        assert manager1 is manager2

    def test_get_excel_manager_creates_new_instance(self):
        """Test get_excel_manager creates new instance when None."""
        import excel_tools
        excel_tools._excel_manager = None

        from excel_tools import get_excel_manager

        manager = get_excel_manager()

        assert manager is not None
        assert excel_tools._excel_manager is manager


class TestExcelStatus:
    """Tests for excel_status tool."""

    @patch("excel_tools.get_excel_manager")
    def test_excel_status_not_running(self, mock_get_manager):
        """Test excel_status when Excel not running."""
        mock_manager = MagicMock()
        mock_manager.is_running.return_value = False
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_status

        # Use .func to access the underlying function
        result = json.loads(excel_status.func())

        assert result["excel_running"] is False
        assert result["workbook_count"] == 0
        assert result["open_workbooks"] == []

    @patch("excel_tools.get_excel_manager")
    def test_excel_status_running(self, mock_get_manager):
        """Test excel_status when Excel is running."""
        mock_manager = MagicMock()
        mock_manager.is_running.return_value = True
        mock_manager.get_workbook_count.return_value = 2
        mock_manager.get_open_workbooks.return_value = ["file1.xlsx", "file2.xlsx"]
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_status

        result = json.loads(excel_status.func())

        assert result["excel_running"] is True
        assert result["workbook_count"] == 2
        assert len(result["open_workbooks"]) == 2

    @patch("excel_tools.get_excel_manager", side_effect=Exception("Test error"))
    def test_excel_status_error(self, mock_get_manager):
        """Test excel_status handles errors."""
        from excel_tools import excel_status

        result = json.loads(excel_status.func())

        assert "error" in result


class TestExcelOpenWorkbook:
    """Tests for excel_open_workbook tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.open_workbook")
    def test_excel_open_workbook_success(self, mock_open_wb, mock_get_manager, tmp_path):
        """Test excel_open_workbook opens workbook successfully."""
        mock_manager = MagicMock()
        mock_get_manager.return_value = mock_manager

        mock_workbook = MagicMock()
        mock_workbook.Name = "test.xlsx"
        mock_open_wb.return_value = (mock_workbook, ["Sheet1", "Sheet2"])

        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        from excel_tools import excel_open_workbook

        result = json.loads(excel_open_workbook.func(str(test_file)))

        assert result["success"] is True
        assert result["workbook_name"] == "test.xlsx"
        assert result["worksheets"] == ["Sheet1", "Sheet2"]

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.open_workbook", side_effect=FileNotFoundError("Not found"))
    def test_excel_open_workbook_file_not_found(self, mock_open_wb, mock_get_manager):
        """Test excel_open_workbook handles FileNotFoundError."""
        mock_manager = MagicMock()
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_open_workbook

        result = json.loads(excel_open_workbook.func("/nonexistent.xlsx"))

        assert "error" in result
        assert "not found" in result["error"].lower()

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.open_workbook", side_effect=Exception("Test error"))
    def test_excel_open_workbook_error(self, mock_open_wb, mock_get_manager):
        """Test excel_open_workbook handles errors."""
        mock_manager = MagicMock()
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_open_workbook

        result = json.loads(excel_open_workbook.func("/path/to/file.xlsx"))

        assert "error" in result

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.open_workbook")
    def test_excel_open_workbook_read_only(self, mock_open_wb, mock_get_manager, tmp_path):
        """Test excel_open_workbook with read_only flag."""
        mock_manager = MagicMock()
        mock_get_manager.return_value = mock_manager

        mock_workbook = MagicMock()
        mock_workbook.Name = "test.xlsx"
        mock_open_wb.return_value = (mock_workbook, [])

        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        from excel_tools import excel_open_workbook

        result = json.loads(excel_open_workbook.func(str(test_file), read_only=True))

        assert result["read_only"] is True


class TestExcelListWorksheets:
    """Tests for excel_list_worksheets tool."""

    @patch("excel_tools.get_excel_manager")
    def test_excel_list_worksheets_success(self, mock_get_manager):
        """Test excel_list_worksheets returns worksheets."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_manager.list_worksheets.return_value = ["Sheet1", "Sheet2", "Sheet3"]
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_list_worksheets

        result = json.loads(excel_list_worksheets.func("/path/to/file.xlsx"))

        assert result["success"] is True
        assert result["worksheets"] == ["Sheet1", "Sheet2", "Sheet3"]
        assert result["count"] == 3

    @patch("excel_tools.get_excel_manager")
    def test_excel_list_worksheets_workbook_not_open(self, mock_get_manager):
        """Test excel_list_worksheets handles closed workbook."""
        mock_manager = MagicMock()
        mock_manager.get_workbook.side_effect = ValueError("Not open")
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_list_worksheets

        result = json.loads(excel_list_worksheets.func("/path/to/file.xlsx"))

        assert "error" in result


class TestExcelReadRange:
    """Tests for excel_read_range tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.read_range")
    def test_excel_read_range_success(self, mock_read, mock_get_manager):
        """Test excel_read_range reads data successfully."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        sample_data = [["A1", "B1"], ["A2", "B2"]]
        mock_read.return_value = sample_data

        from excel_tools import excel_read_range

        result = json.loads(excel_read_range.func("/path/file.xlsx", "Sheet1", "A1:B2"))

        assert result["success"] is True
        assert result["data"] == sample_data
        assert result["rows"] == 2
        assert result["columns"] == 2

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.read_range", side_effect=ValueError("Sheet not found"))
    def test_excel_read_range_error(self, mock_read, mock_get_manager):
        """Test excel_read_range handles errors."""
        mock_manager = MagicMock()
        mock_manager.get_workbook.return_value = MagicMock()
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_read_range

        result = json.loads(excel_read_range.func("/path/file.xlsx", "Sheet1", "A1"))

        assert "error" in result


class TestExcelWriteRange:
    """Tests for excel_write_range tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.write_range")
    def test_excel_write_range_success(self, mock_write, mock_get_manager):
        """Test excel_write_range writes data successfully."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_write_range

        data = json.dumps([["A1", "B1"], ["A2", "B2"]])
        result = json.loads(excel_write_range.func("/path/file.xlsx", "Sheet1", "A1", data))

        assert result["success"] is True
        assert result["rows_written"] == 2
        assert result["columns_written"] == 2

    @patch("excel_tools.get_excel_manager")
    def test_excel_write_range_invalid_json(self, mock_get_manager):
        """Test excel_write_range handles invalid JSON."""
        mock_manager = MagicMock()
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_write_range

        result = json.loads(excel_write_range.func("/path/file.xlsx", "Sheet1", "A1", "not json"))

        assert "error" in result
        assert "Invalid JSON" in result["error"]

    @patch("excel_tools.get_excel_manager")
    def test_excel_write_range_not_2d_array(self, mock_get_manager):
        """Test excel_write_range handles non-2D array."""
        mock_manager = MagicMock()
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_write_range

        result = json.loads(excel_write_range.func("/path/file.xlsx", "Sheet1", "A1", "[1, 2, 3]"))

        assert "error" in result
        assert "2D array" in result["error"]


class TestExcelSaveWorkbook:
    """Tests for excel_save_workbook tool."""

    @patch("excel_tools.get_excel_manager")
    def test_excel_save_workbook_success(self, mock_get_manager):
        """Test excel_save_workbook saves successfully."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_save_workbook

        result = json.loads(excel_save_workbook.func("/path/file.xlsx"))

        assert result["success"] is True
        mock_manager.save_workbook.assert_called_once()

    @patch("excel_tools.get_excel_manager")
    def test_excel_save_workbook_as(self, mock_get_manager):
        """Test excel_save_workbook with save_as parameter."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_save_workbook

        result = json.loads(excel_save_workbook.func("/path/file.xlsx", save_as="/new/path/new.xlsx"))

        assert result["success"] is True
        assert result["saved_as"] == "/new/path/new.xlsx"


class TestExcelCloseWorkbook:
    """Tests for excel_close_workbook tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.close_workbook")
    def test_excel_close_workbook_success(self, mock_close, mock_get_manager):
        """Test excel_close_workbook closes successfully."""
        mock_manager = MagicMock()
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_close_workbook

        result = json.loads(excel_close_workbook.func("/path/file.xlsx", save=True))

        assert result["success"] is True
        assert result["saved"] is True

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.close_workbook", side_effect=ValueError("Not open"))
    def test_excel_close_workbook_error(self, mock_close, mock_get_manager):
        """Test excel_close_workbook handles errors."""
        mock_manager = MagicMock()
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_close_workbook

        result = json.loads(excel_close_workbook.func("/path/file.xlsx"))

        assert "error" in result


class TestExcelAddWorksheet:
    """Tests for excel_add_worksheet tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.add_worksheet")
    def test_excel_add_worksheet_success(self, mock_add, mock_get_manager):
        """Test excel_add_worksheet adds worksheet."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_add_worksheet

        result = json.loads(excel_add_worksheet.func("/path/file.xlsx", "NewSheet"))

        assert result["success"] is True
        assert result["worksheet_name"] == "NewSheet"


class TestExcelDeleteWorksheet:
    """Tests for excel_delete_worksheet tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.delete_worksheet")
    def test_excel_delete_worksheet_success(self, mock_delete, mock_get_manager):
        """Test excel_delete_worksheet deletes worksheet."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_delete_worksheet

        result = json.loads(excel_delete_worksheet.func("/path/file.xlsx", "OldSheet"))

        assert result["success"] is True
        assert result["deleted_worksheet"] == "OldSheet"


class TestExcelRenameWorksheet:
    """Tests for excel_rename_worksheet tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.rename_worksheet")
    def test_excel_rename_worksheet_success(self, mock_rename, mock_get_manager):
        """Test excel_rename_worksheet renames worksheet."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_rename_worksheet

        result = json.loads(excel_rename_worksheet.func("/path/file.xlsx", "OldName", "NewName"))

        assert result["success"] is True
        assert result["old_name"] == "OldName"
        assert result["new_name"] == "NewName"


class TestExcelCopyWorksheet:
    """Tests for excel_copy_worksheet tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.copy_worksheet")
    def test_excel_copy_worksheet_success(self, mock_copy, mock_get_manager):
        """Test excel_copy_worksheet copies worksheet."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        mock_new_sheet = MagicMock()
        mock_new_sheet.Name = "Sheet1 (2)"
        mock_copy.return_value = mock_new_sheet

        from excel_tools import excel_copy_worksheet

        result = json.loads(excel_copy_worksheet.func("/path/file.xlsx", "Sheet1"))

        assert result["success"] is True
        assert result["source_worksheet"] == "Sheet1"


class TestExcelGetUsedRange:
    """Tests for excel_get_used_range tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.workbook_ops.get_used_range")
    def test_excel_get_used_range_success(self, mock_get_range, mock_get_manager):
        """Test excel_get_used_range returns range info."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        mock_get_range.return_value = ("$A$1:$D$10", 10, 4)

        from excel_tools import excel_get_used_range

        result = json.loads(excel_get_used_range.func("/path/file.xlsx", "Sheet1"))

        assert result["success"] is True
        assert result["used_range"] == "$A$1:$D$10"
        assert result["rows"] == 10
        assert result["columns"] == 4


class TestExcelSetFontFormat:
    """Tests for excel_set_font_format tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formatting_ops.set_font_format")
    def test_excel_set_font_format_success(self, mock_set_font, mock_get_manager):
        """Test excel_set_font_format applies formatting."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_set_font_format

        result = json.loads(excel_set_font_format.func(
            "/path/file.xlsx", "Sheet1", "A1:D10",
            font_name="Arial", size=12, bold=True
        ))

        assert result["success"] is True
        assert "font_name" in result["formatting_applied"]


class TestExcelSetCellFormat:
    """Tests for excel_set_cell_format tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formatting_ops.set_cell_format")
    def test_excel_set_cell_format_success(self, mock_set_cell, mock_get_manager):
        """Test excel_set_cell_format applies formatting."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_set_cell_format

        result = json.loads(excel_set_cell_format.func(
            "/path/file.xlsx", "Sheet1", "A1:D10",
            horizontal_alignment="center", wrap_text=True
        ))

        assert result["success"] is True


class TestExcelSetBorderFormat:
    """Tests for excel_set_border_format tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formatting_ops.set_border_format")
    def test_excel_set_border_format_success(self, mock_set_border, mock_get_manager):
        """Test excel_set_border_format applies formatting."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_set_border_format

        result = json.loads(excel_set_border_format.func(
            "/path/file.xlsx", "Sheet1", "A1:D10",
            edge="all", style="continuous"
        ))

        assert result["success"] is True


class TestExcelSetBackgroundColor:
    """Tests for excel_set_background_color tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formatting_ops.set_background_color")
    def test_excel_set_background_color_success(self, mock_set_bg, mock_get_manager):
        """Test excel_set_background_color applies color."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_set_background_color

        result = json.loads(excel_set_background_color.func(
            "/path/file.xlsx", "Sheet1", "A1:D10", "FFFF00"
        ))

        assert result["success"] is True
        assert result["background_color"] == "FFFF00"


class TestExcelSetFormula:
    """Tests for excel_set_formula tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formula_ops.set_formula")
    def test_excel_set_formula_success(self, mock_set_formula, mock_get_manager):
        """Test excel_set_formula sets formula."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_set_formula

        result = json.loads(excel_set_formula.func(
            "/path/file.xlsx", "Sheet1", "A1", "=SUM(B1:B10)"
        ))

        assert result["success"] is True
        assert result["formula"] == "=SUM(B1:B10)"


class TestExcelGetFormula:
    """Tests for excel_get_formula tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formula_ops.get_formula")
    def test_excel_get_formula_success(self, mock_get_formula, mock_get_manager):
        """Test excel_get_formula retrieves formula."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        mock_get_formula.return_value = "=SUM(A1:A10)"

        from excel_tools import excel_get_formula

        result = json.loads(excel_get_formula.func("/path/file.xlsx", "Sheet1", "A1"))

        assert result["success"] is True
        assert result["formula"] == "=SUM(A1:A10)"
        assert result["is_formula"] is True

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formula_ops.get_formula")
    def test_excel_get_formula_value_not_formula(self, mock_get_formula, mock_get_manager):
        """Test excel_get_formula with value not formula."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        mock_get_formula.return_value = "42"  # Just a value

        from excel_tools import excel_get_formula

        result = json.loads(excel_get_formula.func("/path/file.xlsx", "Sheet1", "A1"))

        assert result["success"] is True
        assert result["is_formula"] is False


class TestExcelAutoFitColumns:
    """Tests for excel_auto_fit_columns tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formatting_ops.auto_fit_columns")
    def test_excel_auto_fit_columns_success(self, mock_autofit, mock_get_manager):
        """Test excel_auto_fit_columns autofits columns."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_auto_fit_columns

        result = json.loads(excel_auto_fit_columns.func("/path/file.xlsx", "Sheet1", "A:D"))

        assert result["success"] is True
        assert result["columns_auto_fitted"] == "A:D"


class TestExcelSetColumnWidth:
    """Tests for excel_set_column_width tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formatting_ops.set_column_width")
    def test_excel_set_column_width_success(self, mock_set_width, mock_get_manager):
        """Test excel_set_column_width sets width."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_set_column_width

        result = json.loads(excel_set_column_width.func(
            "/path/file.xlsx", "Sheet1", "A:C", 15.0
        ))

        assert result["success"] is True
        assert result["width"] == 15.0


class TestExcelSetRowHeight:
    """Tests for excel_set_row_height tool."""

    @patch("excel_tools.get_excel_manager")
    @patch("excel_tools.formatting_ops.set_row_height")
    def test_excel_set_row_height_success(self, mock_set_height, mock_get_manager):
        """Test excel_set_row_height sets height."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_manager.get_workbook.return_value = mock_workbook
        mock_get_manager.return_value = mock_manager

        from excel_tools import excel_set_row_height

        result = json.loads(excel_set_row_height.func(
            "/path/file.xlsx", "Sheet1", "1:5", 20.0
        ))

        assert result["success"] is True
        assert result["height"] == 20.0


class TestExcelToolsList:
    """Tests for EXCEL_TOOLS list."""

    def test_excel_tools_list_complete(self):
        """Test EXCEL_TOOLS contains all expected tools."""
        from excel_tools import EXCEL_TOOLS

        tool_names = [tool.name for tool in EXCEL_TOOLS]

        # Status tool
        assert "excel_status" in tool_names

        # Workbook tools
        assert "excel_open_workbook" in tool_names
        assert "excel_list_worksheets" in tool_names
        assert "excel_read_range" in tool_names
        assert "excel_write_range" in tool_names
        assert "excel_save_workbook" in tool_names
        assert "excel_close_workbook" in tool_names

        # Worksheet tools
        assert "excel_add_worksheet" in tool_names
        assert "excel_delete_worksheet" in tool_names
        assert "excel_rename_worksheet" in tool_names
        assert "excel_copy_worksheet" in tool_names
        assert "excel_get_used_range" in tool_names

        # Format tools
        assert "excel_set_font_format" in tool_names
        assert "excel_set_cell_format" in tool_names
        assert "excel_set_border_format" in tool_names
        assert "excel_set_background_color" in tool_names

        # Formula tools
        assert "excel_set_formula" in tool_names
        assert "excel_get_formula" in tool_names

        # Column/Row tools
        assert "excel_auto_fit_columns" in tool_names
        assert "excel_set_column_width" in tool_names
        assert "excel_set_row_height" in tool_names

    def test_excel_tools_count(self):
        """Test EXCEL_TOOLS has correct number of tools."""
        from excel_tools import EXCEL_TOOLS

        # Total: 1 status + 6 workbook + 5 worksheet + 4 format + 2 formula + 3 col/row = 21
        assert len(EXCEL_TOOLS) == 21
