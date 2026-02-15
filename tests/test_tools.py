"""Tests for tools package."""

import json
from unittest.mock import MagicMock, patch
from contextlib import contextmanager


@contextmanager
def _fake_workbook_context(manager=None, workbook=None, path="/path/file.xlsx"):
    yield manager or MagicMock(), workbook or MagicMock(), path



class TestGetExcelManager:
    """Tests for get_excel_manager function."""

    def test_get_excel_manager_singleton(self):
        """Test get_excel_manager returns singleton instance."""
        import tools
        tools.excel_tool._excel_manager = None

        from tools.excel_tool import get_excel_manager

        manager1 = get_excel_manager()
        manager2 = get_excel_manager()

        assert manager1 is manager2

    def test_get_excel_manager_creates_new_instance(self):
        """Test get_excel_manager creates new instance when None."""
        import tools
        tools.excel_tool._excel_manager = None

        from tools.excel_tool import get_excel_manager

        manager = get_excel_manager()

        assert manager is not None
        assert tools.excel_tool._excel_manager is manager


class TestExcelStatus:
    """Tests for excel_status tool."""

    def test_excel_status_not_running(self):
        """Test excel_status when Excel not running."""
        import tools
        mock_manager = MagicMock()
        mock_manager.is_running.return_value = False
        mock_manager.is_app_alive.return_value = True
        tools.excel_tool._excel_manager = mock_manager

        from tools.excel_tool import excel_status

        # Use .func to access the underlying function
        result = json.loads(excel_status.func())

        assert result["success"] is True
        assert result["data"]["excel_running"] is False
        assert result["data"]["workbook_count"] == 0
        assert result["data"]["open_workbooks"] == []

    def test_excel_status_running(self):
        """Test excel_status when Excel is running."""
        import tools
        mock_manager = MagicMock()
        mock_manager.is_running.return_value = True
        mock_manager.is_app_alive.return_value = True
        mock_manager.get_workbook_count.return_value = 2
        mock_manager.get_open_workbooks.return_value = ["file1.xlsx", "file2.xlsx"]
        tools.excel_tool._excel_manager = mock_manager

        from tools.excel_tool import excel_status

        result = json.loads(excel_status.func())

        assert result["success"] is True
        assert result["data"]["excel_running"] is True
        assert result["data"]["workbook_count"] == 2
        assert len(result["data"]["open_workbooks"]) == 2

    def test_excel_status_error(self):
        """Test excel_status handles errors."""
        import tools
        mock_manager = MagicMock()
        mock_manager.is_running.side_effect = Exception("Test error")
        tools.excel_tool._excel_manager = mock_manager

        from tools.excel_tool import excel_status

        result = json.loads(excel_status.func())

        assert "error" in result


class TestExcelOpenWorkbook:
    """Tests for excel_open_workbook tool."""

    def test_excel_open_workbook_not_supported(self):
        """Test excel_open_workbook returns not supported error."""
        from tools.excel_tool import excel_open_workbook

        result = json.loads(excel_open_workbook.func("/path/to/file.xlsx"))

        assert result["success"] is False
        assert "no longer supported" in result["summary"].lower()


class TestExcelListWorksheets:
    """Tests for excel_list_worksheets tool."""

    @patch("tools.excel_tool.workbook_context")
    def test_excel_list_worksheets_success(self, mock_context):
        """Test excel_list_worksheets returns worksheets."""
        mock_manager = MagicMock()
        mock_manager.list_worksheets.return_value = ["Sheet1", "Sheet2", "Sheet3"]
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/to/file.xlsx")

        from tools.excel_tool import excel_list_worksheets

        result = json.loads(excel_list_worksheets.func("/path/to/file.xlsx"))

        assert result["success"] is True
        assert result["data"]["worksheets"] == ["Sheet1", "Sheet2", "Sheet3"]
        assert result["data"]["count"] == 3

    @patch("tools.excel_tool.workbook_context", side_effect=Exception("Not open"))
    def test_excel_list_worksheets_workbook_not_open(self, mock_context):
        """Test excel_list_worksheets handles errors."""
        from tools.excel_tool import excel_list_worksheets

        result = json.loads(excel_list_worksheets.func("/path/to/file.xlsx"))

        assert result["success"] is False


class TestExcelReadRange:
    """Tests for excel_read_range tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.read_range")
    def test_excel_read_range_success(self, mock_read, mock_context):
        """Test excel_read_range reads data successfully."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, mock_workbook, "/path/file.xlsx")

        sample_data = [["A1", "B1"], ["A2", "B2"]]
        mock_read.return_value = sample_data

        from tools.excel_tool import excel_read_range

        result = json.loads(excel_read_range.func("/path/file.xlsx", "Sheet1", "A1:B2"))

        assert result["success"] is True
        assert result["data"]["data"] == sample_data
        assert result["data"]["rows"] == 2
        assert result["data"]["columns"] == 2

    @patch("tools.excel_tool.workbook_context", side_effect=Exception("Sheet not found"))
    def test_excel_read_range_error(self, mock_context):
        """Test excel_read_range handles errors."""
        from tools.excel_tool import excel_read_range

        result = json.loads(excel_read_range.func("/path/file.xlsx", "Sheet1", "A1"))

        assert result["success"] is False


class TestExcelWriteRange:
    """Tests for excel_write_range tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.write_range")
    def test_excel_write_range_success(self, mock_write, mock_context):
        """Test excel_write_range writes data successfully."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, mock_workbook, "/path/file.xlsx")

        from tools.excel_tool import excel_write_range

        data = json.dumps([["A1", "B1"], ["A2", "B2"]])
        result = json.loads(excel_write_range.func("/path/file.xlsx", "Sheet1", "A1", data))

        assert result["success"] is True
        assert result["data"]["rows_written"] == 2
        assert result["data"]["columns_written"] == 2

    def test_excel_write_range_invalid_json(self):
        """Test excel_write_range handles invalid JSON."""
        from tools.excel_tool import excel_write_range

        result = json.loads(excel_write_range.func("/path/file.xlsx", "Sheet1", "A1", "not json"))

        assert result["success"] is False
        assert "invalid json" in result.get("summary", "").lower() or "invalid json" in result.get("error", "").lower()

    def test_excel_write_range_not_2d_array(self):
        """Test excel_write_range handles non-2D array."""
        from tools.excel_tool import excel_write_range

        result = json.loads(excel_write_range.func("/path/file.xlsx", "Sheet1", "A1", "[1, 2, 3]"))

        assert result["success"] is False
        assert "2d array" in result.get("summary", "").lower() or "2d array" in result.get("error", "").lower()


class TestExcelSaveWorkbook:
    """Tests for excel_save_workbook tool."""

    @patch("tools.excel_tool.workbook_context")
    def test_excel_save_workbook_success(self, mock_context):
        """Test excel_save_workbook saves successfully."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, mock_workbook, "/path/file.xlsx")

        from tools.excel_tool import excel_save_workbook

        result = json.loads(excel_save_workbook.func("/path/file.xlsx"))

        assert result["success"] is True
        mock_manager.save_workbook.assert_called_once()

    @patch("tools.excel_tool.workbook_context")
    def test_excel_save_workbook_as(self, mock_context):
        """Test excel_save_workbook with save_as parameter."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, mock_workbook, "/path/file.xlsx")

        from tools.excel_tool import excel_save_workbook

        result = json.loads(excel_save_workbook.func("/path/file.xlsx", save_as="/new/path/new.xlsx"))

        assert result["success"] is True
        assert result["data"]["saved_as"] == "/new/path/new.xlsx"


class TestExcelCloseWorkbook:
    """Tests for excel_close_workbook tool."""

    def test_excel_close_workbook_not_supported(self):
        """Test excel_close_workbook returns not supported error."""
        from tools.excel_tool import excel_close_workbook

        result = json.loads(excel_close_workbook.func("/path/file.xlsx", save=True))

        assert result["success"] is False
        assert "no longer supported" in result["summary"].lower()


class TestExcelAddWorksheet:
    """Tests for excel_add_worksheet tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.add_worksheet")
    def test_excel_add_worksheet_success(self, mock_add, mock_context):
        """Test excel_add_worksheet adds worksheet."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_add_worksheet

        result = json.loads(excel_add_worksheet.func("/path/file.xlsx", "NewSheet"))

        assert result["success"] is True
        assert result["data"]["worksheet_name"] == "NewSheet"


class TestExcelDeleteWorksheet:
    """Tests for excel_delete_worksheet tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.delete_worksheet")
    def test_excel_delete_worksheet_success(self, mock_delete, mock_context):
        """Test excel_delete_worksheet deletes worksheet."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_delete_worksheet

        result = json.loads(excel_delete_worksheet.func("/path/file.xlsx", "OldSheet"))

        assert result["success"] is True
        assert result["data"]["deleted_worksheet"] == "OldSheet"


class TestExcelRenameWorksheet:
    """Tests for excel_rename_worksheet tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.rename_worksheet")
    def test_excel_rename_worksheet_success(self, mock_rename, mock_context):
        """Test excel_rename_worksheet renames worksheet."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_rename_worksheet

        result = json.loads(excel_rename_worksheet.func("/path/file.xlsx", "OldName", "NewName"))

        assert result["success"] is True
        assert result["data"]["old_name"] == "OldName"
        assert result["data"]["new_name"] == "NewName"


class TestExcelCopyWorksheet:
    """Tests for excel_copy_worksheet tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.copy_worksheet")
    def test_excel_copy_worksheet_success(self, mock_copy, mock_context):
        """Test excel_copy_worksheet copies worksheet."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        mock_new_sheet = MagicMock()
        mock_new_sheet.Name = "Sheet1 (2)"
        mock_copy.return_value = mock_new_sheet

        from tools.excel_tool import excel_copy_worksheet

        result = json.loads(excel_copy_worksheet.func("/path/file.xlsx", "Sheet1"))

        assert result["success"] is True
        assert result["data"]["source_worksheet"] == "Sheet1"


class TestExcelGetUsedRange:
    """Tests for excel_get_used_range tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.get_used_range")
    def test_excel_get_used_range_success(self, mock_get_range, mock_context):
        """Test excel_get_used_range returns range info."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        mock_get_range.return_value = ("$A$1:$D$10", 10, 4)

        from tools.excel_tool import excel_get_used_range

        result = json.loads(excel_get_used_range.func("/path/file.xlsx", "Sheet1"))

        assert result["success"] is True
        assert result["data"]["used_range"] == "$A$1:$D$10"
        assert result["data"]["rows"] == 10
        assert result["data"]["columns"] == 4


class TestExcelSetFontFormat:
    """Tests for excel_set_font_format tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.set_font_format")
    def test_excel_set_font_format_success(self, mock_set_font, mock_context):
        """Test excel_set_font_format applies formatting."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_set_font_format

        result = json.loads(excel_set_font_format.func(
            "/path/file.xlsx", "Sheet1", "A1:D10",
            font_name="Arial", size=12, bold=True
        ))

        assert result["success"] is True
        assert "font_name" in result["data"]["formatting_applied"]


class TestExcelSetCellFormat:
    """Tests for excel_set_cell_format tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.set_cell_format")
    def test_excel_set_cell_format_success(self, mock_set_cell, mock_context):
        """Test excel_set_cell_format applies formatting."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_set_cell_format

        result = json.loads(excel_set_cell_format.func(
            "/path/file.xlsx", "Sheet1", "A1:D10",
            horizontal_alignment="center", wrap_text=True
        ))

        assert result["success"] is True


class TestExcelSetBorderFormat:
    """Tests for excel_set_border_format tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.set_border_format")
    def test_excel_set_border_format_success(self, mock_set_border, mock_context):
        """Test excel_set_border_format applies formatting."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_set_border_format

        result = json.loads(excel_set_border_format.func(
            "/path/file.xlsx", "Sheet1", "A1:D10",
            edge="all", style="continuous"
        ))

        assert result["success"] is True


class TestExcelSetBackgroundColor:
    """Tests for excel_set_background_color tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.set_background_color")
    def test_excel_set_background_color_success(self, mock_set_bg, mock_context):
        """Test excel_set_background_color applies color."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_set_background_color

        result = json.loads(excel_set_background_color.func(
            "/path/file.xlsx", "Sheet1", "A1:D10", "FFFF00"
        ))

        assert result["success"] is True
        assert result["data"]["background_color"] == "FFFF00"


class TestExcelSetFormula:
    """Tests for excel_set_formula tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formula_ops.set_formula")
    def test_excel_set_formula_success(self, mock_set_formula, mock_context):
        """Test excel_set_formula sets formula."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_set_formula

        result = json.loads(excel_set_formula.func(
            "/path/file.xlsx", "Sheet1", "A1", "=SUM(B1:B10)"
        ))

        assert result["success"] is True
        assert result["data"]["formula"] == "=SUM(B1:B10)"


class TestExcelGetFormula:
    """Tests for excel_get_formula tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formula_ops.get_formula")
    def test_excel_get_formula_success(self, mock_get_formula, mock_context):
        """Test excel_get_formula retrieves formula."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        mock_get_formula.return_value = "=SUM(A1:A10)"

        from tools.excel_tool import excel_get_formula

        result = json.loads(excel_get_formula.func("/path/file.xlsx", "Sheet1", "A1"))

        assert result["success"] is True
        assert result["data"]["formula"] == "=SUM(A1:A10)"
        assert result["data"]["is_formula"] is True

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formula_ops.get_formula")
    def test_excel_get_formula_value_not_formula(self, mock_get_formula, mock_context):
        """Test excel_get_formula with value not formula."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        mock_get_formula.return_value = "42"  # Just a value

        from tools.excel_tool import excel_get_formula

        result = json.loads(excel_get_formula.func("/path/file.xlsx", "Sheet1", "A1"))

        assert result["success"] is True
        assert result["data"]["is_formula"] is False


class TestExcelAutoFitColumns:
    """Tests for excel_auto_fit_columns tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.auto_fit_columns")
    def test_excel_auto_fit_columns_success(self, mock_autofit, mock_context):
        """Test excel_auto_fit_columns autofits columns."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_auto_fit_columns

        result = json.loads(excel_auto_fit_columns.func("/path/file.xlsx", "Sheet1", "A:D"))

        assert result["success"] is True
        assert result["data"]["columns_auto_fitted"] == "A:D"


class TestExcelSetColumnWidth:
    """Tests for excel_set_column_width tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.set_column_width")
    def test_excel_set_column_width_success(self, mock_set_width, mock_context):
        """Test excel_set_column_width sets width."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_set_column_width

        result = json.loads(excel_set_column_width.func(
            "/path/file.xlsx", "Sheet1", "A:C", 15.0
        ))

        assert result["success"] is True
        assert result["data"]["width"] == 15.0


class TestExcelSetRowHeight:
    """Tests for excel_set_row_height tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.set_row_height")
    def test_excel_set_row_height_success(self, mock_set_height, mock_context):
        """Test excel_set_row_height sets height."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock(), "/path/file.xlsx")

        from tools.excel_tool import excel_set_row_height

        result = json.loads(excel_set_row_height.func(
            "/path/file.xlsx", "Sheet1", "1:5", 20.0
        ))

        assert result["success"] is True
        assert result["data"]["height"] == 20.0


class TestExcelToolsList:
    """Tests for EXCEL_TOOLS list."""

    def test_tools_list_complete(self):
        """Test EXCEL_TOOLS contains all expected tools."""
        from tools.excel_tool import EXCEL_TOOLS

        tool_names = [tool.name for tool in EXCEL_TOOLS]

        # Status tool
        assert "excel_status" in tool_names

        # Workbook tools
        assert "excel_create_workbook" in tool_names
        assert "excel_list_worksheets" in tool_names
        assert "excel_read_range" in tool_names
        assert "excel_write_range" in tool_names
        assert "excel_save_workbook" in tool_names

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

    def test_tools_count(self):
        """Test EXCEL_TOOLS has correct number of tools."""
        from tools.excel_tool import EXCEL_TOOLS

        # Total: 1 status + 3 workbook + 2 range + 5 worksheet + 4 format + 2 formula + 3 col/row = 20
        assert len(EXCEL_TOOLS) == 20
