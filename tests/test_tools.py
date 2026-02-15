"""Tests for tools package - CRUD-based API."""

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
        from tools import excel_tool
        excel_tool._excel_manager = None

        from tools.excel_tool import get_excel_manager

        manager1 = get_excel_manager()
        manager2 = get_excel_manager()

        assert manager1 is manager2

    def test_get_excel_manager_creates_new_instance(self):
        """Test get_excel_manager creates new instance when None."""
        from tools import excel_tool
        excel_tool._excel_manager = None

        from tools.excel_tool import get_excel_manager

        manager = get_excel_manager()

        assert manager is not None
        assert excel_tool._excel_manager is manager


# =============================================================================
# Workbook Lifecycle Tools
# =============================================================================

class TestExcelOpen:
    """Tests for excel_open tool."""

    def test_excel_open_existing_file(self):
        """Test excel_open opens an existing workbook."""
        with patch("tools.excel_tool.Path") as mock_path, \
             patch("tools.excel_tool.workbook_context") as mock_context:
            mock_path_instance = MagicMock()
            mock_path_instance.exists.return_value = True
            mock_path.return_value = mock_path_instance

            mock_manager = MagicMock()
            mock_manager.list_worksheets.return_value = ["Sheet1", "Sheet2"]
            mock_workbook = MagicMock()
            mock_workbook.Name = "test.xlsx"
            mock_context.return_value = _fake_workbook_context(mock_manager, mock_workbook)

            from tools.excel_tool import excel_open

            result = json.loads(excel_open.func("/path/to/file.xlsx"))

            assert result["success"] is True
            assert result["data"]["created"] is False
            assert result["data"]["worksheets"] == ["Sheet1", "Sheet2"]

    def test_excel_open_file_not_found(self):
        """Test excel_open returns error when file not found."""
        with patch("tools.excel_tool.Path") as mock_path:
            mock_path_instance = MagicMock()
            mock_path_instance.exists.return_value = False
            mock_path.return_value = mock_path_instance

            from tools.excel_tool import excel_open

            result = json.loads(excel_open.func("/path/to/missing.xlsx"))

            assert result["success"] is False
            assert "does not exist" in result["error"]


class TestExcelSave:
    """Tests for excel_save tool."""

    @patch("tools.excel_tool.workbook_context")
    def test_excel_save_success(self, mock_context):
        """Test excel_save saves successfully."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, mock_workbook)

        from tools.excel_tool import excel_save

        result = json.loads(excel_save.func("/path/file.xlsx"))

        assert result["success"] is True
        mock_manager.save_workbook.assert_called_once()

    @patch("tools.excel_tool.workbook_context")
    def test_excel_save_as(self, mock_context):
        """Test excel_save with save_as parameter."""
        mock_manager = MagicMock()
        mock_workbook = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, mock_workbook)

        from tools.excel_tool import excel_save

        result = json.loads(excel_save.func("/path/file.xlsx", save_as="/new/path/new.xlsx"))

        assert result["success"] is True
        assert result["data"]["saved_as"] == "/new/path/new.xlsx"


class TestExcelInfo:
    """Tests for excel_info tool."""

    def test_excel_info_not_initialized(self):
        """Test excel_info when Excel not initialized."""
        from tools import excel_tool
        excel_tool._excel_manager = None

        from tools.excel_tool import excel_info

        result = json.loads(excel_info.func())

        assert result["success"] is True
        assert result["data"]["excel_running"] is False

    def test_excel_info_running(self):
        """Test excel_info when Excel is running."""
        from tools import excel_tool
        mock_manager = MagicMock()
        mock_manager.is_running.return_value = True
        mock_manager.is_app_alive.return_value = True
        mock_manager.get_workbook_count.return_value = 2
        mock_manager.get_open_workbooks.return_value = ["file1.xlsx", "file2.xlsx"]
        excel_tool._excel_manager = mock_manager

        from tools.excel_tool import excel_info

        result = json.loads(excel_info.func())

        assert result["success"] is True
        assert result["data"]["excel_running"] is True
        assert result["data"]["workbook_count"] == 2


# =============================================================================
# CRUD: Create
# =============================================================================

class TestExcelCreate:
    """Tests for excel_create tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.add_worksheet")
    def test_create_worksheet(self, mock_add, mock_context):
        """Test excel_create creates a worksheet."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_create

        result = json.loads(excel_create.func(
            "/path/file.xlsx",
            "worksheet",
            '{"name": "NewSheet"}'
        ))

        assert result["success"] is True
        assert result["data"]["worksheet_name"] == "NewSheet"

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.insert_rows")
    def test_create_rows(self, mock_insert, mock_context):
        """Test excel_create inserts rows."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_create

        result = json.loads(excel_create.func(
            "/path/file.xlsx",
            "rows",
            '{"worksheet": "Sheet1", "index": 5, "count": 3}'
        ))

        assert result["success"] is True
        assert result["data"]["count"] == 3

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.insert_columns")
    def test_create_columns(self, mock_insert, mock_context):
        """Test excel_create inserts columns."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_create

        result = json.loads(excel_create.func(
            "/path/file.xlsx",
            "columns",
            '{"worksheet": "Sheet1", "index": "C", "count": 2}'
        ))

        assert result["success"] is True
        assert result["data"]["count"] == 2

    def test_create_invalid_json(self):
        """Test excel_create handles invalid JSON."""
        from tools.excel_tool import excel_create

        result = json.loads(excel_create.func("/path/file.xlsx", "worksheet", "not json"))

        assert result["success"] is False

    def test_create_invalid_target(self):
        """Test excel_create handles invalid target."""
        from tools.excel_tool import excel_create

        result = json.loads(excel_create.func("/path/file.xlsx", "invalid", '{}'))

        assert result["success"] is False


# =============================================================================
# CRUD: Read
# =============================================================================

class TestExcelRead:
    """Tests for excel_read tool."""

    @patch("tools.excel_tool.workbook_context")
    def test_read_worksheets(self, mock_context):
        """Test excel_read lists worksheets."""
        mock_manager = MagicMock()
        mock_manager.list_worksheets.return_value = ["Sheet1", "Sheet2", "Sheet3"]
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_read

        result = json.loads(excel_read.func("/path/file.xlsx", "worksheets", '{}'))

        assert result["success"] is True
        assert result["data"]["worksheets"] == ["Sheet1", "Sheet2", "Sheet3"]
        assert result["data"]["count"] == 3

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.read_range")
    def test_read_range(self, mock_read, mock_context):
        """Test excel_read reads a range."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        sample_data = [["A1", "B1"], ["A2", "B2"]]
        mock_read.return_value = sample_data

        from tools.excel_tool import excel_read

        result = json.loads(excel_read.func(
            "/path/file.xlsx",
            "range",
            '{"worksheet": "Sheet1", "range": "A1:B2"}'
        ))

        assert result["success"] is True
        assert result["data"]["data"] == sample_data
        assert result["data"]["rows"] == 2
        assert result["data"]["columns"] == 2

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formula_ops.get_formula")
    def test_read_formula(self, mock_get_formula, mock_context):
        """Test excel_read retrieves formula."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())
        mock_get_formula.return_value = "=SUM(A1:A10)"

        from tools.excel_tool import excel_read

        result = json.loads(excel_read.func(
            "/path/file.xlsx",
            "formula",
            '{"worksheet": "Sheet1", "range": "A1"}'
        ))

        assert result["success"] is True
        assert result["data"]["formula"] == "=SUM(A1:A10)"
        assert result["data"]["is_formula"] is True

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.get_used_range")
    def test_read_used_range(self, mock_get_range, mock_context):
        """Test excel_read gets used range."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())
        mock_get_range.return_value = ("$A$1:$D$10", 10, 4)

        from tools.excel_tool import excel_read

        result = json.loads(excel_read.func(
            "/path/file.xlsx",
            "used_range",
            '{"worksheet": "Sheet1"}'
        ))

        assert result["success"] is True
        assert result["data"]["used_range"] == "$A$1:$D$10"
        assert result["data"]["rows"] == 10
        assert result["data"]["columns"] == 4

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.get_format")
    def test_read_format(self, mock_get_format, mock_context):
        """Test excel_read retrieves format."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())
        mock_get_format.return_value = {"font": {"bold": True}}

        from tools.excel_tool import excel_read

        result = json.loads(excel_read.func(
            "/path/file.xlsx",
            "format",
            '{"worksheet": "Sheet1", "range": "A1:D10"}'
        ))

        assert result["success"] is True
        assert "font" in result["data"]["format"]


# =============================================================================
# CRUD: Update
# =============================================================================

class TestExcelUpdate:
    """Tests for excel_update tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.write_range")
    def test_update_range(self, mock_write, mock_context):
        """Test excel_update writes data."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_update

        result = json.loads(excel_update.func(
            "/path/file.xlsx",
            "range",
            '{"worksheet": "Sheet1", "range": "A1", "data": [[1, 2], [3, 4]]}'
        ))

        assert result["success"] is True
        assert result["data"]["rows_written"] == 2

    def test_update_range_invalid_json(self):
        """Test excel_update handles invalid JSON."""
        from tools.excel_tool import excel_update

        result = json.loads(excel_update.func("/path/file.xlsx", "range", "not json"))

        assert result["success"] is False

    def test_update_range_not_2d_array(self):
        """Test excel_update handles non-2D array."""
        from tools.excel_tool import excel_update

        result = json.loads(excel_update.func(
            "/path/file.xlsx",
            "range",
            '{"worksheet": "Sheet1", "range": "A1", "data": [1, 2, 3]}'
        ))

        assert result["success"] is False

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.set_font_format")
    @patch("libs.excel_com.formatting_ops.set_background_color")
    def test_update_format(self, mock_set_bg, mock_set_font, mock_context):
        """Test excel_update applies format."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_update

        result = json.loads(excel_update.func(
            "/path/file.xlsx",
            "format",
            '{"worksheet": "Sheet1", "ranges": [{"range": "A1:B2", "font": {"bold": true}, "fill": {"color": "FFFF00"}}]}'
        ))

        assert result["success"] is True
        assert result["data"]["ranges_formatted"] == 1

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formula_ops.set_formula")
    def test_update_formula(self, mock_set_formula, mock_context):
        """Test excel_update sets formula."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_update

        result = json.loads(excel_update.func(
            "/path/file.xlsx",
            "formula",
            '{"worksheet": "Sheet1", "range": "A1", "formula": "=SUM(B1:B10)"}'
        ))

        assert result["success"] is True
        assert result["data"]["formula"] == "=SUM(B1:B10)"

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.rename_worksheet")
    def test_update_worksheet_rename(self, mock_rename, mock_context):
        """Test excel_update renames worksheet."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_update

        result = json.loads(excel_update.func(
            "/path/file.xlsx",
            "worksheet",
            '{"name": "OldName", "new_name": "NewName"}'
        ))

        assert result["success"] is True
        assert result["data"]["old_name"] == "OldName"
        assert result["data"]["new_name"] == "NewName"

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.copy_worksheet")
    def test_update_worksheet_copy(self, mock_copy, mock_context):
        """Test excel_update copies worksheet."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        mock_new_sheet = MagicMock()
        mock_new_sheet.Name = "Sheet1 (2)"
        mock_copy.return_value = mock_new_sheet

        from tools.excel_tool import excel_update

        result = json.loads(excel_update.func(
            "/path/file.xlsx",
            "worksheet",
            '{"name": "Sheet1", "copy_to": "Sheet1 Copy"}'
        ))

        assert result["success"] is True
        assert result["data"]["source_worksheet"] == "Sheet1"

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.auto_fit_columns")
    def test_update_structure_autofit(self, mock_autofit, mock_context):
        """Test excel_update auto-fits columns."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_update

        result = json.loads(excel_update.func(
            "/path/file.xlsx",
            "structure",
            '{"worksheet": "Sheet1", "columns": {"A:D": "auto"}}'
        ))

        assert result["success"] is True


# =============================================================================
# CRUD: Delete
# =============================================================================

class TestExcelDelete:
    """Tests for excel_delete tool."""

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.delete_worksheet")
    def test_delete_worksheet(self, mock_delete, mock_context):
        """Test excel_delete deletes a worksheet."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_delete

        result = json.loads(excel_delete.func(
            "/path/file.xlsx",
            "worksheet",
            '{"name": "OldSheet"}'
        ))

        assert result["success"] is True
        assert result["data"]["deleted_worksheet"] == "OldSheet"

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.clear_range")
    def test_delete_range(self, mock_clear, mock_context):
        """Test excel_delete clears a range."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_delete

        result = json.loads(excel_delete.func(
            "/path/file.xlsx",
            "range",
            '{"worksheet": "Sheet1", "range": "A1:D10"}'
        ))

        assert result["success"] is True
        assert result["data"]["cleared_range"] == "A1:D10"

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.delete_rows")
    def test_delete_rows(self, mock_delete_rows, mock_context):
        """Test excel_delete deletes rows."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_delete

        result = json.loads(excel_delete.func(
            "/path/file.xlsx",
            "rows",
            '{"worksheet": "Sheet1", "index": 5, "count": 3}'
        ))

        assert result["success"] is True
        assert result["data"]["count"] == 3

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.workbook_ops.delete_columns")
    def test_delete_columns(self, mock_delete_cols, mock_context):
        """Test excel_delete deletes columns."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_delete

        result = json.loads(excel_delete.func(
            "/path/file.xlsx",
            "columns",
            '{"worksheet": "Sheet1", "index": "C", "count": 2}'
        ))

        assert result["success"] is True
        assert result["data"]["count"] == 2

    @patch("tools.excel_tool.workbook_context")
    @patch("libs.excel_com.formatting_ops.clear_format")
    def test_delete_format(self, mock_clear_format, mock_context):
        """Test excel_delete clears format."""
        mock_manager = MagicMock()
        mock_context.return_value = _fake_workbook_context(mock_manager, MagicMock())

        from tools.excel_tool import excel_delete

        result = json.loads(excel_delete.func(
            "/path/file.xlsx",
            "format",
            '{"worksheet": "Sheet1", "range": "A1:D10"}'
        ))

        assert result["success"] is True


# =============================================================================
# Tools List
# =============================================================================

class TestExcelToolsList:
    """Tests for EXCEL_TOOLS list."""

    def test_tools_list_complete(self):
        """Test EXCEL_TOOLS contains all expected tools."""
        from tools.excel_tool import EXCEL_TOOLS

        tool_names = [tool.name for tool in EXCEL_TOOLS]

        # Lifecycle tools
        assert "excel_open" in tool_names
        assert "excel_save" in tool_names
        assert "excel_info" in tool_names

        # CRUD tools
        assert "excel_create" in tool_names
        assert "excel_read" in tool_names
        assert "excel_update" in tool_names
        assert "excel_delete" in tool_names

    def test_tools_count(self):
        """Test EXCEL_TOOLS has correct number of tools."""
        from tools.excel_tool import EXCEL_TOOLS

        # Total: 3 lifecycle + 4 CRUD = 7
        assert len(EXCEL_TOOLS) == 7
