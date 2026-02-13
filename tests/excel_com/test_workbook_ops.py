"""Tests for excel_com/workbook_ops.py."""

from unittest.mock import MagicMock, patch

import pytest

from excel_com.manager import ExcelAppManager
from excel_com import workbook_ops


class TestOpenWorkbook:
    """Tests for open_workbook function."""

    def test_open_workbook(self, tmp_path):
        """Test open_workbook opens a workbook and returns info."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()
        mock_workbook.Name = "test.xlsx"

        manager.open_workbook.return_value = mock_workbook
        manager.list_worksheets.return_value = ["Sheet1", "Sheet2", "Sheet3"]

        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        result_workbook, worksheets = workbook_ops.open_workbook(
            manager, str(test_file), read_only=False
        )

        assert result_workbook == mock_workbook
        assert worksheets == ["Sheet1", "Sheet2", "Sheet3"]
        manager.open_workbook.assert_called_once_with(str(test_file), False)

    def test_open_workbook_read_only(self, tmp_path):
        """Test open_workbook with read_only=True."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        manager.open_workbook.return_value = mock_workbook
        manager.list_worksheets.return_value = []

        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        workbook_ops.open_workbook(manager, str(test_file), read_only=True)

        manager.open_workbook.assert_called_once_with(str(test_file), True)


class TestCloseWorkbook:
    """Tests for close_workbook function."""

    def test_close_workbook(self):
        """Test close_workbook closes the workbook."""
        manager = MagicMock(spec=ExcelAppManager)

        workbook_ops.close_workbook(manager, "/path/to/file.xlsx", save=True)

        manager.close_workbook.assert_called_once_with("/path/to/file.xlsx", True)

    def test_close_workbook_without_save(self):
        """Test close_workbook without saving."""
        manager = MagicMock(spec=ExcelAppManager)

        workbook_ops.close_workbook(manager, "/path/to/file.xlsx", save=False)

        manager.close_workbook.assert_called_once_with("/path/to/file.xlsx", False)


class TestSaveWorkbook:
    """Tests for save_workbook function."""

    def test_save_workbook(self):
        """Test save_workbook saves the workbook."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        workbook_ops.save_workbook(manager, mock_workbook)

        manager.save_workbook.assert_called_once_with(mock_workbook, None)

    def test_save_workbook_as(self):
        """Test save_workbook with new path."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        workbook_ops.save_workbook(manager, mock_workbook, "/new/path/file.xlsx")

        manager.save_workbook.assert_called_once_with(mock_workbook, "/new/path/file.xlsx")


class TestAddWorksheet:
    """Tests for add_worksheet function."""

    def test_add_worksheet_at_end(self):
        """Test add_worksheet adds at end by default."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()
        mock_workbook.Worksheets.Count = 2

        mock_new_sheet = MagicMock()
        mock_workbook.Worksheets.Add.return_value = mock_new_sheet

        result = workbook_ops.add_worksheet(manager, mock_workbook, "NewSheet")

        mock_workbook.Worksheets.Add.assert_called_once()
        assert mock_new_sheet.Name == "NewSheet"
        assert result == mock_new_sheet

    def test_add_worksheet_after_specific(self):
        """Test add_worksheet adds after specific sheet."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        mock_after_sheet = MagicMock()
        manager.get_worksheet.return_value = mock_after_sheet

        mock_new_sheet = MagicMock()
        mock_workbook.Worksheets.Add.return_value = mock_new_sheet

        result = workbook_ops.add_worksheet(
            manager, mock_workbook, "NewSheet", after="Sheet1"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        mock_workbook.Worksheets.Add.assert_called_once_with(After=mock_after_sheet)
        assert result == mock_new_sheet


class TestDeleteWorksheet:
    """Tests for delete_worksheet function."""

    def test_delete_worksheet(self):
        """Test delete_worksheet deletes the worksheet."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        mock_worksheet = MagicMock()
        manager.get_worksheet.return_value = mock_worksheet

        workbook_ops.delete_worksheet(manager, mock_workbook, "SheetToDelete")

        manager.get_worksheet.assert_called_once_with(mock_workbook, "SheetToDelete")
        mock_worksheet.Delete.assert_called_once()


class TestRenameWorksheet:
    """Tests for rename_worksheet function."""

    def test_rename_worksheet(self):
        """Test rename_worksheet renames the worksheet."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        mock_worksheet = MagicMock()
        manager.get_worksheet.return_value = mock_worksheet

        result = workbook_ops.rename_worksheet(
            manager, mock_workbook, "OldName", "NewName"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "OldName")
        assert mock_worksheet.Name == "NewName"
        assert result == mock_worksheet


class TestCopyWorksheet:
    """Tests for copy_worksheet function."""

    def test_copy_worksheet_with_new_name(self):
        """Test copy_worksheet with new name."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        mock_worksheet = MagicMock()
        manager.get_worksheet.return_value = mock_worksheet

        mock_new_sheet = MagicMock()
        mock_workbook.ActiveSheet = mock_new_sheet

        result = workbook_ops.copy_worksheet(
            manager, mock_workbook, "SheetToCopy", new_name="CopiedSheet"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "SheetToCopy")
        mock_worksheet.Copy.assert_called_once_with(After=mock_worksheet)
        assert result == mock_new_sheet

    def test_copy_worksheet_without_new_name(self):
        """Test copy_worksheet without new name (auto-name)."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        mock_worksheet = MagicMock()
        manager.get_worksheet.return_value = mock_worksheet

        mock_new_sheet = MagicMock()
        mock_workbook.ActiveSheet = mock_new_sheet

        result = workbook_ops.copy_worksheet(
            manager, mock_workbook, "SheetToCopy"
        )

        mock_worksheet.Copy.assert_called_once()
        # Name should not be set when new_name is None
        assert result == mock_new_sheet


class TestGetUsedRange:
    """Tests for get_used_range function."""

    def test_get_used_range(self):
        """Test get_used_range returns range info."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        mock_worksheet = MagicMock()
        mock_used_range = MagicMock()
        mock_used_range.Address = "$A$1:$D$10"
        mock_used_range.Rows.Count = 10
        mock_used_range.Columns.Count = 4
        mock_worksheet.UsedRange = mock_used_range

        manager.get_worksheet.return_value = mock_worksheet

        address, rows, cols = workbook_ops.get_used_range(
            manager, mock_workbook, "Sheet1"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        assert address == "$A$1:$D$10"
        assert rows == 10
        assert cols == 4


class TestReadRange:
    """Tests for read_range function."""

    def test_read_range(self):
        """Test read_range reads data from range."""
        manager = MagicMock(spec=ExcelAppManager)
        mock_workbook = MagicMock()

        mock_worksheet = MagicMock()
        manager.get_worksheet.return_value = mock_worksheet

        sample_data = [[1, 2], [3, 4]]
        manager.read_range.return_value = sample_data

        result = workbook_ops.read_range(
            manager, mock_workbook, "Sheet1", "A1:B2"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        manager.read_range.assert_called_once_with(mock_worksheet, "A1:B2")
        assert result == sample_data


class TestWriteRange:
    """Tests for write_range function."""

    def test_write_range(self):
        """Test write_range writes data to range."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()

        # Set up the context manager mock
        mock_preserve = MagicMock()
        mock_preserve.__enter__ = MagicMock(return_value=None)
        mock_preserve.__exit__ = MagicMock(return_value=None)

        manager.get_worksheet.return_value = mock_worksheet

        sample_data = [[1, 2], [3, 4]]

        with patch("excel_com.workbook_ops.preserve_user_state", return_value=mock_preserve):
            workbook_ops.write_range(
                manager, mock_workbook, "Sheet1", "A1", sample_data
            )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        manager.write_range.assert_called_once_with(mock_worksheet, "A1", sample_data)

    def test_write_range_preserves_user_state(self):
        """Test write_range uses preserve_user_state context manager."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        manager.get_worksheet.return_value = mock_worksheet

        sample_data = [[1, 2]]

        with patch("excel_com.workbook_ops.preserve_user_state") as mock_preserve_ctx:
            mock_preserve_ctx.return_value.__enter__ = MagicMock(return_value=None)
            mock_preserve_ctx.return_value.__exit__ = MagicMock(return_value=None)

            workbook_ops.write_range(
                manager, mock_workbook, "Sheet1", "A1", sample_data
            )

            mock_preserve_ctx.assert_called_once_with(manager.app)
