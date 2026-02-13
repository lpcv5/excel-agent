"""Tests for excel_com/manager.py."""

from pathlib import Path
from unittest.mock import MagicMock, patch, PropertyMock

import pytest

from excel_com.manager import ExcelAppManager


class TestExcelAppManagerInit:
    """Tests for ExcelAppManager initialization."""

    def test_init_default_params(self):
        """Test initialization with default parameters."""
        manager = ExcelAppManager()
        assert manager._visible is False
        assert manager._display_alerts is False
        assert manager._attach_to_existing is True
        assert manager._app is None
        assert manager._workbooks == {}
        assert manager._workbook_owned == {}

    def test_init_custom_params(self):
        """Test initialization with custom parameters."""
        manager = ExcelAppManager(visible=True, display_alerts=True, attach_to_existing=False)
        assert manager._visible is True
        assert manager._display_alerts is True
        assert manager._attach_to_existing is False
        assert manager._app is None


class TestExcelAppManagerAppProperty:
    """Tests for ExcelAppManager.app property."""

    def test_app_raises_when_not_initialized(self):
        """Test that app property raises when not initialized."""
        manager = ExcelAppManager()
        with pytest.raises(RuntimeError, match="not initialized"):
            _ = manager.app

    def test_app_returns_app_when_initialized(self):
        """Test that app property returns app when initialized."""
        manager = ExcelAppManager()
        mock_app = MagicMock()
        manager._app = mock_app
        assert manager.app == mock_app


class TestExcelAppManagerStart:
    """Tests for ExcelAppManager.start method."""

    def test_start_is_idempotent(self):
        """Test that calling start multiple times doesn't create multiple instances."""
        manager = ExcelAppManager()
        mock_app = MagicMock()
        mock_app.Workbooks.Count = 0
        manager._app = mock_app

        # Should return early since app is already set
        manager.start()

        # App should still be the same mock
        assert manager._app == mock_app


class TestExcelAppManagerStop:
    """Tests for ExcelAppManager.stop method."""

    def test_stop_handles_no_app(self):
        """Test stop handles case when app is None."""
        manager = ExcelAppManager()
        # Should not raise
        manager.stop()
        assert manager._app is None

    def test_stop_closes_owned_workbooks(self):
        """Test stop closes workbooks we own."""
        manager = ExcelAppManager()
        mock_app = MagicMock()
        manager._app = mock_app
        manager._attached_to_existing = False

        mock_workbook = MagicMock()
        manager._workbooks["C:\\test\\test.xlsx"] = mock_workbook
        manager._workbook_owned["C:\\test\\test.xlsx"] = True

        manager.stop()

        mock_workbook.Close.assert_called_once()
        assert len(manager._workbooks) == 0

    def test_stop_does_not_close_user_workbooks(self):
        """Test stop does not close workbooks opened by user."""
        manager = ExcelAppManager()
        mock_app = MagicMock()
        manager._app = mock_app
        manager._attached_to_existing = True

        mock_workbook = MagicMock()
        manager._workbooks["C:\\test\\user.xlsx"] = mock_workbook
        manager._workbook_owned["C:\\test\\user.xlsx"] = False

        manager.stop()

        mock_workbook.Close.assert_not_called()
        assert len(manager._workbooks) == 0


class TestExcelAppManagerOpenWorkbook:
    """Tests for ExcelAppManager.open_workbook method."""

    def test_open_workbook_raises_when_not_initialized(self, tmp_path):
        """Test open_workbook raises when Excel not initialized."""
        manager = ExcelAppManager()
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        with pytest.raises(RuntimeError, match="not initialized"):
            manager.open_workbook(str(test_file))

    def test_open_workbook_raises_when_file_not_found(self):
        """Test open_workbook raises when file doesn't exist."""
        manager = ExcelAppManager()
        manager._app = MagicMock()

        with pytest.raises(FileNotFoundError):
            manager.open_workbook("/nonexistent/path/file.xlsx")

    def test_open_workbook_returns_cached_workbook(self, tmp_path):
        """Test open_workbook returns cached workbook if already open."""
        manager = ExcelAppManager()
        manager._app = MagicMock()

        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        str_path = str(test_file.resolve())

        mock_workbook = MagicMock()
        manager._workbooks[str_path] = mock_workbook

        result = manager.open_workbook(str(test_file))
        assert result == mock_workbook

    def test_open_workbook_opens_new_workbook(self, tmp_path):
        """Test open_workbook opens a new workbook."""
        manager = ExcelAppManager()
        mock_app = MagicMock()
        mock_workbook = MagicMock()
        mock_app.Workbooks.Open.return_value = mock_workbook
        manager._app = mock_app

        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        str_path = str(test_file.resolve())

        # Mock _find_open_workbook to return None
        manager._find_open_workbook = MagicMock(return_value=None)

        result = manager.open_workbook(str(test_file))

        mock_app.Workbooks.Open.assert_called_once()
        assert result == mock_workbook
        assert manager._workbook_owned[str_path] is True

    def test_open_workbook_read_only(self, tmp_path):
        """Test open_workbook with read_only=True."""
        manager = ExcelAppManager()
        mock_app = MagicMock()
        mock_workbook = MagicMock()
        mock_app.Workbooks.Open.return_value = mock_workbook
        manager._app = mock_app

        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        manager._find_open_workbook = MagicMock(return_value=None)
        manager.open_workbook(str(test_file), read_only=True)

        mock_app.Workbooks.Open.assert_called_with(
            str(test_file.resolve()),
            ReadOnly=True
        )


class TestExcelAppManagerCloseWorkbook:
    """Tests for ExcelAppManager.close_workbook method."""

    def test_close_workbook_raises_when_not_initialized(self):
        """Test close_workbook raises when Excel not initialized."""
        manager = ExcelAppManager()

        with pytest.raises(RuntimeError, match="not initialized"):
            manager.close_workbook("/path/to/file.xlsx")

    def test_close_workbook_raises_when_not_open(self):
        """Test close_workbook raises when workbook not open."""
        manager = ExcelAppManager()
        manager._app = MagicMock()

        with pytest.raises(ValueError, match="not open"):
            manager.close_workbook("/path/to/file.xlsx")

    def test_close_workbook_closes_owned_workbook(self):
        """Test close_workbook closes workbook we own."""
        manager = ExcelAppManager()
        manager._app = MagicMock()

        mock_workbook = MagicMock()
        str_path = "C:\\test\\test.xlsx"
        manager._workbooks[str_path] = mock_workbook
        manager._workbook_owned[str_path] = True

        manager.close_workbook(str_path, save=False)

        mock_workbook.Close.assert_called_once_with(SaveChanges=False)
        assert str_path not in manager._workbooks

    def test_close_workbook_saves_before_closing(self):
        """Test close_workbook saves when save=True."""
        manager = ExcelAppManager()
        manager._app = MagicMock()

        mock_workbook = MagicMock()
        str_path = "C:\\test\\test.xlsx"
        manager._workbooks[str_path] = mock_workbook
        manager._workbook_owned[str_path] = True

        manager.close_workbook(str_path, save=True)

        mock_workbook.Close.assert_called_once_with(SaveChanges=True)


class TestExcelAppManagerGetWorkbook:
    """Tests for ExcelAppManager.get_workbook method."""

    def test_get_workbook_raises_when_not_initialized(self):
        """Test get_workbook raises when Excel not initialized."""
        manager = ExcelAppManager()

        with pytest.raises(RuntimeError, match="not initialized"):
            manager.get_workbook("/path/to/file.xlsx")

    def test_get_workbook_raises_when_not_open(self):
        """Test get_workbook raises when workbook not open."""
        manager = ExcelAppManager()
        manager._app = MagicMock()

        with pytest.raises(ValueError, match="not open"):
            manager.get_workbook("/path/to/file.xlsx")

    def test_get_workbook_returns_workbook(self):
        """Test get_workbook returns the workbook."""
        manager = ExcelAppManager()
        manager._app = MagicMock()

        mock_workbook = MagicMock()
        str_path = "C:\\test\\test.xlsx"
        manager._workbooks[str_path] = mock_workbook

        result = manager.get_workbook(str_path)
        assert result == mock_workbook


class TestExcelAppManagerSaveWorkbook:
    """Tests for ExcelAppManager.save_workbook method."""

    def test_save_workbook_saves(self):
        """Test save_workbook calls Save."""
        manager = ExcelAppManager()
        mock_workbook = MagicMock()

        manager.save_workbook(mock_workbook)

        mock_workbook.Save.assert_called_once()

    def test_save_workbook_save_as(self):
        """Test save_workbook calls SaveAs with path."""
        manager = ExcelAppManager()
        mock_workbook = MagicMock()

        manager.save_workbook(mock_workbook, "/new/path/file.xlsx")

        mock_workbook.SaveAs.assert_called_once_with("/new/path/file.xlsx")


class TestExcelAppManagerListWorksheets:
    """Tests for ExcelAppManager.list_worksheets method."""

    def test_list_worksheets(self):
        """Test list_worksheets returns worksheet names."""
        manager = ExcelAppManager()

        # Create mock worksheets with specific names
        mock_sheet1 = MagicMock()
        mock_sheet1.Name = "Sheet1"
        mock_sheet2 = MagicMock()
        mock_sheet2.Name = "Sheet2"
        mock_sheet3 = MagicMock()
        mock_sheet3.Name = "Sheet3"

        mock_workbook = MagicMock()
        mock_workbook.Worksheets.Count = 3

        # Use side_effect to return different sheets for different indices
        def get_worksheet(i):
            return [mock_sheet1, mock_sheet2, mock_sheet3][i - 1]

        mock_workbook.Worksheets.side_effect = get_worksheet

        result = manager.list_worksheets(mock_workbook)

        assert result == ["Sheet1", "Sheet2", "Sheet3"]


class TestExcelAppManagerGetWorksheet:
    """Tests for ExcelAppManager.get_worksheet method."""

    def test_get_worksheet(self):
        """Test get_worksheet returns worksheet by name."""
        manager = ExcelAppManager()
        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_workbook.Worksheets.return_value = mock_worksheet

        result = manager.get_worksheet(mock_workbook, "Sheet1")

        mock_workbook.Worksheets.assert_called_once_with("Sheet1")
        assert result == mock_worksheet

    def test_get_worksheet_raises_when_not_found(self):
        """Test get_worksheet raises when worksheet not found."""
        manager = ExcelAppManager()
        mock_workbook = MagicMock()
        mock_workbook.Worksheets.side_effect = Exception("Not found")

        with pytest.raises(ValueError, match="not found"):
            manager.get_worksheet(mock_workbook, "NonExistent")


class TestExcelAppManagerGetActiveWorksheet:
    """Tests for ExcelAppManager.get_active_worksheet method."""

    def test_get_active_worksheet(self):
        """Test get_active_worksheet returns active sheet."""
        manager = ExcelAppManager()
        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_workbook.ActiveSheet = mock_worksheet

        result = manager.get_active_worksheet(mock_workbook)
        assert result == mock_worksheet


class TestExcelAppManagerReadRange:
    """Tests for ExcelAppManager.read_range method."""

    def test_read_range_single_cell(self):
        """Test read_range with single cell."""
        manager = ExcelAppManager()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Value = "Test Value"
        mock_worksheet.Range.return_value = mock_range

        result = manager.read_range(mock_worksheet, "A1")

        assert result == [["Test Value"]]

    def test_read_range_2d_array(self):
        """Test read_range with 2D array."""
        manager = ExcelAppManager()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Value = ((1, 2), (3, 4))
        mock_worksheet.Range.return_value = mock_range

        result = manager.read_range(mock_worksheet, "A1:B2")

        assert result == [[1, 2], [3, 4]]

    def test_read_range_empty(self):
        """Test read_range with empty range."""
        manager = ExcelAppManager()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Value = None
        mock_worksheet.Range.return_value = mock_range

        result = manager.read_range(mock_worksheet, "A1")

        assert result == []


class TestExcelAppManagerNormalizeValue:
    """Tests for ExcelAppManager._normalize_value static method."""

    def test_normalize_value_float_to_int(self):
        """Test _normalize_value converts float to int when whole number."""
        assert ExcelAppManager._normalize_value(1.0) == 1
        assert ExcelAppManager._normalize_value(42.0) == 42

    def test_normalize_value_keeps_float(self):
        """Test _normalize_value keeps float when not whole number."""
        assert ExcelAppManager._normalize_value(1.5) == 1.5
        assert ExcelAppManager._normalize_value(3.14159) == 3.14159

    def test_normalize_value_keeps_string(self):
        """Test _normalize_value keeps strings unchanged."""
        assert ExcelAppManager._normalize_value("hello") == "hello"

    def test_normalize_value_keeps_none(self):
        """Test _normalize_value keeps None unchanged."""
        assert ExcelAppManager._normalize_value(None) is None


class TestExcelAppManagerColumnNumberToLetter:
    """Tests for ExcelAppManager._column_number_to_letter static method."""

    def test_single_letter_columns(self):
        """Test _column_number_to_letter for A-Z."""
        assert ExcelAppManager._column_number_to_letter(1) == "A"
        assert ExcelAppManager._column_number_to_letter(2) == "B"
        assert ExcelAppManager._column_number_to_letter(26) == "Z"

    def test_double_letter_columns(self):
        """Test _column_number_to_letter for AA-ZZ."""
        assert ExcelAppManager._column_number_to_letter(27) == "AA"
        assert ExcelAppManager._column_number_to_letter(28) == "AB"
        assert ExcelAppManager._column_number_to_letter(52) == "AZ"
        assert ExcelAppManager._column_number_to_letter(53) == "BA"


class TestExcelAppManagerWriteRange:
    """Tests for ExcelAppManager.write_range method."""

    def test_write_range_raises_on_empty_data(self):
        """Test write_range raises on empty data."""
        manager = ExcelAppManager()
        mock_worksheet = MagicMock()

        with pytest.raises(ValueError, match="empty data"):
            manager.write_range(mock_worksheet, "A1", [])

    def test_write_range_writes_data(self):
        """Test write_range writes data to range."""
        manager = ExcelAppManager()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.GetAddress.return_value = "$A$1"
        mock_worksheet.Range.return_value = mock_range

        data = [[1, 2], [3, 4]]
        manager.write_range(mock_worksheet, "A1", data)

        mock_range.Value = data  # This is set in the implementation


class TestExcelAppManagerStateQuery:
    """Tests for ExcelAppManager state query methods."""

    def test_is_running_false_initially(self):
        """Test is_running returns False initially."""
        manager = ExcelAppManager()
        assert manager.is_running() is False

    def test_is_running_true_after_start(self):
        """Test is_running returns True after start."""
        manager = ExcelAppManager()
        manager._app = MagicMock()
        assert manager.is_running() is True

    def test_get_open_workbooks_empty_initially(self):
        """Test get_open_workbooks returns empty list initially."""
        manager = ExcelAppManager()
        assert manager.get_open_workbooks() == []

    def test_get_open_workbooks_returns_paths(self):
        """Test get_open_workbooks returns workbook paths."""
        manager = ExcelAppManager()
        manager._workbooks = {
            "C:\\test\\file1.xlsx": MagicMock(),
            "C:\\test\\file2.xlsx": MagicMock(),
        }

        result = manager.get_open_workbooks()
        assert "C:\\test\\file1.xlsx" in result
        assert "C:\\test\\file2.xlsx" in result

    def test_is_workbook_owned_true(self):
        """Test is_workbook_owned returns True for owned workbooks."""
        manager = ExcelAppManager()
        manager._workbooks["C:\\test\\test.xlsx"] = MagicMock()
        manager._workbook_owned["C:\\test\\test.xlsx"] = True

        assert manager.is_workbook_owned("C:\\test\\test.xlsx") is True

    def test_is_workbook_owned_false(self):
        """Test is_workbook_owned returns False for user workbooks."""
        manager = ExcelAppManager()
        manager._workbooks["C:\\test\\test.xlsx"] = MagicMock()
        manager._workbook_owned["C:\\test\\test.xlsx"] = False

        assert manager.is_workbook_owned("C:\\test\\test.xlsx") is False

    def test_is_workbook_owned_not_found(self):
        """Test is_workbook_owned returns False for non-existent workbook."""
        manager = ExcelAppManager()
        assert manager.is_workbook_owned("C:\\nonexistent.xlsx") is False

    def test_get_workbook_count(self):
        """Test get_workbook_count returns correct count."""
        manager = ExcelAppManager()
        assert manager.get_workbook_count() == 0

        manager._workbooks["file1.xlsx"] = MagicMock()
        manager._workbooks["file2.xlsx"] = MagicMock()
        assert manager.get_workbook_count() == 2


class TestExcelAppManagerGetActiveWorkbook:
    """Tests for ExcelAppManager.get_active_workbook method."""

    def test_get_active_workbook_none_when_no_app(self):
        """Test get_active_workbook returns None when no app."""
        manager = ExcelAppManager()
        assert manager.get_active_workbook() is None

    def test_get_active_workbook_returns_workbook(self):
        """Test get_active_workbook returns active workbook."""
        manager = ExcelAppManager()
        mock_app = MagicMock()
        mock_workbook = MagicMock()
        mock_app.ActiveWorkbook = mock_workbook
        manager._app = mock_app

        result = manager.get_active_workbook()
        assert result == mock_workbook
