"""Tests for excel_com/context.py."""

from unittest.mock import MagicMock

import pytest

from libs.excel_com.context import UserStatePreserver, preserve_user_state


class TestUserStatePreserver:
    """Tests for UserStatePreserver class."""

    def test_init(self):
        """Test initialization."""
        mock_app = MagicMock()
        preserver = UserStatePreserver(mock_app)
        assert preserver._app == mock_app
        assert preserver._saved_active_workbook is None
        assert preserver._saved_active_worksheet is None
        assert preserver._saved_selection is None
        assert preserver._saved_active_cell is None
        assert preserver._saved_window_scroll_row is None
        assert preserver._saved_window_scroll_column is None

    def test_save_state_with_active_workbook(self):
        """Test save_state when there's an active workbook."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = MagicMock()
        mock_app.ActiveWorkbook.Name = "test.xlsx"
        mock_app.ActiveSheet = MagicMock()
        mock_app.ActiveSheet.Name = "Sheet1"
        mock_app.Selection = MagicMock()
        mock_app.ActiveCell = MagicMock()
        mock_app.ActiveWindow = MagicMock()
        mock_app.ActiveWindow.ScrollRow = 5
        mock_app.ActiveWindow.ScrollColumn = 3

        preserver = UserStatePreserver(mock_app)
        preserver.save_state()

        assert preserver._saved_active_workbook is not None
        assert preserver._saved_active_worksheet is not None
        assert preserver._saved_selection is not None
        assert preserver._saved_active_cell is not None
        assert preserver._saved_window_scroll_row == 5
        assert preserver._saved_window_scroll_column == 3

    def test_save_state_without_active_workbook(self):
        """Test save_state when there's no active workbook."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = None

        preserver = UserStatePreserver(mock_app)
        preserver.save_state()

        assert preserver._saved_active_workbook is None
        assert preserver._saved_active_worksheet is None

    def test_save_state_handles_selection_exception(self):
        """Test save_state handles exceptions when getting Selection."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = MagicMock()
        mock_app.Selection = MagicMock(side_effect=Exception("No selection"))
        # Use property mock to raise exception on access
        type(mock_app).Selection = property(lambda self: (_ for _ in ()).throw(Exception("No selection")))

        preserver = UserStatePreserver(mock_app)
        # Should not raise
        preserver.save_state()
        assert preserver._saved_selection is None

    def test_save_state_handles_active_cell_exception(self):
        """Test save_state handles exceptions when getting ActiveCell."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = MagicMock()
        mock_app.ActiveCell = None  # Simulate no active cell

        preserver = UserStatePreserver(mock_app)
        preserver.save_state()
        # Should complete without error
        assert preserver._saved_active_cell is None

    def test_save_state_handles_window_exception(self):
        """Test save_state handles exceptions when accessing ActiveWindow."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = MagicMock()
        mock_app.ActiveWindow = None

        preserver = UserStatePreserver(mock_app)
        preserver.save_state()
        # Should complete without error
        assert preserver._saved_window_scroll_row is None
        assert preserver._saved_window_scroll_column is None

    def test_restore_state_with_saved_workbook(self):
        """Test restore_state restores saved workbook and worksheet."""
        mock_app = MagicMock()
        mock_workbook = MagicMock()
        mock_workbook.Name = "test.xlsx"
        mock_workbook.Activate = MagicMock()

        mock_worksheet = MagicMock()
        mock_worksheet.Name = "Sheet1"
        mock_worksheet.Activate = MagicMock()

        mock_app.ActiveWindow = MagicMock()

        preserver = UserStatePreserver(mock_app)
        preserver._saved_active_workbook = mock_workbook
        preserver._saved_active_worksheet = mock_worksheet
        preserver._saved_window_scroll_row = 10
        preserver._saved_window_scroll_column = 5

        preserver.restore_state()

        mock_workbook.Activate.assert_called_once()
        mock_worksheet.Activate.assert_called_once()
        assert mock_app.ActiveWindow.ScrollRow == 10
        assert mock_app.ActiveWindow.ScrollColumn == 5

    def test_restore_state_without_saved_workbook(self):
        """Test restore_state when no workbook was saved."""
        mock_app = MagicMock()

        preserver = UserStatePreserver(mock_app)
        preserver.restore_state()
        # Should not raise any errors

    def test_restore_state_handles_closed_workbook(self):
        """Test restore_state handles case where workbook was closed."""
        mock_app = MagicMock()
        mock_workbook = MagicMock()
        # Simulate workbook was closed - accessing Name raises exception
        mock_workbook.Name = MagicMock(side_effect=Exception("Workbook closed"))
        type(mock_workbook).Name = property(lambda self: (_ for _ in ()).throw(Exception("Workbook closed")))

        preserver = UserStatePreserver(mock_app)
        preserver._saved_active_workbook = mock_workbook
        # Should not raise
        preserver.restore_state()

    def test_restore_state_handles_deleted_worksheet(self):
        """Test restore_state handles case where worksheet was deleted."""
        mock_app = MagicMock()
        mock_workbook = MagicMock()
        mock_workbook.Name = "test.xlsx"
        mock_workbook.Activate = MagicMock()

        mock_worksheet = MagicMock()
        mock_worksheet.Name = MagicMock(side_effect=Exception("Worksheet deleted"))
        type(mock_worksheet).Name = property(lambda self: (_ for _ in ()).throw(Exception("Worksheet deleted")))

        preserver = UserStatePreserver(mock_app)
        preserver._saved_active_workbook = mock_workbook
        preserver._saved_active_worksheet = mock_worksheet
        # Should not raise
        preserver.restore_state()


class TestPreserveUserStateContextManager:
    """Tests for preserve_user_state context manager."""

    def test_context_manager_saves_and_restores_state(self):
        """Test that context manager saves and restores state."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = MagicMock()
        mock_app.ActiveWorkbook.Name = "test.xlsx"
        mock_app.ActiveWorkbook.Activate = MagicMock()
        mock_app.ActiveSheet = MagicMock()
        mock_app.ActiveSheet.Name = "Sheet1"
        mock_app.ActiveSheet.Activate = MagicMock()
        mock_app.ActiveWindow = MagicMock()
        mock_app.ActiveWindow.ScrollRow = 1
        mock_app.ActiveWindow.ScrollColumn = 1

        with preserve_user_state(mock_app):
            # Inside context, state should be saved
            pass
        # After context, state should be restored
        mock_app.ActiveWorkbook.Activate.assert_called()

    def test_context_manager_restores_on_exception(self):
        """Test that state is restored even if exception occurs."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = MagicMock()
        mock_app.ActiveWorkbook.Name = "test.xlsx"
        mock_app.ActiveWorkbook.Activate = MagicMock()
        mock_app.ActiveWindow = MagicMock()

        with pytest.raises(ValueError):
            with preserve_user_state(mock_app):
                raise ValueError("Test exception")

        # State should still be restored
        mock_app.ActiveWorkbook.Activate.assert_called()

    def test_context_manager_yields_none(self):
        """Test that context manager yields None."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = MagicMock()
        mock_app.ActiveWindow = MagicMock()

        with preserve_user_state(mock_app) as result:
            assert result is None

    def test_context_manager_with_no_active_workbook(self):
        """Test context manager when there's no active workbook."""
        mock_app = MagicMock()
        mock_app.ActiveWorkbook = None

        # Should not raise
        with preserve_user_state(mock_app):
            pass
