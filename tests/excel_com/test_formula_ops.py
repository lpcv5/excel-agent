"""Tests for excel_com/formula_ops.py."""

from unittest.mock import MagicMock, patch

import pytest

from excel_com.manager import ExcelAppManager
from excel_com import formula_ops


class TestSetFormula:
    """Tests for set_formula function."""

    def _create_mock_preserve_context(self):
        """Create a mock context manager for preserve_user_state."""
        mock_ctx = MagicMock()
        mock_ctx.__enter__ = MagicMock(return_value=None)
        mock_ctx.__exit__ = MagicMock(return_value=None)
        return mock_ctx

    def test_set_formula(self):
        """Test set_formula sets formula in cell."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("excel_com.formula_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formula_ops.set_formula(
                manager, mock_workbook, "Sheet1", "A1", "=SUM(B1:B10)"
            )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        mock_worksheet.Range.assert_called_once_with("A1")
        assert mock_range.Formula == "=SUM(B1:B10)"

    def test_set_formula_preserves_user_state(self):
        """Test that set_formula uses preserve_user_state."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        with patch("excel_com.formula_ops.preserve_user_state") as mock_preserve:
            mock_preserve.return_value.__enter__ = MagicMock(return_value=None)
            mock_preserve.return_value.__exit__ = MagicMock(return_value=None)

            formula_ops.set_formula(
                manager, mock_workbook, "Sheet1", "A1", "=1+1"
            )

            mock_preserve.assert_called_once_with(manager.app)

    def test_set_formula_for_range(self):
        """Test set_formula for a range of cells."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("excel_com.formula_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formula_ops.set_formula(
                manager, mock_workbook, "Sheet1", "A1:A10", "=ROW()"
            )

        mock_worksheet.Range.assert_called_once_with("A1:A10")
        assert mock_range.Formula == "=ROW()"


class TestGetFormula:
    """Tests for get_formula function."""

    def test_get_formula(self):
        """Test get_formula returns formula from cell."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Formula = "=SUM(A1:A10)"
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        result = formula_ops.get_formula(
            manager, mock_workbook, "Sheet1", "A1"
        )

        manager.get_worksheet.assert_called_once_with(mock_workbook, "Sheet1")
        mock_worksheet.Range.assert_called_once_with("A1")
        assert result == "=SUM(A1:A10)"

    def test_get_formula_empty_cell(self):
        """Test get_formula returns empty string for empty cell."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Formula = None
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        result = formula_ops.get_formula(
            manager, mock_workbook, "Sheet1", "A1"
        )

        assert result == ""

    def test_get_formula_value_not_formula(self):
        """Test get_formula returns value when cell has value not formula."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Formula = "42"  # Just a value, not a formula
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        result = formula_ops.get_formula(
            manager, mock_workbook, "Sheet1", "A1"
        )

        assert result == "42"

    def test_get_formula_from_range(self):
        """Test get_formula handles range formula result."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        # When reading from a range, Formula can be a 2D array
        mock_range.Formula = [["=A1+B1", "=C1+D1"], ["=A2+B2", "=C2+D2"]]
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        result = formula_ops.get_formula(
            manager, mock_workbook, "Sheet1", "A1:B2"
        )

        # Should return first formula from the 2D array
        assert result == "=A1+B1"

    def test_get_formula_from_empty_range(self):
        """Test get_formula handles empty range."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Formula = []  # Empty range result
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        result = formula_ops.get_formula(
            manager, mock_workbook, "Sheet1", "A1"
        )

        assert result == ""

    def test_get_formula_from_range_with_empty_first_row(self):
        """Test get_formula handles range with empty first row."""
        manager = MagicMock(spec=ExcelAppManager)

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Formula = [[]]  # Empty first row
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        result = formula_ops.get_formula(
            manager, mock_workbook, "Sheet1", "A1"
        )

        assert result == ""
