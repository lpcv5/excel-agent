"""Tests for excel_com/formatting_ops.py."""

from unittest.mock import MagicMock, patch


from libs.excel_com.manager import ExcelAppManager
from libs.excel_com import formatting_ops
from libs.excel_com.constants import (
    XL_UNDERLINE_STYLE_SINGLE,
    XL_UNDERLINE_STYLE_NONE,
    HORIZONTAL_ALIGNMENT_MAP,
    VERTICAL_ALIGNMENT_MAP,
    BORDER_STYLE_MAP,
    BORDER_EDGE_MAP,
)


class TestSetFontFormat:
    """Tests for set_font_format function."""

    def _create_mock_preserve_context(self):
        """Create a mock context manager for preserve_user_state."""
        mock_ctx = MagicMock()
        mock_ctx.__enter__ = MagicMock(return_value=None)
        mock_ctx.__exit__ = MagicMock(return_value=None)
        return mock_ctx

    def test_set_font_format_full(self):
        """Test set_font_format with all options."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_font = MagicMock()
        mock_range.Font = mock_font
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_font_format(
                manager, mock_workbook, "Sheet1", "A1:D10",
                font_name="Arial",
                size=12,
                bold=True,
                italic=True,
                underline=True,
                color="FF0000"
            )

        assert mock_font.Name == "Arial"
        assert mock_font.Size == 12
        assert mock_font.Bold is True
        assert mock_font.Italic is True
        assert mock_font.Underline == XL_UNDERLINE_STYLE_SINGLE
        assert mock_font.Color == int("FF0000", 16)

    def test_set_font_format_partial(self):
        """Test set_font_format with partial options."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_font = MagicMock()
        mock_range.Font = mock_font
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_font_format(
                manager, mock_workbook, "Sheet1", "A1:D10",
                bold=True,
                color="0000FF"
            )

        # Only bold and color should be set
        assert mock_font.Bold is True
        assert mock_font.Color == int("0000FF", 16)
        # Others should not have been called
        assert not hasattr(mock_font, 'Name') or mock_font.Name != "Arial"

    def test_set_font_format_underline_false(self):
        """Test set_font_format with underline=False."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_font = MagicMock()
        mock_range.Font = mock_font
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_font_format(
                manager, mock_workbook, "Sheet1", "A1:D10",
                underline=False
            )

        assert mock_font.Underline == XL_UNDERLINE_STYLE_NONE

    def test_set_font_format_preserves_user_state(self):
        """Test that set_font_format uses preserve_user_state."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_range.Font = MagicMock()
        mock_worksheet.Range.return_value = mock_range
        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state") as mock_preserve:
            mock_preserve.return_value.__enter__ = MagicMock(return_value=None)
            mock_preserve.return_value.__exit__ = MagicMock(return_value=None)

            formatting_ops.set_font_format(
                manager, mock_workbook, "Sheet1", "A1", bold=True
            )

            mock_preserve.assert_called_once_with(manager.app)


class TestSetCellFormat:
    """Tests for set_cell_format function."""

    def _create_mock_preserve_context(self):
        """Create a mock context manager."""
        mock_ctx = MagicMock()
        mock_ctx.__enter__ = MagicMock(return_value=None)
        mock_ctx.__exit__ = MagicMock(return_value=None)
        return mock_ctx

    def test_set_cell_format(self):
        """Test set_cell_format with all options."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_cell_format(
                manager, mock_workbook, "Sheet1", "A1:D10",
                horizontal_alignment="center",
                vertical_alignment="top",
                wrap_text=True,
                number_format="#,##0.00"
            )

        assert mock_range.HorizontalAlignment == HORIZONTAL_ALIGNMENT_MAP["center"]
        assert mock_range.VerticalAlignment == VERTICAL_ALIGNMENT_MAP["top"]
        assert mock_range.WrapText is True
        assert mock_range.NumberFormat == "#,##0.00"

    def test_set_cell_format_partial(self):
        """Test set_cell_format with partial options."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_cell_format(
                manager, mock_workbook, "Sheet1", "A1:D10",
                wrap_text=True
            )

        assert mock_range.WrapText is True


class TestSetBorderFormat:
    """Tests for set_border_format function."""

    def _create_mock_preserve_context(self):
        """Create a mock context manager."""
        mock_ctx = MagicMock()
        mock_ctx.__enter__ = MagicMock(return_value=None)
        mock_ctx.__exit__ = MagicMock(return_value=None)
        return mock_ctx

    def test_set_border_format_single_edge(self):
        """Test set_border_format with single edge."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_borders = MagicMock()
        mock_border = MagicMock()
        mock_borders.Item.return_value = mock_border
        mock_range.Borders = mock_borders
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_border_format(
                manager, mock_workbook, "Sheet1", "A1:D10",
                edge="left",
                style="continuous",
                weight=2,
                color="000000"
            )

        mock_borders.Item.assert_called_once_with(BORDER_EDGE_MAP["left"])
        assert mock_border.LineStyle == BORDER_STYLE_MAP["continuous"]
        assert mock_border.Weight == 2
        assert mock_border.Color == int("000000", 16)

    def test_set_border_format_all_edges(self):
        """Test set_border_format with all edges."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_borders = MagicMock()
        mock_range.Borders = mock_borders
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_border_format(
                manager, mock_workbook, "Sheet1", "A1:D10",
                edge="all",
                style="continuous",
                weight=2,
                color="000000"
            )

        assert mock_borders.LineStyle == BORDER_STYLE_MAP["continuous"]
        assert mock_borders.Weight == 2
        assert mock_borders.Color == int("000000", 16)


class TestSetBackgroundColor:
    """Tests for set_background_color function."""

    def _create_mock_preserve_context(self):
        """Create a mock context manager."""
        mock_ctx = MagicMock()
        mock_ctx.__enter__ = MagicMock(return_value=None)
        mock_ctx.__exit__ = MagicMock(return_value=None)
        return mock_ctx

    def test_set_background_color(self):
        """Test set_background_color sets color."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_interior = MagicMock()
        mock_range.Interior = mock_interior
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_background_color(
                manager, mock_workbook, "Sheet1", "A1:D10", "FFFF00"
            )

        assert mock_interior.Color == int("FFFF00", 16)


class TestAutoFitColumns:
    """Tests for auto_fit_columns function."""

    def _create_mock_preserve_context(self):
        """Create a mock context manager."""
        mock_ctx = MagicMock()
        mock_ctx.__enter__ = MagicMock(return_value=None)
        mock_ctx.__exit__ = MagicMock(return_value=None)
        return mock_ctx

    def test_auto_fit_columns_with_range(self):
        """Test auto_fit_columns with specific range."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_columns = MagicMock()
        mock_range.Columns = mock_columns
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.auto_fit_columns(
                manager, mock_workbook, "Sheet1", "A:D"
            )

        mock_columns.AutoFit.assert_called_once()

    def test_auto_fit_columns_without_range(self):
        """Test auto_fit_columns without range (all used columns)."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_used_range = MagicMock()
        mock_columns = MagicMock()
        mock_used_range.Columns = mock_columns
        mock_worksheet.UsedRange = mock_used_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.auto_fit_columns(
                manager, mock_workbook, "Sheet1"
            )

        mock_columns.AutoFit.assert_called_once()


class TestSetColumnWidth:
    """Tests for set_column_width function."""

    def _create_mock_preserve_context(self):
        """Create a mock context manager."""
        mock_ctx = MagicMock()
        mock_ctx.__enter__ = MagicMock(return_value=None)
        mock_ctx.__exit__ = MagicMock(return_value=None)
        return mock_ctx

    def test_set_column_width(self):
        """Test set_column_width sets width."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_column_width(
                manager, mock_workbook, "Sheet1", "A:C", 15.5
            )

        assert mock_range.ColumnWidth == 15.5


class TestSetRowHeight:
    """Tests for set_row_height function."""

    def _create_mock_preserve_context(self):
        """Create a mock context manager."""
        mock_ctx = MagicMock()
        mock_ctx.__enter__ = MagicMock(return_value=None)
        mock_ctx.__exit__ = MagicMock(return_value=None)
        return mock_ctx

    def test_set_row_height(self):
        """Test set_row_height sets height."""
        manager = MagicMock(spec=ExcelAppManager)
        manager.app = MagicMock()

        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_range = MagicMock()
        mock_worksheet.Range.return_value = mock_range

        manager.get_worksheet.return_value = mock_worksheet

        with patch("libs.excel_com.formatting_ops.preserve_user_state", return_value=self._create_mock_preserve_context()):
            formatting_ops.set_row_height(
                manager, mock_workbook, "Sheet1", "1:5", 20.0
            )

        assert mock_range.RowHeight == 20.0
