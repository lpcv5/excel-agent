"""Tests for excel_com/constants.py."""

import pytest

from excel_com.constants import (
    # Underline styles
    XL_UNDERLINE_STYLE_SINGLE,
    XL_UNDERLINE_STYLE_NONE,
    # Horizontal alignment
    XL_LEFT,
    XL_CENTER,
    XL_RIGHT,
    XL_GENERAL,
    # Vertical alignment
    XL_TOP,
    XL_BOTTOM,
    XL_JUSTIFY,
    XL_CENTER_V,
    # Border styles
    XL_CONTINUOUS,
    XL_DASH,
    XL_DOT,
    XL_DOUBLE,
    XL_NONE,
    XL_SLANT_DASH_DOT,
    # Border edges
    XL_EDGE_LEFT,
    XL_EDGE_RIGHT,
    XL_EDGE_TOP,
    XL_EDGE_BOTTOM,
    XL_INSIDE_HORIZONTAL,
    XL_INSIDE_VERTICAL,
    # Border weights
    XL_HAIRLINE,
    XL_THIN,
    XL_MEDIUM,
    XL_THICK,
    # Lookup dictionaries
    HORIZONTAL_ALIGNMENT_MAP,
    VERTICAL_ALIGNMENT_MAP,
    BORDER_STYLE_MAP,
    BORDER_EDGE_MAP,
    BORDER_WEIGHT_MAP,
)


class TestUnderlineStyleConstants:
    """Tests for underline style constants."""

    def test_underline_style_single(self):
        """Test XL_UNDERLINE_STYLE_SINGLE value."""
        assert XL_UNDERLINE_STYLE_SINGLE == 2

    def test_underline_style_none(self):
        """Test XL_UNDERLINE_STYLE_NONE value."""
        assert XL_UNDERLINE_STYLE_NONE == -4142


class TestHorizontalAlignmentConstants:
    """Tests for horizontal alignment constants."""

    def test_xl_left(self):
        """Test XL_LEFT value."""
        assert XL_LEFT == -4131

    def test_xl_center(self):
        """Test XL_CENTER value."""
        assert XL_CENTER == -4108

    def test_xl_right(self):
        """Test XL_RIGHT value."""
        assert XL_RIGHT == -4152

    def test_xl_general(self):
        """Test XL_GENERAL value."""
        assert XL_GENERAL == 1


class TestVerticalAlignmentConstants:
    """Tests for vertical alignment constants."""

    def test_xl_top(self):
        """Test XL_TOP value."""
        assert XL_TOP == -4160

    def test_xl_bottom(self):
        """Test XL_BOTTOM value."""
        assert XL_BOTTOM == -4107

    def test_xl_justify(self):
        """Test XL_JUSTIFY value."""
        assert XL_JUSTIFY == -4130

    def test_xl_center_v(self):
        """Test XL_CENTER_V value."""
        assert XL_CENTER_V == -4108


class TestBorderStyleConstants:
    """Tests for border style constants."""

    def test_xl_continuous(self):
        """Test XL_CONTINUOUS value."""
        assert XL_CONTINUOUS == 1

    def test_xl_dash(self):
        """Test XL_DASH value."""
        assert XL_DASH == -4115

    def test_xl_dot(self):
        """Test XL_DOT value."""
        assert XL_DOT == -4118

    def test_xl_double(self):
        """Test XL_DOUBLE value."""
        assert XL_DOUBLE == -4119

    def test_xl_none(self):
        """Test XL_NONE value."""
        assert XL_NONE == -4142

    def test_xl_slant_dash_dot(self):
        """Test XL_SLANT_DASH_DOT value."""
        assert XL_SLANT_DASH_DOT == -4135


class TestBorderEdgeConstants:
    """Tests for border edge position constants."""

    def test_xl_edge_left(self):
        """Test XL_EDGE_LEFT value."""
        assert XL_EDGE_LEFT == 7

    def test_xl_edge_right(self):
        """Test XL_EDGE_RIGHT value."""
        assert XL_EDGE_RIGHT == 10

    def test_xl_edge_top(self):
        """Test XL_EDGE_TOP value."""
        assert XL_EDGE_TOP == 8

    def test_xl_edge_bottom(self):
        """Test XL_EDGE_BOTTOM value."""
        assert XL_EDGE_BOTTOM == 9

    def test_xl_inside_horizontal(self):
        """Test XL_INSIDE_HORIZONTAL value."""
        assert XL_INSIDE_HORIZONTAL == 12

    def test_xl_inside_vertical(self):
        """Test XL_INSIDE_VERTICAL value."""
        assert XL_INSIDE_VERTICAL == 11


class TestBorderWeightConstants:
    """Tests for border weight constants."""

    def test_xl_hairline(self):
        """Test XL_HAIRLINE value."""
        assert XL_HAIRLINE == 1

    def test_xl_thin(self):
        """Test XL_THIN value."""
        assert XL_THIN == 2

    def test_xl_medium(self):
        """Test XL_MEDIUM value."""
        assert XL_MEDIUM == -4138

    def test_xl_thick(self):
        """Test XL_THICK value."""
        assert XL_THICK == 4


class TestHorizontalAlignmentMap:
    """Tests for HORIZONTAL_ALIGNMENT_MAP dictionary."""

    def test_map_contains_left(self):
        """Test that 'left' key exists and maps correctly."""
        assert "left" in HORIZONTAL_ALIGNMENT_MAP
        assert HORIZONTAL_ALIGNMENT_MAP["left"] == XL_LEFT

    def test_map_contains_center(self):
        """Test that 'center' key exists and maps correctly."""
        assert "center" in HORIZONTAL_ALIGNMENT_MAP
        assert HORIZONTAL_ALIGNMENT_MAP["center"] == XL_CENTER

    def test_map_contains_right(self):
        """Test that 'right' key exists and maps correctly."""
        assert "right" in HORIZONTAL_ALIGNMENT_MAP
        assert HORIZONTAL_ALIGNMENT_MAP["right"] == XL_RIGHT

    def test_map_contains_general(self):
        """Test that 'general' key exists and maps correctly."""
        assert "general" in HORIZONTAL_ALIGNMENT_MAP
        assert HORIZONTAL_ALIGNMENT_MAP["general"] == XL_GENERAL

    def test_map_has_correct_count(self):
        """Test that the map has exactly 4 entries."""
        assert len(HORIZONTAL_ALIGNMENT_MAP) == 4


class TestVerticalAlignmentMap:
    """Tests for VERTICAL_ALIGNMENT_MAP dictionary."""

    def test_map_contains_top(self):
        """Test that 'top' key exists and maps correctly."""
        assert "top" in VERTICAL_ALIGNMENT_MAP
        assert VERTICAL_ALIGNMENT_MAP["top"] == XL_TOP

    def test_map_contains_center(self):
        """Test that 'center' key exists and maps correctly."""
        assert "center" in VERTICAL_ALIGNMENT_MAP
        assert VERTICAL_ALIGNMENT_MAP["center"] == XL_CENTER_V

    def test_map_contains_bottom(self):
        """Test that 'bottom' key exists and maps correctly."""
        assert "bottom" in VERTICAL_ALIGNMENT_MAP
        assert VERTICAL_ALIGNMENT_MAP["bottom"] == XL_BOTTOM

    def test_map_contains_justify(self):
        """Test that 'justify' key exists and maps correctly."""
        assert "justify" in VERTICAL_ALIGNMENT_MAP
        assert VERTICAL_ALIGNMENT_MAP["justify"] == XL_JUSTIFY

    def test_map_has_correct_count(self):
        """Test that the map has exactly 4 entries."""
        assert len(VERTICAL_ALIGNMENT_MAP) == 4


class TestBorderStyleMap:
    """Tests for BORDER_STYLE_MAP dictionary."""

    def test_map_contains_continuous(self):
        """Test that 'continuous' key exists and maps correctly."""
        assert "continuous" in BORDER_STYLE_MAP
        assert BORDER_STYLE_MAP["continuous"] == XL_CONTINUOUS

    def test_map_contains_dash(self):
        """Test that 'dash' key exists and maps correctly."""
        assert "dash" in BORDER_STYLE_MAP
        assert BORDER_STYLE_MAP["dash"] == XL_DASH

    def test_map_contains_dot(self):
        """Test that 'dot' key exists and maps correctly."""
        assert "dot" in BORDER_STYLE_MAP
        assert BORDER_STYLE_MAP["dot"] == XL_DOT

    def test_map_contains_double(self):
        """Test that 'double' key exists and maps correctly."""
        assert "double" in BORDER_STYLE_MAP
        assert BORDER_STYLE_MAP["double"] == XL_DOUBLE

    def test_map_contains_none(self):
        """Test that 'none' key exists and maps correctly."""
        assert "none" in BORDER_STYLE_MAP
        assert BORDER_STYLE_MAP["none"] == XL_NONE

    def test_map_contains_slant_dash_dot(self):
        """Test that 'slant_dash_dot' key exists and maps correctly."""
        assert "slant_dash_dot" in BORDER_STYLE_MAP
        assert BORDER_STYLE_MAP["slant_dash_dot"] == XL_SLANT_DASH_DOT

    def test_map_has_correct_count(self):
        """Test that the map has exactly 6 entries."""
        assert len(BORDER_STYLE_MAP) == 6


class TestBorderEdgeMap:
    """Tests for BORDER_EDGE_MAP dictionary."""

    def test_map_contains_left(self):
        """Test that 'left' key exists and maps correctly."""
        assert "left" in BORDER_EDGE_MAP
        assert BORDER_EDGE_MAP["left"] == XL_EDGE_LEFT

    def test_map_contains_right(self):
        """Test that 'right' key exists and maps correctly."""
        assert "right" in BORDER_EDGE_MAP
        assert BORDER_EDGE_MAP["right"] == XL_EDGE_RIGHT

    def test_map_contains_top(self):
        """Test that 'top' key exists and maps correctly."""
        assert "top" in BORDER_EDGE_MAP
        assert BORDER_EDGE_MAP["top"] == XL_EDGE_TOP

    def test_map_contains_bottom(self):
        """Test that 'bottom' key exists and maps correctly."""
        assert "bottom" in BORDER_EDGE_MAP
        assert BORDER_EDGE_MAP["bottom"] == XL_EDGE_BOTTOM

    def test_map_has_correct_count(self):
        """Test that the map has exactly 4 entries."""
        assert len(BORDER_EDGE_MAP) == 4


class TestBorderWeightMap:
    """Tests for BORDER_WEIGHT_MAP dictionary."""

    def test_map_contains_hairline(self):
        """Test that 'hairline' key exists and maps correctly."""
        assert "hairline" in BORDER_WEIGHT_MAP
        assert BORDER_WEIGHT_MAP["hairline"] == XL_HAIRLINE

    def test_map_contains_thin(self):
        """Test that 'thin' key exists and maps correctly."""
        assert "thin" in BORDER_WEIGHT_MAP
        assert BORDER_WEIGHT_MAP["thin"] == XL_THIN

    def test_map_contains_medium(self):
        """Test that 'medium' key exists and maps correctly."""
        assert "medium" in BORDER_WEIGHT_MAP
        assert BORDER_WEIGHT_MAP["medium"] == XL_MEDIUM

    def test_map_contains_thick(self):
        """Test that 'thick' key exists and maps correctly."""
        assert "thick" in BORDER_WEIGHT_MAP
        assert BORDER_WEIGHT_MAP["thick"] == XL_THICK

    def test_map_has_correct_count(self):
        """Test that the map has exactly 4 entries."""
        assert len(BORDER_WEIGHT_MAP) == 4
