"""
Excel COM Constants.

This module defines all Excel COM constant values used for formatting
and style operations. These are the numeric values that Excel uses
internally for various formatting and style options.

Reference: https://docs.microsoft.com/en-us/office/vba/api/overview/excel
"""


# =============================================================================
# Underline Styles
# =============================================================================

XL_UNDERLINE_STYLE_SINGLE = 2      # xlUnderlineStyleSingle
XL_UNDERLINE_STYLE_NONE = -4142    # xlUnderlineStyleNone


# =============================================================================
# Horizontal Alignment
# =============================================================================

XL_LEFT = -4131        # xlLeft
XL_CENTER = -4108      # xlCenter
XL_RIGHT = -4152       # xlRight
XL_GENERAL = 1         # xlGeneral


# =============================================================================
# Vertical Alignment
# =============================================================================

XL_TOP = -4160         # xlTop
XL_BOTTOM = -4107      # xlBottom
XL_JUSTIFY = -4130     # xlJustify
XL_CENTER_V = -4108    # xlCenter (same as horizontal)


# =============================================================================
# Border Styles
# =============================================================================

XL_CONTINUOUS = 1      # xlContinuous
XL_DASH = -4115        # xlDash
XL_DOT = -4118         # xlDot
XL_DOUBLE = -4119      # xlDouble
XL_NONE = -4142        # xlNone
XL_SLANT_DASH_DOT = -4135  # xlSlantDashDot


# =============================================================================
# Border Edge Positions
# =============================================================================

XL_EDGE_LEFT = 7       # xlEdgeLeft
XL_EDGE_RIGHT = 10     # xlEdgeRight
XL_EDGE_TOP = 8        # xlEdgeTop
XL_EDGE_BOTTOM = 9     # xlEdgeBottom
XL_INSIDE_HORIZONTAL = 12   # xlInsideHorizontal
XL_INSIDE_VERTICAL = 11     # xlInsideVertical


# =============================================================================
# Border Weights
# =============================================================================

XL_HAIRLINE = 1        # xlHairline
XL_THIN = 2            # xlThin
XL_MEDIUM = -4138      # xlMedium
XL_THICK = 4           # xlThick


# =============================================================================
# Lookup Dictionaries
# =============================================================================

# Horizontal alignment mapping
HORIZONTAL_ALIGNMENT_MAP = {
    "left": XL_LEFT,
    "center": XL_CENTER,
    "right": XL_RIGHT,
    "general": XL_GENERAL,
}

# Vertical alignment mapping
VERTICAL_ALIGNMENT_MAP = {
    "top": XL_TOP,
    "center": XL_CENTER_V,
    "bottom": XL_BOTTOM,
    "justify": XL_JUSTIFY,
}

# Border style mapping
BORDER_STYLE_MAP = {
    "continuous": XL_CONTINUOUS,
    "dash": XL_DASH,
    "dot": XL_DOT,
    "double": XL_DOUBLE,
    "none": XL_NONE,
    "slant_dash_dot": XL_SLANT_DASH_DOT,
}

# Border edge position mapping
BORDER_EDGE_MAP = {
    "left": XL_EDGE_LEFT,
    "right": XL_EDGE_RIGHT,
    "top": XL_EDGE_TOP,
    "bottom": XL_EDGE_BOTTOM,
}

# Border weight mapping
BORDER_WEIGHT_MAP = {
    "hairline": XL_HAIRLINE,
    "thin": XL_THIN,
    "medium": XL_MEDIUM,
    "thick": XL_THICK,
}
