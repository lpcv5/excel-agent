"""
Excel COM Module - Windows COM interface for Excel operations.

This module provides the core components for working with Excel
through the Windows COM interface.
"""

from .manager import ExcelAppManager
from .context import preserve_user_state, UserStatePreserver
from .exceptions import (
    ExcelMCPError,
    WorkbookError,
    WorkbookNotFoundError,
    WorksheetError,
    WorksheetNotFoundError,
    RangeError,
    InvalidRangeError,
    ExcelApplicationError,
    ExcelNotRunningError,
    ValidationError,
)

__all__ = [
    "ExcelAppManager",
    "preserve_user_state",
    "UserStatePreserver",
    "ExcelMCPError",
    "WorkbookError",
    "WorkbookNotFoundError",
    "WorksheetError",
    "WorksheetNotFoundError",
    "RangeError",
    "InvalidRangeError",
    "ExcelApplicationError",
    "ExcelNotRunningError",
    "ValidationError",
]
