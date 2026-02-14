"""
Excel MCP Server Exceptions.

This module defines custom exceptions for Excel-related operations.
These exceptions provide specific error types that can be caught
and handled appropriately in the MCP server layer.
"""


# =============================================================================
# Base Exception
# =============================================================================

class ExcelMCPError(Exception):
    """Base exception for all Excel MCP errors."""

    def __init__(self, message: str, details: dict | None = None):
        """Initialize Excel MCP error.

        Args:
            message: Human-readable error message
            details: Optional dictionary with additional error details
        """
        super().__init__(message)
        self.message = message
        self.details = details or {}

    def to_dict(self) -> dict:
        """Convert exception to dictionary for JSON response.

        Returns:
            Dictionary with error information
        """
        result = {"error": self.message}
        if self.details:
            result.update(self.details)
        return result


# =============================================================================
# Workbook Exceptions
# =============================================================================

class WorkbookError(ExcelMCPError):
    """Exception raised for workbook-related errors."""

    pass


class WorkbookNotFoundError(WorkbookError):
    """Exception raised when a workbook file cannot be found."""

    def __init__(self, filepath: str):
        super().__init__(
            f"Workbook not found: {filepath}",
            {"filepath": filepath, "error_type": "file_not_found"}
        )
        self.filepath = filepath


class WorkbookAlreadyOpenError(WorkbookError):
    """Exception raised when trying to open a workbook that is already open."""

    def __init__(self, filepath: str):
        super().__init__(
            f"Workbook is already open: {filepath}",
            {"filepath": filepath, "error_type": "already_open"}
        )
        self.filepath = filepath


class WorkbookSaveError(WorkbookError):
    """Exception raised when saving a workbook fails."""

    def __init__(self, filepath: str, reason: str):
        super().__init__(
            f"Failed to save workbook: {filepath}",
            {"filepath": filepath, "reason": reason, "error_type": "save_failed"}
        )
        self.filepath = filepath
        self.reason = reason


# =============================================================================
# Worksheet Exceptions
# =============================================================================

class WorksheetError(ExcelMCPError):
    """Exception raised for worksheet-related errors."""

    pass


class WorksheetNotFoundError(WorksheetError):
    """Exception raised when a worksheet cannot be found."""

    def __init__(self, worksheet_name: str, workbook_name: str | None = None):
        message = f"Worksheet not found: {worksheet_name}"
        if workbook_name:
            message += f" (in workbook: {workbook_name})"
        super().__init__(
            message,
            {
                "worksheet_name": worksheet_name,
                "workbook_name": workbook_name,
                "error_type": "worksheet_not_found"
            }
        )
        self.worksheet_name = worksheet_name
        self.workbook_name = workbook_name


class WorksheetAlreadyExistsError(WorksheetError):
    """Exception raised when trying to create a worksheet that already exists."""

    def __init__(self, worksheet_name: str):
        super().__init__(
            f"Worksheet already exists: {worksheet_name}",
            {"worksheet_name": worksheet_name, "error_type": "worksheet_exists"}
        )
        self.worksheet_name = worksheet_name


class InvalidWorksheetNameError(WorksheetError):
    """Exception raised when a worksheet name is invalid."""

    def __init__(self, worksheet_name: str, reason: str):
        super().__init__(
            f"Invalid worksheet name: {worksheet_name}",
            {"worksheet_name": worksheet_name, "reason": reason, "error_type": "invalid_name"}
        )
        self.worksheet_name = worksheet_name
        self.reason = reason


# =============================================================================
# Range Exceptions
# =============================================================================

class RangeError(ExcelMCPError):
    """Exception raised for range-related errors."""

    pass


class InvalidRangeError(RangeError):
    """Exception raised when a range specification is invalid."""

    def __init__(self, range_spec: str, reason: str):
        super().__init__(
            f"Invalid range: {range_spec}",
            {"range": range_spec, "reason": reason, "error_type": "invalid_range"}
        )
        self.range_spec = range_spec
        self.reason = reason


class RangeWriteError(RangeError):
    """Exception raised when writing to a range fails."""

    def __init__(self, range_spec: str, reason: str):
        super().__init__(
            f"Failed to write to range: {range_spec}",
            {"range": range_spec, "reason": reason, "error_type": "write_failed"}
        )
        self.range_spec = range_spec
        self.reason = reason


# =============================================================================
# Excel Application Exceptions
# =============================================================================

class ExcelApplicationError(ExcelMCPError):
    """Exception raised for Excel application-related errors."""

    pass


class ExcelNotRunningError(ExcelApplicationError):
    """Exception raised when Excel is not running."""

    def __init__(self):
        super().__init__(
            "Excel application is not running",
            {"error_type": "excel_not_running"}
        )


class ExcelCOMError(ExcelApplicationError):
    """Exception raised for Excel COM errors."""

    def __init__(self, operation: str, com_error: Exception):
        super().__init__(
            f"Excel COM error during {operation}: {str(com_error)}",
            {"operation": operation, "com_error": str(com_error), "error_type": "com_error"}
        )
        self.operation = operation
        self.com_error = com_error


# =============================================================================
# Validation Exceptions
# =============================================================================

class ValidationError(ExcelMCPError):
    """Exception raised for input validation errors."""

    def __init__(self, field: str, reason: str):
        super().__init__(
            f"Validation error for field '{field}': {reason}",
            {"field": field, "reason": reason, "error_type": "validation_error"}
        )
        self.field = field
        self.reason = reason
