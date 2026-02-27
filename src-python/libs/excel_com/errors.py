"""Custom exception hierarchy for Excel COM operations."""


class ExcelComError(Exception):
    """Base exception for all Excel COM errors."""

    def __init__(self, user_message: str, raw_error: Exception | None = None):
        self.user_message = user_message
        self.raw_error = raw_error
        super().__init__(user_message)


class ExcelInstanceError(ExcelComError):
    """Excel Application could not be obtained or created."""


class WorkbookNotFoundError(ExcelComError):
    """Workbook not found in registry."""


class WorkbookReadOnlyError(ExcelComError):
    """Workbook is open in read-only mode."""


class WorkbookLockedError(ExcelComError):
    """Workbook is locked by another process."""


class SheetNotFoundError(ExcelComError):
    """Worksheet does not exist."""


class RangeError(ExcelComError):
    """Invalid range address."""


class FormulaError(ExcelComError):
    """Formula syntax or evaluation error."""


class TableError(ExcelComError):
    """Table (ListObject) operation error."""


class QueryError(ExcelComError):
    """PowerQuery operation error."""


class ComCallError(ExcelComError):
    """Low-level COM call failure."""


class StaleReferenceError(ExcelComError):
    """COM reference is no longer valid."""
