"""Tests for excel_com/exceptions.py."""


from libs.excel_com.exceptions import (
    # Base exception
    ExcelMCPError,
    # Workbook exceptions
    WorkbookError,
    WorkbookNotFoundError,
    WorkbookAlreadyOpenError,
    WorkbookSaveError,
    # Worksheet exceptions
    WorksheetError,
    WorksheetNotFoundError,
    WorksheetAlreadyExistsError,
    InvalidWorksheetNameError,
    # Range exceptions
    RangeError,
    InvalidRangeError,
    RangeWriteError,
    # Excel application exceptions
    ExcelApplicationError,
    ExcelNotRunningError,
    ExcelCOMError,
    # Validation exceptions
    ValidationError,
)


class TestExcelMCPError:
    """Tests for ExcelMCPError base exception."""

    def test_init_with_message_only(self):
        """Test initialization with message only."""
        error = ExcelMCPError("Test error message")
        assert error.message == "Test error message"
        assert error.details == {}
        assert str(error) == "Test error message"

    def test_init_with_message_and_details(self):
        """Test initialization with message and details."""
        details = {"key": "value", "number": 42}
        error = ExcelMCPError("Test error", details=details)
        assert error.message == "Test error"
        assert error.details == details

    def test_to_dict_without_details(self):
        """Test to_dict method without details."""
        error = ExcelMCPError("Simple error")
        result = error.to_dict()
        assert result == {"error": "Simple error"}

    def test_to_dict_with_details(self):
        """Test to_dict method with details."""
        details = {"extra": "info", "code": 123}
        error = ExcelMCPError("Error with details", details=details)
        result = error.to_dict()
        assert result["error"] == "Error with details"
        assert result["extra"] == "info"
        assert result["code"] == 123


class TestWorkbookNotFoundError:
    """Tests for WorkbookNotFoundError exception."""

    def test_init(self):
        """Test initialization."""
        error = WorkbookNotFoundError("/path/to/workbook.xlsx")
        assert error.filepath == "/path/to/workbook.xlsx"
        assert "not found" in error.message.lower()
        assert error.details["filepath"] == "/path/to/workbook.xlsx"
        assert error.details["error_type"] == "file_not_found"

    def test_inheritance(self):
        """Test that it inherits from WorkbookError."""
        error = WorkbookNotFoundError("/path/to/file.xlsx")
        assert isinstance(error, WorkbookError)
        assert isinstance(error, ExcelMCPError)
        assert isinstance(error, Exception)


class TestWorkbookAlreadyOpenError:
    """Tests for WorkbookAlreadyOpenError exception."""

    def test_init(self):
        """Test initialization."""
        error = WorkbookAlreadyOpenError("/path/to/workbook.xlsx")
        assert error.filepath == "/path/to/workbook.xlsx"
        assert "already open" in error.message.lower()
        assert error.details["error_type"] == "already_open"

    def test_inheritance(self):
        """Test that it inherits from WorkbookError."""
        error = WorkbookAlreadyOpenError("/path/to/file.xlsx")
        assert isinstance(error, WorkbookError)
        assert isinstance(error, ExcelMCPError)


class TestWorkbookSaveError:
    """Tests for WorkbookSaveError exception."""

    def test_init(self):
        """Test initialization."""
        error = WorkbookSaveError("/path/to/workbook.xlsx", "Permission denied")
        assert error.filepath == "/path/to/workbook.xlsx"
        assert error.reason == "Permission denied"
        assert "Failed to save" in error.message
        assert error.details["reason"] == "Permission denied"
        assert error.details["error_type"] == "save_failed"

    def test_inheritance(self):
        """Test that it inherits from WorkbookError."""
        error = WorkbookSaveError("/path/to/file.xlsx", "reason")
        assert isinstance(error, WorkbookError)
        assert isinstance(error, ExcelMCPError)


class TestWorksheetNotFoundError:
    """Tests for WorksheetNotFoundError exception."""

    def test_init_with_workbook_name(self):
        """Test initialization with workbook name."""
        error = WorksheetNotFoundError("Sheet1", workbook_name="test.xlsx")
        assert error.worksheet_name == "Sheet1"
        assert error.workbook_name == "test.xlsx"
        assert "not found" in error.message.lower()
        assert "test.xlsx" in error.message

    def test_init_without_workbook_name(self):
        """Test initialization without workbook name."""
        error = WorksheetNotFoundError("Sheet1")
        assert error.worksheet_name == "Sheet1"
        assert error.workbook_name is None
        assert "not found" in error.message.lower()

    def test_inheritance(self):
        """Test that it inherits from WorksheetError."""
        error = WorksheetNotFoundError("Sheet1")
        assert isinstance(error, WorksheetError)
        assert isinstance(error, ExcelMCPError)


class TestWorksheetAlreadyExistsError:
    """Tests for WorksheetAlreadyExistsError exception."""

    def test_init(self):
        """Test initialization."""
        error = WorksheetAlreadyExistsError("Sheet1")
        assert error.worksheet_name == "Sheet1"
        assert "already exists" in error.message.lower()
        assert error.details["error_type"] == "worksheet_exists"

    def test_inheritance(self):
        """Test that it inherits from WorksheetError."""
        error = WorksheetAlreadyExistsError("Sheet1")
        assert isinstance(error, WorksheetError)
        assert isinstance(error, ExcelMCPError)


class TestInvalidWorksheetNameError:
    """Tests for InvalidWorksheetNameError exception."""

    def test_init(self):
        """Test initialization."""
        error = InvalidWorksheetNameError("Invalid/Name", "Contains invalid character")
        assert error.worksheet_name == "Invalid/Name"
        assert error.reason == "Contains invalid character"
        assert "Invalid worksheet name" in error.message
        assert error.details["error_type"] == "invalid_name"

    def test_inheritance(self):
        """Test that it inherits from WorksheetError."""
        error = InvalidWorksheetNameError("Name", "reason")
        assert isinstance(error, WorksheetError)
        assert isinstance(error, ExcelMCPError)


class TestInvalidRangeError:
    """Tests for InvalidRangeError exception."""

    def test_init(self):
        """Test initialization."""
        error = InvalidRangeError("A1:Z999", "Range too large")
        assert error.range_spec == "A1:Z999"
        assert error.reason == "Range too large"
        assert "Invalid range" in error.message
        assert error.details["error_type"] == "invalid_range"

    def test_inheritance(self):
        """Test that it inherits from RangeError."""
        error = InvalidRangeError("A1", "reason")
        assert isinstance(error, RangeError)
        assert isinstance(error, ExcelMCPError)


class TestRangeWriteError:
    """Tests for RangeWriteError exception."""

    def test_init(self):
        """Test initialization."""
        error = RangeWriteError("A1:B2", "Cell is locked")
        assert error.range_spec == "A1:B2"
        assert error.reason == "Cell is locked"
        assert "Failed to write" in error.message
        assert error.details["error_type"] == "write_failed"

    def test_inheritance(self):
        """Test that it inherits from RangeError."""
        error = RangeWriteError("A1", "reason")
        assert isinstance(error, RangeError)
        assert isinstance(error, ExcelMCPError)


class TestExcelNotRunningError:
    """Tests for ExcelNotRunningError exception."""

    def test_init(self):
        """Test initialization."""
        error = ExcelNotRunningError()
        assert "not running" in error.message.lower()
        assert error.details["error_type"] == "excel_not_running"

    def test_inheritance(self):
        """Test that it inherits from ExcelApplicationError."""
        error = ExcelNotRunningError()
        assert isinstance(error, ExcelApplicationError)
        assert isinstance(error, ExcelMCPError)


class TestExcelCOMError:
    """Tests for ExcelCOMError exception."""

    def test_init(self):
        """Test initialization."""
        original_error = Exception("COM error occurred")
        error = ExcelCOMError("Open workbook", original_error)
        assert error.operation == "Open workbook"
        assert error.com_error == original_error
        assert "COM error" in error.message
        assert "Open workbook" in error.message
        assert error.details["error_type"] == "com_error"

    def test_inheritance(self):
        """Test that it inherits from ExcelApplicationError."""
        error = ExcelCOMError("operation", Exception("error"))
        assert isinstance(error, ExcelApplicationError)
        assert isinstance(error, ExcelMCPError)


class TestValidationError:
    """Tests for ValidationError exception."""

    def test_init(self):
        """Test initialization."""
        error = ValidationError("filepath", "Cannot be empty")
        assert error.field == "filepath"
        assert error.reason == "Cannot be empty"
        assert "Validation error" in error.message
        assert "filepath" in error.message
        assert error.details["error_type"] == "validation_error"

    def test_inheritance(self):
        """Test that it inherits from ExcelMCPError."""
        error = ValidationError("field", "reason")
        assert isinstance(error, ExcelMCPError)


class TestExceptionInheritance:
    """Tests for exception inheritance hierarchy."""

    def test_workbook_error_inherits_from_excel_mcp_error(self):
        """Test WorkbookError inheritance."""
        assert issubclass(WorkbookError, ExcelMCPError)

    def test_worksheet_error_inherits_from_excel_mcp_error(self):
        """Test WorksheetError inheritance."""
        assert issubclass(WorksheetError, ExcelMCPError)

    def test_range_error_inherits_from_excel_mcp_error(self):
        """Test RangeError inheritance."""
        assert issubclass(RangeError, ExcelMCPError)

    def test_excel_application_error_inherits_from_excel_mcp_error(self):
        """Test ExcelApplicationError inheritance."""
        assert issubclass(ExcelApplicationError, ExcelMCPError)

    def test_all_exceptions_inherit_from_exception(self):
        """Test that all custom exceptions inherit from Exception."""
        exceptions_to_test = [
            ExcelMCPError,
            WorkbookError,
            WorkbookNotFoundError,
            WorkbookAlreadyOpenError,
            WorkbookSaveError,
            WorksheetError,
            WorksheetNotFoundError,
            WorksheetAlreadyExistsError,
            InvalidWorksheetNameError,
            RangeError,
            InvalidRangeError,
            RangeWriteError,
            ExcelApplicationError,
            ExcelNotRunningError,
            ExcelCOMError,
            ValidationError,
        ]
        for exc_class in exceptions_to_test:
            assert issubclass(exc_class, Exception)
