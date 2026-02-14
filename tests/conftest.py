"""
Shared pytest fixtures for excel-agent tests.

This module provides common fixtures used across all test modules.
All Excel-related tests share a single mock COM object.
"""

import json
from unittest.mock import MagicMock, patch

import pytest


# =============================================================================
# Shared Mock Excel COM Objects (Singleton Pattern)
# =============================================================================

class MockExcelCOMFactory:
    """Factory for creating shared mock Excel COM objects."""

    _instance = None
    _excel_app = None
    _workbooks = {}
    _worksheets = {}

    @classmethod
    def get_instance(cls):
        """Get the singleton factory instance."""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    @classmethod
    def reset(cls):
        """Reset all mock objects (call this in fixtures if needed)."""
        cls._excel_app = None
        cls._workbooks = {}
        cls._worksheets = {}

    def create_excel_app(self):
        """Create or return the shared mock Excel Application."""
        if self._excel_app is None:
            app = MagicMock()
            app.Visible = False
            app.DisplayAlerts = False
            app.Workbooks = MagicMock()
            app.Workbooks.Count = 0
            app.ActiveWorkbook = None
            app.ActiveSheet = None
            app.ActiveWindow = MagicMock()
            app.ActiveWindow.ScrollRow = 1
            app.ActiveWindow.ScrollColumn = 1
            app.ActiveCell = None
            app.Selection = None
            app.Quit = MagicMock()
            self._excel_app = app
        return self._excel_app

    def create_workbook(self, name="test_workbook.xlsx", path="C:\\test"):
        """Create a mock workbook with the given name."""
        key = f"{path}\\{name}"
        if key not in self._workbooks:
            workbook = MagicMock()
            workbook.Name = name
            workbook.FullName = f"{path}\\{name}"
            workbook.Path = path

            # Create mock worksheets collection
            worksheets = MagicMock()
            worksheets.Count = 3

            # Create default sheets
            sheets = []
            for sheet_name in ["Sheet1", "Sheet2", "Sheet3"]:
                sheet = self.create_worksheet(sheet_name)
                sheets.append(sheet)

            def get_worksheet(idx):
                if isinstance(idx, int):
                    return sheets[idx - 1] if 1 <= idx <= len(sheets) else None
                # Find by name
                for s in sheets:
                    if s.Name == idx:
                        return s
                raise Exception(f"Worksheet '{idx}' not found")

            worksheets.__getitem__ = get_worksheet
            worksheets.__iter__ = lambda: iter(sheets)
            worksheets.Count = len(sheets)

            workbook.Worksheets = worksheets
            workbook.ActiveSheet = sheets[0] if sheets else None

            # Add workbook methods
            workbook.Save = MagicMock()
            workbook.SaveAs = MagicMock()
            workbook.Close = MagicMock()

            self._workbooks[key] = workbook

        return self._workbooks[key]

    def create_worksheet(self, name="Sheet1"):
        """Create a mock worksheet with the given name."""
        if name not in self._worksheets:
            worksheet = MagicMock()
            worksheet.Name = name

            # Create mock range
            range_obj = MagicMock()
            range_obj.Value = None
            range_obj.Formula = None
            range_obj.Address = "$A$1"
            range_obj.GetAddress.return_value = "$A$1"
            range_obj.Font = MagicMock()
            range_obj.Interior = MagicMock()
            range_obj.Borders = MagicMock()
            range_obj.Columns = MagicMock()
            range_obj.Rows = MagicMock()
            range_obj.AutoFilter = MagicMock()
            range_obj.ClearContents = MagicMock()
            range_obj.Clear = MagicMock()

            worksheet.Range.return_value = range_obj

            # Mock UsedRange
            used_range = MagicMock()
            used_range.Address = "$A$1:$D$10"
            used_range.Rows.Count = 10
            used_range.Columns.Count = 4
            used_range.Columns.AutoFit = MagicMock()
            worksheet.UsedRange = used_range

            # Mock ChartObjects
            worksheet.ChartObjects = MagicMock()

            # Add worksheet methods
            worksheet.Activate = MagicMock()
            worksheet.Delete = MagicMock()
            worksheet.Copy = MagicMock()

            self._worksheets[name] = worksheet

        return self._worksheets[name]

    def create_range(self, value=None, formula=None):
        """Create a mock range object."""
        range_obj = MagicMock()
        range_obj.Value = value
        range_obj.Formula = formula
        range_obj.Address = "$A$1"
        range_obj.GetAddress.return_value = "$A$1"
        range_obj.Font = MagicMock()
        range_obj.Interior = MagicMock()
        range_obj.Borders = MagicMock()
        range_obj.Columns = MagicMock()
        range_obj.Rows = MagicMock()
        range_obj.AutoFilter = MagicMock()
        range_obj.ClearContents = MagicMock()
        range_obj.Clear = MagicMock()
        return range_obj


# =============================================================================
# Pytest Fixtures
# =============================================================================

@pytest.fixture(scope="session")
def mock_excel_factory():
    """Session-scoped factory for mock Excel COM objects."""
    return MockExcelCOMFactory.get_instance()


@pytest.fixture
def mock_excel_app(mock_excel_factory):
    """Create a mock Excel Application COM object (shared across tests)."""
    mock_excel_factory.reset()
    return mock_excel_factory.create_excel_app()


@pytest.fixture
def mock_workbook(mock_excel_factory):
    """Create a mock Workbook COM object."""
    return mock_excel_factory.create_workbook()


@pytest.fixture
def mock_worksheet(mock_excel_factory):
    """Create a mock Worksheet COM object."""
    return mock_excel_factory.create_worksheet()


@pytest.fixture
def mock_range(mock_excel_factory):
    """Create a mock Range COM object."""
    return mock_excel_factory.create_range()


@pytest.fixture
def mock_manager(mock_excel_app):
    """Create a mock ExcelAppManager instance with shared COM object."""
    with patch("libs.excel_com.manager.win32com.client") as mock_win32com:
        mock_win32com.client.Dispatch.return_value = mock_excel_app
        mock_win32com.client.DispatchEx.return_value = mock_excel_app

        from libs.excel_com.manager import ExcelAppManager
        manager = ExcelAppManager(visible=False, display_alerts=False)
        manager._app = mock_excel_app
        yield manager


@pytest.fixture
def temp_excel_file(tmp_path):
    """Create a temporary Excel file for testing.

    Note: This creates an empty file, not a real Excel file.
    For real Excel operations, use actual Excel files.
    """
    excel_file = tmp_path / "test.xlsx"
    excel_file.touch()
    return str(excel_file)


@pytest.fixture
def temp_dir(tmp_path):
    """Create a temporary directory for testing."""
    return str(tmp_path)


@pytest.fixture
def sample_data():
    """Provide sample 2D data for testing."""
    return [
        ["Header1", "Header2", "Header3"],
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9],
    ]


@pytest.fixture
def sample_json_data(sample_data):
    """Provide sample data as JSON string."""
    return json.dumps(sample_data)


# =============================================================================
# Auto-reset fixture to ensure test isolation
# =============================================================================

@pytest.fixture(autouse=True)
def reset_excel_com_state():
    """Automatically reset COM state before each test."""
    MockExcelCOMFactory.reset()
    yield
    # Cleanup after test if needed
    MockExcelCOMFactory.reset()
