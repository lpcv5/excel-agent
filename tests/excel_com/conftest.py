"""
Shared pytest fixtures for excel_com module tests.

This module extends the parent conftest.py with excel_com-specific fixtures.
All tests share the mock COM objects from the parent conftest.py.
"""

from unittest.mock import MagicMock, patch

import pytest


# =============================================================================
# excel_com Module-specific Fixtures
# =============================================================================

@pytest.fixture
def mock_win32com(mock_excel_app):
    """Mock the win32com.client module with shared Excel app."""
    with patch("excel_com.manager.win32com.client") as mock_client:
        mock_client.Dispatch.return_value = mock_excel_app
        mock_client.DispatchEx.return_value = mock_excel_app
        yield mock_client


@pytest.fixture
def mock_excel_app_for_manager(mock_excel_factory):
    """Create a comprehensive mock Excel Application for manager tests."""
    return mock_excel_factory.create_excel_app()


@pytest.fixture
def mock_workbook_for_ops(mock_excel_factory):
    """Create a mock workbook for workbook_ops tests."""
    return mock_excel_factory.create_workbook("test_ops.xlsx")


@pytest.fixture
def mock_worksheet_for_ops(mock_excel_factory):
    """Create a mock worksheet for operations tests."""
    return mock_excel_factory.create_worksheet("TestSheet")


@pytest.fixture
def mock_range_for_ops(mock_excel_factory):
    """Create a mock range for operations tests."""
    range_obj = mock_excel_factory.create_range()
    # Set up more specific values for operations tests
    range_obj.Value = [["A1", "B1"], ["A2", "B2"]]
    range_obj.Formula = "=SUM(A1:A10)"
    range_obj.RemoveDuplicates = MagicMock(return_value=True)
    range_obj.Sort = MagicMock()
    return range_obj


@pytest.fixture
def mock_preserve_context():
    """Create a mock context manager for preserve_user_state."""
    mock_ctx = MagicMock()
    mock_ctx.__enter__ = MagicMock(return_value=None)
    mock_ctx.__exit__ = MagicMock(return_value=None)
    return mock_ctx
