"""Excel COM automation library."""

from .com_thread import run_on_com_thread, shutdown_com_thread
from .instance_manager import ExcelInstanceManager, WorkbookEntry
from .errors import (
    ExcelComError,
    ExcelInstanceError,
    WorkbookNotFoundError,
    WorkbookReadOnlyError,
    SheetNotFoundError,
    RangeError,
    FormulaError,
    TableError,
    QueryError,
    ComCallError,
    StaleReferenceError,
)

__all__ = [
    "run_on_com_thread",
    "shutdown_com_thread",
    "ExcelInstanceManager",
    "WorkbookEntry",
    "ExcelComError",
    "ExcelInstanceError",
    "WorkbookNotFoundError",
    "WorkbookReadOnlyError",
    "SheetNotFoundError",
    "RangeError",
    "FormulaError",
    "TableError",
    "QueryError",
    "ComCallError",
    "StaleReferenceError",
]
