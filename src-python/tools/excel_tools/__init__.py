"""Excel tools — composite provider aggregating all Excel tool providers."""

from libs.excel_com import ExcelInstanceManager

from .workbook_tool import WorkbookToolProvider
from .read_tool import ReadToolProvider
from .write_tool import WriteToolProvider
from .formula_tool import FormulaToolProvider
from .table_tool import TableToolProvider
from .sheet_tool import SheetToolProvider
from .format_tool import FormatToolProvider


class CompositeExcelToolProvider:
    """Aggregates all Excel tool providers into a single interface."""

    def __init__(self, manager: ExcelInstanceManager | None = None) -> None:
        self._manager = manager or ExcelInstanceManager()
        self._providers = [
            WorkbookToolProvider(self._manager),
            ReadToolProvider(self._manager),
            WriteToolProvider(self._manager),
            FormulaToolProvider(self._manager),
            TableToolProvider(self._manager),
            SheetToolProvider(self._manager),
            FormatToolProvider(self._manager),
        ]

    def get_tools(self) -> list:
        return [t for p in self._providers for t in p.get_tools()]
