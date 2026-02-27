"""Write tool — write data to Excel ranges."""

import logging

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from libs.excel_com import ExcelInstanceManager
from libs.excel_com import range_ops
from ._common import format_result, safe_excel_call

logger = logging.getLogger("app.excel")


class WriteRangeInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称，默认活动工作表")
    start_cell: str = Field(description="写入起始单元格(如'A1')")
    values: list[list] = Field(description="要写入的二维数组数据")
    auto_fit: bool = Field(default=True, description="写入后是否自动调整列宽")
    save: bool = Field(default=True, description="写入后是否保存文件")


class WriteToolProvider:
    def __init__(self, manager: ExcelInstanceManager):
        self._mgr = manager

    def get_tools(self) -> list:
        return [
            StructuredTool(
                name="write_excel_range",
                description=(
                    "向Excel工作表写入数据。"
                    "提供起始单元格和二维数组数据，自动写入对应区域。"
                ),
                args_schema=WriteRangeInput,
                func=self._write,
            ),
        ]

    @safe_excel_call
    def _write(
        self,
        file_path: str,
        start_cell: str,
        values: list[list],
        sheet: str | None = None,
        auto_fit: bool = True,
        save: bool = True,
    ) -> str:
        with self._mgr.batch_operation(file_path):
            ws = self._mgr.get_sheet(file_path, sheet)
            written_addr = range_ops.write_range(ws, start_cell, values)

            if auto_fit:
                range_ops.auto_fit_columns(ws, written_addr)

        if save:
            self._mgr.save_workbook(file_path)

        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已写入 {len(values)}行x{len(values[0])}列 到 {written_addr}",
                    "range": written_addr,
                    "rows": len(values),
                    "cols": len(values[0]),
                }
            ],
            file_saved=save,
        )
