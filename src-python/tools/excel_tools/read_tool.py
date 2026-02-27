"""Read tool — read ranges and used ranges from Excel."""

import logging

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from libs.excel_com import ExcelInstanceManager
from libs.excel_com import range_ops
from ._common import format_result, safe_excel_call

logger = logging.getLogger("app.excel")

MAX_RETURN_ROWS = 100


class ReadRangeInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称，默认活动工作表")
    range_address: str | None = Field(
        default=None, description="要读取的区域地址(如'A1:D10')，为空则读取已用区域"
    )


class ReadToolProvider:
    def __init__(self, manager: ExcelInstanceManager):
        self._mgr = manager

    def get_tools(self) -> list:
        return [
            StructuredTool(
                name="read_excel_range",
                description=(
                    "读取Excel工作表中的数据。"
                    "可指定区域地址，不指定则读取整个已用区域。"
                    "返回二维数组数据，最多返回100行。"
                ),
                args_schema=ReadRangeInput,
                func=self._read,
            ),
        ]

    @safe_excel_call
    def _read(
        self,
        file_path: str,
        sheet: str | None = None,
        range_address: str | None = None,
    ) -> str:
        ws = self._mgr.get_sheet(file_path, sheet)

        if range_address is None:
            addr, data, total_rows, total_cols = range_ops.read_used_range(ws)
        else:
            data = range_ops.read_range(ws, range_address)
            addr = range_address
            total_rows = len(data)
            total_cols = len(data[0]) if data else 0

        truncated = False
        if len(data) > MAX_RETURN_ROWS:
            data = data[:MAX_RETURN_ROWS]
            truncated = True

        result: dict = {
            "status": "ok",
            "range": addr,
            "total_rows": total_rows,
            "total_cols": total_cols,
            "data": data,
        }
        if truncated:
            result["truncated"] = True
            result["message"] = (
                f"数据已截断，仅返回前{MAX_RETURN_ROWS}行(共{total_rows}行)"
            )

        return format_result(True, [result])
