"""Format tool — number formats, column widths, etc."""

import logging

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from libs.excel_com import ExcelInstanceManager
from libs.excel_com import range_ops
from ._common import format_result, safe_excel_call

logger = logging.getLogger("app.excel")


class SetNumberFormatInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称")
    range_address: str = Field(description="要设置格式的区域")
    number_format: str = Field(description="数字格式字符串(如'¥#,##0.00')")
    save: bool = Field(default=True, description="操作后是否保存")


class AutoFitInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称")
    range_address: str | None = Field(
        default=None, description="区域，为空则自动调整所有列"
    )
    save: bool = Field(default=True, description="操作后是否保存")


class FormatToolProvider:
    def __init__(self, manager: ExcelInstanceManager):
        self._mgr = manager

    def get_tools(self) -> list:
        return [
            StructuredTool(
                name="set_number_format",
                description=(
                    "设置Excel区域的数字格式。"
                    "常用格式：'¥#,##0.00'(货币), '#,##0'(千分位整数), "
                    "'0.00%'(百分比), 'yyyy-mm-dd'(日期)"
                ),
                args_schema=SetNumberFormatInput,
                func=self._set_format,
            ),
            StructuredTool(
                name="auto_fit_columns",
                description="自动调整列宽以适应内容。",
                args_schema=AutoFitInput,
                func=self._auto_fit,
            ),
        ]

    @safe_excel_call
    def _set_format(
        self,
        file_path: str,
        range_address: str,
        number_format: str,
        sheet: str | None = None,
        save: bool = True,
    ) -> str:
        ws = self._mgr.get_sheet(file_path, sheet)
        range_ops.set_number_format(ws, range_address, number_format)
        if save:
            self._mgr.save_workbook(file_path)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已设置 {range_address} 格式为: {number_format}",
                }
            ],
            file_saved=save,
        )

    @safe_excel_call
    def _auto_fit(
        self,
        file_path: str,
        sheet: str | None = None,
        range_address: str | None = None,
        save: bool = True,
    ) -> str:
        ws = self._mgr.get_sheet(file_path, sheet)
        range_ops.auto_fit_columns(ws, range_address)
        if save:
            self._mgr.save_workbook(file_path)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": "已自动调整列宽",
                }
            ],
            file_saved=save,
        )
