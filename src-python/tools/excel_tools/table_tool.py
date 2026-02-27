"""Table tool — create, list, style, add column, delete tables."""

import logging

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from libs.excel_com import ExcelInstanceManager
from libs.excel_com import table_ops
from ._common import format_result, safe_excel_call

logger = logging.getLogger("app.excel")


class CreateTableInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称")
    range_address: str = Field(description="数据区域(如'A1:D10')")
    table_name: str = Field(description="表格名称")
    has_headers: bool = Field(default=True, description="第一行是否为表头")
    style: str | None = Field(
        default=None, description="表格样式名(如'TableStyleMedium2')"
    )
    save: bool = Field(default=True, description="操作后是否保存")


class ListTablesInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称")


class AddTableColumnInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称")
    table_name: str = Field(description="表格名称")
    column_name: str = Field(description="新列名称")
    formula: str | None = Field(default=None, description="列公式(可选)")
    save: bool = Field(default=True, description="操作后是否保存")


class TableToolProvider:
    def __init__(self, manager: ExcelInstanceManager):
        self._mgr = manager

    def get_tools(self) -> list:
        return [
            StructuredTool(
                name="table_create",
                description=(
                    "将Excel数据区域转换为表格(ListObject)。可指定表格名称和样式。"
                ),
                args_schema=CreateTableInput,
                func=self._create,
            ),
            StructuredTool(
                name="list_excel_tables",
                description="列出工作表上的所有表格。",
                args_schema=ListTablesInput,
                func=self._list,
            ),
            StructuredTool(
                name="add_table_column",
                description=(
                    "向已有表格添加计算列。"
                    "可指定公式，支持结构化引用如[@Price]*[@Qty]。"
                ),
                args_schema=AddTableColumnInput,
                func=self._add_column,
            ),
        ]

    @safe_excel_call
    def _create(
        self,
        file_path: str,
        range_address: str,
        table_name: str,
        sheet: str | None = None,
        has_headers: bool = True,
        style: str | None = None,
        save: bool = True,
    ) -> str:
        ws = self._mgr.get_sheet(file_path, sheet)
        tbl = table_ops.create_table(ws, range_address, table_name, has_headers)
        if style:
            table_ops.set_table_style(ws, table_name, style)
        if save:
            self._mgr.save_workbook(file_path)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已创建表格 '{table_name}' 在 {range_address}",
                    "table_range": tbl.Range.Address,
                }
            ],
            file_saved=save,
        )

    @safe_excel_call
    def _list(
        self,
        file_path: str,
        sheet: str | None = None,
    ) -> str:
        ws = self._mgr.get_sheet(file_path, sheet)
        tables = table_ops.list_tables(ws)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "tables": tables,
                }
            ],
        )

    @safe_excel_call
    def _add_column(
        self,
        file_path: str,
        table_name: str,
        column_name: str,
        sheet: str | None = None,
        formula: str | None = None,
        save: bool = True,
    ) -> str:
        ws = self._mgr.get_sheet(file_path, sheet)
        table_ops.add_table_column(ws, table_name, column_name, formula)
        if save:
            self._mgr.save_workbook(file_path)
        msg = f"已添加列 '{column_name}' 到表格 '{table_name}'"
        if formula:
            msg += f"，公式: {formula}"
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": msg,
                }
            ],
            file_saved=save,
        )
