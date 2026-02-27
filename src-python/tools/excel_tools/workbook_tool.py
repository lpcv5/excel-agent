"""Workbook tool — open, save, close, create, info, list."""

import logging

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from libs.excel_com import ExcelInstanceManager
from libs.excel_com import sheet_ops, table_ops
from ._common import format_result, safe_excel_call

logger = logging.getLogger("app.excel")


class CreateWorkbookInput(BaseModel):
    file_path: str = Field(description="保存新工作簿的完整路径")


class OpenWorkbookInput(BaseModel):
    file_path: str = Field(description="要打开的Excel文件路径")


class SaveWorkbookInput(BaseModel):
    file_path: str = Field(description="要保存的Excel文件路径")


class WorkbookInfoInput(BaseModel):
    file_path: str = Field(description="要查询信息的Excel文件路径")


class WorkbookToolProvider:
    def __init__(self, manager: ExcelInstanceManager):
        self._mgr = manager

    def get_tools(self) -> list:
        return [
            StructuredTool(
                name="create_workbook",
                description=(
                    "创建一个新的Excel工作簿并保存到指定路径。返回工作簿信息。"
                ),
                args_schema=CreateWorkbookInput,
                func=self._create,
            ),
            StructuredTool(
                name="open_workbook",
                description="打开一个已有的Excel文件。",
                args_schema=OpenWorkbookInput,
                func=self._open,
            ),
            StructuredTool(
                name="save_workbook",
                description="保存一个已打开的Excel文件。",
                args_schema=SaveWorkbookInput,
                func=self._save,
            ),
            StructuredTool(
                name="workbook_info",
                description=(
                    "获取工作簿概要信息：工作表列表、表格列表等。"
                    "如果文件未打开会自动打开。"
                ),
                args_schema=WorkbookInfoInput,
                func=self._info,
            ),
        ]

    @safe_excel_call
    def _create(self, file_path: str) -> str:
        entry = self._mgr.create_workbook(file_path)
        wb = entry.workbook
        sheets = sheet_ops.list_sheets(wb)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已创建工作簿: {file_path}",
                    "sheets": sheets,
                }
            ],
        )

    @safe_excel_call
    def _open(self, file_path: str) -> str:
        entry = self._mgr.open_workbook(file_path)
        wb = entry.workbook
        sheets = sheet_ops.list_sheets(wb)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已打开工作簿: {file_path}",
                    "sheets": sheets,
                }
            ],
        )

    @safe_excel_call
    def _save(self, file_path: str) -> str:
        self._mgr.save_workbook(file_path)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已保存: {file_path}",
                }
            ],
            file_saved=True,
        )

    @safe_excel_call
    def _info(self, file_path: str) -> str:
        entry = self._mgr.get_workbook(file_path)
        wb = entry.workbook
        sheets = sheet_ops.list_sheets(wb)
        tables_info = []
        for s_name in sheets:
            ws = wb.Sheets(s_name)
            for t in table_ops.list_tables(ws):
                t["sheet"] = s_name
                tables_info.append(t)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "file_path": file_path,
                    "sheets": sheets,
                    "tables": tables_info,
                }
            ],
        )
