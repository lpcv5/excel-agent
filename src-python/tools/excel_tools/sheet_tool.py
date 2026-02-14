"""Sheet tool — add, delete, rename, copy, list worksheets."""

import logging

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from libs.excel_com import ExcelInstanceManager
from libs.excel_com import sheet_ops
from ._common import format_result, safe_excel_call

logger = logging.getLogger("app.excel")


class AddSheetInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet_name: str = Field(description="新工作表名称")
    save: bool = Field(default=True, description="操作后是否保存")


class DeleteSheetInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet_name: str = Field(description="要删除的工作表名称")
    save: bool = Field(default=True, description="操作后是否保存")


class RenameSheetInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    old_name: str = Field(description="原工作表名称")
    new_name: str = Field(description="新工作表名称")
    save: bool = Field(default=True, description="操作后是否保存")


class ListSheetsInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")


class SheetToolProvider:
    def __init__(self, manager: ExcelInstanceManager):
        self._mgr = manager

    def get_tools(self) -> list:
        return [
            StructuredTool(
                name="add_excel_sheet",
                description="添加新工作表。",
                args_schema=AddSheetInput,
                func=self._add,
            ),
            StructuredTool(
                name="delete_excel_sheet",
                description="删除工作表。",
                args_schema=DeleteSheetInput,
                func=self._delete,
            ),
            StructuredTool(
                name="rename_excel_sheet",
                description="重命名工作表。",
                args_schema=RenameSheetInput,
                func=self._rename,
            ),
            StructuredTool(
                name="list_excel_sheets",
                description="列出工作簿中所有工作表名称。",
                args_schema=ListSheetsInput,
                func=self._list,
            ),
        ]

    @safe_excel_call
    def _add(self, file_path: str, sheet_name: str, save: bool = True) -> str:
        entry = self._mgr.get_workbook(file_path)
        name = sheet_ops.add_sheet(entry.workbook, sheet_name)
        if save:
            self._mgr.save_workbook(file_path)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已添加工作表: {name}",
                }
            ],
            file_saved=save,
        )

    @safe_excel_call
    def _delete(self, file_path: str, sheet_name: str, save: bool = True) -> str:
        entry = self._mgr.get_workbook(file_path)
        sheet_ops.delete_sheet(self._mgr.app, entry.workbook, sheet_name)
        if save:
            self._mgr.save_workbook(file_path)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已删除工作表: {sheet_name}",
                }
            ],
            file_saved=save,
        )

    @safe_excel_call
    def _rename(
        self,
        file_path: str,
        old_name: str,
        new_name: str,
        save: bool = True,
    ) -> str:
        entry = self._mgr.get_workbook(file_path)
        sheet_ops.rename_sheet(entry.workbook, old_name, new_name)
        if save:
            self._mgr.save_workbook(file_path)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已重命名: {old_name} -> {new_name}",
                }
            ],
            file_saved=save,
        )

    @safe_excel_call
    def _list(self, file_path: str) -> str:
        entry = self._mgr.get_workbook(file_path)
        sheets = sheet_ops.list_sheets(entry.workbook)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "sheets": sheets,
                }
            ],
        )
