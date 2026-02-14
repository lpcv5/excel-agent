"""Formula tool — read and set formulas."""

import logging

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from libs.excel_com import ExcelInstanceManager
from libs.excel_com import formula_ops
from ._common import format_result, safe_excel_call

logger = logging.getLogger("app.excel")


class SetFormulaInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称")
    range_address: str = Field(description="设置公式的区域(如'F2:F100')")
    formula: str = Field(description="公式字符串(如'=D2*12')，使用英文函数名")
    save: bool = Field(default=True, description="设置后是否保存")


class GetFormulaInput(BaseModel):
    file_path: str = Field(description="Excel文件路径")
    sheet: str | None = Field(default=None, description="工作表名称")
    range_address: str = Field(description="读取公式的区域")


class FormulaToolProvider:
    def __init__(self, manager: ExcelInstanceManager):
        self._mgr = manager

    def get_tools(self) -> list:
        return [
            StructuredTool(
                name="range_set_formula",
                description=(
                    "在Excel单元格或区域设置公式。使用英文函数名(如SUM, IF, VLOOKUP)。"
                    "对区域设置公式时Excel会自动调整行引用。"
                ),
                args_schema=SetFormulaInput,
                func=self._set_formula,
            ),
            StructuredTool(
                name="get_excel_formula",
                description="读取Excel单元格或区域中的公式。",
                args_schema=GetFormulaInput,
                func=self._get_formula,
            ),
        ]

    @safe_excel_call
    def _set_formula(
        self,
        file_path: str,
        range_address: str,
        formula: str,
        sheet: str | None = None,
        save: bool = True,
    ) -> str:
        ws = self._mgr.get_sheet(file_path, sheet)
        formula_ops.set_formula(ws, range_address, formula)
        if save:
            self._mgr.save_workbook(file_path)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "message": f"已在 {range_address} 设置公式: {formula}",
                }
            ],
            file_saved=save,
        )

    @safe_excel_call
    def _get_formula(
        self,
        file_path: str,
        range_address: str,
        sheet: str | None = None,
    ) -> str:
        ws = self._mgr.get_sheet(file_path, sheet)
        formulas = formula_ops.get_formula(ws, range_address)
        return format_result(
            True,
            [
                {
                    "status": "ok",
                    "range": range_address,
                    "formulas": formulas,
                }
            ],
        )
