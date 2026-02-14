"""Schema tool — read data source schema files from .excel-agent/schema/."""

import json
import logging
from pathlib import Path
from typing import Optional

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

logger = logging.getLogger("app.schema")

SCHEMA_DIR_NAME = "schema"


class ReadSchemaInput(BaseModel):
    source_name: Optional[str] = Field(
        default=None,
        description=(
            "数据源名称（如 'sales.xlsx'）。"
            "为空则返回所有数据源的摘要列表。"
        ),
    )


class SchemaToolProvider:
    def __init__(self, project_root: Path) -> None:
        self._schema_dir = project_root / ".excel-agent" / SCHEMA_DIR_NAME

    def get_tools(self) -> list:
        return [
            StructuredTool(
                name="read_datasource_schema",
                description=(
                    "读取数据源的结构信息（表头、列名、数据类型、样本值等）。\n"
                    "- source_name 为空：返回所有数据源的摘要，用于了解项目有哪些数据\n"
                    "- source_name='sales.xlsx'：返回该数据源的完整 schema\n"
                    "在操作任何数据源之前，应先调用此工具了解其结构。"
                ),
                args_schema=ReadSchemaInput,
                func=self._read,
            )
        ]

    def _read(self, source_name: Optional[str] = None) -> str:
        if not self._schema_dir.exists():
            return json.dumps({"success": False, "message": "暂无数据源分析结果，请等待后台分析完成或手动触发分析。"}, ensure_ascii=False)

        if source_name is None:
            return self._read_summary()
        return self._read_source(source_name)

    def _read_summary(self) -> str:
        summary_file = self._schema_dir / "summary.md"
        if not summary_file.exists():
            files = [f.stem for f in self._schema_dir.glob("*.md") if f.name != "summary.md"]
            if not files:
                return json.dumps({"success": False, "message": "暂无数据源分析结果。"}, ensure_ascii=False)
            return json.dumps({
                "success": True,
                "message": "未找到摘要文件，以下数据源有详细 schema 可读取",
                "available_sources": files,
            }, ensure_ascii=False)

        content = summary_file.read_text(encoding="utf-8")
        return json.dumps({
            "success": True,
            "summary": content,
            "hint": "Use read_datasource_schema(source_name=...) to get full schema. Schema files contain absolute file paths — use them directly with Excel tools.",
        }, ensure_ascii=False)

    def _read_source(self, source_name: str) -> str:
        import re
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", source_name)

        # Try exact match first, then stem match
        candidates = [
            self._schema_dir / f"{safe_name}.md",
            self._schema_dir / f"{source_name}.md",
        ]
        # Also try matching by stem (without extension)
        stem = Path(source_name).stem
        candidates += list(self._schema_dir.glob(f"{stem}*.md"))

        for path in candidates:
            if path.exists() and path.name != "summary.md":
                content = path.read_text(encoding="utf-8")
                return json.dumps({"success": True, "source": source_name, "schema": content}, ensure_ascii=False)

        # List available
        available = [f.stem for f in self._schema_dir.glob("*.md") if f.name != "summary.md"]
        return json.dumps({
            "success": False,
            "message": f"未找到 '{source_name}' 的 schema，可用数据源: {available}",
        }, ensure_ascii=False)
