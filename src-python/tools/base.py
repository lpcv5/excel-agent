from typing import Protocol, runtime_checkable
from langchain_core.tools import BaseTool


@runtime_checkable
class ToolProvider(Protocol):
    def get_tools(self) -> list[BaseTool]: ...
