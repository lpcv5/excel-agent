"""Subagent-as-tool provider.

Wraps a DeepAgent as a LangChain StructuredTool so the main agent
can delegate to specialized subagents via tool calls.
"""

from typing import Optional

from langchain_core.tools import BaseTool, StructuredTool
from langchain_core.language_models import BaseChatModel
from pydantic import BaseModel, Field


class SubAgentInput(BaseModel):
    query: str = Field(description="The task or question to delegate to this subagent")


def create_subagent_tool(
    name: str,
    description: str,
    model: BaseChatModel,
    tools: list[BaseTool],
    system_prompt: Optional[str] = None,
) -> BaseTool:
    """Create a tool that delegates to a subagent."""
    from libs.deepagents import create_deep_agent
    from langgraph.checkpoint.memory import MemorySaver

    agent = create_deep_agent(
        model=model,
        tools=tools,
        system_prompt=system_prompt or "",
        checkpointer=MemorySaver(),
    )

    def run(query: str) -> str:
        result = agent.invoke(
            {"messages": [("user", query)]},
            config={"configurable": {"thread_id": "subagent"}},
        )
        messages = result.get("messages", [])
        if messages:
            content = messages[-1].content
            if isinstance(content, list):
                return "".join(
                    p.get("text", "") for p in content if isinstance(p, dict)
                )
            return str(content)
        return ""

    async def arun(query: str) -> str:
        result = await agent.ainvoke(
            {"messages": [("user", query)]},
            config={"configurable": {"thread_id": "subagent"}},
        )
        messages = result.get("messages", [])
        if messages:
            content = messages[-1].content
            if isinstance(content, list):
                return "".join(
                    p.get("text", "") for p in content if isinstance(p, dict)
                )
            return str(content)
        return ""

    return StructuredTool(
        name=name,
        description=description,
        args_schema=SubAgentInput,
        func=run,
        coroutine=arun,
    )


class SubAgentToolProvider:
    """Provides a subagent as a tool for the main agent."""

    def __init__(
        self,
        name: str,
        description: str,
        model: BaseChatModel,
        tools: list[BaseTool],
        system_prompt: Optional[str] = None,
    ):
        self.name = name
        self.description = description
        self.model = model
        self.tools = tools
        self.system_prompt = system_prompt
        self._tool: Optional[BaseTool] = None

    def get_tools(self) -> list[BaseTool]:
        if self._tool is None:
            self._tool = create_subagent_tool(
                name=self.name,
                description=self.description,
                model=self.model,
                tools=self.tools,
                system_prompt=self.system_prompt,
            )
        return [self._tool]
