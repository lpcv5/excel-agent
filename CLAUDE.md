# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Agent is a DeepAgents-based intelligent Excel processing agent that uses Windows COM to interact directly with Microsoft Excel. It provides natural language interaction for data analysis, formula generation, and report creation.

**Platform Requirement**: Windows with Microsoft Excel installed.

## Key Commands

```bash
# Install dependencies
uv sync

# Run in interactive CLI mode (default)
uv run python agent.py

# Run single query
uv run python agent.py "Read sales.xlsx and show the first 10 rows"

# Run with Web UI
uv run python agent.py --web

# List available tools
uv run python agent.py --list-tools

# Run with specific model
uv run python agent.py --model openai:gpt-4 "Analyze data.xlsx"

# Run with LLM call logging (for debugging)
uv run python agent.py --log-level DEBUG "Read data.xlsx"

# Run tests
uv run pytest

# Run specific test file
uv run pytest tests/test_agent.py

# Run tests with coverage
uv run pytest --cov
```

## Architecture

The codebase has four layers:

```
agent.py               # CLI entry point with UI mode selection
excel_agent/           # UI-agnostic agent core
├── core.py            # AgentCore - agent creation, streaming, events
├── config.py          # AgentConfig dataclass with logging settings
├── events.py          # Event types for UI consumption
├── session.py         # Session management
└── callbacks/         # LangChain callbacks
tools/
└── excel_tool.py      # LangChain @tool wrappers, returns JSON strings
ui/                    # UI implementations
├── cli/runner.py      # CLI interface (default)
└── web/server.py      # Web UI interface (--web flag)
libs/
├── excel_com/         # COM interface layer for Excel operations
│   ├── manager.py     # ExcelAppManager - COM lifecycle, workbook tracking
│   ├── workbook_ops.py
│   ├── formatting_ops.py
│   ├── formula_ops.py
│   ├── advanced_ops.py
│   ├── constants.py   # Excel enum mappings (e.g., xlCenter)
│   ├── context.py     # Context managers for preserving Excel state
│   └── exceptions.py
├── stream_msg_parser/ # Streaming message parser for LangGraph
│   ├── parser.py
│   └── events.py
└── deepagents/        # DeepAgents framework (submodule)
```

**Data Flow**: Natural language query → UI (CLI/Web) → AgentCore → DeepAgent → excel_tool.py → libs/excel_com/ → Windows COM → Excel

### AgentCore Pattern

The `AgentCore` class in `excel_agent/core.py` is the UI-agnostic interface:
- `astream_query()`: Async streaming with event-based output
- `invoke()`: Single query, returns string
- Emits events: `TextEvent`, `ToolCallStartEvent`, `ToolResultEvent`, `ErrorEvent`, etc.

### COM Layer Patterns

The excel_com layer uses thread-local storage for COM objects (COM cannot cross threads):

```python
# Thread-local manager pattern (from tools/excel_tool.py)
_thread_local = threading.local()

def get_excel_manager() -> ExcelAppManager:
    if not hasattr(_thread_local, 'manager'):
        _thread_local.manager = ExcelAppManager(visible=False, display_alerts=False)
    return _thread_local.manager
```

ExcelAppManager tracks workbook ownership - workbooks opened by the user are NOT closed by the agent.

### Tool Implementation Pattern

Tools use LangChain's `@tool` decorator and return JSON strings:

```python
@tool
def my_tool(param: str) -> str:
    """Tool description for the LLM."""
    try:
        result = {"success": True, "data": ...}
        return json.dumps(result, indent=2, default=str, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})
```

### Testing

Tests use a shared mock COM factory (tests/conftest.py) with pytest fixtures. All COM-related tests mock the win32com.client.Dispatch calls.

## Dependencies

- Python 3.12+
- deepagents (local: libs/deepagents/libs/deepagents)
- stream_msg_parser (local: libs/stream_msg_parser)
- pywin32 (Windows COM interface)
- langgraph (state graph, checkpointing)
- langchain-openai (LLM provider)
