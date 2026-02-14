# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Agent is a DeepAgents-based intelligent Excel processing agent that uses Windows COM to interact directly with Microsoft Excel. It provides natural language interaction for data analysis, formula generation, and report creation.

**Platform Requirement**: Windows with Microsoft Excel installed.

## Key Commands

```bash
# Install dependencies
uv sync

# Run in interactive mode
uv run python agent.py

# Run single query
uv run python agent.py "Read sales.xlsx and show the first 10 rows"

# List available tools
uv run python agent.py --list-tools

# Run tests
uv run pytest

# Run specific test file
uv run pytest tests/test_agent.py

# Run tests with coverage
uv run pytest --cov
```

## Architecture

The codebase has three layers:

```
agent.py              # CLI entry point, creates DeepAgent
excel_tools.py        # LangChain @tool wrappers, returns JSON strings
excel_com/            # COM interface layer for Excel operations
├── manager.py        # ExcelAppManager - COM lifecycle, workbook tracking
├── workbook_ops.py   # Open/close/read/write workbooks
├── formatting_ops.py # Font, cell, border, background formatting
├── formula_ops.py    # Get/set formulas
├── constants.py      # Excel enum mappings (e.g., xlCenter)
├── context.py        # Context managers for preserving Excel state
└── exceptions.py     # Custom exceptions
```

**Data Flow**: Natural language query → DeepAgent → excel_tools.py → excel_com/ → Windows COM → Excel

### Agent Configuration

The agent is configured in `create_excel_agent()` using `create_deep_agent()` from the DeepAgents harness:

- `model`: LLM provider (default: `openai:gpt-5-mini`)
- `tools`: EXCEL_TOOLS list from excel_tools.py
- `memory`: AGENTS.md file defining agent identity and safety rules
- `skills`: skills/ directory for specialized workflows
- `backend`: FilesystemBackend for file operations
- `checkpointer`: MemorySaver for conversation persistence

### COM Layer Patterns

The excel_com layer uses thread-local storage for COM objects (COM cannot cross threads):

```python
# Thread-local manager pattern (from excel_tools.py)
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
- pywin32 (Windows COM interface)
- langgraph (state graph, checkpointing)
- langchain-openai (LLM provider)
