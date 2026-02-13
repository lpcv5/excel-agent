# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Agent is a DeepAgents-based intelligent Excel processing agent. It uses the DeepAgents framework (built on LangChain/LangGraph) to provide natural language interaction with Excel files for data analysis, formula generation, and report creation.

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
```

## Architecture

The agent follows the DeepAgents framework pattern with these components:

- **agent.py**: CLI entry point that creates the DeepAgent using `create_deep_agent()` from the deepagents package
- **excel_tools.py**: LangChain tools decorated with `@tool` - all tools return JSON strings for structured output
- **AGENTS.md**: Defines agent identity, capabilities, safety rules, and workflow guidelines (loaded as memory)
- **skills/**: Contains SKILL.md files with specialized workflow guidance using YAML frontmatter

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

### DeepAgents Configuration

The agent is configured in `create_excel_agent()` with:
- `model`: LLM provider (default: anthropic:claude-sonnet-4-20250514)
- `tools`: EXCEL_TOOLS list from excel_tools.py
- `memory`: AGENTS.md path
- `skills`: skills/ directory path
- `backend`: FilesystemBackend for file operations

## Dependencies

- Python 3.12+
- deepagents (DeepAgents framework)
- pandas + openpyxl (Excel processing)
- langchain-anthropic (LLM integration)
