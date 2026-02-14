# Excel Agent

A DeepAgents-based intelligent Excel processing agent that uses Windows COM to interact directly with Microsoft Excel. It provides natural language interaction for data analysis, formula generation, and report creation.

**Platform Requirement**: Windows with Microsoft Excel installed.

## Features

- **Direct Excel Integration**: Works with real Excel files via Windows COM interface
- **Natural Language Interface**: Interact with Excel using plain English commands
- **Multiple UI Modes**: CLI (default) and Web UI (`--web` flag)
- **Comprehensive Operations**: Read, write, format, and manipulate Excel data
- **Formula Support**: Set and read Excel formulas
- **Formatting Tools**: Font, cell, border, and background formatting

## Installation

```bash
# Clone the repository
git clone <repository-url>
cd excel-agent

# Install dependencies with uv
uv sync
```

## Usage

### Interactive CLI Mode (Default)

```bash
uv run python agent.py
```

### Single Query

```bash
uv run python agent.py "Read sales.xlsx and show the first 10 rows"
```

### Web UI Mode

```bash
uv run python agent.py --web
```

### List Available Tools

```bash
uv run python agent.py --list-tools
```

### With Specific Model

```bash
uv run python agent.py --model openai:gpt-4 "Analyze data.xlsx"
```

### With LLM Call Logging (Debugging)

```bash
uv run python agent.py --log-level DEBUG "Read data.xlsx"
uv run python agent.py --log-level DEBUG --log-file logs/llm.log "Analyze data.xlsx"
```

## Examples

### Data Operations

```bash
# Open and read data
uv run python agent.py "Open sales.xlsx and show me the first 10 rows"

# Write data
uv run python agent.py "Write the values 'Q1', 'Q2', 'Q3', 'Q4' in cells A1:A4 of report.xlsx"
```

### Formatting

```bash
# Format headers
uv run python agent.py "Format the header row in data.xlsx with bold text and yellow background"

# Adjust columns
uv run python agent.py "Auto-fit all columns in sales.xlsx"
```

### Formulas

```bash
# Add a formula
uv run python agent.py "Add a SUM formula in cell C10 to sum C1:C9 in budget.xlsx"

# Read formulas
uv run python agent.py "Show me the formula in cell D5 of calculations.xlsx"
```

## Project Structure

```
excel-agent/
├── agent.py              # CLI entry point with UI mode selection
├── excel_agent/          # UI-agnostic agent core
│   ├── core.py           # AgentCore - agent creation and streaming
│   ├── config.py         # Configuration dataclass
│   └── events.py         # Event types for UI consumption
├── tools/
│   └── excel_tool.py     # LangChain @tool wrappers for Excel operations
├── ui/                   # UI implementations
│   ├── cli/              # CLI interface (default)
│   └── web/              # Web UI interface
├── libs/
│   ├── excel_com/        # COM interface layer for Excel operations
│   ├── stream_msg_parser/ # Streaming message parser
│   └── deepagents/       # DeepAgents framework (submodule)
├── skills/               # Specialized skill definitions
├── tests/                # Test suite
├── AGENTS.md             # Agent identity and behavior rules
└── pyproject.toml
```

## Available Tools

### Status (1)
| Tool | Description |
|------|-------------|
| `excel_status` | Get current Excel application status and open workbooks |

### Workbook Operations (7)
| Tool | Description |
|------|-------------|
| `excel_open_workbook` | Open an Excel workbook file |
| `excel_create_workbook` | Create a new Excel workbook |
| `excel_list_worksheets` | List all worksheets in a workbook |
| `excel_read_range` | Read data from a cell range |
| `excel_write_range` | Write data to a cell range |
| `excel_save_workbook` | Save the workbook (optionally as new file) |
| `excel_close_workbook` | Close an open workbook |

### Worksheet Operations (5)
| Tool | Description |
|------|-------------|
| `excel_add_worksheet` | Add a new worksheet |
| `excel_delete_worksheet` | Delete a worksheet |
| `excel_rename_worksheet` | Rename a worksheet |
| `excel_copy_worksheet` | Copy a worksheet |
| `excel_get_used_range` | Get the used range of a worksheet |

### Formatting (4)
| Tool | Description |
|------|-------------|
| `excel_set_font_format` | Set font name, size, bold, italic, underline, color |
| `excel_set_cell_format` | Set alignment, wrap text, number format |
| `excel_set_border_format` | Set border style, weight, and color |
| `excel_set_background_color` | Set cell background color |

### Formula Operations (2)
| Tool | Description |
|------|-------------|
| `excel_set_formula` | Set a formula in a cell |
| `excel_get_formula` | Get the formula from a cell |

### Column/Row Operations (3)
| Tool | Description |
|------|-------------|
| `excel_auto_fit_columns` | Auto-fit column widths to content |
| `excel_set_column_width` | Set specific column width |
| `excel_set_row_height` | Set specific row height |

## Requirements

- Python 3.12+
- Windows operating system
- Microsoft Excel installed
- OpenAI API key (or compatible LLM provider)

## Dependencies

- [deepagents](https://github.com/anthropics/deepagents) - Agent framework
- [pywin32](https://github.com/mhammond/pywin32) - Windows COM interface
- [langchain-openai](https://github.com/langchain-ai/langchain) - LLM provider
- [langgraph](https://github.com/langchain-ai/langgraph) - State graph and checkpointing

## License

MIT
