# Excel Agent

A DeepAgents-based intelligent Excel processing agent that understands natural language commands for data analysis, formula generation, and report creation.

## Features

- **Data Analysis**: Read, explore, and analyze Excel files with statistical methods
- **Formula Generation**: Generate Excel formulas from natural language descriptions
- **Report Generation**: Create formatted summaries and export reports
- **File Operations**: Read, write, filter, and transform Excel data

## Installation

```bash
# Clone the repository
git clone <repository-url>
cd excel-agent

# Install dependencies with uv
uv sync
```

## Usage

### Interactive Mode

```bash
uv run python agent.py
```

### Single Query

```bash
uv run python agent.py "Read sales.xlsx and show the first 10 rows"
```

### List Available Tools

```bash
uv run python agent.py --list-tools
```

## Examples

### Data Analysis

```bash
# Analyze a file
uv run python agent.py "Analyze the data in report.xlsx and give me a summary"

# Check for missing values
uv run python agent.py "Check data quality in customers.xlsx"
```

### Formula Generation

```bash
# Generate a formula
uv run python agent.py "Create a formula to calculate 10% bonus for sales over 5000"

# Lookup formula
uv run python agent.py "I need a formula to look up product prices from another sheet"
```

### Report Generation

```bash
# Create a summary report
uv run python agent.py "Create a monthly sales summary report from data.xlsx"
```

## Project Structure

```
excel-agent/
├── agent.py              # Main CLI entry point
├── excel_tools.py        # Custom Excel tools
├── AGENTS.md             # Agent identity and behavior rules
├── skills/               # Specialized skill definitions
│   ├── data-analysis/
│   │   └── SKILL.md
│   ├── formula-writing/
│   │   └── SKILL.md
│   └── report-generation/
│       └── SKILL.md
├── docs/
│   └── deepagents-guide.md   # DeepAgents development guide
├── pyproject.toml
└── README.md
```

## Available Tools

| Tool | Description |
|------|-------------|
| `read_excel` | Read data from Excel files |
| `list_sheets` | List all sheets in a workbook |
| `write_excel` | Write data to Excel files |
| `analyze_data` | Perform statistical analysis |
| `generate_formula` | Generate Excel formula suggestions |
| `create_pivot_summary` | Create pivot table summaries |
| `filter_data` | Filter data based on conditions |

## Development

See [docs/deepagents-guide.md](docs/deepagents-guide.md) for detailed documentation on DeepAgents framework and development best practices.

## Requirements

- Python 3.12+
- deepagents
- pandas
- openpyxl
- langchain-anthropic

## License

MIT
