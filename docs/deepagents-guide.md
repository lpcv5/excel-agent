# DeepAgents 开发指南

本指南详细介绍如何使用 DeepAgents 框架构建智能代理应用。DeepAgents 是基于 LangChain 和 LangGraph 的"开箱即用"的 Agent 框架，提供内置的规划、文件操作、Shell 访问和子代理支持。

## 目录

1. [概述与架构](#概述与架构)
2. [核心概念详解](#核心概念详解)
3. [目录结构规范](#目录结构规范)
4. [AGENTS.md 编写指南](#agentsmd-编写指南)
5. [Skills 编写指南](#skills-编写指南)
6. [Subagents 配置指南](#subagents-配置指南)
7. [自定义工具开发](#自定义工具开发)
8. [完整示例代码](#完整示例代码)
9. [最佳实践](#最佳实践)

---

## 概述与架构

### 什么是 DeepAgents

DeepAgents 是一个生产级的 AI Agent 框架，它封装了 LangChain 和 LangGraph 的复杂性，提供简洁的 API 来创建功能强大的智能代理。

**核心特性：**

| 特性 | 说明 |
|------|------|
| **Planning** | `write_todos` / `read_todos` 任务分解与追踪 |
| **Filesystem** | `read_file`, `write_file`, `edit_file`, `ls`, `glob`, `grep` 文件操作 |
| **Shell access** | `execute` 执行命令（带沙箱选项） |
| **Sub-agents** | `task` 委托子代理完成复杂任务 |
| **Context management** | 自动摘要，大输出保存到文件 |

### 架构图

```
┌─────────────────────────────────────────────────────────┐
│                    DeepAgents                           │
├─────────────────────────────────────────────────────────┤
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐     │
│  │   Memory    │  │   Skills    │  │ Subagents   │     │
│  │ (AGENTS.md) │  │ (*.md files)│  │ (YAML/code) │     │
│  └──────┬──────┘  └──────┬──────┘  └──────┬──────┘     │
│         │                │                │             │
│         └────────────────┼────────────────┘             │
│                          ▼                              │
│  ┌───────────────────────────────────────────────────┐  │
│  │              create_deep_agent()                  │  │
│  │  ┌─────────┐ ┌─────────┐ ┌─────────┐ ┌────────┐  │  │
│  │  │ Planning│ │  Tools  │ │ Middleware│ │ Backend│  │  │
│  │  └─────────┘ └─────────┘ └─────────┘ └────────┘  │  │
│  └───────────────────────────────────────────────────┘  │
│                          │                              │
│                          ▼                              │
│  ┌───────────────────────────────────────────────────┐  │
│  │                 LangGraph Runtime                 │  │
│  └───────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────┘
```

---

## 核心概念详解

### `create_deep_agent` API

这是创建 Agent 的主要入口函数：

```python
from deepagents import create_deep_agent
from deepagents.backends import FilesystemBackend

agent = create_deep_agent(
    # 模型配置
    model="anthropic:claude-sonnet-4-20250514",

    # 自定义工具
    tools=[my_custom_tool],

    # 系统提示（可选，通常从 AGENTS.md 加载）
    system_prompt="You are a helpful assistant...",

    # Memory 文件列表
    memory=["./AGENTS.md", "./context.md"],

    # Skills 目录列表
    skills=["./skills/"],

    # 子代理配置
    subagents=[...],

    # 后端存储
    backend=FilesystemBackend(root_dir="./workspace"),

    # 中间件配置
    middleware=[
        SkillsMiddleware(),
        MemoryMiddleware(),
    ],
)
```

### 参数说明

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `model` | str | 是 | LLM 模型标识符，格式：`provider:model-name` |
| `tools` | list | 否 | 自定义 LangChain 工具列表 |
| `system_prompt` | str | 否 | 覆盖默认的系统提示 |
| `memory` | list | 否 | Memory 文件路径列表 |
| `skills` | list | 否 | Skills 目录路径列表 |
| `subagents` | list | 否 | 子代理配置列表 |
| `backend` | Backend | 否 | 存储后端实例 |
| `middleware` | list | 否 | 中间件实例列表 |

### 支持的模型

```python
# Anthropic (推荐)
model="anthropic:claude-sonnet-4-20250514"
model="anthropic:claude-opus-4-20250514"

# OpenAI
model="openai:gpt-4o"
model="openai:gpt-4-turbo"

# Google
model="google:gemini-1.5-pro"

# 本地模型 (通过 Ollama)
model="ollama:llama3"
```

### Middleware 机制

中间件用于扩展 Agent 的能力：

```python
from deepagents.middleware import (
    SkillsMiddleware,      # 加载和处理 Skills
    MemoryMiddleware,      # 管理 Memory 文件
    ContextMiddleware,     # 上下文管理
    OutputMiddleware,      # 输出处理
)

agent = create_deep_agent(
    model="anthropic:claude-sonnet-4-20250514",
    middleware=[
        SkillsMiddleware(skills_dir="./skills"),
        MemoryMiddleware(memory_files=["./AGENTS.md"]),
    ],
)
```

### Backend 类型

DeepAgents 支持多种存储后端：

```python
from deepagents.backends import (
    FilesystemBackend,    # 本地文件系统
    MemoryBackend,        # 内存存储（测试用）
    S3Backend,           # AWS S3 存储
)

# 本地文件系统
backend = FilesystemBackend(
    root_dir="./workspace",
    allowed_extensions=[".xlsx", ".csv", ".txt"],
)

# 内存存储（用于测试）
backend = MemoryBackend()

# S3 存储
backend = S3Backend(
    bucket="my-agent-data",
    prefix="workspace/",
)
```

---

## 目录结构规范

### 标准项目结构

```
my-agent/
├── agent.py              # 主程序入口
├── tools.py              # 自定义工具
├── AGENTS.md             # Agent 身份定义
├── skills/               # 技能目录
│   ├── skill-one/
│   │   └── SKILL.md
│   └── skill-two/
│       └── SKILL.md
├── subagents/            # 子代理目录（可选）
│   └── helper.yaml
├── tests/                # 测试目录
│   └── test_agent.py
├── docs/                 # 文档目录
├── pyproject.toml        # 项目配置
└── README.md
```

### 目录说明

| 目录/文件 | 用途 |
|-----------|------|
| `agent.py` | CLI 入口，创建和运行 Agent |
| `tools.py` | 自定义 LangChain 工具定义 |
| `AGENTS.md` | Agent 的身份、行为规范、约束 |
| `skills/` | 专门化的工作流指导文件 |
| `subagents/` | 子代理配置 |
| `pyproject.toml` | Python 项目配置和依赖 |

---

## AGENTS.md 编写指南

AGENTS.md 定义了 Agent 的"人格"——它是什么角色，能做什么，应该怎么表现。

### 基本结构

```markdown
# [Agent 名称]

[一句话描述 Agent 的角色]

## Identity

[详细描述 Agent 的专业领域和能力]

## Capabilities

### [能力类别 1]
- [具体能力]
- [具体能力]

### [能力类别 2]
- [具体能力]

## Safety Rules

### Read Operations
- [只读操作规则]

### Write Operations
- [写入操作规则，需要确认的场景]

## Workflow Guidelines

### 1. [步骤名称]
- [详细说明]

### 2. [步骤名称]
- [详细说明]

## Communication Style

- [沟通风格要求]

## Available Tools

| Tool | Purpose |
|------|---------|
| `tool_name` | [用途] |

## Example Interactions

**User:** "[示例用户输入]"

**Agent:** [示例响应]
```

### 示例：数据分析 Agent

```markdown
# Data Analyst Agent

You are an expert data analyst specializing in statistical analysis
and business intelligence.

## Identity

You are a data analysis expert with deep knowledge of:
- Statistical methods and hypothesis testing
- Data visualization best practices
- Business metrics and KPIs
- Data quality assessment

## Capabilities

### Statistical Analysis
- Descriptive statistics (mean, median, variance)
- Correlation and regression analysis
- Time series analysis
- Hypothesis testing

### Data Visualization
- Chart type recommendations
- Color and design guidance
- Interactive dashboard design

## Safety Rules

### Data Privacy
- Never expose sensitive personal data in outputs
- Aggregate data to protect individual records
- Follow data retention policies

### Write Operations
- Require confirmation before:
  - Deleting any data
  - Modifying original source files
  - Sharing data externally

## Workflow Guidelines

### 1. Understand the Question
- Clarify the business objective
- Identify relevant data sources
- Define success criteria

### 2. Explore the Data
- Check data quality and completeness
- Identify patterns and anomalies
- Document assumptions

### 3. Analyze and Interpret
- Apply appropriate statistical methods
- Validate results
- Draw actionable conclusions

### 4. Communicate Results
- Use clear, non-technical language
- Visualize key findings
- Provide recommendations

## Communication Style

- Lead with insights, not methodology
- Use visuals when possible
- Quantify uncertainty in conclusions
- Acknowledge data limitations
```

---

## Skills 编写指南

Skills 是专门化的工作流指导，帮助 Agent 更好地完成特定类型的任务。

### 目录结构

```
skills/
├── data-cleaning/
│   └── SKILL.md
├── visualization/
│   └── SKILL.md
└── forecasting/
    └── SKILL.md
```

### YAML Frontmatter 规范

每个 SKILL.md 文件应以 YAML frontmatter 开头：

```markdown
---
name: skill-name              # 技能标识符（必需）
description: 简短描述          # 技能描述（必需）
triggers:                     # 触发关键词（可选）
  - keyword1
  - keyword2
  - 中文关键词
priority: 0                   # 优先级，数字越大优先级越高（可选）
---
```

### Progressive Disclosure 模式

Skills 使用"渐进式披露"模式——首先提供概述，需要时展开细节：

```markdown
---
name: data-validation
description: Validate data quality and identify issues
triggers:
  - validate
  - quality
  - check
  - 验证
---

# Data Validation Skill

Quick guide for validating data quality.

## Quick Reference

| Check | Tool | Common Issues |
|-------|------|---------------|
| Missing values | `analyze_data(type="missing")` | Null, empty, NaN |
| Duplicates | `analyze_data(type="unique")` | Repeated rows |
| Outliers | `analyze_data(type="summary")` | Extreme values |
| Types | `analyze_data(type="summary")` | Wrong data types |

## Detailed Workflow

### Step 1: Load and Preview

[详细说明...]

### Step 2: Check Data Types

[详细说明...]

### Step 3: Identify Missing Values

[详细说明...]

## Common Issues and Fixes

### Issue: Date Parsing Errors

**Symptoms:**
- Dates appear as strings
- Inconsistent formats

**Solution:**
```python
# Use pandas to parse dates
df['date'] = pd.to_datetime(df['date'], errors='coerce')
```

[更多问题...]
```

### 完整 Skill 示例

```markdown
---
name: sql-generation
description: Generate SQL queries from natural language
triggers:
  - sql
  - query
  - database
  - 数据库
  - 查询
---

# SQL Generation Skill

Guide for generating correct and efficient SQL queries.

## Quick Reference

| Need | SQL Clause |
|------|------------|
| Filter rows | WHERE |
| Group data | GROUP BY |
| Sort results | ORDER BY |
| Limit rows | LIMIT / TOP |
| Join tables | JOIN |

## Query Patterns

### Basic SELECT

```sql
SELECT column1, column2
FROM table_name
WHERE condition
ORDER BY column1
LIMIT 10;
```

### Aggregation

```sql
SELECT category,
       COUNT(*) as count,
       SUM(amount) as total
FROM sales
GROUP BY category
HAVING SUM(amount) > 1000
ORDER BY total DESC;
```

### JOIN

```sql
SELECT a.name, b.order_date
FROM customers a
JOIN orders b ON a.id = b.customer_id
WHERE b.order_date >= '2024-01-01';
```

## Best Practices

### 1. Always Use Aliases
- Makes output readable
- Required for calculated fields

### 2. Qualify Column Names
- Use table aliases: `t.column_name`
- Prevents ambiguity in joins

### 3. Filter Early
- Put filters in WHERE, not HAVING when possible
- Improves query performance

### 4. Use Parameterized Queries
- Never concatenate user input
- Prevents SQL injection

## Common Mistakes

❌ **Wrong:**
```sql
SELECT * FROM table
```
✅ **Better:**
```sql
SELECT id, name, created_at FROM table
```

❌ **Wrong:**
```sql
WHERE date LIKE '2024-01%'
```
✅ **Better:**
```sql
WHERE date >= '2024-01-01' AND date < '2024-02-01'
```
```

---

## Subagents 配置指南

Subagents 允许将复杂任务委托给专门的子代理。

### YAML 配置方式

```yaml
# subagents/data-processor.yaml
name: data-processor
description: Handles large data processing tasks
model: anthropic:claude-sonnet-4-20250514
tools:
  - read_file
  - write_file
  - execute
system_prompt: |
  You are a data processing specialist.
  Focus on efficiency and accuracy.
memory:
  - ./subagents/data-processor/CONTEXT.md
```

### 代码配置方式

```python
from deepagents import create_deep_agent, SubagentConfig

# 定义子代理配置
data_processor = SubagentConfig(
    name="data-processor",
    description="Processes and transforms data files",
    model="anthropic:claude-sonnet-4-20250514",
    tools=[read_excel, write_excel, analyze_data],
)

# 创建主代理时包含子代理
agent = create_deep_agent(
    model="anthropic:claude-opus-4-20250514",
    subagents=[data_processor],
    # ... other config
)
```

### 使用子代理

主代理会自动使用 `task` 工具委托工作给子代理：

```python
# 用户请求
"Process this large CSV file and generate a summary report"

# 主代理自动委托
# 1. 识别需要数据处理
# 2. 调用 task 工具委托给 data-processor
# 3. 等待结果并整合到最终响应
```

---

## 自定义工具开发

### 基本工具定义

使用 LangChain 的 `@tool` 装饰器：

```python
from langchain_core.tools import tool
from typing import Optional

@tool
def my_custom_tool(
    param1: str,
    param2: Optional[int] = None,
) -> str:
    """
    Brief description of what the tool does.

    Args:
        param1: Description of param1
        param2: Description of param2 (optional)

    Returns:
        Description of return value
    """
    # Implementation
    result = do_something(param1, param2)
    return f"Result: {result}"
```

### 工具最佳实践

#### 1. 返回 JSON 字符串

```python
import json

@tool
def analyze_data(file_path: str) -> str:
    """Analyze data and return structured results."""
    result = {
        "status": "success",
        "data": {...},
        "metadata": {...}
    }
    return json.dumps(result, indent=2, ensure_ascii=False)
```

#### 2. 处理错误优雅

```python
@tool
def safe_operation(file_path: str) -> str:
    """Perform operation with error handling."""
    try:
        # Operation
        result = perform_operation(file_path)
        return json.dumps({"success": True, "result": result})
    except FileNotFoundError:
        return json.dumps({"error": f"File not found: {file_path}"})
    except Exception as e:
        return json.dumps({"error": f"Operation failed: {str(e)}"})
```

#### 3. 提供详细文档

```python
@tool
def complex_tool(
    data: str,
    options: Optional[str] = None,
) -> str:
    """
    Perform a complex data transformation.

    This tool transforms input data according to the specified options.
    It handles various data formats including JSON and CSV.

    Args:
        data: Input data as a JSON string. Expected format:
            {"records": [{"id": 1, "value": "a"}, ...]}
        options: Optional configuration as JSON string:
            - format: Output format ("json", "csv", "table")
            - sort: Sort key for output
            - filter: Filter condition

    Returns:
        JSON string containing:
            - success: boolean
            - data: transformed data
            - count: number of records processed
            - errors: list of any errors encountered

    Example:
        >>> complex_tool(
        ...     '{"records": [{"id": 1, "value": "a"}]}',
        ...     '{"format": "table", "sort": "id"}'
        ... )
    """
    # Implementation
```

### 工具注册

```python
# tools.py
from langchain_core.tools import tool

@tool
def tool_one(): pass

@tool
def tool_two(): pass

# 导出工具列表
MY_TOOLS = [tool_one, tool_two]

# agent.py
from tools import MY_TOOLS
from deepagents import create_deep_agent

agent = create_deep_agent(
    model="anthropic:claude-sonnet-4-20250514",
    tools=MY_TOOLS,
    # ...
)
```

---

## 完整示例代码

### 最小可行 Agent

```python
# agent.py
from deepagents import create_deep_agent
from deepagents.backends import FilesystemBackend

def main():
    agent = create_deep_agent(
        model="anthropic:claude-sonnet-4-20250514",
        memory=["./AGENTS.md"],
        backend=FilesystemBackend(root_dir="./"),
    )

    result = agent.invoke({"input": "Hello!"})
    print(result["output"])

if __name__ == "__main__":
    main()
```

### 完整 CLI Agent

```python
#!/usr/bin/env python3
"""
My Custom Agent - A DeepAgents-based assistant.

Usage:
    python agent.py                    # Interactive mode
    python agent.py "Your query"       # Single query
    python agent.py --help             # Show help
"""

import argparse
import sys
from pathlib import Path

from deepagents import create_deep_agent
from deepagents.backends import FilesystemBackend

from my_tools import MY_TOOLS


def create_my_agent(working_dir: str = None):
    """Create and configure the agent."""
    if working_dir is None:
        working_dir = Path.cwd()

    return create_deep_agent(
        model="anthropic:claude-sonnet-4-20250514",
        tools=MY_TOOLS,
        memory=["./AGENTS.md"],
        skills=["./skills/"],
        backend=FilesystemBackend(root_dir=str(working_dir)),
    )


def interactive_mode(agent):
    """Run interactive REPL."""
    print("My Agent - Interactive Mode")
    print("Type 'quit' to exit.\n")

    while True:
        try:
            query = input("You: ").strip()
            if not query:
                continue
            if query.lower() in ("quit", "exit", "q"):
                break

            result = agent.invoke({"input": query})
            print(f"\nAgent: {result['output']}\n")

        except KeyboardInterrupt:
            print("\nGoodbye!")
            break


def main():
    parser = argparse.ArgumentParser(
        description="My Custom Agent",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("query", nargs="?", help="Query to execute")
    parser.add_argument("--dir", "-d", help="Working directory")
    parser.add_argument("--model", "-m", default="anthropic:claude-sonnet-4-20250514")

    args = parser.parse_args()

    agent = create_my_agent(args.dir)

    if args.query:
        result = agent.invoke({"input": args.query})
        print(result["output"])
    else:
        interactive_mode(agent)


if __name__ == "__main__":
    main()
```

---

## 最佳实践

### 1. Agent 设计原则

- **单一职责**: 每个 Agent 专注于一个领域
- **清晰的边界**: 明确定义 Agent 能做什么和不能做什么
- **优雅降级**: 当工具不可用时提供有用的反馈

### 2. 安全考虑

```python
# 总是验证文件路径
def safe_read(file_path: str) -> str:
    path = Path(file_path).resolve()
    if not str(path).startswith(ALLOWED_DIR):
        return json.dumps({"error": "Access denied"})
    # ... read file

# 使用沙箱执行命令
agent = create_deep_agent(
    model="...",
    sandbox=True,  # Enable command sandboxing
)
```

### 3. 性能优化

```python
# 使用 chunked 处理大数据
@tool
def process_large_file(file_path: str, chunk_size: int = 1000) -> str:
    """Process large files in chunks."""
    results = []
    for chunk in pd.read_csv(file_path, chunksize=chunk_size):
        results.append(process_chunk(chunk))
    return json.dumps({"chunks_processed": len(results)})
```

### 4. 测试策略

```python
# tests/test_agent.py
import pytest
from deepagents.backends import MemoryBackend
from my_agent import create_my_agent

@pytest.fixture
def agent():
    return create_my_agent(
        backend=MemoryBackend()
    )

def test_basic_query(agent):
    result = agent.invoke({"input": "Hello"})
    assert "output" in result
    assert result["output"]  # Non-empty response

def test_tool_usage(agent):
    result = agent.invoke({"input": "Read file test.csv"})
    assert "error" not in result["output"].lower()
```

### 5. 监控和日志

```python
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('agent.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# 在工具中记录
@tool
def my_tool(param: str) -> str:
    logger.info(f"Tool called with param: {param}")
    try:
        result = do_work(param)
        logger.info(f"Tool completed successfully")
        return result
    except Exception as e:
        logger.error(f"Tool failed: {e}")
        raise
```

---

## 参考资源

- [DeepAgents GitHub 仓库](https://github.com/langchain-ai/deepagents)
- [LangChain 文档](https://python.langchain.com/)
- [LangGraph 文档](https://langchain-ai.github.io/langgraph/)
- [本项目 Excel Agent 示例](../)
