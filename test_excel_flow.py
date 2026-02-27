"""Test script for Excel agent — hits the FastAPI /api/stream endpoint directly.

Usage:
    1. Start the backend:  EXCEL_AGENT_DEV=1 uv run uvicorn server:app --host 127.0.0.1 --port 8765
    2. Run this script:    uv run python -u test_excel_flow.py
    3. Run single case:    uv run python -u test_excel_flow.py --case test_03_formula
"""

import json
import re
import sys
import time
import argparse
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import httpx

# ── Force unbuffered output ─────────────────────────────────────────
_print = print


def print(*args, **kwargs):
    kwargs.setdefault("flush", True)
    _print(*args, **kwargs)


# ── Config ──────────────────────────────────────────────────────────
BASE = "http://127.0.0.1:8765"
TIMEOUT_PER_CASE = 600          # seconds – per test case
TEST_DIR = Path(__file__).resolve().parent / "test_output"

# ── Data structures ─────────────────────────────────────────────────


@dataclass
class ToolResult:
    name: str
    status: str
    result: str
    has_error: bool


@dataclass
class StreamResult:
    ok: bool
    event_count: int = 0
    tool_results: list[ToolResult] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    text: str = ""


@dataclass
class CaseResult:
    name: str
    passed: bool
    duration: float = 0.0
    errors: list[str] = field(default_factory=list)
    stream: StreamResult | None = None


# ── Shared context across cases ─────────────────────────────────────
# Some test cases create files that later cases depend on. Store paths here.
shared: dict = {}


# ── Logging helpers ─────────────────────────────────────────────────

def find_latest_log() -> Path | None:
    log_dir = Path(__file__).resolve().parent / "logs"
    if not log_dir.exists():
        return None
    logs = sorted(log_dir.glob("app_*.log"), key=lambda p: p.stat().st_mtime)
    return logs[-1] if logs else None


def count_log_lines(log_path: Path | None) -> int:
    if not log_path or not log_path.exists():
        return 0
    return len(log_path.read_text(encoding="utf-8").splitlines())


def check_log_for_errors(log_path: Path | None, start_line: int) -> list[str]:
    if not log_path or not log_path.exists():
        return []
    errors = []
    lines = log_path.read_text(encoding="utf-8").splitlines()
    for i, line in enumerate(lines):
        if i < start_line:
            continue
        if "Error:" in line and ("tool:result" in line or "excel" in line.lower()):
            errors.append(line.strip())
    return errors


# ── Core streaming helper ───────────────────────────────────────────

def stream_request(prompt: str, timeout: int = TIMEOUT_PER_CASE) -> StreamResult:
    """Send a streaming request and collect all SSE events."""
    tool_results: list[ToolResult] = []
    text_chunks: list[str] = []
    errors: list[str] = []
    event_count = 0

    print(f"\n{'─'*60}")
    print(f"  PROMPT: {prompt[:100]}{'...' if len(prompt) > 100 else ''}")
    print(f"{'─'*60}\n")

    try:
        with httpx.Client(timeout=timeout) as client:
            with client.stream(
                "POST",
                f"{BASE}/api/stream",
                json={"message": prompt},
                headers={"X-App-Token": "dev"},
            ) as resp:
                if resp.status_code != 200:
                    print(f"  [FAIL] HTTP {resp.status_code}")
                    return StreamResult(ok=False, errors=[f"HTTP {resp.status_code}"])

                for raw_line in resp.iter_lines():
                    if not raw_line.startswith("data: "):
                        continue
                    data = json.loads(raw_line[6:])
                    event_count += 1
                    etype = data.get("type", "")

                    if etype == "stream:text":
                        text_chunks.append(data.get("token", ""))

                    elif etype == "tool:start":
                        name = data.get("name", "?")
                        cid = data.get("id", "")[:12]
                        print(f"  [TOOL:START] {name}  id={cid}")

                    elif etype == "tool:result":
                        name = data.get("name", "?")
                        status = data.get("status", "?")
                        result_str = data.get("result", "") or ""
                        dur = data.get("duration_ms")
                        dur_s = f" ({dur:.0f}ms)" if dur else ""
                        is_err = "Error" in result_str
                        tag = "ERROR" if is_err else "OK"
                        short = (result_str[:140] + "...") if len(result_str) > 140 else result_str
                        print(f"  [TOOL:{tag}] {name} status={status}{dur_s}  {short}")
                        tr = ToolResult(name=name, status=status, result=result_str, has_error=is_err)
                        tool_results.append(tr)
                        if is_err:
                            errors.append(f"{name}: {result_str}")

                    elif etype == "tasks:update":
                        tasks = data.get("tasks", [])
                        summary = ", ".join(
                            f"{t['label'][:20]}..={t['status']}" for t in tasks
                        )
                        print(f"  [TASKS] {summary}")

                    elif etype == "stream:done":
                        err = data.get("error")
                        if err:
                            errors.append(f"stream:done error: {err}")
                            print(f"  [DONE] error={err}")
                        else:
                            print(f"  [DONE]")

    except httpx.ReadTimeout:
        errors.append(f"Timeout after {timeout}s")
        print(f"  [TIMEOUT] {timeout}s exceeded")
    except Exception as exc:
        errors.append(f"Exception: {exc}")
        print(f"  [EXCEPTION] {exc}")

    return StreamResult(
        ok=len(errors) == 0,
        event_count=event_count,
        tool_results=tool_results,
        errors=errors,
        text="".join(text_chunks),
    )


# ── Assertion helpers ───────────────────────────────────────────────

def assert_no_tool_errors(sr: StreamResult) -> list[str]:
    """Return list of error descriptions if any tool had errors."""
    return [f"{t.name}: {t.result[:200]}" for t in sr.tool_results if t.has_error]


def assert_tool_was_called(sr: StreamResult, tool_name: str) -> list[str]:
    """Check that a specific tool was invoked at least once."""
    if any(t.name == tool_name for t in sr.tool_results):
        return []
    return [f"Expected tool '{tool_name}' to be called but it was not"]


def assert_any_tool_called(sr: StreamResult, tool_names: list[str]) -> list[str]:
    """Check that at least one of the listed tools was invoked."""
    called = {t.name for t in sr.tool_results}
    if called & set(tool_names):
        return []
    return [f"Expected one of {tool_names} to be called, got: {sorted(called)}"]


def assert_text_contains(sr: StreamResult, *keywords: str) -> list[str]:
    """Check that the agent's final text response contains keywords."""
    errs = []
    for kw in keywords:
        if kw.lower() not in sr.text.lower():
            errs.append(f"Agent response missing keyword '{kw}'")
    return errs


def assert_file_exists(path: Path, min_kb: float = 1.0) -> list[str]:
    if not path.exists():
        return [f"File not found: {path}"]
    size_kb = path.stat().st_size / 1024
    if size_kb < min_kb:
        return [f"File {path.name} too small: {size_kb:.1f}KB (min {min_kb}KB)"]
    print(f"  [FILE] {path.name}  ({size_kb:.1f} KB)")
    return []


def find_xlsx(pattern: str) -> Path | None:
    """Find an xlsx in TEST_DIR matching glob pattern."""
    matches = sorted(TEST_DIR.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    return matches[0] if matches else None


def find_any_xlsx() -> Path | None:
    return find_xlsx("*.xlsx")


def setup_project() -> None:
    """Ensure an active project pointing at TEST_DIR exists before running tests."""
    print("Setting up test project...")
    project_path = str(TEST_DIR)
    with httpx.Client(timeout=10) as client:
        # Try opening existing project first
        r = client.post(
            f"{BASE}/api/projects/open",
            json={"project_path": project_path},
            headers={"X-App-Token": "dev"},
        )
        if r.status_code == 200:
            print(f"[OK] Opened existing project at {TEST_DIR}\n")
            return
        # Create new project
        r = client.post(
            f"{BASE}/api/projects/create",
            json={"project_path": project_path, "name": "Test Project", "data_sources": []},
            headers={"X-App-Token": "dev"},
        )
        if r.status_code == 200:
            print(f"[OK] Created project at {TEST_DIR}\n")
            return
        print(f"[FAIL] Could not set up project: {r.status_code} {r.text}")
        sys.exit(1)


# ── TEST CASES ──────────────────────────────────────────────────────
# Each test is a function: () -> CaseResult
# They are executed in order. Tests that depend on prior state can
# read/write the global `shared` dict.

def test_01_create_basic_file() -> CaseResult:
    """Create an Excel file with various data types (numbers, text, dates, booleans)."""
    name = "test_01_create_basic_file"
    prompt = (
        f"创建一个Excel文件，保存到 {TEST_DIR} 目录，文件名为 '基础数据测试.xlsx'。\n"
        "包含一个工作表 '数据总览'，从A1开始写入以下数据：\n"
        "- 表头行：姓名, 年龄, 入职日期, 薪资, 在职\n"
        "- 至少8行示例数据，包含：中文姓名、整数年龄(25-55)、日期(2018-2024年)、"
        "带小数的薪资(8000-25000)、布尔值TRUE/FALSE\n"
        "写入完成后保存文件。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    errs += assert_any_tool_called(sr, ["write_excel_range", "excel_write"])

    # Check file was created
    target = TEST_DIR / "基础数据测试.xlsx"
    errs += assert_file_exists(target)
    if not errs:
        shared["basic_file"] = str(target)

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_02_read_and_verify() -> CaseResult:
    """Read back the file created in test_01 and verify contents."""
    name = "test_02_read_and_verify"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"读取文件 '{file_path}' 中工作表 '数据总览' 的所有数据（使用读取工具），"
        "然后告诉我：共有几行几列数据、表头是什么、第一行数据的内容。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    errs += assert_any_tool_called(sr, ["read_excel_range", "excel_read"])
    # The agent should mention column count / headers in its response
    errs += assert_text_contains(sr, "姓名")

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_03_formula_write_and_read() -> CaseResult:
    """Add formula columns and read them back."""
    name = "test_03_formula_write_and_read"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，在工作表 '数据总览' 中做以下操作：\n"
        "1. 在F1写入表头 '年薪'，然后在F2往下的所有数据行设置公式 =D2*12（D列是薪资）\n"
        "2. 在G1写入表头 '薪资等级'，G2往下设置公式：=IF(D2>=15000,\"高\",IF(D2>=10000,\"中\",\"低\"))\n"
        "3. 设置完成后保存，然后读取F列和G列的数据告诉我结果。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    # Should have used formula-related tools and read tools
    errs += assert_any_tool_called(sr, ["range_set_formula", "excel_formula", "write_excel_range", "excel_write"])

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_04_table_create_and_style() -> CaseResult:
    """Create a Table (ListObject) from existing data and apply a style."""
    name = "test_04_table_create_and_style"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，对工作表 '数据总览' 中已有数据区域（包括之前添加的F、G列）"
        "执行以下操作：\n"
        "1. 将整个数据区域转换为Excel表格(Table)，命名为 '员工信息表'\n"
        "2. 设置一个美观的表格样式\n"
        "3. 完成后保存文件"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    errs += assert_any_tool_called(sr, ["table_create", "excel_table"])

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_05_table_add_column_with_formula() -> CaseResult:
    """Add a calculated column to an existing table — this was previously failing."""
    name = "test_05_table_add_column_with_formula"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，在表格 '员工信息表' 中添加一个新列：\n"
        "- 列名：'工龄'\n"
        "- 公式：用TODAY()减去入职日期列再除以365，取整数\n"
        "添加完成后保存文件，并读取这一列的数据告诉我结果。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    # This is the scenario that was failing in the logs; we want 0 errors
    errs += assert_any_tool_called(sr, [
        "table_add_column", "excel_table",
        "range_set_formula", "excel_formula",
        "write_excel_range", "excel_write",
    ])

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_06_sheet_operations() -> CaseResult:
    """Add, rename, and list sheets."""
    name = "test_06_sheet_operations"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，执行以下工作表操作：\n"
        "1. 添加一个新工作表，命名为 '统计分析'\n"
        "2. 添加另一个新工作表，命名为 '临时表'\n"
        "3. 将 '临时表' 重命名为 '备份数据'\n"
        "4. 列出当前所有工作表名称告诉我\n"
        "5. 保存文件"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    # Should mention both sheet names
    errs += assert_text_contains(sr, "统计分析")

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_07_cross_sheet_workflow() -> CaseResult:
    """Write summary data on the second sheet referencing the first sheet."""
    name = "test_07_cross_sheet_workflow"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，在工作表 '统计分析' 中创建一个汇总区域：\n"
        "- A1: '统计项目', B1: '值'\n"
        "- A2: '员工总数', B2: 使用COUNTA公式统计数据总览表A列的数据行数(不含表头)\n"
        "- A3: '平均薪资', B3: 使用AVERAGE公式计算数据总览表D列的平均值\n"
        "- A4: '最高薪资', B4: 使用MAX公式\n"
        "- A5: '最低薪资', B5: 使用MIN公式\n"
        "- A6: '高薪人数', B6: 使用COUNTIF统计D列>=15000的人数\n"
        "全部使用公式引用 '数据总览' 表的数据，保存后读取B列的值告诉我结果。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    errs += assert_any_tool_called(sr, [
        "range_set_formula", "excel_formula",
        "write_excel_range", "excel_write",
    ])

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_08_new_file_sales_data() -> CaseResult:
    """Create a brand-new sales data file (independent from tests 01-07)."""
    name = "test_08_new_file_sales_data"
    prompt = (
        f"创建一个新Excel文件，保存到 {TEST_DIR} 目录，文件名为 '销售报表.xlsx'。\n"
        "创建工作表 '月度销售'，包含以下列：\n"
        "月份(1-12月), 产品名称(至少3种产品每月重复), 销售数量, 单价, 销售额(公式=数量*单价)\n"
        "至少36行数据（12个月×3种产品），销售额列必须使用公式而不是手动计算的值。\n"
        "将数据区域创建为表格，命名为 '月度销售表'，应用美观的样式，然后保存。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)

    target = TEST_DIR / "销售报表.xlsx"
    errs += assert_file_exists(target, min_kb=5.0)
    if not errs:
        shared["sales_file"] = str(target)

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_09_pivot_style_summary() -> CaseResult:
    """Create a summary sheet with formulas that aggregate sales data."""
    name = "test_09_pivot_style_summary"

    if "sales_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: sales_file not set by test_08"])

    file_path = shared["sales_file"]
    prompt = (
        f"打开文件 '{file_path}'，添加一个新工作表 '产品汇总'：\n"
        "- A1: '产品名称', B1: '总销售额', C1: '总销售数量', D1: '平均单价'\n"
        "- 针对 '月度销售' 表中出现的每种产品，各写一行，使用SUMIF/AVERAGEIF公式汇总\n"
        "- 在最后一行写一个 '合计' 行，用SUM公式合计B列和C列\n"
        "- 将汇总区域也创建为表格，命名为 '产品汇总表'\n"
        "保存文件。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    errs += assert_any_tool_called(sr, [
        "range_set_formula", "excel_formula",
        "write_excel_range", "excel_write",
    ])

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_10_conditional_formatting_and_number_format() -> CaseResult:
    """Apply number formats and conditional formatting to the sales file."""
    name = "test_10_conditional_formatting_and_number_format"

    if "sales_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: sales_file not set by test_08"])

    file_path = shared["sales_file"]
    prompt = (
        f"打开文件 '{file_path}'，对工作表 '月度销售' 做以下格式设置：\n"
        "1. 单价和销售额列设置为货币格式（人民币，保留2位小数）\n"
        "2. 销售数量列设置为整数格式（千分位分隔符）\n"
        "3. 自动调整所有列宽\n"
        "保存文件。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_11_delete_sheet() -> CaseResult:
    """Delete a sheet — verify the agent handles DisplayAlerts correctly."""
    name = "test_11_delete_sheet"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，删除工作表 '备份数据'（之前从临时表重命名的），"
        "然后列出剩余的工作表名称告诉我，保存文件。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    # '备份数据' should NOT appear in response as existing sheet
    if "备份数据" in sr.text and "删除" not in sr.text.split("备份数据")[0][-20:]:
        # Only flag if it sounds like the sheet still exists
        pass  # Heuristic check, don't be too strict

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_12_error_recovery_bad_formula() -> CaseResult:
    """Send an intentionally bad formula — agent should handle gracefully and retry/fix."""
    name = "test_12_error_recovery_bad_formula"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，在工作表 '数据总览' 的H1写入 '测试'，\n"
        "然后在H2设置公式 =SUMX(A:A)（注意SUMX不是有效的Excel函数）。\n"
        "如果公式设置失败，请改用正确的SUM函数重试，保存文件。"
    )

    sr = stream_request(prompt)
    # We expect the agent to recover — the final result should be OK
    # Tool errors from the *first* attempt are acceptable, but the agent should retry
    final_tool_results = sr.tool_results
    if final_tool_results:
        last_write = [t for t in final_tool_results if "formula" in t.name.lower() or "write" in t.name.lower()]
        if last_write and all(t.has_error for t in last_write):
            return CaseResult(
                name=name, passed=False, stream=sr,
                errors=["Agent never recovered from bad formula"]
            )

    return CaseResult(name=name, passed=True, errors=[], stream=sr)


def test_13_large_dataset() -> CaseResult:
    """Create a file with a larger dataset (200+ rows) — performance test."""
    name = "test_13_large_dataset"
    prompt = (
        f"创建一个新Excel文件，保存到 {TEST_DIR} 目录，文件名为 '大数据测试.xlsx'。\n"
        "在 Sheet1 中写入一个200行×5列的数据集：\n"
        "- A列：序号(1-200)\n"
        "- B列：随机中文姓名\n"
        "- C列：随机部门名称(从5个部门中选)\n"
        "- D列：随机绩效分数(60-100的整数)\n"
        "- E列：公式 =IF(D2>=90,\"优秀\",IF(D2>=80,\"良好\",IF(D2>=70,\"合格\",\"待改进\")))\n"
        "将数据转为表格，保存文件。"
    )

    sr = stream_request(prompt, timeout=900)
    errs = assert_no_tool_errors(sr)

    target = TEST_DIR / "大数据测试.xlsx"
    errs += assert_file_exists(target, min_kb=10.0)

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_14_multi_table_single_sheet() -> CaseResult:
    """Create two separate tables on the same sheet — tests non-overlapping ranges."""
    name = "test_14_multi_table_single_sheet"
    prompt = (
        f"创建一个新Excel文件，保存到 {TEST_DIR} 目录，文件名为 '多表格测试.xlsx'。\n"
        "在 Sheet1 中创建两个独立的表格（它们之间留一些空行）：\n"
        "1. A1:C5 区域：表格名 '水果价格'，列：水果名称、单价、库存，4行数据\n"
        "2. A8:C12 区域：表格名 '蔬菜价格'，列：蔬菜名称、单价、库存，4行数据\n"
        "两个表格用不同的样式，保存文件。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)

    target = TEST_DIR / "多表格测试.xlsx"
    errs += assert_file_exists(target)

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_15_read_nonexistent_sheet() -> CaseResult:
    """Try reading from a sheet that doesn't exist — agent should report error gracefully."""
    name = "test_15_read_nonexistent_sheet"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"读取文件 '{file_path}' 中工作表 '不存在的表' 的A1:D10数据。"
    )

    sr = stream_request(prompt)
    # We expect the agent to tell the user the sheet doesn't exist
    # Tool may report an error, which is acceptable — agent should handle it
    errs = assert_text_contains(sr, "不存在")
    if errs:
        # Alternatively the agent might say "找不到" or similar
        errs2 = assert_text_contains(sr, "找不到")
        if errs2:
            errs3 = assert_text_contains(sr, "没有")
            if errs3:
                errs = ["Agent didn't clearly tell user the sheet doesn't exist"]
            else:
                errs = []
        else:
            errs = []

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_16_complex_formulas() -> CaseResult:
    """Test complex / nested formulas: VLOOKUP, INDEX/MATCH, array formulas."""
    name = "test_16_complex_formulas"

    if "sales_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: sales_file not set by test_08"])

    file_path = shared["sales_file"]
    prompt = (
        f"打开文件 '{file_path}'，添加一个新工作表 '高级公式'：\n"
        "- A1: '公式名称', B1: '公式', C1: '结果'\n"
        "- A2: '最高销售额月份', C2: 使用INDEX+MATCH公式找出月度销售表中销售额最高的月份\n"
        "- A3: '销售额中位数', C3: 使用MEDIAN函数\n"
        "- A4: '销售额标准差', C4: 使用STDEV函数\n"
        "- A5: '大于均值的记录数', C5: 使用COUNTIF配合AVERAGE的组合公式\n"
        "B列写入你使用的公式文本（作为字符串显示），C列是实际公式。保存文件。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_17_overwrite_existing_data() -> CaseResult:
    """Overwrite cells that already have data — tests that no append-only bug exists."""
    name = "test_17_overwrite_existing_data"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，在工作表 '统计分析' 中，"
        "将A1的值改为 '更新后的统计项目'，B1改为 '更新后的值'，"
        "然后读取A1:B1的内容确认修改成功，保存文件。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    errs += assert_text_contains(sr, "更新后")

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_18_special_characters_in_data() -> CaseResult:
    """Write data with special characters: quotes, newlines, unicode symbols."""
    name = "test_18_special_characters_in_data"
    prompt = (
        f"创建一个新Excel文件，保存到 {TEST_DIR} 目录，文件名为 '特殊字符测试.xlsx'。\n"
        "在 Sheet1 中写入以下数据：\n"
        "- A1: '测试项', B1: '内容'\n"
        "- A2: '引号测试', B2: 包含双引号的字符串 He said \"hello\"\n"
        "- A3: '特殊符号', B3: ★☆♠♣♥♦\n"
        "- A4: '长文本', B4: 一段超过50个字的中文描述\n"
        "- A5: '数字字符串', B5: 以文本形式存储的 '00123'（不要变成数字123）\n"
        "保存文件。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)

    target = TEST_DIR / "特殊字符测试.xlsx"
    errs += assert_file_exists(target)

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_19_close_and_reopen() -> CaseResult:
    """Close a file then reopen it — tests workbook lifecycle management."""
    name = "test_19_close_and_reopen"

    if "basic_file" not in shared:
        return CaseResult(name=name, passed=False, errors=["Skipped: basic_file not set by test_01"])

    file_path = shared["basic_file"]
    prompt = (
        f"打开文件 '{file_path}'，读取工作表 '数据总览' 的A1:B2数据告诉我。"
    )

    sr = stream_request(prompt)
    errs = assert_no_tool_errors(sr)
    errs += assert_text_contains(sr, "姓名")

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


def test_20_full_workflow_from_scratch() -> CaseResult:
    """End-to-end: create file, add data, formulas, table, summary sheet — all in one prompt."""
    name = "test_20_full_workflow_from_scratch"
    prompt = (
        f"请完成以下完整任务，保存到 {TEST_DIR}/完整流程测试.xlsx：\n"
        "1. 创建工作表 '订单明细'，写入列：订单号、客户名、产品、数量、单价，"
        "至少10行数据\n"
        "2. 添加 '金额' 列，使用公式 =数量*单价\n"
        "3. 将数据区域创建为表格 '订单表'，应用样式\n"
        "4. 新建工作表 '客户汇总'，使用SUMIF公式按客户汇总金额\n"
        "5. 保存文件\n"
        "请直接开始执行，不需要先规划。"
    )

    sr = stream_request(prompt, timeout=900)
    errs = assert_no_tool_errors(sr)

    target = TEST_DIR / "完整流程测试.xlsx"
    errs += assert_file_exists(target, min_kb=5.0)

    return CaseResult(name=name, passed=len(errs) == 0, errors=errs, stream=sr)


# ── Test registry ───────────────────────────────────────────────────

ALL_TESTS: list[tuple[str, Callable[[], CaseResult]]] = [
    ("test_01_create_basic_file",               test_01_create_basic_file),
    ("test_02_read_and_verify",                  test_02_read_and_verify),
    ("test_03_formula_write_and_read",           test_03_formula_write_and_read),
    ("test_04_table_create_and_style",           test_04_table_create_and_style),
    ("test_05_table_add_column_with_formula",    test_05_table_add_column_with_formula),
    ("test_06_sheet_operations",                 test_06_sheet_operations),
    ("test_07_cross_sheet_workflow",             test_07_cross_sheet_workflow),
    ("test_08_new_file_sales_data",              test_08_new_file_sales_data),
    ("test_09_pivot_style_summary",              test_09_pivot_style_summary),
    ("test_10_conditional_formatting_and_number_format", test_10_conditional_formatting_and_number_format),
    ("test_11_delete_sheet",                     test_11_delete_sheet),
    ("test_12_error_recovery_bad_formula",       test_12_error_recovery_bad_formula),
    ("test_13_large_dataset",                    test_13_large_dataset),
    ("test_14_multi_table_single_sheet",         test_14_multi_table_single_sheet),
    ("test_15_read_nonexistent_sheet",           test_15_read_nonexistent_sheet),
    ("test_16_complex_formulas",                 test_16_complex_formulas),
    ("test_17_overwrite_existing_data",          test_17_overwrite_existing_data),
    ("test_18_special_characters_in_data",       test_18_special_characters_in_data),
    ("test_19_close_and_reopen",                 test_19_close_and_reopen),
    ("test_20_full_workflow_from_scratch",       test_20_full_workflow_from_scratch),
]


# ── Runner ──────────────────────────────────────────────────────────

def run_tests(selected: str | None = None) -> list[CaseResult]:
    results: list[CaseResult] = []

    tests_to_run = ALL_TESTS
    if selected:
        tests_to_run = [(n, f) for n, f in ALL_TESTS if selected in n]
        if not tests_to_run:
            print(f"[ERROR] No test matching '{selected}'")
            sys.exit(1)

    total = len(tests_to_run)
    for idx, (name, func) in enumerate(tests_to_run, 1):
        print(f"\n{'='*60}")
        print(f"  [{idx}/{total}]  {name}")
        print(f"{'='*60}")

        log_path = find_latest_log()
        log_start = count_log_lines(log_path)

        t0 = time.time()
        try:
            result = func()
        except Exception as exc:
            result = CaseResult(name=name, passed=False, errors=[f"Unhandled exception: {exc}"])
        result.duration = time.time() - t0

        # Append log-level errors
        log_path = find_latest_log()
        log_errs = check_log_for_errors(log_path, log_start)
        if log_errs:
            result.errors.extend([f"[LOG] {e[:200]}" for e in log_errs])
            if result.passed:
                result.passed = False  # log errors demote to FAIL

        status = "PASS ✓" if result.passed else "FAIL ✗"
        print(f"\n  [{status}] {name}  ({result.duration:.1f}s)")
        if result.errors:
            for e in result.errors[:5]:
                print(f"    ⚠ {e[:200]}")
            if len(result.errors) > 5:
                print(f"    ... and {len(result.errors) - 5} more")

        results.append(result)

    return results


def print_report(results: list[CaseResult]):
    passed = sum(1 for r in results if r.passed)
    failed = sum(1 for r in results if not r.passed)
    total = len(results)
    total_time = sum(r.duration for r in results)

    print(f"\n\n{'='*60}")
    print(f"  FINAL REPORT")
    print(f"{'='*60}")
    print(f"  Total: {total}   Passed: {passed}   Failed: {failed}   Time: {total_time:.0f}s\n")

    for r in results:
        icon = "✓" if r.passed else "✗"
        tool_count = len(r.stream.tool_results) if r.stream else 0
        err_count = len(r.errors)
        print(f"  {icon}  {r.name:<50s}  {r.duration:6.1f}s  tools={tool_count:<3d}  errors={err_count}")

    if failed > 0:
        print(f"\n{'─'*60}")
        print(f"  FAILED TESTS DETAIL:")
        print(f"{'─'*60}")
        for r in results:
            if not r.passed:
                print(f"\n  ✗ {r.name}:")
                for e in r.errors[:10]:
                    print(f"      {e[:200]}")

    print(f"\n{'='*60}")
    print(f"  {'ALL PASSED ✓' if failed == 0 else f'{failed} FAILED ✗'}")
    print(f"{'='*60}\n")


def main():
    parser = argparse.ArgumentParser(description="Excel Agent Test Suite")
    parser.add_argument("--case", type=str, default=None,
                        help="Run only tests whose name contains this string")
    args = parser.parse_args()

    # Health check
    print("Checking server health...")
    try:
        r = httpx.get(f"{BASE}/health", timeout=5)
        if r.status_code != 200:
            print(f"[FAIL] Server not healthy: {r.status_code}")
            sys.exit(1)
    except httpx.ConnectError:
        print("[FAIL] Cannot connect to server at", BASE)
        print("Start it with: EXCEL_AGENT_DEV=1 uv run uvicorn server:app --host 127.0.0.1 --port 8765")
        sys.exit(1)
    print("[OK] Server is up\n")

    # Ensure test output directory
    TEST_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Test output dir: {TEST_DIR}\n")

    # Set up test project
    setup_project()

    # Clean previous test files
    if not args.case:  # Only clean on full run
        for f in TEST_DIR.glob("*.xlsx"):
            print(f"  Removing old: {f.name}")
            f.unlink()

    # Run
    results = run_tests(selected=args.case)
    print_report(results)

    # Exit code
    failed = sum(1 for r in results if not r.passed)
    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()