"""Background data source analysis agent.

Reads Excel/CSV structure via COM (Excel) or stdlib csv, then uses a
one-shot LLM call to format the raw structure into a concise markdown
summary. Results are written to .excel-agent/schema/<source_name>.md
(one file per data source) plus a summary.md for all sources.
"""

from __future__ import annotations

import csv
import logging
import os
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger("app.analysis")

_SYSTEM_PROMPT = """\
You are a data structure analyst. Given raw Excel/CSV structure data, produce a concise markdown summary.

Rules:
- Document ONLY: sheet names, table regions, header structure (including multi-level headers), column names, inferred data types, 2-3 sample values per column, approximate row count
- For Chinese-style reports with merged headers: describe the header hierarchy clearly
- For sheets with multiple tables: document each table region separately
- Do NOT interpret data, draw conclusions, or make recommendations
- Output clean markdown suitable for embedding in a project memory file
- Be concise — one table per data region, no prose
"""


# ── CSV analysis ──────────────────────────────────────────────────────────────


def _analyze_csv(path: str) -> dict:
    """Read headers and sample values from a CSV file."""
    result: dict = {"path": path, "headers": [], "samples": [], "row_count": 0}
    try:
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.reader(f)
            rows = []
            count = 0
            for i, row in enumerate(reader):
                count = i
                if i == 0:
                    result["headers"] = row
                elif i <= 4:
                    rows.append(row)
                else:
                    break
            result["samples"] = rows
            # count remaining rows for estimate
            for _ in reader:
                count += 1
            result["row_count"] = count
    except Exception as e:
        result["error"] = str(e)
    return result


# ── Excel COM analysis ────────────────────────────────────────────────────────


def _get_merged_cells_info(sheet: Any) -> list[dict]:
    """Return list of merged cell regions in the used range."""
    merged = []
    try:
        used = sheet.UsedRange
        for cell in used:
            try:
                if cell.MergeCells:
                    area = cell.MergeArea
                    addr = area.Address
                    if not any(m["address"] == addr for m in merged):
                        merged.append({
                            "address": addr,
                            "row_span": area.Rows.Count,
                            "col_span": area.Columns.Count,
                            "value": str(cell.Value) if cell.Value is not None else "",
                        })
            except Exception:
                continue
    except Exception:
        pass
    return merged


def _detect_table_regions(data: list[list]) -> list[dict]:
    """Find separate table blocks by scanning for empty row separators."""
    if not data:
        return []

    regions = []
    n_rows = len(data)
    n_cols = max(len(r) for r in data) if data else 0

    start = None
    for i, row in enumerate(data):
        row_empty = all(v is None or str(v).strip() == "" for v in row)
        if not row_empty and start is None:
            start = i
        elif row_empty and start is not None:
            regions.append({"start_row": start, "end_row": i - 1, "start_col": 0, "end_col": n_cols - 1})
            start = None

    if start is not None:
        regions.append({"start_row": start, "end_row": n_rows - 1, "start_col": 0, "end_col": n_cols - 1})

    return regions if regions else [{"start_row": 0, "end_row": n_rows - 1, "start_col": 0, "end_col": n_cols - 1}]


def _detect_header_rows(data: list[list], region: dict) -> int:
    """Return number of header rows (heuristic: rows where >60% non-empty cells are strings)."""
    header_count = 0
    for i in range(region["start_row"], min(region["start_row"] + 5, region["end_row"] + 1)):
        row = data[i][region["start_col"]:region["end_col"] + 1]
        non_empty = [v for v in row if v is not None and str(v).strip() != ""]
        if not non_empty:
            break
        string_count = sum(1 for v in non_empty if isinstance(v, str) and not _is_numeric_str(str(v)))
        if string_count / len(non_empty) > 0.6:
            header_count += 1
        else:
            break
    return max(header_count, 1)


def _is_numeric_str(s: str) -> bool:
    try:
        float(s.replace(",", "").replace("%", ""))
        return True
    except ValueError:
        return False


def _analyze_excel_sheet(sheet: Any, sheet_name: str) -> dict:
    """Read sheet data and return structured info."""
    result: dict = {"sheet_name": sheet_name, "regions": [], "merged_cells": [], "error": None}
    try:
        used = sheet.UsedRange
        if used is None:
            return result

        # Read all values as 2D list
        raw = used.Value
        if raw is None:
            return result

        # Normalize to list of lists
        if not isinstance(raw, tuple):
            data: list[list] = [[raw]]
        else:
            data = [list(row) if isinstance(row, tuple) else [row] for row in raw]

        result["row_count"] = len(data)
        result["col_count"] = max(len(r) for r in data) if data else 0
        result["merged_cells"] = _get_merged_cells_info(sheet)

        regions = _detect_table_regions(data)
        for region in regions:
            n_headers = _detect_header_rows(data, region)
            header_rows = []
            for hi in range(region["start_row"], region["start_row"] + n_headers):
                header_rows.append(data[hi][region["start_col"]:region["end_col"] + 1])

            # Collect sample data rows (up to 3)
            sample_rows = []
            data_start = region["start_row"] + n_headers
            for si in range(data_start, min(data_start + 3, region["end_row"] + 1)):
                sample_rows.append(data[si][region["start_col"]:region["end_col"] + 1])

            data_row_count = region["end_row"] - data_start + 1

            result["regions"].append({
                "start_row": region["start_row"],
                "end_row": region["end_row"],
                "header_rows": header_rows,
                "sample_rows": sample_rows,
                "data_row_count": max(data_row_count, 0),
            })
    except Exception as e:
        result["error"] = str(e)
    return result


# ── Raw structure builder ─────────────────────────────────────────────────────


def _build_raw_structure(data_sources: list, mgr: Any) -> str:
    """Build markdown from all sources using COM for Excel, csv module for CSV."""
    lines: list[str] = []

    for source in data_sources:
        path: str = source.path
        src_type: str = source.type
        name: str = source.name

        lines.append(f"### {name} ({src_type})")
        lines.append(f"Path: `{path}`")
        lines.append("")

        if src_type == "csv":
            info = _analyze_csv(path)
            if "error" in info:
                lines.append(f"Error reading file: {info['error']}")
            else:
                lines.append(f"Rows: ~{info['row_count']}")
                lines.append(f"Headers: {', '.join(str(h) for h in info['headers'])}")
                if info["samples"]:
                    lines.append("Sample rows:")
                    for row in info["samples"]:
                        lines.append(f"  {row}")
        elif src_type in ("excel", "folder"):
            if src_type == "folder":
                # Scan folder for Excel/CSV files
                folder_path = Path(path)
                excel_files = list(folder_path.rglob("*.xlsx")) + list(folder_path.rglob("*.xls")) + list(folder_path.rglob("*.xlsm"))
                csv_files = list(folder_path.rglob("*.csv"))
                lines.append(f"Folder contains {len(excel_files)} Excel file(s), {len(csv_files)} CSV file(s)")
                for ef in excel_files[:3]:
                    lines.append(f"\n#### {ef.name}")
                    _append_excel_info(lines, str(ef), mgr)
                for cf in csv_files[:3]:
                    lines.append(f"\n#### {cf.name}")
                    info = _analyze_csv(str(cf))
                    if "error" not in info:
                        lines.append(f"Headers: {', '.join(str(h) for h in info['headers'])}")
            else:
                _append_excel_info(lines, path, mgr)

        lines.append("")

    return "\n".join(lines)


def _append_excel_info(lines: list[str], path: str, mgr: Any) -> None:
    """Append Excel workbook structure info to lines list."""
    try:
        entry = mgr.open_workbook(path)
        wb = entry.workbook
        sheet_count = wb.Sheets.Count
        lines.append(f"Sheets ({sheet_count}):")
        for i in range(1, min(sheet_count + 1, 11)):  # cap at 10 sheets
            sheet = wb.Sheets(i)
            sheet_name = sheet.Name
            info = _analyze_excel_sheet(sheet, sheet_name)
            lines.append(f"\n**Sheet: {sheet_name}**")
            if info.get("error"):
                lines.append(f"  Error: {info['error']}")
                continue
            lines.append(f"  Rows: ~{info.get('row_count', 0)}, Cols: {info.get('col_count', 0)}")
            if info.get("merged_cells"):
                lines.append(f"  Merged cells: {len(info['merged_cells'])} region(s)")
            for region in info.get("regions", []):
                if region["header_rows"]:
                    lines.append(f"  Headers: {region['header_rows']}")
                if region["sample_rows"]:
                    lines.append(f"  Sample data: {region['sample_rows'][:2]}")
                lines.append(f"  Data rows: ~{region['data_row_count']}")
        # Close workbook after reading (read-only analysis)
        try:
            if not entry.was_already_open:
                wb.Close(SaveChanges=False)
                from libs.excel_com.utils import normalize_path
                norm = normalize_path(os.path.abspath(path))
                mgr._registry.pop(norm, None)
        except Exception:
            pass
    except Exception as e:
        lines.append(f"Error opening workbook: {e}")


# ── LLM formatting ────────────────────────────────────────────────────────────


async def _format_with_llm(raw_structure: str, model_spec: str, api_key: str, provider: str) -> str:
    """Call LLM to format raw structure into clean markdown. Falls back to raw on error."""
    try:
        from agent.model_provider import PREDEFINED_PROVIDERS
        from langchain_openai import ChatOpenAI

        provider_cfg = PREDEFINED_PROVIDERS.get(provider)

        # Resolve api_key: explicit arg > env var
        resolved_key = api_key
        if not resolved_key and provider_cfg:
            resolved_key = os.environ.get(provider_cfg.api_key_env, "")
        if not resolved_key:
            logger.warning("LLM formatting skipped: no api_key for provider=%s", provider)
            return raw_structure

        model_name = model_spec.split(":", 1)[1] if ":" in model_spec else model_spec
        base_url = provider_cfg.api_base if provider_cfg else None

        llm_kwargs: dict = {"model": model_name, "temperature": 0, "openai_api_key": resolved_key}
        if base_url:
            llm_kwargs["openai_api_base"] = base_url

        logger.debug("LLM format: model=%s base_url=%s input_chars=%d", model_name, base_url, len(raw_structure))
        llm = ChatOpenAI(**llm_kwargs)
        from langchain_core.messages import HumanMessage, SystemMessage
        response = await llm.ainvoke([
            SystemMessage(content=_SYSTEM_PROMPT),
            HumanMessage(content=f"Format this data source structure:\n\n{raw_structure}"),
        ])
        result = str(response.content)
        logger.debug("LLM format done: output_chars=%d", len(result))
        return result
    except Exception as e:
        logger.warning("LLM formatting failed, using raw structure: %s", e)
        return raw_structure


# ── Public entry point ────────────────────────────────────────────────────────

SCHEMA_DIR_NAME = "schema"


def _safe_filename(name: str) -> str:
    """Convert a data source name to a safe filename."""
    import re
    return re.sub(r'[\\/:*?"<>|]', "_", name)


async def run_analysis(
    data_sources: list,
    model_spec: str,
    api_key: str,
    provider: str,
    project_root: Optional[Path],
) -> str:
    """Analyze all data sources, write per-source schema files, return summary markdown."""
    import time
    from libs.excel_com.instance_manager import ExcelInstanceManager  # type: ignore[import]

    logger.info("Analysis started: %d source(s), model=%s, project=%s",
                len(data_sources), model_spec or "(none)", project_root)

    mgr = ExcelInstanceManager()

    # Build schema dir
    schema_dir: Optional[Path] = None
    if project_root:
        schema_dir = project_root / ".excel-agent" / SCHEMA_DIR_NAME
        schema_dir.mkdir(parents=True, exist_ok=True)
        logger.debug("Schema dir: %s", schema_dir)

    summary_lines: list[str] = []
    use_llm = bool(model_spec and model_spec.strip())

    for source in data_sources:
        t0 = time.perf_counter()
        logger.debug("Analyzing source: name=%s type=%s path=%s", source.name, source.type, source.path)

        try:
            raw = _build_raw_structure([source], mgr)
            logger.debug("Raw structure built: source=%s chars=%d", source.name, len(raw))
        except Exception as e:
            logger.error("Failed to build raw structure for %s: %s", source.name, e)
            raw = f"Error reading {source.name}: {e}"

        if use_llm:
            content = await _format_with_llm(raw, model_spec, api_key, provider)
        else:
            content = raw

        # Prepend absolute path header so agent can use it directly
        abs_path = os.path.abspath(source.path)
        content = f"**File:** `{abs_path}`\n\n{content}"

        # Write per-source schema file
        if schema_dir is not None:
            fname = _safe_filename(source.name) + ".md"
            schema_file = schema_dir / fname
            schema_file.write_text(content, encoding="utf-8")
            logger.debug("Schema written: %s", schema_file)

        elapsed = (time.perf_counter() - t0) * 1000
        logger.info("Source analyzed: %s (%.0fms)", source.name, elapsed)

        # Use absolute path to schema file in summary
        if schema_dir is not None:
            schema_abs = str(schema_dir / (_safe_filename(source.name) + ".md"))
            summary_lines.append(
                f"- **{source.name}** (`{source.type}`) "
                f"| file: `{abs_path}` "
                f"| schema: `{schema_abs}`"
            )
        else:
            summary_lines.append(f"- **{source.name}** (`{source.type}`) | file: `{abs_path}`")

    summary = "\n".join(summary_lines)

    # Write summary.md
    if schema_dir is not None:
        (schema_dir / "summary.md").write_text(summary, encoding="utf-8")
        logger.debug("Summary written: %s", schema_dir / "summary.md")

    logger.info("Analysis complete: %d source(s) processed", len(data_sources))
    return summary
