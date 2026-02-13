"""
Excel processing tools for DeepAgents.

This module provides custom tools for reading, writing, and analyzing Excel files.
"""

import json
from pathlib import Path
from typing import Any, Optional

import pandas as pd
from langchain_core.tools import tool


@tool
def read_excel(
    file_path: str,
    sheet_name: Optional[str] = None,
    header: int = 0,
    nrows: Optional[int] = None,
) -> str:
    """
    Read data from an Excel file and return it as a formatted string.

    Args:
        file_path: Path to the Excel file (.xlsx, .xls)
        sheet_name: Name of the sheet to read. If None, reads the first sheet.
        header: Row number to use as column headers (0-indexed). Default is 0.
        nrows: Number of rows to read. If None, reads all rows.

    Returns:
        JSON string containing the data with metadata (columns, shape, preview).
    """
    try:
        path = Path(file_path)
        if not path.exists():
            return json.dumps({"error": f"File not found: {file_path}"})

        # Read the Excel file
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name if sheet_name else 0,
            header=header,
            nrows=nrows,
        )

        # Handle case where sheet_name returns a dict
        if isinstance(df, dict):
            df = list(df.values())[0]

        result = {
            "file_path": str(path.absolute()),
            "sheet_name": sheet_name or "first sheet",
            "columns": list(df.columns),
            "shape": {"rows": len(df), "columns": len(df.columns)},
            "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
            "data_preview": df.head(10).to_dict(orient="records"),
            "has_more_data": len(df) > 10,
        }

        return json.dumps(result, indent=2, default=str, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": f"Failed to read Excel file: {str(e)}"})


@tool
def list_sheets(file_path: str) -> str:
    """
    List all sheet names in an Excel file.

    Args:
        file_path: Path to the Excel file (.xlsx, .xls)

    Returns:
        JSON string containing list of sheet names and their basic info.
    """
    try:
        path = Path(file_path)
        if not path.exists():
            return json.dumps({"error": f"File not found: {file_path}"})

        # Get sheet names and info
        xl_file = pd.ExcelFile(file_path)
        sheets_info = []

        for sheet in xl_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            sheets_info.append(
                {
                    "name": sheet,
                    "rows": len(df),
                    "columns": len(df.columns),
                    "column_names": list(df.columns)[:20],  # First 20 columns
                }
            )

        result = {
            "file_path": str(path.absolute()),
            "total_sheets": len(sheets_info),
            "sheets": sheets_info,
        }

        return json.dumps(result, indent=2, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": f"Failed to list sheets: {str(e)}"})


@tool
def write_excel(
    file_path: str,
    data: str,
    sheet_name: str = "Sheet1",
    mode: str = "write",
) -> str:
    """
    Write data to an Excel file.

    Args:
        file_path: Path where the Excel file will be saved (.xlsx)
        data: JSON string representing the data (list of dicts or dict of lists)
        sheet_name: Name of the sheet to write to. Default is "Sheet1".
        mode: "write" to create new file, "append" to add sheet to existing file.

    Returns:
        JSON string with operation result and file info.
    """
    try:
        # Parse the data
        parsed_data = json.loads(data)

        # Convert to DataFrame
        if isinstance(parsed_data, list):
            df = pd.DataFrame(parsed_data)
        elif isinstance(parsed_data, dict):
            # Check if it's a dict of columns
            if all(isinstance(v, list) for v in parsed_data.values()):
                df = pd.DataFrame(parsed_data)
            else:
                df = pd.DataFrame([parsed_data])
        else:
            return json.dumps({"error": "Data must be a JSON array or object"})

        path = Path(file_path)

        if mode == "append" and path.exists():
            # Append to existing file
            with pd.ExcelWriter(
                file_path, mode="a", engine="openpyxl", if_sheet_exists="replace"
            ) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            # Create new file
            df.to_excel(file_path, sheet_name=sheet_name, index=False)

        result = {
            "success": True,
            "file_path": str(path.absolute()),
            "sheet_name": sheet_name,
            "rows_written": len(df),
            "columns_written": len(df.columns),
            "mode": mode,
        }

        return json.dumps(result, indent=2, ensure_ascii=False)

    except json.JSONDecodeError as e:
        return json.dumps({"error": f"Invalid JSON data: {str(e)}"})
    except Exception as e:
        return json.dumps({"error": f"Failed to write Excel file: {str(e)}"})


@tool
def analyze_data(
    file_path: str,
    sheet_name: Optional[str] = None,
    analysis_type: str = "summary",
    columns: Optional[str] = None,
) -> str:
    """
    Perform data analysis on an Excel file.

    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet to analyze. If None, uses first sheet.
        analysis_type: Type of analysis to perform:
            - "summary": Basic statistics (count, mean, std, min, max, quartiles)
            - "correlation": Correlation matrix for numeric columns
            - "missing": Missing value analysis
            - "unique": Count of unique values per column
            - "distribution": Value distribution for categorical columns
        columns: Comma-separated list of specific columns to analyze. If None, analyzes all.

    Returns:
        JSON string with analysis results.
    """
    try:
        path = Path(file_path)
        if not path.exists():
            return json.dumps({"error": f"File not found: {file_path}"})

        # Read the data
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name if sheet_name else 0,
        )
        if isinstance(df, dict):
            df = list(df.values())[0]

        # Filter columns if specified
        if columns:
            col_list = [c.strip() for c in columns.split(",")]
            df = df[col_list]

        result = {
            "file_path": str(path.absolute()),
            "sheet_name": sheet_name or "first sheet",
            "analysis_type": analysis_type,
            "shape": {"rows": len(df), "columns": len(df.columns)},
        }

        if analysis_type == "summary":
            # Get numeric and non-numeric columns
            numeric_df = df.select_dtypes(include=["number"])
            result["numeric_summary"] = (
                numeric_df.describe().to_dict() if len(numeric_df.columns) > 0 else {}
            )
            result["non_numeric_columns"] = list(
                df.select_dtypes(exclude=["number"]).columns
            )
            result["dtypes"] = {col: str(dtype) for col, dtype in df.dtypes.items()}

        elif analysis_type == "correlation":
            numeric_df = df.select_dtypes(include=["number"])
            if len(numeric_df.columns) < 2:
                result["error"] = "Need at least 2 numeric columns for correlation"
            else:
                result["correlation_matrix"] = numeric_df.corr().to_dict()

        elif analysis_type == "missing":
            missing = df.isnull().sum()
            result["missing_counts"] = missing.to_dict()
            result["missing_percentages"] = (
                (missing / len(df) * 100).round(2).to_dict()
            )
            result["total_cells"] = df.size
            result["total_missing"] = missing.sum()

        elif analysis_type == "unique":
            result["unique_counts"] = {col: df[col].nunique() for col in df.columns}
            result["total_rows"] = len(df)

        elif analysis_type == "distribution":
            cat_cols = df.select_dtypes(exclude=["number"]).columns
            if len(cat_cols) == 0:
                result["error"] = "No categorical columns found"
            else:
                distributions = {}
                for col in cat_cols[:10]:  # Limit to first 10 categorical columns
                    value_counts = df[col].value_counts().head(20)
                    distributions[col] = {
                        "values": value_counts.to_dict(),
                        "unique_count": df[col].nunique(),
                    }
                result["distributions"] = distributions

        else:
            result["error"] = f"Unknown analysis type: {analysis_type}"

        return json.dumps(result, indent=2, default=str, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": f"Analysis failed: {str(e)}"})


@tool
def generate_formula(
    description: str,
    context: Optional[str] = None,
) -> str:
    """
    Generate Excel formula suggestions based on a natural language description.

    Args:
        description: Natural language description of what the formula should do.
        context: Optional context about the data structure (column names, ranges, etc.)

    Returns:
        JSON string with formula suggestions and explanations.
    """
    # Common formula patterns
    formula_templates = {
        "sum": {
            "syntax": "=SUM(range)",
            "description": "Adds all numbers in a range",
            "example": "=SUM(A1:A10)",
        },
        "average": {
            "syntax": "=AVERAGE(range)",
            "description": "Calculates the arithmetic mean",
            "example": "=AVERAGE(B1:B10)",
        },
        "count": {
            "syntax": "=COUNT(range) or =COUNTA(range)",
            "description": "COUNT counts numbers, COUNTA counts non-empty cells",
            "example": "=COUNT(A1:A10)",
        },
        "if": {
            "syntax": "=IF(condition, value_if_true, value_if_false)",
            "description": "Returns one value if condition is true, another if false",
            "example": '=IF(A1>100, "High", "Low")',
        },
        "vlookup": {
            "syntax": "=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
            "description": "Looks up a value in the first column and returns a value from another column",
            "example": "=VLOOKUP(A1, B:C, 2, FALSE)",
        },
        "xlookup": {
            "syntax": "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])",
            "description": "Modern lookup function (Excel 365/2021)",
            "example": "=XLOOKUP(A1, B:B, C:C, \"Not Found\")",
        },
        "sumif": {
            "syntax": "=SUMIF(range, criteria, [sum_range])",
            "description": "Sums values that meet a condition",
            "example": "=SUMIF(A:A, \">100\", B:B)",
        },
        "countif": {
            "syntax": "=COUNTIF(range, criteria)",
            "description": "Counts cells that meet a condition",
            "example": '=COUNTIF(A:A, "Yes")',
        },
        "concatenate": {
            "syntax": "=CONCATENATE(text1, text2, ...) or =text1 & text2",
            "description": "Joins text strings together",
            "example": "=A1 & \" \" & B1",
        },
        "index_match": {
            "syntax": "=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))",
            "description": "Powerful combination for flexible lookups",
            "example": "=INDEX(C:C, MATCH(A1, B:B, 0))",
        },
        "pivot": {
            "syntax": "Insert > PivotTable (or use GETPIVOTDATA)",
            "description": "For complex data summarization, consider using a PivotTable",
            "example": "Select data range, Insert > PivotTable",
        },
        "text_to_columns": {
            "syntax": "Data > Text to Columns",
            "description": "For splitting text into multiple columns",
            "example": "Select column, Data > Text to Columns > Delimited",
        },
    }

    # Keywords to detect formula intent
    keywords_map = {
        "sum": ["sum", "add", "total", "合计", "求和", "加"],
        "average": ["average", "mean", "avg", "平均", "均值"],
        "count": ["count", "number of", "how many", "数量", "计数"],
        "if": ["if", "condition", "conditional", "如果", "条件"],
        "vlookup": ["lookup", "find", "search", "match", "查找", "搜索"],
        "xlookup": ["xlookup", "modern lookup"],
        "sumif": ["sum if", "sum where", "conditional sum", "条件求和"],
        "countif": ["count if", "count where", "conditional count", "条件计数"],
        "concatenate": ["concat", "join", "combine", "merge", "连接", "合并"],
        "index_match": ["index match", "flexible lookup", "精确查找"],
        "pivot": ["pivot", "summarize", "group by", "透视", "汇总"],
        "text_to_columns": ["split", "separate", "divide", "拆分", "分割"],
    }

    description_lower = description.lower()

    # Find matching formulas
    matched_formulas = []
    for formula_type, keywords in keywords_map.items():
        if any(kw in description_lower for kw in keywords):
            matched_formulas.append(formula_type)

    result = {
        "description": description,
        "context": context,
        "matched_formulas": matched_formulas,
        "suggestions": [
            formula_templates[ft] for ft in matched_formulas if ft in formula_templates
        ],
        "all_available_formulas": list(formula_templates.keys()),
        "usage_note": "Provide specific cell references or column names for more precise formula suggestions.",
    }

    return json.dumps(result, indent=2, ensure_ascii=False)


@tool
def create_pivot_summary(
    file_path: str,
    index_columns: str,
    value_columns: str,
    aggfunc: str = "sum",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None,
) -> str:
    """
    Create a pivot table summary from Excel data.

    Args:
        file_path: Path to the source Excel file
        index_columns: Comma-separated column names to use as row index
        value_columns: Comma-separated column names to aggregate
        aggfunc: Aggregation function: "sum", "mean", "count", "min", "max"
        sheet_name: Source sheet name. If None, uses first sheet.
        output_path: Path to save the result. If None, returns data only.

    Returns:
        JSON string with pivot table result.
    """
    try:
        path = Path(file_path)
        if not path.exists():
            return json.dumps({"error": f"File not found: {file_path}"})

        # Read the data
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name if sheet_name else 0,
        )
        if isinstance(df, dict):
            df = list(df.values())[0]

        # Parse column lists
        index_cols = [c.strip() for c in index_columns.split(",")]
        value_cols = [c.strip() for c in value_columns.split(",")]

        # Validate columns exist
        missing_index = [c for c in index_cols if c not in df.columns]
        missing_value = [c for c in value_cols if c not in df.columns]
        if missing_index or missing_value:
            return json.dumps(
                {
                    "error": f"Columns not found: index={missing_index}, values={missing_value}",
                    "available_columns": list(df.columns),
                }
            )

        # Map aggregation function
        agg_map = {
            "sum": "sum",
            "mean": "mean",
            "count": "count",
            "min": "min",
            "max": "max",
            "average": "mean",
        }
        agg = agg_map.get(aggfunc.lower(), "sum")

        # Create pivot table
        pivot = df.pivot_table(
            index=index_cols,
            values=value_cols,
            aggfunc=agg,
        )

        result = {
            "source_file": str(path.absolute()),
            "pivot_shape": {"rows": len(pivot), "columns": len(pivot.columns)},
            "aggregation": agg,
            "index_columns": index_cols,
            "value_columns": value_cols,
            "data": pivot.reset_index().to_dict(orient="records"),
        }

        # Save to file if output path provided
        if output_path:
            pivot.reset_index().to_excel(output_path, index=False)
            result["output_file"] = str(Path(output_path).absolute())

        return json.dumps(result, indent=2, default=str, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": f"Pivot table creation failed: {str(e)}"})


@tool
def filter_data(
    file_path: str,
    filter_conditions: str,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None,
) -> str:
    """
    Filter Excel data based on conditions.

    Args:
        file_path: Path to the source Excel file
        filter_conditions: JSON string with filter conditions.
            Format: {"column_name": {"operator": "value"}}
            Operators: "eq", "ne", "gt", "lt", "gte", "lte", "contains", "in"
            Example: '{"Age": {"gt": 18}, "Name": {"contains": "John"}}'
        sheet_name: Source sheet name. If None, uses first sheet.
        output_path: Path to save filtered data. If None, returns data only.

    Returns:
        JSON string with filtered data.
    """
    try:
        path = Path(file_path)
        if not path.exists():
            return json.dumps({"error": f"File not found: {file_path}"})

        # Read the data
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name if sheet_name else 0,
        )
        if isinstance(df, dict):
            df = list(df.values())[0]

        # Parse filter conditions
        conditions = json.loads(filter_conditions)
        original_count = len(df)

        # Apply filters
        for column, condition in conditions.items():
            if column not in df.columns:
                return json.dumps(
                    {
                        "error": f"Column '{column}' not found",
                        "available_columns": list(df.columns),
                    }
                )

            for operator, value in condition.items():
                if operator == "eq":
                    df = df[df[column] == value]
                elif operator == "ne":
                    df = df[df[column] != value]
                elif operator == "gt":
                    df = df[df[column] > value]
                elif operator == "lt":
                    df = df[df[column] < value]
                elif operator == "gte":
                    df = df[df[column] >= value]
                elif operator == "lte":
                    df = df[df[column] <= value]
                elif operator == "contains":
                    df = df[df[column].astype(str).str.contains(value, case=False, na=False)]
                elif operator == "in":
                    df = df[df[column].isin(value if isinstance(value, list) else [value])]

        result = {
            "source_file": str(path.absolute()),
            "original_rows": original_count,
            "filtered_rows": len(df),
            "filter_conditions": conditions,
            "data": df.head(100).to_dict(orient="records"),
            "has_more_data": len(df) > 100,
        }

        # Save to file if output path provided
        if output_path:
            df.to_excel(output_path, index=False)
            result["output_file"] = str(Path(output_path).absolute())

        return json.dumps(result, indent=2, default=str, ensure_ascii=False)

    except json.JSONDecodeError as e:
        return json.dumps({"error": f"Invalid filter conditions JSON: {str(e)}"})
    except Exception as e:
        return json.dumps({"error": f"Filter failed: {str(e)}"})


# Export all tools
EXCEL_TOOLS = [
    read_excel,
    list_sheets,
    write_excel,
    analyze_data,
    generate_formula,
    create_pivot_summary,
    filter_data,
]
