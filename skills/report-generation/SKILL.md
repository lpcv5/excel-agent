---
name: report-generation
description: Create formatted reports, summaries, and data exports from Excel data
---

# Report Generation Skill

You are helping users create reports from Excel data. Generate summaries, formatted outputs, and export data in structured ways.

## Report Generation Workflow

### Step 1: Understand Report Requirements

Clarify the report purpose:
- What is the target audience?
- What decisions will this report support?
- What time period does it cover?
- What format is expected?

### Step 2: Data Preparation

Before creating the report:

1. **Verify Data Quality**
   ```
   analyze_data(file_path, analysis_type="missing")
   analyze_data(file_path, analysis_type="summary")
   ```

2. **Filter Relevant Data**
   ```
   filter_data(
       file_path,
       filter_conditions='{"Date": {"gte": "2024-01-01"}}'
   )
   ```

3. **Create Aggregations**
   ```
   create_pivot_summary(
       file_path,
       index_columns="Category",
       value_columns="Sales,Quantity",
       aggfunc="sum"
   )
   ```

### Step 3: Generate Report Structure

Organize the output logically:

```
📊 Report Title
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📌 Executive Summary
- Key finding 1
- Key finding 2
- Key finding 3

📈 Detailed Analysis
[Tables and data]

📋 Recommendations
- Action item 1
- Action item 2

📎 Data Sources
- File: [filename]
- Generated: [timestamp]
```

### Step 4: Export Results

Save the report:
```
write_excel(
    file_path="output/report.xlsx",
    data="[formatted data]",
    sheet_name="Report"
)
```

## Report Types

### Summary Report

Quick overview of key metrics:

```python
# Generate summary statistics
analyze_data(file_path, analysis_type="summary")

# Create pivot summary
create_pivot_summary(
    file_path,
    index_columns="Category",
    value_columns="Amount",
    aggfunc="sum"
)
```

**Output Structure:**
```
📊 Summary Report
─────────────────
Period: [date range]
Total Records: [count]

🔢 Key Metrics:
- Total: $XXX,XXX
- Average: $X,XXX
- Min: $XXX
- Max: $XX,XXX

📊 Breakdown by Category:
| Category | Amount | % of Total |
|----------|--------|------------|
| A        | $X,XXX | XX%        |
| B        | $X,XXX | XX%        |
```

### Comparison Report

Compare data across periods or categories:

```python
# Create multiple pivot tables
create_pivot_summary(file_path, index_columns="Region", value_columns="Sales")
create_pivot_summary(file_path, index_columns="Product", value_columns="Sales")
```

**Output Structure:**
```
📊 Comparison Report
────────────────────

📍 By Region:
| Region | Q1 | Q2 | Change |
|--------|-----|-----|--------|
| North  | $X | $Y | +Z%    |

📍 By Product:
| Product | Current | Previous | Trend |
|---------|---------|----------|-------|
| Item A  | $X      | $Y       | ↑     |
```

### Exception Report

Highlight outliers and issues:

```python
# Filter for anomalies
filter_data(
    file_path,
    filter_conditions='{"Status": {"eq": "Exception"}}'
)
```

**Output Structure:**
```
⚠️ Exception Report
───────────────────
Generated: [timestamp]

🔴 Critical Issues: X items
| ID | Issue | Value | Threshold |
|----|-------|-------|-----------|
| 1  | ...   | ...   | ...       |

🟡 Warnings: Y items
| ID | Issue | Value | Expected |
|----|-------|-------|----------|
```

### Trend Report

Show changes over time:

```python
# Create time-based pivot
create_pivot_summary(
    file_path,
    index_columns="Month",
    value_columns="Revenue,Cost",
    aggfunc="sum"
)
```

**Output Structure:**
```
📈 Trend Report
───────────────
Period: [start] to [end]

Monthly Trend:
| Month    | Revenue | Cost | Profit | Margin |
|----------|---------|------|--------|--------|
| Jan 2024 | $X      | $Y   | $Z     | XX%    |
| Feb 2024 | $X      | $Y   | $Z     | XX%    |

📉 Key Observations:
- Revenue trend: [increasing/decreasing/stable]
- Best month: [month]
- Growth rate: X%
```

## Formatting Best Practices

### Number Formatting
- Currency: $1,234.56
- Percentages: 12.3%
- Large numbers: 1.2M, 1.5K
- Decimals: consistent (2 decimal places)

### Table Formatting
- Left-align text, right-align numbers
- Use consistent column widths
- Highlight headers
- Add totals row when appropriate

### Visual Hierarchy
- Use headers for sections
- Bold key metrics
- Use symbols for quick scanning (↑↓→)
- Add whitespace for readability

## Export Options

### Single Sheet Export
```python
write_excel(
    file_path="report.xlsx",
    data=json_data,
    sheet_name="Summary"
)
```

### Multi-Sheet Report
```python
# Create workbook with multiple sheets
write_excel("report.xlsx", summary_data, "Summary", mode="write")
write_excel("report.xlsx", detail_data, "Details", mode="append")
write_excel("report.xlsx", raw_data, "Raw Data", mode="append")
```

## Report Templates

### Executive Summary Template
```
📊 [Report Name] - Executive Summary
═══════════════════════════════════════
Generated: [timestamp]
Data Period: [start] to [end]

📌 KEY FINDINGS
───────────────
1. [Primary insight]
2. [Secondary insight]
3. [Notable observation]

🔢 METRICS AT A GLANCE
──────────────────────
| Metric       | Value    | vs Prior |
|--------------|----------|----------|
| Total Sales  | $XXX,XXX | +X%      |
| Transactions | X,XXX    | +X%      |
| Average      | $XX.XX   | -X%      |

📋 RECOMMENDATIONS
──────────────────
1. [Action recommendation]
2. [Action recommendation]

📎 DATA SOURCE
──────────────
File: [filename]
Records: [count]
Last Updated: [date]
```

### Data Quality Report Template
```
🔍 Data Quality Report
═══════════════════════
File: [filename]
Analyzed: [timestamp]

📊 DATASET OVERVIEW
───────────────────
- Total Rows: X,XXX
- Total Columns: XX
- Numeric Columns: XX
- Text Columns: XX

⚠️ DATA QUALITY ISSUES
──────────────────────
| Column    | Missing | Duplicates | Issues |
|-----------|---------|------------|--------|
| Email     | X%      | X          | [list] |
| Phone     | X%      | X          | [list] |

✅ RECOMMENDATIONS
──────────────────
1. [Data cleaning suggestion]
2. [Validation rule suggestion]
```

## Common Report Scenarios

### Sales Report
Focus on: revenue, units, growth rates, top performers

### Financial Report
Focus on: income, expenses, profit margins, variances

### Operations Report
Focus on: efficiency metrics, throughput, resource utilization

### Customer Report
Focus on: acquisition, retention, satisfaction scores

## Example Dialogue

**User:** "Create a monthly sales report for Q4"

**Steps:**
1. Filter data for Q4 (Oct-Dec)
2. Create pivot by month
3. Calculate totals and averages
4. Identify trends and top performers
5. Format and export

**Response:**
```
📊 Q4 2024 Monthly Sales Report
═══════════════════════════════════
Generated: [timestamp]
Data Period: October - December 2024

📌 EXECUTIVE SUMMARY
────────────────────
Q4 total revenue: $1,234,567
Best performing month: November (+15% vs Oct)
Top region: West Coast ($456,789)

📈 MONTHLY BREAKDOWN
────────────────────
| Month     | Revenue   | Orders | Avg Order |
|-----------|-----------|--------|-----------|
| October   | $380,000  | 1,234  | $308      |
| November  | $437,000  | 1,456  | $300      |
| December  | $417,567  | 1,389  | $301      |
| TOTAL     | $1,234,567| 4,079  | $303      |

🏆 TOP PERFORMERS
─────────────────
| Region      | Revenue  | % of Total |
|-------------|----------|------------|
| West Coast  | $456,789 | 37%        |
| Northeast   | $345,678 | 28%        |
| Midwest     | $234,567 | 19%        |

📎 Report saved to: q4_sales_report.xlsx
```
