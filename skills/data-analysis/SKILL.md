---
name: data-analysis
description: Analyze Excel data with statistical methods, identify patterns, and provide insights
triggers:
  - analyze
  - statistics
  - summary
  - patterns
  - insights
  - 数据分析
  - 统计
---

# Data Analysis Skill

You are performing data analysis on Excel files. Follow this structured approach to provide comprehensive insights.

## Analysis Workflow

### Step 1: Initial Data Assessment

Before any analysis, understand the data:

```
1. Load and preview the data
2. Check data types and column names
3. Identify missing values
4. Note the data size and complexity
```

### Step 2: Choose Analysis Type

Select the appropriate analysis based on user needs:

| Analysis Type | Best For | Tool Parameter |
|---------------|----------|----------------|
| Summary | Quick overview of numeric data | `summary` |
| Correlation | Finding relationships between variables | `correlation` |
| Missing Values | Data quality assessment | `missing` |
| Unique Values | Understanding categorical data | `unique` |
| Distribution | Examining value patterns | `distribution` |

### Step 3: Execute Analysis

Use the `analyze_data` tool with appropriate parameters:

```
analyze_data(
    file_path="path/to/file.xlsx",
    sheet_name="Sheet1",  # optional
    analysis_type="summary",  # or correlation, missing, unique, distribution
    columns="Col1,Col2"  # optional, for specific columns
)
```

### Step 4: Interpret Results

Provide meaningful insights:

1. **Summary Statistics**
   - Explain what the mean, median, and std dev tell us
   - Identify outliers using min/max vs quartiles
   - Note any surprising findings

2. **Correlation Analysis**
   - Highlight strong correlations (|r| > 0.7)
   - Note potential causal relationships
   - Warn about correlation ≠ causation

3. **Missing Values**
   - Identify problematic columns (>20% missing)
   - Suggest handling strategies
   - Note impact on analysis validity

4. **Distributions**
   - Identify skewed distributions
   - Note dominant categories
   - Suggest data transformation if needed

## Best Practices

### For Large Datasets
- Use `nrows` parameter to preview first
- Consider sampling for initial exploration
- Be aware of memory limitations

### For Time Series Data
- Check for date/time columns
- Look for seasonality patterns
- Consider trend analysis

### For Categorical Data
- Check cardinality of categories
- Look for imbalanced classes
- Consider grouping rare categories

## Example Dialogue

**User:** "Analyze the sales data and tell me about any patterns"

**Analysis Steps:**
1. List sheets to understand file structure
2. Read and preview the data
3. Run summary statistics
4. Check correlations between numeric columns
5. Examine categorical distributions

**Response Structure:**
```
Based on the analysis of your sales data:

📊 Overview:
- X rows, Y columns
- Date range: [start] to [end]

📈 Key Findings:
1. [Insight about central tendency]
2. [Insight about correlation]
3. [Insight about distribution]

⚠️ Data Quality Notes:
- [Missing value info]
- [Outlier info]

💡 Recommendations:
- [Actionable suggestion 1]
- [Actionable suggestion 2]
```

## Common Analysis Scenarios

### Sales Data Analysis
Focus on: revenue trends, product performance, regional comparisons

### Survey Data Analysis
Focus on: response distributions, correlations between questions, demographic breakdowns

### Financial Data Analysis
Focus on: trends over time, variance analysis, ratio calculations

### Operational Data Analysis
Focus on: efficiency metrics, bottlenecks, resource utilization
