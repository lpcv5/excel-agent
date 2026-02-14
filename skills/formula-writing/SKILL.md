---
name: formula-writing
description: Generate and explain Excel formulas based on natural language descriptions
---

# Formula Writing Skill

You are helping users create Excel formulas. Translate natural language requirements into working Excel formulas with clear explanations.

## Formula Generation Workflow

### Step 1: Understand the Requirement

Ask clarifying questions when needed:
- What cells/ranges are involved?
- What is the expected output type?
- Are there any conditions or exceptions?
- What Excel version is being used?

### Step 2: Select the Right Function

Match user intent to Excel functions:

| User Intent | Recommended Functions |
|-------------|----------------------|
| "Add up" | SUM, SUMIF, SUMIFS |
| "Count" | COUNT, COUNTA, COUNTIF, COUNTIFS |
| "Average" | AVERAGE, AVERAGEIF, AVERAGEIFS |
| "Find/Lookup" | VLOOKUP, XLOOKUP, INDEX/MATCH |
| "If/Condition" | IF, IFS, SWITCH, IFERROR |
| "Text manipulation" | LEFT, RIGHT, MID, CONCATENATE, TEXT |
| "Date calculations" | DATEDIF, EDATE, EOMONTH, NETWORKDAYS |
| "Round numbers" | ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR |

### Step 3: Generate Formula

Use the `generate_formula` tool:

```
generate_formula(
    description="Calculate 5% commission for sales over $1000",
    context="Sales data in column B, commission in column C"
)
```

### Step 4: Explain the Formula

Always provide:
1. **Syntax**: The exact formula to use
2. **Explanation**: What each part does
3. **Example**: With sample cell references
4. **Variations**: Common alternatives

## Formula Categories

### Lookup Formulas

#### VLOOKUP (Traditional)
```
=VLOOKUP(lookup_value, table_range, column_num, [exact_match])
```
- Best for simple left-to-right lookups
- Requires lookup column to be first in range
- Use FALSE for exact match

#### XLOOKUP (Modern - Excel 365/2021)
```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
```
- More flexible, can look in any direction
- Built-in error handling
- Can return multiple values

#### INDEX/MATCH (Universal)
```
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
```
- Works in all Excel versions
- More flexible than VLOOKUP
- Can handle two-dimensional lookups

### Conditional Formulas

#### Simple IF
```
=IF(A1>100, "High", "Low")
```

#### Nested IF (avoid when possible)
```
=IF(A1>100, "High", IF(A1>50, "Medium", "Low"))
```

#### IFS (Excel 2019+)
```
=IFS(A1>100, "High", A1>50, "Medium", TRUE, "Low")
```

#### SUMIF / SUMIFS
```
=SUMIF(range, criteria, [sum_range])
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

### Text Formulas

#### Concatenation
```
=A1 & " " & B1
=CONCATENATE(A1, " ", B1)
=TEXTJOIN(", ", TRUE, A1:A10)
```

#### Text Extraction
```
=LEFT(A1, 3)
=RIGHT(A1, 3)
=MID(A1, start, length)
```

#### Text Cleaning
```
=TRIM(A1)           // Remove extra spaces
=CLEAN(A1)          // Remove non-printable chars
=UPPER/LOWER/PROPER(A1)
```

### Date Formulas

#### Date Arithmetic
```
=TODAY()
=NOW()
=A1 + 30            // Add 30 days
=EDATE(A1, 3)       // Add 3 months
=EOMONTH(A1, 0)     // End of month
```

#### Date Differences
```
=DATEDIF(start, end, "D")  // Days
=DATEDIF(start, end, "M")  // Months
=DATEDIF(start, end, "Y")  // Years
```

#### Working Days
```
=NETWORKDAYS(start, end, [holidays])
=WORKDAY(start, days, [holidays])
```

## Formula Optimization Tips

### 1. Avoid Volatile Functions
These recalculate on every change (slow for large sheets):
- INDIRECT, OFFSET, NOW, TODAY, RAND, RANDBETWEEN

### 2. Use Table References
Convert ranges to tables for maintainable formulas:
```
=SUM(Table1[Sales])
instead of
=SUM(B2:B1000)
```

### 3. Error Handling
Wrap formulas that might fail:
```
=IFERROR(VLOOKUP(...), "Not Found")
=IFNA(VLOOKUP(...), "Not Found")
```

### 4. Named Ranges
Use descriptive names:
```
=SUM(Q1_Sales)
instead of
=SUM(B2:B50)
```

## Example Dialogues

### Basic Formula Request
**User:** "I need to add up all values in column B"

**Response:**
```
Here's the formula to sum column B:

📊 Formula:
=SUM(B:B)

💡 Explanation:
- SUM adds all numeric values
- B:B references the entire column B
- Empty cells and text are automatically ignored

📌 Example:
=SUM(B2:B100)  // If you want specific rows only
```

### Conditional Formula Request
**User:** "Calculate 10% bonus for employees with sales over $5000"

**Response:**
```
Here's the formula for conditional bonus calculation:

📊 Formula:
=IF(B2>5000, B2*0.1, 0)

💡 Explanation:
- IF checks if sales (B2) exceed 5000
- If true: calculate 10% of sales
- If false: return 0

🔄 Alternative using SUMIF (for entire column):
=SUMIF(B:B, ">5000")*0.1

📌 Usage:
- Place in first row of bonus column
- Copy down for all employees
```

### Lookup Formula Request
**User:** "Look up product price from another sheet based on product code"

**Response:**
```
Here are options for looking up product prices:

📊 Option 1 - XLOOKUP (Recommended for Excel 365):
=XLOOKUP(A2, Products!A:A, Products!B:B, "Not Found")

📊 Option 2 - VLOOKUP:
=VLOOKUP(A2, Products!A:B, 2, FALSE)

📊 Option 3 - INDEX/MATCH (Most compatible):
=INDEX(Products!B:B, MATCH(A2, Products!A:A, 0))

💡 Recommendation:
- Use XLOOKUP if available (simplest)
- Use INDEX/MATCH for older Excel versions

⚠️ Note:
- Product codes in column A, prices in column B
- FALSE/0 ensures exact match only
```

## Common Formula Errors

| Error | Meaning | Common Fix |
|-------|---------|------------|
| #N/A | Value not found | Check lookup value exists |
| #REF! | Invalid reference | Check deleted cells/ranges |
| #VALUE! | Wrong data type | Check numeric vs text |
| #DIV/0! | Division by zero | Use IFERROR wrapper |
| #NAME? | Unknown function | Check spelling |
| #NUM! | Invalid numeric value | Check range/parameters |
