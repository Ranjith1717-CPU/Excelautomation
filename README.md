# Excel Automation Toolkit v1.0

All-in-one Excel automation powered by Python + pandas.
One interactive CLI covering every common Excel task.

---

## Quick Start

**Windows:** Double-click `run.bat`
**Manual:**
```bash
pip install -r requirements.txt
python main.py
```

---

## What It Does

### 1. Consolidate Files
| Operation | Description |
|-----------|-------------|
| Stack vertically | Append rows from 10+ files into one (union) |
| Join by key column | SQL-style INNER/OUTER/LEFT/RIGHT JOIN across files |
| Extract specific columns | Pull selected columns from multiple files |
| Merge all sheets in one file | Consolidate every sheet tab into a single sheet |
| Same sheet across files | Combine `Sheet1` from 10 files into one |

### 2. Calculate & Analyze
| Calculation | Formula |
|-------------|---------|
| Efficiency | `(Actual / Target) × 100` |
| Productivity | `Output / Input` |
| Utilization | `(Used / Available) × 100` |
| Variance | `Actual − Budget` and `%` |
| Growth Rate | `(Current − Previous) / Previous × 100` |
| Summary Stats | Count, Sum, Mean, Median, Std, Min, Max |
| % of Total | Each row's share of the column total |
| Moving Average | Rolling window average |
| KPI Dashboard | Multi-column key metrics in one sheet |
| Weighted Average | `sum(value × weight) / sum(weight)` |

### 3. Clean Data
- Remove duplicate rows (by all or specific columns)
- Drop empty rows and/or columns
- Trim whitespace from text cells
- Standardize date formats (e.g., all to `YYYY-MM-DD`)
- Fill missing values (mean / median / mode / forward-fill / custom)
- Auto-fix data types (text that is actually numbers/dates)
- Normalize text case (UPPER / lower / Title / Sentence)
- Remove special characters from columns
- Remove statistical outliers (±N standard deviations)
- **Full Auto-Clean** — runs all steps in one shot

### 4. Transform Data
- Create Pivot Tables (any aggregation function)
- Unpivot / Melt — wide → long format
- Transpose — flip rows and columns
- Split file by column value (one Excel per category)
- Split sheets into separate files
- Split large files into N-row chunks
- Reshape wide → long (repeated column stubs like Q1_Sales, Q2_Sales)
- Reshape long → wide (crosstab / unstack)
- Add running / cumulative total column
- Rank rows by a numeric column

### 5. Compare Files
- Full side-by-side diff of two files (new, deleted, changed, same)
- Find rows in File2 that don't exist in File1
- Find rows deleted from File1 in File2
- Find changed cell values (same key, different data)
- Find duplicates within a single file
- Find rows common to both files
- Cross-file duplicate check across N files

### 6. Column Operations
- Rename columns (interactive mapping)
- Merge / concatenate multiple columns into one
- Split a column by any delimiter
- Reorder columns in any desired sequence
- Drop / delete columns
- Add calculated columns (formula expressions like `Revenue - Cost`)
- Extract text using regex patterns
- Map / replace column values using a lookup dict
- Expand multi-value cells to separate rows
- Normalize all column headers (snake_case, Title Case, UPPER, lower)

### 7. Generate Reports
- **Summary Report** — stats for one or more files in one workbook
- **Data Profile** — detailed per-column analysis (types, nulls, unique, min/max)
- **KPI Report** — key metrics with group-level breakdown
- **Top-N / Bottom-N** — ranked records by any column
- **Frequency Report** — value counts and % for categorical columns
- **Monthly Summary** — aggregate numeric data by month

---

## Folder Structure
```
excel-automation/
├── main.py            ← Interactive CLI (the main entry point)
├── run.bat            ← Windows launcher (auto-installs deps)
├── requirements.txt   ← Python dependencies
├── modules/
│   ├── consolidator.py
│   ├── calculator.py
│   ├── cleaner.py
│   ├── transformer.py
│   ├── comparator.py
│   ├── reporter.py
│   └── column_ops.py
├── output/            ← All results saved here (auto-timestamped)
└── sample_data/       ← Place your test Excel files here
```

---

## Input Formats Accepted
- Single file path: `C:\data\sales.xlsx`
- Comma-separated: `file1.xlsx, file2.xlsx, file3.xlsx`
- Folder path: `C:\data\reports\`  (loads all .xlsx files)
- Glob pattern: `C:\data\*.xlsx`

---

## Dependencies
```
pandas, openpyxl, xlrd, colorama, tabulate, matplotlib, numpy
```
All installed automatically by `run.bat`.
