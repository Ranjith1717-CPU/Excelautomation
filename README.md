# Excel Automation Toolkit v2.0

All-in-one Excel automation powered by Python + pandas.
One interactive CLI covering every common Excel task — 18 modules, 100+ operations.

---

## Quick Start

**Windows:** Double-click `run.bat`
**Jump straight to a module:**
```bat
run_finance.bat
run_hr.bat
run_sales.bat
run_analytics.bat
```
**Manual:**
```bash
pip install -r requirements.txt
python main.py
python main.py finance       # jump directly to a module
```

---

## Standalone Modules

Each module also lives in `standalone/<name>/` as a fully self-contained folder.
Zip any folder and share it — the recipient only needs that one folder, no toolkit required.

```
standalone/
├── finance/      → finance.py + cli.py + run.bat
├── hr/
├── sales/
├── inventory/
├── analytics/
├── formatter/
├── validator/
├── converter/
├── lookup/       (requires rapidfuzz — auto-installed)
├── consolidate/
├── calculate/
├── clean/
├── transform/
├── compare/
├── columns/
└── reports/
```

Double-click `run.bat` inside any folder to launch that module standalone.

To regenerate all standalone folders after editing a module:
```bash
python generate_standalone.py
```

---

## 17 Modules

### 1. Consolidate Files
| Operation | Description |
|-----------|-------------|
| Stack vertically | Append rows from N files into one |
| Join by key column | SQL-style INNER/OUTER/LEFT/RIGHT JOIN |
| Extract specific columns | Pull selected columns from multiple files |
| Merge all sheets | Consolidate every sheet tab into one |
| Same sheet across files | Combine `Sheet1` from N files |

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
- Split file by column value (one file per category)
- Split sheets into separate files
- Split large files into N-row chunks
- Reshape wide → long (repeated column stubs)
- Reshape long → wide (crosstab / unstack)
- Add running / cumulative total column
- Rank rows by a numeric column

### 5. Compare Files
- Full side-by-side diff of two files (new, deleted, changed, same)
- Find rows added / deleted between two versions
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
- Add calculated columns (`Revenue - Cost`, `Units * Price`, etc.)
- Extract text using regex patterns
- Map / replace column values
- Expand multi-value cells to separate rows
- Normalize all column headers (snake_case, Title Case, UPPER, lower)

### 7. Generate Reports
- **Summary Report** — stats for one or more files
- **Data Profile** — per-column analysis (types, nulls, unique, min/max)
- **KPI Report** — key metrics with group-level breakdown
- **Top-N / Bottom-N** — ranked records by any column
- **Frequency Report** — value counts and % for categorical columns
- **Monthly Summary** — aggregate numeric data by month

### 9. Finance
- AR/AP Aging Analysis — 0-30, 31-60, 61-90, 90+ day buckets
- Loan Amortization Schedule — EMI, principal, interest, balance
- Depreciation Schedule — Straight-line + declining balance
- Financial Ratios — Gross margin, ROI, current ratio
- Payroll Calculator — Gross→Net with HRA/PF/ESI/TDS
- Budget vs Actual — Variance + % variance report
- Compound Interest Schedule — Future value growth table

### 10. HR Analytics
- Attrition Analysis — Turnover rate by department
- Headcount Summary — Count/% by group columns
- Tenure Analysis — Years-of-service bands
- Age Band Analysis — Workforce age demographics
- Salary Analysis — Min/max/median/percentiles
- Performance Distribution — Rating bell-curve + ranking
- Salary Increment Calculator — Apply % increment to salary

### 11. Sales Analytics
- Commission Calculator — Flat % or tiered slab
- RFM Segmentation — Recency, Frequency, Monetary
- Quota Attainment — % attainment + status labels
- Pipeline Analysis — Funnel by stage + conversion
- Sales by Territory — Territory summary + rank
- Customer ABC Analysis — A/B/C by revenue contribution
- Discount Analysis — Discount % + revenue leakage

### 12. Inventory Management
- ABC Analysis — Classify items by value contribution
- Reorder Point — ROP = usage × lead time + safety stock
- Stock Aging — Age-bucket inventory by receipt date
- Inventory Turnover — Turnover ratio + days on hand
- OEE Calculator — Availability × Performance × Quality
- Dead Stock Analysis — Items with no movement beyond N days

### 13. Format & Style
- Add Bar / Line / Pie Charts (embedded in Excel)
- Traffic Light Colors — Red/Yellow/Green conditional fill
- Color Scale — Gradient min→max fill
- Format as Table — Apply Excel table style
- Freeze header + auto-filter
- Auto-fit column widths
- Add Totals row
- Highlight duplicate cells
- Apply number formats (currency, %, comma)

### 14. Validate Data
- Check Mandatory Fields — Flag rows missing required values
- Validate Email — Regex format check
- Validate Phone — Phone number format check
- Numeric Range Check — Flag out-of-range values
- Date Range Check — Flag dates outside boundaries
- Referential Integrity — Values must exist in a lookup file
- Data Quality Report — Comprehensive quality score 0-100
- Detect PII — Flag columns with sensitive data

### 15. Statistical Analytics
- Correlation Matrix — Pairwise Pearson correlation
- Pareto Analysis — 80/20 cumulative % chart data
- Linear Regression — OLS + R² + predictions
- Trend Forecast — Extrapolate trend N periods ahead
- Frequency Distribution — Histogram bin counts
- Z-Score Analysis — Outlier detection via std deviation
- Cohort Retention — Monthly customer retention matrix

### 16. Convert Files
- Excel → CSV (each sheet to separate CSV)
- CSV → Excel (multiple CSVs into one workbook)
- Excel → JSON / JSON → Excel
- XLS → XLSX batch conversion
- Excel → Tab/pipe/custom delimited text
- Merge multiple CSV files into one Excel

### 18. Project Management
- **Team Consolidator** — Merge team member data from multiple files/sheets; optional dedup on ID column
- **Split by Team** — One file per department/team from a master sheet
- **Timesheet Rollup** — Consolidate N timesheet files → Detail + By Person + By Project + Person×Project pivot
- **Resource Allocation** — Allocation % per resource per project, over-allocation flags
- **Milestone Tracker** — RAG status (Red/Amber/Green), slippage days, overdue count, owner summary
- **RACI Matrix** — Build and validate R/A/C/I matrix, flag tasks with no Accountable/Responsible
- **Risk Register** — Score by Probability×Impact (1-5), rank Critical/High/Medium/Low, heat map data
- **Action Tracker** — Consolidate meeting action items, flag overdue, Days_Overdue, priority levels
- **Capacity Planner** — Available vs allocated hours, utilisation %, over-allocation alerts by team
- **Sprint Tracker** — Velocity per sprint, completion %, backlog health, sprints-to-clear estimate

### 17. Lookup & Match
- VLOOKUP — Join columns from a reference file
- Fuzzy Match — Approximate string matching (rapidfuzz)
- Multi-Key Lookup — Multi-column JOIN
- Reverse Lookup — Find key by value
- Enrich from Reference — Add columns from a master file

---

## Folder Structure
```
excel-automation/
├── main.py                  ← Full toolkit (all 17 modules)
├── generate_standalone.py   ← Re-generates standalone/ from main.py
├── run.bat                  ← Full toolkit launcher
├── run_finance.bat          ← Jump directly to Finance
├── run_hr.bat               ← Jump directly to HR
├── run_*.bat                ← (one per module)
├── requirements.txt
├── modules/                 ← Shared module pool
│   ├── consolidator.py
│   ├── calculator.py
│   ├── cleaner.py
│   ├── transformer.py
│   ├── comparator.py
│   ├── reporter.py
│   ├── column_ops.py
│   ├── finance.py
│   ├── hr.py
│   ├── sales.py
│   ├── inventory.py
│   ├── formatter.py
│   ├── validator.py
│   ├── analytics.py
│   ├── converter.py
│   └── lookup.py
├── standalone/              ← Self-contained shareable folders
│   ├── finance/             → finance.py + cli.py + run.bat
│   ├── hr/
│   ├── project_mgmt/
│   └── ... (17 folders)
├── output/                  ← All results saved here (auto-timestamped)
└── sample_data/             ← Place your test Excel files here
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
pandas, openpyxl, xlrd, colorama, tabulate, numpy, rapidfuzz (lookup only)
```
All installed automatically by `run.bat` or each module's own `run.bat`.
