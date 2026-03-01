# Excel Automation Toolkit v3.0

All-in-one Excel automation powered by Python + pandas.
18 modules, 100+ operations — now with a **Natural Language Interface**: just tell it what you want.

---

## What's New in v3.0 — Natural Language Interface

No more menus. Type what you want in plain English and the toolkit figures out which operation to run.

```bash
python ask.py sales.xlsx "remove duplicates"
python ask.py reports/ "consolidate all files into one"
python ask.py orders.xlsx "rfm customer segmentation"
python ask.py q1.xlsx q2.xlsx q3.xlsx "compare and find what changed"
```

Or use the browser UI:

```bash
streamlit run ask_web.py
# → opens http://localhost:8501
```

Windows: double-click `ask.bat` or `run_ask_web.bat`.

### How it works

- **100% offline** — no LLM, no API key, no internet required
- **133 mapped intents** across all 17 modules
- **Keyword scoring** — phrase matches score higher than single words
- **Synonym expansion** — "tidy" → "clean", "tidy up" → matches cleaner operations
- **File inspection** — reads sheet names, column headers, row counts before suggesting anything
- **Scenario detection** — structural context clues boost the right intent:

```
> python ask.py team_data.xlsx "three quarterly sheets, give me a consolidated one"

📊 File: team_data.xlsx
   Sheets: Q1_2025, Q2_2025, Q3_2025  ✓ matches "three sheets"
   Columns: Name, Dept, Hours, Project, Date
   Domain: HR / Project Management

Matched operations:
  [1] ✅  project_mgmt → timesheet_rollup       75%
  [2] 🔸  consolidator → merge_sheets_in_file   45%
  [3] 🔸  comparator   → compare_two_files      33%
```

### Confidence levels

| Score | Behaviour |
|-------|-----------|
| ≥ 60% | Auto-selects — asks Y/n to confirm |
| 30–60% | Shows ranked list — user picks |
| < 30% | Falls back to full module menu |

---

## Quick Start

**Natural Language (new):**
```bash
python ask.py file.xlsx "what to do"   # CLI
streamlit run ask_web.py               # browser UI
ask.bat file.xlsx "what to do"         # Windows CLI
run_ask_web.bat                        # Windows browser UI
```

**Classic menu:**
```bash
pip install -r requirements.txt
python main.py
```

**Windows (classic):** double-click `run.bat`

**Jump straight to a module:**
```bat
run_finance.bat
run_hr.bat
run_sales.bat
run_analytics.bat
```

---

## Natural Language Examples

| You type | Operation run |
|----------|---------------|
| `"remove duplicates"` | cleaner → remove_duplicates |
| `"fill missing values with mean"` | cleaner → fill_missing_values |
| `"pivot by department and sum sales"` | transformer → create_pivot_table |
| `"compare two files and find changes"` | comparator → compare_two_files |
| `"rfm customer segmentation"` | sales → rfm_segmentation |
| `"attrition analysis by department"` | hr → attrition_analysis |
| `"aging report for accounts receivable"` | finance → aging_analysis |
| `"sprint velocity backlog"` | project_mgmt → sprint_tracker |
| `"forecast next 12 months"` | analytics → trend_forecast |
| `"data quality report"` | validator → data_quality_report |
| `"consolidate all sheets in the file"` | consolidator → merge_sheets_in_file |
| `"vlookup from reference file"` | lookup → vlookup |

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
├── reports/
└── project_mgmt/
```

Double-click `run.bat` inside any folder to launch that module standalone.

To regenerate all standalone folders after editing a module:
```bash
python generate_standalone.py
```

---

## 18 Modules

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

### 17. Lookup & Match
- VLOOKUP — Join columns from a reference file
- Fuzzy Match — Approximate string matching (rapidfuzz)
- Multi-Key Lookup — Multi-column JOIN
- Reverse Lookup — Find key by value
- Enrich from Reference — Add columns from a master file

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

---

## Folder Structure

```
excel-automation/
├── nl_router.py             ← NL intent engine (133 intents, offline)
├── ask.py                   ← Natural language CLI
├── ask_web.py               ← Streamlit browser UI
├── ask.bat                  ← Windows NL CLI launcher
├── run_ask_web.bat          ← Windows browser UI launcher
├── main.py                  ← Classic full-menu CLI (all 18 modules)
├── generate_standalone.py   ← Re-generates standalone/ from main.py
├── run.bat                  ← Classic toolkit launcher
├── run_finance.bat          ← Jump directly to Finance
├── run_hr.bat               ← Jump directly to HR
├── run_*.bat                ← (one per module)
├── requirements.txt
├── modules/                 ← Shared module pool (17 files)
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
│   ├── lookup.py
│   └── project_mgmt.py
├── standalone/              ← Self-contained shareable folders
│   ├── finance/             → finance.py + cli.py + run.bat
│   ├── hr/
│   ├── project_mgmt/
│   └── ... (17 folders total)
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
streamlit (browser UI only — installed automatically by run_ask_web.bat)
```

All core dependencies installed automatically by `run.bat` or each module's own `run.bat`.
