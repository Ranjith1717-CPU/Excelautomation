"""
nl_router.py — Natural Language Intent Router for Excel Automation Toolkit
===========================================================================
No LLM / API required. 100% offline keyword + context matching.

Usage:
    from nl_router import parse_intent, inspect_file
    results = parse_intent("remove duplicates", files=["sales.xlsx"])
    # → [{"intent": {...}, "score": 0.92, "confidence": "high"}, ...]
"""

import re
import datetime
from pathlib import Path
from typing import List, Dict, Optional, Any

# ── Synonym expansion map ─────────────────────────────────────────────────────
SYNONYMS: Dict[str, List[str]] = {
    "clean":       ["tidy", "fix", "sanitize", "scrub", "purge", "polish"],
    "duplicate":   ["dupe", "dupes", "repeat", "repeated", "dup", "redundant"],
    "remove":      ["delete", "drop", "eliminate", "get rid of", "strip", "clear"],
    "merge":       ["combine", "consolidate", "stack", "join", "union", "append", "concat"],
    "pivot":       ["group by", "aggregate", "summarize", "rollup", "crosstab", "tabulate"],
    "compare":     ["diff", "difference", "changes", "what changed", "delta", "contrast"],
    "split":       ["divide", "separate", "break", "partition", "chunk", "segment"],
    "format":      ["style", "colour", "color", "beautify", "highlight", "decorate"],
    "validate":    ["check", "verify", "audit", "inspect", "flag", "catch errors"],
    "convert":     ["transform", "export", "change format", "save as"],
    "lookup":      ["find", "match", "search", "vlookup", "fetch"],
    "report":      ["summarize", "overview", "summary", "analyze", "profile"],
    "forecast":    ["predict", "project", "trend", "future"],
    "rank":        ["sort", "order", "top", "best", "worst", "bottom"],
    "calculate":   ["compute", "derive", "work out", "figure out"],
    "missing":     ["null", "blank", "empty", "nan", "na", "no value"],
    "whitespace":  ["spaces", "trim", "leading spaces", "trailing spaces", "padding"],
    "outlier":     ["anomaly", "extreme", "unusual", "weird value", "spike"],
    "attrition":   ["turnover", "churn", "employee leaving", "resignation"],
    "headcount":   ["employee count", "staff count", "workforce size", "team size"],
    "commission":  ["incentive", "bonus", "sales reward"],
    "inventory":   ["stock", "warehouse", "items", "sku"],
    "sprint":      ["iteration", "agile", "scrum", "velocity"],
    "risk":        ["threat", "hazard", "concern"],
    "milestone":   ["deadline", "delivery", "checkpoint", "due date"],
    "timesheet":   ["hours", "time tracking", "working hours", "effort"],
    "aging":       ["overdue", "outstanding", "past due", "accounts receivable", "accounts payable"],
    "payroll":     ["salary", "wages", "compensation", "pay", "remuneration"],
    "correlation": ["relationship", "association", "linked", "connected"],
    "regression":  ["linear model", "trend line", "prediction model"],
    "frequency":   ["count", "occurrences", "how many", "distribution"],
    "cohort":      ["retention", "cohort analysis", "customer retention"],
    "rfm":         ["recency frequency monetary", "customer segments", "customer value"],
}

# ── Domain hint patterns (column name → domain) ───────────────────────────────
DOMAIN_HINTS = {
    "hr_project":  ["hours", "timesheet", "project", "sprint", "velocity", "backlog"],
    "finance":     ["invoice", "amount due", "aging", "principal", "interest", "budget"],
    "sales":       ["quota", "pipeline", "territory", "commission", "revenue", "deal"],
    "inventory":   ["qty", "sku", "reorder", "stock", "warehouse", "unit"],
    "hr":          ["department", "employee", "headcount", "attrition", "hire date", "salary"],
    "analytics":   ["score", "rating", "metric", "kpi", "target"],
}

# ── Quantity word map for context extraction ──────────────────────────────────
QUANTITY_WORDS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "dozen": 12,
    "multiple": 2, "several": 3, "few": 3, "many": 5, "some": 2,
}

# =============================================================================
# INTENT MAP — 100+ operations across 17 modules
# Each entry:
#   id       : unique string
#   module   : module name (matches modules/<module>.py)
#   fn       : exact function name
#   desc     : one-line human description
#   keywords : list of trigger phrases/words
#   anti     : penalize if these words appear
#   multi    : True = needs multiple input files
#   params   : list of {name, type, prompt, options, default, optional}
#
# Param types:
#   file       → first provided file (any arg name)
#   files      → all provided files as list
#   file1      → first file (for two-file ops)
#   file2      → second file (prompt if not provided)
#   ref_file   → reference/lookup file (always prompted)
#   output     → auto-generated .xlsx output path
#   output_dir → auto-generated output directory
#   output_csv → auto-generated .csv output path
#   output_json→ auto-generated .json output path
#   col_req    → required column name (show available columns)
#   col_opt    → optional column name (Enter=skip)
#   cols_req   → required list of columns (comma-sep)
#   cols_opt   → optional list of columns (Enter=all)
#   number     → integer (try extract from query, else prompt)
#   float_val  → float with default
#   choice     → pick from options list
#   string     → free text with optional default
#   bool_val   → Y/N prompt
#   mapping    → dict of key→value pairs (prompted interactively)
# =============================================================================

INTENT_MAP = [

    # ── CLEANER ──────────────────────────────────────────────────────────────
    {
        "id": "clean_duplicates",
        "module": "cleaner", "fn": "remove_duplicates",
        "desc": "Remove duplicate rows from a file",
        "keywords": ["duplicate", "dupe", "dedup", "unique rows", "remove duplicate",
                     "duplicates", "dupes", "drop duplicate", "deduplicate"],
        "anti": ["find duplicate", "highlight duplicate", "detect duplicate"],
        "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
            {"name": "subset",      "type": "cols_opt", "prompt": "Columns to check (Enter=all)"},
            {"name": "keep",        "type": "choice",   "prompt": "Keep which duplicate?",
             "options": ["first", "last"], "default": "first"},
        ]
    },
    {
        "id": "clean_empty",
        "module": "cleaner", "fn": "remove_empty_rows_cols",
        "desc": "Drop empty rows and/or columns",
        "keywords": ["empty rows", "blank rows", "empty columns", "blank columns",
                     "drop empty", "remove empty", "drop blank", "null rows"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "clean_whitespace",
        "module": "cleaner", "fn": "trim_whitespace",
        "desc": "Trim leading/trailing whitespace from text columns",
        "keywords": ["trim", "whitespace", "leading spaces", "trailing spaces",
                     "strip spaces", "remove spaces", "extra spaces"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "clean_dates",
        "module": "cleaner", "fn": "standardize_dates",
        "desc": "Parse and reformat date columns to a standard format",
        "keywords": ["standardize dates", "date format", "fix dates", "clean dates",
                     "date column", "reformat date", "parse date"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "output_path",  "type": "output"},
            {"name": "date_columns", "type": "cols_opt", "prompt": "Date columns (Enter=auto-detect)"},
            {"name": "date_format",  "type": "string", "prompt": "Output date format",
             "default": "%Y-%m-%d"},
        ]
    },
    {
        "id": "clean_missing",
        "module": "cleaner", "fn": "fill_missing_values",
        "desc": "Fill missing/NaN values using a chosen strategy",
        "keywords": ["fill missing", "missing values", "fill null", "impute",
                     "fill blank", "replace null", "fill na", "handle missing"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
            {"name": "strategy",    "type": "choice", "prompt": "Fill strategy",
             "options": ["mean", "median", "mode", "zero", "forward", "backward"],
             "default": "mean"},
            {"name": "columns",     "type": "cols_opt", "prompt": "Columns to fill (Enter=all)"},
        ]
    },
    {
        "id": "clean_types",
        "module": "cleaner", "fn": "fix_data_types",
        "desc": "Auto-detect and fix column data types",
        "keywords": ["fix types", "data types", "coerce types", "fix data type",
                     "wrong types", "type errors", "numeric columns"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "clean_case",
        "module": "cleaner", "fn": "normalize_text_case",
        "desc": "Convert text columns to upper/lower/title case",
        "keywords": ["uppercase", "lowercase", "title case", "capitalize", "normalize case",
                     "text case", "upper case", "lower case"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
            {"name": "case",        "type": "choice", "prompt": "Case style",
             "options": ["lower", "upper", "title", "sentence"], "default": "lower"},
            {"name": "columns",     "type": "cols_opt", "prompt": "Columns (Enter=all text cols)"},
        ]
    },
    {
        "id": "clean_special_chars",
        "module": "cleaner", "fn": "remove_special_characters",
        "desc": "Strip special/non-alphanumeric characters from text columns",
        "keywords": ["special characters", "special chars", "non alphanumeric",
                     "remove symbols", "clean text", "strip special"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
            {"name": "columns",     "type": "cols_opt", "prompt": "Columns (Enter=all text cols)"},
        ]
    },
    {
        "id": "clean_outliers",
        "module": "cleaner", "fn": "remove_outliers",
        "desc": "Remove rows with extreme/outlier values",
        "keywords": ["outlier", "outliers", "extreme values", "anomaly", "anomalies",
                     "remove outliers", "drop outliers", "unusual values"],
        "anti": ["detect outlier", "find outlier", "flag outlier"],
        "multi": False,
        "params": [
            {"name": "file",          "type": "file"},
            {"name": "output_path",   "type": "output"},
            {"name": "columns",       "type": "cols_opt", "prompt": "Columns to check (Enter=all numeric)"},
            {"name": "std_threshold", "type": "float_val", "prompt": "Std dev threshold",
             "default": 3.0},
        ]
    },
    {
        "id": "clean_full",
        "module": "cleaner", "fn": "full_clean",
        "desc": "Run all cleaning steps in one shot",
        "keywords": ["full clean", "clean everything", "clean all", "complete clean",
                     "clean the file", "clean up", "all cleaning"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },

    # ── CONSOLIDATOR ─────────────────────────────────────────────────────────
    {
        "id": "consolidate_stack",
        "module": "consolidator", "fn": "merge_files_stack",
        "desc": "Stack multiple Excel files vertically (append rows)",
        "keywords": ["stack files", "consolidate files", "combine files", "append files",
                     "merge files", "stack together", "union files", "merge all"],
        "anti": ["join on", "by key", "same sheet", "sheets"],
        "multi": True,
        "params": [
            {"name": "files",       "type": "files"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "consolidate_by_key",
        "module": "consolidator", "fn": "merge_files_by_key",
        "desc": "Join multiple files on a shared key column (SQL-style)",
        "keywords": ["join on key", "merge on key", "sql join", "join files by",
                     "merge by key", "lookup join", "key column merge"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",       "type": "files"},
            {"name": "key_column",  "type": "col_req", "prompt": "Key column to join on"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "consolidate_columns",
        "module": "consolidator", "fn": "merge_specific_columns",
        "desc": "Extract specific columns from multiple files and stack them",
        "keywords": ["extract columns from files", "specific columns", "selected columns",
                     "merge specific columns", "pick columns from files"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",       "type": "files"},
            {"name": "columns",     "type": "cols_req", "prompt": "Columns to extract"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "consolidate_sheets",
        "module": "consolidator", "fn": "merge_sheets_in_file",
        "desc": "Consolidate all sheets in a single file into one sheet",
        "keywords": ["merge sheets", "combine sheets", "consolidate sheets", "all sheets",
                     "stack sheets", "sheets into one", "multi sheet"],
        "anti": ["multiple files", "separate files"],
        "multi": False,
        "params": [
            {"name": "file_path",   "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "consolidate_same_sheet",
        "module": "consolidator", "fn": "merge_same_sheet_cross_files",
        "desc": "Extract and stack the same sheet from multiple files",
        "keywords": ["same sheet across files", "sheet from multiple files",
                     "extract same sheet", "pull sheet from files"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",       "type": "files"},
            {"name": "sheet_name",  "type": "string", "prompt": "Sheet name to extract",
             "default": "Sheet1"},
            {"name": "output_path", "type": "output"},
        ]
    },

    # ── CALCULATOR ───────────────────────────────────────────────────────────
    {
        "id": "calc_efficiency",
        "module": "calculator", "fn": "calculate_efficiency",
        "desc": "Efficiency = (Actual / Target) × 100",
        "keywords": ["efficiency", "actual vs target", "performance ratio",
                     "target achievement", "achievement rate"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "actual_col",  "type": "col_req", "prompt": "Actual column"},
            {"name": "target_col",  "type": "col_req", "prompt": "Target column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "calc_productivity",
        "module": "calculator", "fn": "calculate_productivity",
        "desc": "Productivity = Output / Input (e.g. units per hour)",
        "keywords": ["productivity", "output per input", "throughput", "units per hour"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_col",  "type": "col_req", "prompt": "Output/production column"},
            {"name": "input_col",   "type": "col_req", "prompt": "Input/resource column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "calc_utilization",
        "module": "calculator", "fn": "calculate_utilization",
        "desc": "Utilization = (Used / Available) × 100",
        "keywords": ["utilization", "utilisation", "capacity used", "resource used",
                     "used vs available"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",          "type": "file"},
            {"name": "used_col",      "type": "col_req", "prompt": "Used/actual column"},
            {"name": "available_col", "type": "col_req", "prompt": "Available/capacity column"},
            {"name": "output_path",   "type": "output"},
        ]
    },
    {
        "id": "calc_variance",
        "module": "calculator", "fn": "calculate_variance",
        "desc": "Variance = Actual − Budget",
        "keywords": ["variance", "actual vs budget", "budget variance", "over budget",
                     "budget difference", "actuals vs plan"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "actual_col",  "type": "col_req", "prompt": "Actual column"},
            {"name": "budget_col",  "type": "col_req", "prompt": "Budget/plan column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "calc_growth",
        "module": "calculator", "fn": "calculate_growth_rate",
        "desc": "Period-over-period growth rate",
        "keywords": ["growth rate", "period growth", "yoy", "year over year",
                     "growth percentage", "month over month"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "value_col",   "type": "col_req", "prompt": "Value column for growth"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "calc_stats",
        "module": "calculator", "fn": "calculate_summary_stats",
        "desc": "Descriptive statistics (count, sum, mean, median, std, min, max)",
        "keywords": ["summary stats", "statistics", "describe", "descriptive stats",
                     "mean median", "average std", "statistical summary"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "columns",     "type": "cols_opt", "prompt": "Columns (Enter=all numeric)"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "calc_pct_total",
        "module": "calculator", "fn": "calculate_percentage_of_total",
        "desc": "Each row's percentage share of the column total",
        "keywords": ["percentage of total", "pct share", "percent contribution",
                     "share of total", "% of total", "contribution percentage"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "value_col",   "type": "col_req", "prompt": "Value column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "calc_moving_avg",
        "module": "calculator", "fn": "calculate_moving_average",
        "desc": "Rolling/moving average for a numeric column",
        "keywords": ["moving average", "rolling average", "rolling mean",
                     "moving mean", "smoothing"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "value_col",   "type": "col_req", "prompt": "Value column"},
            {"name": "window",      "type": "number", "prompt": "Window size", "default": 7},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "calc_kpi",
        "module": "calculator", "fn": "calculate_kpi_dashboard",
        "desc": "KPI dashboard with key metrics for multiple columns",
        "keywords": ["kpi", "key performance indicator", "kpi dashboard",
                     "performance metrics", "kpis"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "kpi_columns",  "type": "cols_req", "prompt": "KPI columns to include"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "calc_weighted_avg",
        "module": "calculator", "fn": "calculate_weighted_average",
        "desc": "Weighted average = sum(value × weight) / sum(weight)",
        "keywords": ["weighted average", "weighted mean", "weighted avg",
                     "weight by", "weighted score"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "value_col",   "type": "col_req", "prompt": "Value column"},
            {"name": "weight_col",  "type": "col_req", "prompt": "Weight column"},
            {"name": "output_path", "type": "output"},
        ]
    },

    # ── TRANSFORMER ──────────────────────────────────────────────────────────
    {
        "id": "transform_pivot",
        "module": "transformer", "fn": "create_pivot_table",
        "desc": "Create a pivot table (group by + aggregate)",
        "keywords": ["pivot", "pivot table", "pivot by", "group by", "aggregate",
                     "summarize by", "rollup", "crosstab"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "index_cols",  "type": "cols_req", "prompt": "Row grouping columns"},
            {"name": "values_cols", "type": "cols_req", "prompt": "Value columns to aggregate"},
            {"name": "aggfunc",     "type": "choice", "prompt": "Aggregation function",
             "options": ["sum", "mean", "count", "max", "min"], "default": "sum"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "transform_unpivot",
        "module": "transformer", "fn": "unpivot_data",
        "desc": "Unpivot (melt) from wide to long format",
        "keywords": ["unpivot", "melt", "wide to long", "reshape wide", "normalize table"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "id_vars",     "type": "cols_req", "prompt": "ID columns to keep"},
            {"name": "value_name",  "type": "string", "prompt": "Value column name",
             "default": "Value"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "transform_transpose",
        "module": "transformer", "fn": "transpose_data",
        "desc": "Flip rows and columns (transpose)",
        "keywords": ["transpose", "flip rows columns", "rotate table", "swap rows cols"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "transform_split_by_value",
        "module": "transformer", "fn": "split_by_column_value",
        "desc": "Split file into separate files, one per unique value",
        "keywords": ["split by value", "split by column", "one file per",
                     "separate by", "divide by column"],
        "anti": ["sheets", "rows", "chunk"],
        "multi": False,
        "params": [
            {"name": "file",          "type": "file"},
            {"name": "split_column",  "type": "col_req", "prompt": "Column to split by"},
            {"name": "output_folder", "type": "output_dir"},
        ]
    },
    {
        "id": "transform_split_sheets",
        "module": "transformer", "fn": "split_sheets_to_files",
        "desc": "Extract each sheet into its own separate file",
        "keywords": ["split sheets to files", "one file per sheet", "extract sheets",
                     "sheets to separate files"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",       "type": "file"},
            {"name": "output_dir", "type": "output_dir"},
        ]
    },
    {
        "id": "transform_chunk",
        "module": "transformer", "fn": "split_file_by_rows",
        "desc": "Split a large file into smaller chunks by row count",
        "keywords": ["chunk", "split by rows", "break into chunks", "large file split",
                     "batch split", "chunk size"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",          "type": "file"},
            {"name": "chunk_size",    "type": "number", "prompt": "Rows per chunk",
             "default": 1000},
            {"name": "output_folder", "type": "output_dir"},
        ]
    },
    {
        "id": "transform_wide_to_long",
        "module": "transformer", "fn": "reshape_wide_to_long",
        "desc": "Reshape wide-format data to long format",
        "keywords": ["wide to long", "reshape wide", "normalize wide", "wide format to long"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "transform_long_to_wide",
        "module": "transformer", "fn": "reshape_long_to_wide",
        "desc": "Reshape long-format data to wide format",
        "keywords": ["long to wide", "reshape long", "wide pivot", "long format to wide"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "index_cols",   "type": "cols_req", "prompt": "ID/index columns"},
            {"name": "columns_col",  "type": "col_req",  "prompt": "Column whose values become headers"},
            {"name": "values_col",   "type": "col_req",  "prompt": "Values column"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "transform_running_total",
        "module": "transformer", "fn": "add_running_total",
        "desc": "Add a cumulative/running total column",
        "keywords": ["running total", "cumulative sum", "cumulative total", "running sum",
                     "cumulative", "running balance"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "value_col",   "type": "col_req", "prompt": "Column to accumulate"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "transform_rank",
        "module": "transformer", "fn": "rank_column",
        "desc": "Add a rank column sorted by a numeric value",
        "keywords": ["rank", "ranking", "add rank", "rank by", "top ranked",
                     "order by rank", "percentile rank"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "value_col",   "type": "col_req", "prompt": "Column to rank"},
            {"name": "output_path", "type": "output"},
        ]
    },

    # ── COMPARATOR ───────────────────────────────────────────────────────────
    {
        "id": "compare_files",
        "module": "comparator", "fn": "compare_two_files",
        "desc": "Full side-by-side comparison of two Excel files",
        "keywords": ["compare", "diff", "difference between files", "what changed",
                     "compare two files", "file comparison", "changes between"],
        "anti": [], "multi": True,
        "params": [
            {"name": "file1",       "type": "file1"},
            {"name": "file2",       "type": "file2"},
            {"name": "output_path", "type": "output"},
            {"name": "key_col",     "type": "col_opt", "prompt": "Key column (Enter=skip)"},
        ]
    },
    {
        "id": "compare_new_rows",
        "module": "comparator", "fn": "find_new_rows",
        "desc": "Find rows in file2 that don't exist in file1 (new additions)",
        "keywords": ["new rows", "added rows", "new records", "additions",
                     "what was added", "rows added"],
        "anti": [], "multi": True,
        "params": [
            {"name": "file1",       "type": "file1"},
            {"name": "file2",       "type": "file2"},
            {"name": "output_path", "type": "output"},
            {"name": "key_col",     "type": "col_opt", "prompt": "Key column (Enter=skip)"},
        ]
    },
    {
        "id": "compare_deleted_rows",
        "module": "comparator", "fn": "find_deleted_rows",
        "desc": "Find rows removed from file1 compared to file2",
        "keywords": ["deleted rows", "removed rows", "what was deleted",
                     "rows removed", "missing rows"],
        "anti": [], "multi": True,
        "params": [
            {"name": "file1",       "type": "file1"},
            {"name": "file2",       "type": "file2"},
            {"name": "output_path", "type": "output"},
            {"name": "key_col",     "type": "col_opt", "prompt": "Key column (Enter=skip)"},
        ]
    },
    {
        "id": "compare_changed",
        "module": "comparator", "fn": "find_changed_values",
        "desc": "Find rows with the same key but changed values",
        "keywords": ["changed values", "modified values", "updated values",
                     "what changed", "value changes", "modified rows"],
        "anti": [], "multi": True,
        "params": [
            {"name": "file1",        "type": "file1"},
            {"name": "file2",        "type": "file2"},
            {"name": "key_column",   "type": "col_req", "prompt": "Key column for matching"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "compare_find_dupes",
        "module": "comparator", "fn": "find_duplicates_in_file",
        "desc": "Find and report duplicate rows within a single file",
        "keywords": ["find duplicates", "detect duplicates", "flag duplicates",
                     "identify dupes", "which rows duplicate"],
        "anti": ["remove", "drop", "delete"],
        "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
            {"name": "subset",      "type": "cols_opt", "prompt": "Key columns (Enter=all)"},
        ]
    },
    {
        "id": "compare_common",
        "module": "comparator", "fn": "find_common_rows",
        "desc": "Find rows that exist in BOTH files (intersection)",
        "keywords": ["common rows", "intersection", "rows in both", "shared rows",
                     "in both files"],
        "anti": [], "multi": True,
        "params": [
            {"name": "file1",       "type": "file1"},
            {"name": "file2",       "type": "file2"},
            {"name": "output_path", "type": "output"},
            {"name": "key_col",     "type": "col_opt", "prompt": "Key column (Enter=skip)"},
        ]
    },
    {
        "id": "compare_cross_dupes",
        "module": "comparator", "fn": "cross_file_duplicate_check",
        "desc": "Find records that appear in more than one file",
        "keywords": ["cross file duplicates", "duplicates across files",
                     "same in multiple files", "cross file check"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",        "type": "files"},
            {"name": "key_columns",  "type": "cols_req", "prompt": "Key columns to match on"},
            {"name": "output_path",  "type": "output"},
        ]
    },

    # ── COLUMN OPS ───────────────────────────────────────────────────────────
    {
        "id": "col_rename",
        "module": "column_ops", "fn": "rename_columns",
        "desc": "Rename columns using old→new name mapping",
        "keywords": ["rename columns", "rename column", "column rename",
                     "change column name", "relabel columns"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "rename_map",  "type": "mapping",  "prompt": "Enter renames (old:new, comma-sep)"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "col_merge",
        "module": "column_ops", "fn": "merge_columns",
        "desc": "Concatenate multiple columns into one new column",
        "keywords": ["merge columns", "concatenate columns", "join columns",
                     "combine columns", "concat cols"],
        "anti": ["files", "sheets"],
        "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "columns",     "type": "cols_req", "prompt": "Columns to merge"},
            {"name": "new_col",     "type": "string", "prompt": "New column name",
             "default": "Merged"},
            {"name": "separator",   "type": "string", "prompt": "Separator",
             "default": " "},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "col_split",
        "module": "column_ops", "fn": "split_column",
        "desc": "Split a column into multiple columns by delimiter",
        "keywords": ["split column", "column split", "delimit column",
                     "separate column", "parse column"],
        "anti": ["files", "file"],
        "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column to split"},
            {"name": "delimiter",   "type": "string", "prompt": "Delimiter",
             "default": ","},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "col_reorder",
        "module": "column_ops", "fn": "reorder_columns",
        "desc": "Reorder columns to a specified order",
        "keywords": ["reorder columns", "column order", "rearrange columns",
                     "reorganize columns", "move columns"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "column_order", "type": "cols_req", "prompt": "New column order (all cols)"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "col_drop",
        "module": "column_ops", "fn": "drop_columns",
        "desc": "Remove specified columns from the file",
        "keywords": ["drop columns", "delete columns", "remove columns",
                     "eliminate columns", "hide columns"],
        "anti": ["rows", "empty"],
        "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "columns",     "type": "cols_req", "prompt": "Columns to drop"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "col_calculated",
        "module": "column_ops", "fn": "add_calculated_column",
        "desc": "Add a new column computed from an expression",
        "keywords": ["add column", "calculated column", "formula column",
                     "derived column", "new column", "compute column"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "new_col",     "type": "string", "prompt": "New column name",
             "default": "Calculated"},
            {"name": "expression",  "type": "string", "prompt": "Expression (e.g. ColA + ColB)",
             "default": ""},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "col_regex_extract",
        "module": "column_ops", "fn": "extract_from_column",
        "desc": "Extract text from a column using a regex pattern",
        "keywords": ["regex extract", "extract from column", "extract pattern",
                     "parse column", "regex column"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column to extract from"},
            {"name": "pattern",     "type": "string",  "prompt": "Regex pattern"},
            {"name": "new_col",     "type": "string",  "prompt": "Output column name",
             "default": "Extracted"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "col_map_values",
        "module": "column_ops", "fn": "map_column_values",
        "desc": "Replace values in a column using a lookup mapping",
        "keywords": ["map values", "replace values", "value mapping",
                     "recode", "value lookup", "translate values"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column to remap"},
            {"name": "mapping",     "type": "mapping",  "prompt": "Enter mappings (old:new, comma-sep)"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "col_expand_multi",
        "module": "column_ops", "fn": "pivot_column_to_rows",
        "desc": "Expand a column with multi-value cells into multiple rows",
        "keywords": ["expand multi", "multi value column", "split cell values",
                     "one row per value", "unnest column"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column with multi-values"},
            {"name": "delimiter",   "type": "string",  "prompt": "Delimiter in cells",
             "default": ","},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "col_normalize_headers",
        "module": "column_ops", "fn": "normalize_column_names",
        "desc": "Standardize all column headers to a consistent format",
        "keywords": ["normalize headers", "normalize column names", "header format",
                     "clean headers", "snake case headers", "column names"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
            {"name": "style",       "type": "choice", "prompt": "Naming style",
             "options": ["snake_case", "camel_case", "title_case", "upper", "lower"],
             "default": "snake_case"},
        ]
    },

    # ── REPORTER ─────────────────────────────────────────────────────────────
    {
        "id": "report_summary",
        "module": "reporter", "fn": "generate_summary_report",
        "desc": "Summary statistics report across multiple files",
        "keywords": ["summary report", "summary statistics", "overview report",
                     "generate report", "file summary"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",       "type": "files"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "report_profile",
        "module": "reporter", "fn": "data_profile",
        "desc": "Detailed column-by-column data profiling report",
        "keywords": ["profile", "data profile", "column profile", "data analysis",
                     "data overview", "analyze file", "inspect data"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "report_kpi",
        "module": "reporter", "fn": "generate_kpi_report",
        "desc": "Formatted KPI report with key metrics",
        "keywords": ["kpi report", "key metrics report", "performance report",
                     "metrics dashboard"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "kpi_columns",  "type": "cols_req", "prompt": "KPI columns"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "report_top_n",
        "module": "reporter", "fn": "top_n_report",
        "desc": "Top-N and Bottom-N report sorted by a numeric column",
        "keywords": ["top n", "top 10", "top performers", "bottom n", "bottom 10",
                     "best performers", "worst performers", "top records"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "sort_column", "type": "col_req", "prompt": "Column to sort by"},
            {"name": "n",           "type": "number",  "prompt": "Number of rows", "default": 10},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "report_frequency",
        "module": "reporter", "fn": "frequency_report",
        "desc": "Value frequency count (like a pivot count) for categorical columns",
        "keywords": ["frequency report", "frequency count", "count by", "value counts",
                     "count occurrences", "category count"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "columns",     "type": "cols_req", "prompt": "Columns to count"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "report_monthly",
        "module": "reporter", "fn": "monthly_summary_report",
        "desc": "Aggregate data by month from a date column",
        "keywords": ["monthly summary", "by month", "monthly report",
                     "month by month", "monthly aggregation"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "date_column",  "type": "col_req",  "prompt": "Date column"},
            {"name": "value_cols",   "type": "cols_req", "prompt": "Value columns to aggregate"},
            {"name": "output_path",  "type": "output"},
        ]
    },

    # ── FINANCE ──────────────────────────────────────────────────────────────
    {
        "id": "finance_aging",
        "module": "finance", "fn": "aging_analysis",
        "desc": "AR/AP Aging analysis — overdue buckets (0-30, 31-60, 61-90, 90+ days)",
        "keywords": ["aging", "ar aging", "ap aging", "overdue", "outstanding invoices",
                     "accounts receivable", "accounts payable", "aging report"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "date_col",    "type": "col_req", "prompt": "Invoice/due date column"},
            {"name": "amount_col",  "type": "col_req", "prompt": "Amount column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "finance_amortization",
        "module": "finance", "fn": "loan_amortization",
        "desc": "Full EMI amortization schedule",
        "keywords": ["amortization", "emi", "loan schedule", "loan amortization",
                     "loan payment", "mortgage schedule"],
        "anti": [], "multi": False,
        "params": [
            {"name": "principal",      "type": "float_val", "prompt": "Principal amount",
             "default": 100000.0},
            {"name": "annual_rate",    "type": "float_val", "prompt": "Annual interest rate (%)",
             "default": 10.0},
            {"name": "tenure_months",  "type": "number",    "prompt": "Tenure in months",
             "default": 12},
            {"name": "output_path",    "type": "output"},
        ]
    },
    {
        "id": "finance_depreciation",
        "module": "finance", "fn": "depreciation_schedule",
        "desc": "Straight-line and declining balance depreciation schedule",
        "keywords": ["depreciation", "depreciation schedule", "asset depreciation",
                     "straight line", "declining balance"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "finance_ratios",
        "module": "finance", "fn": "financial_ratios",
        "desc": "Compute common financial ratios per row",
        "keywords": ["financial ratios", "ratios", "liquidity ratio",
                     "profitability ratio", "finance ratios"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "finance_payroll",
        "module": "finance", "fn": "payroll_calculator",
        "desc": "Payroll: Gross → Net after deductions",
        "keywords": ["payroll", "payroll calculator", "salary calculation",
                     "net salary", "gross to net", "payroll processing"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "basic_col",   "type": "col_req", "prompt": "Basic salary column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "finance_budget_vs_actual",
        "module": "finance", "fn": "budget_vs_actual",
        "desc": "Merge budget and actual files to show variance",
        "keywords": ["budget vs actual", "bva", "budget actual comparison",
                     "actuals vs budget", "plan vs actual"],
        "anti": [], "multi": True,
        "params": [
            {"name": "budget_file",  "type": "file1"},
            {"name": "actual_file",  "type": "file2"},
            {"name": "key_col",      "type": "col_req", "prompt": "Key column to join on"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "finance_compound",
        "module": "finance", "fn": "compound_interest_schedule",
        "desc": "Future value compounding growth table",
        "keywords": ["compound interest", "compounding", "future value",
                     "investment growth", "fv calculation"],
        "anti": [], "multi": False,
        "params": [
            {"name": "principal",   "type": "float_val", "prompt": "Principal amount",
             "default": 100000.0},
            {"name": "annual_rate", "type": "float_val", "prompt": "Annual rate (%)",
             "default": 10.0},
            {"name": "years",       "type": "number",    "prompt": "Number of years",
             "default": 5},
            {"name": "output_path", "type": "output"},
        ]
    },

    # ── HR ────────────────────────────────────────────────────────────────────
    {
        "id": "hr_attrition",
        "module": "hr", "fn": "attrition_analysis",
        "desc": "Attrition/turnover rate by department",
        "keywords": ["attrition", "turnover", "employee attrition", "turnover rate",
                     "attrition analysis", "churn rate", "resignation rate"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "status_col",  "type": "col_req", "prompt": "Employment status column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "hr_headcount",
        "module": "hr", "fn": "headcount_summary",
        "desc": "Headcount count and % share grouped by department/location",
        "keywords": ["headcount", "employee count", "staff count", "workforce",
                     "headcount by", "team size", "headcount report"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "group_cols",  "type": "cols_req", "prompt": "Group by columns (e.g. Dept, Location)"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "hr_tenure",
        "module": "hr", "fn": "tenure_analysis",
        "desc": "Years-of-service distribution (tenure bands)",
        "keywords": ["tenure", "years of service", "service duration", "tenure analysis",
                     "how long employees", "seniority"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "join_date_col","type": "col_req", "prompt": "Joining date column"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "hr_age",
        "module": "hr", "fn": "age_band_analysis",
        "desc": "Workforce age demographics by 10-year bands",
        "keywords": ["age analysis", "age bands", "age demographics",
                     "age distribution", "workforce age"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "dob_col",     "type": "col_req", "prompt": "Date of birth column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "hr_salary",
        "module": "hr", "fn": "salary_analysis",
        "desc": "Salary statistics by department (min, max, mean, median, P25, P75)",
        "keywords": ["salary analysis", "salary statistics", "pay analysis",
                     "compensation analysis", "salary by department"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "salary_col",  "type": "col_req", "prompt": "Salary column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "hr_performance",
        "module": "hr", "fn": "performance_distribution",
        "desc": "Performance rating distribution with count and % by band",
        "keywords": ["performance distribution", "performance review", "rating distribution",
                     "performance bands", "appraisal distribution"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "rating_col",  "type": "col_req", "prompt": "Performance rating column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "hr_increment",
        "module": "hr", "fn": "salary_increment_calculator",
        "desc": "Apply salary increment and compute new salary",
        "keywords": ["salary increment", "increment", "raise", "salary hike",
                     "pay raise", "salary increase"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "salary_col",  "type": "col_req", "prompt": "Current salary column"},
            {"name": "output_path", "type": "output"},
        ]
    },

    # ── SALES ─────────────────────────────────────────────────────────────────
    {
        "id": "sales_commission",
        "module": "sales", "fn": "commission_calculator",
        "desc": "Calculate sales commission by rep/territory",
        "keywords": ["commission", "sales commission", "incentive", "commission calculator",
                     "sales incentive"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "sales_col",   "type": "col_req", "prompt": "Sales amount column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "sales_rfm",
        "module": "sales", "fn": "rfm_segmentation",
        "desc": "RFM (Recency, Frequency, Monetary) customer segmentation",
        "keywords": ["rfm", "rfm analysis", "customer segments", "recency frequency monetary",
                     "customer segmentation", "customer value"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "customer_col", "type": "col_req", "prompt": "Customer ID column"},
            {"name": "date_col",     "type": "col_req", "prompt": "Transaction date column"},
            {"name": "amount_col",   "type": "col_req", "prompt": "Purchase amount column"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "sales_quota",
        "module": "sales", "fn": "quota_attainment",
        "desc": "Quota attainment: actual vs quota with Above/Near/Below labels",
        "keywords": ["quota", "quota attainment", "target achievement", "sales vs quota",
                     "quota performance"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "actual_col",  "type": "col_req", "prompt": "Actual sales column"},
            {"name": "quota_col",   "type": "col_req", "prompt": "Quota/target column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "sales_pipeline",
        "module": "sales", "fn": "pipeline_analysis",
        "desc": "Sales pipeline funnel: count and value by stage",
        "keywords": ["pipeline", "sales pipeline", "funnel", "deal stages",
                     "opportunity pipeline", "crm pipeline"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "stage_col",   "type": "col_req", "prompt": "Pipeline stage column"},
            {"name": "value_col",   "type": "col_req", "prompt": "Deal value column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "sales_territory",
        "module": "sales", "fn": "sales_by_territory",
        "desc": "Territory-level sales summary with ranking",
        "keywords": ["territory", "sales by territory", "regional sales",
                     "territory performance", "region breakdown"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",          "type": "file"},
            {"name": "territory_col", "type": "col_req", "prompt": "Territory column"},
            {"name": "sales_col",     "type": "col_req", "prompt": "Sales amount column"},
            {"name": "output_path",   "type": "output"},
        ]
    },
    {
        "id": "sales_customer_abc",
        "module": "sales", "fn": "customer_abc",
        "desc": "A/B/C customer classification by cumulative revenue",
        "keywords": ["customer abc", "abc customer", "customer classification",
                     "top customers", "customer analysis"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "customer_col", "type": "col_req", "prompt": "Customer column"},
            {"name": "revenue_col",  "type": "col_req", "prompt": "Revenue column"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "sales_discount",
        "module": "sales", "fn": "discount_analysis",
        "desc": "Discount % per row and total revenue leakage analysis",
        "keywords": ["discount", "discount analysis", "revenue leakage",
                     "discount report", "price discount"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",              "type": "file"},
            {"name": "list_price_col",    "type": "col_req", "prompt": "List/MRP price column"},
            {"name": "actual_price_col",  "type": "col_req", "prompt": "Actual/sold price column"},
            {"name": "output_path",       "type": "output"},
        ]
    },

    # ── INVENTORY ─────────────────────────────────────────────────────────────
    {
        "id": "inv_abc",
        "module": "inventory", "fn": "abc_analysis",
        "desc": "ABC inventory classification by cumulative value %",
        "keywords": ["abc analysis", "inventory abc", "abc classification",
                     "inventory categorize"],
        "anti": ["customer"], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "item_col",    "type": "col_req", "prompt": "Item/SKU column"},
            {"name": "value_col",   "type": "col_req", "prompt": "Value/cost column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "inv_reorder",
        "module": "inventory", "fn": "reorder_point",
        "desc": "Reorder Point = (Avg Daily Usage × Lead Time) + Safety Stock",
        "keywords": ["reorder point", "reorder", "stock replenishment",
                     "when to order", "replenish stock"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "inv_stock_aging",
        "module": "inventory", "fn": "stock_aging",
        "desc": "Stock age analysis: bucket inventory into aging bands",
        "keywords": ["stock aging", "inventory aging", "old stock",
                     "stock age", "goods age"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",              "type": "file"},
            {"name": "receipt_date_col",  "type": "col_req", "prompt": "Receipt/purchase date column"},
            {"name": "output_path",       "type": "output"},
        ]
    },
    {
        "id": "inv_turnover",
        "module": "inventory", "fn": "inventory_turnover",
        "desc": "Inventory Turnover = COGS / Average Inventory",
        "keywords": ["inventory turnover", "stock turnover", "turnover ratio",
                     "cogs inventory", "inventory efficiency"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "cogs_col",    "type": "col_req", "prompt": "COGS column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "inv_oee",
        "module": "inventory", "fn": "oee_calculator",
        "desc": "OEE = Availability × Performance × Quality",
        "keywords": ["oee", "overall equipment effectiveness", "machine efficiency",
                     "equipment oee"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "inv_dead_stock",
        "module": "inventory", "fn": "dead_stock_analysis",
        "desc": "Identify items with no movement (dead stock / obsolete)",
        "keywords": ["dead stock", "obsolete inventory", "slow moving",
                     "no movement", "stale inventory"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",                "type": "file"},
            {"name": "last_movement_col",   "type": "col_req",
             "prompt": "Last movement/transaction date column"},
            {"name": "output_path",         "type": "output"},
        ]
    },

    # ── FORMATTER ─────────────────────────────────────────────────────────────
    {
        "id": "fmt_bar_chart",
        "module": "formatter", "fn": "add_bar_chart",
        "desc": "Add an embedded bar chart to the worksheet",
        "keywords": ["bar chart", "add bar chart", "bar graph", "column chart"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "x_col",       "type": "col_req", "prompt": "X-axis (category) column"},
            {"name": "y_col",       "type": "col_req", "prompt": "Y-axis (value) column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_line_chart",
        "module": "formatter", "fn": "add_line_chart",
        "desc": "Add an embedded line chart to the worksheet",
        "keywords": ["line chart", "add line chart", "line graph", "trend chart"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "x_col",       "type": "col_req", "prompt": "X-axis column"},
            {"name": "y_col",       "type": "col_req", "prompt": "Y-axis (value) column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_pie_chart",
        "module": "formatter", "fn": "add_pie_chart",
        "desc": "Add an embedded pie chart to the worksheet",
        "keywords": ["pie chart", "add pie chart", "donut chart", "pie graph"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "category_col", "type": "col_req", "prompt": "Category column"},
            {"name": "value_col",    "type": "col_req", "prompt": "Value column"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "fmt_traffic_light",
        "module": "formatter", "fn": "apply_traffic_light",
        "desc": "Apply Red/Yellow/Green colour fills to a numeric column",
        "keywords": ["traffic light", "rag status", "red yellow green",
                     "rag coloring", "traffic light format"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column to colour"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_color_scale",
        "module": "formatter", "fn": "apply_color_scale",
        "desc": "Gradient colour scale (white→blue) from min to max",
        "keywords": ["color scale", "colour scale", "gradient color",
                     "heat map color", "conditional color"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column to colour"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_table",
        "module": "formatter", "fn": "format_as_table",
        "desc": "Apply Excel table style to the used range",
        "keywords": ["format table", "excel table", "table style", "apply table",
                     "format as table"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_freeze",
        "module": "formatter", "fn": "freeze_and_filter",
        "desc": "Freeze the header row and enable auto-filter",
        "keywords": ["freeze header", "freeze row", "auto filter", "freeze pane",
                     "lock header", "filter headers"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_auto_fit",
        "module": "formatter", "fn": "auto_fit_columns",
        "desc": "Auto-fit column widths to content",
        "keywords": ["auto fit", "column width", "adjust columns", "autofit",
                     "fit columns"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_totals",
        "module": "formatter", "fn": "add_totals_row",
        "desc": "Add a SUM totals row at the bottom for numeric columns",
        "keywords": ["totals row", "sum row", "add totals", "total row",
                     "grand total row"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_highlight_dupes",
        "module": "formatter", "fn": "highlight_duplicates",
        "desc": "Highlight duplicate values in a column with yellow fill",
        "keywords": ["highlight duplicates", "highlight dupes", "colour duplicates",
                     "mark duplicates"],
        "anti": ["remove", "drop", "delete"], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column to check"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "fmt_number_format",
        "module": "formatter", "fn": "apply_number_format",
        "desc": "Apply number format (currency, percentage, comma) to columns",
        "keywords": ["number format", "currency format", "percentage format",
                     "format numbers", "decimal format"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",          "type": "file"},
            {"name": "columns",       "type": "cols_req", "prompt": "Columns to format"},
            {"name": "format_string", "type": "string",   "prompt": "Format string",
             "default": "#,##0.00"},
            {"name": "output_path",   "type": "output"},
        ]
    },

    # ── VALIDATOR ─────────────────────────────────────────────────────────────
    {
        "id": "val_mandatory",
        "module": "validator", "fn": "check_mandatory_fields",
        "desc": "Flag rows with missing values in required columns",
        "keywords": ["mandatory fields", "required fields", "missing mandatory",
                     "check required", "null check", "completeness"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",          "type": "file"},
            {"name": "required_cols", "type": "cols_req", "prompt": "Required columns"},
            {"name": "output_path",   "type": "output"},
        ]
    },
    {
        "id": "val_email",
        "module": "validator", "fn": "validate_email",
        "desc": "Validate email addresses with regex",
        "keywords": ["validate email", "email validation", "check email",
                     "invalid email", "email format"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "email_col",   "type": "col_req", "prompt": "Email column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "val_phone",
        "module": "validator", "fn": "validate_phone",
        "desc": "Validate phone number formats",
        "keywords": ["validate phone", "phone validation", "check phone",
                     "phone number format"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "phone_col",   "type": "col_req", "prompt": "Phone column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "val_range",
        "module": "validator", "fn": "validate_numeric_range",
        "desc": "Flag rows where a numeric column is outside [min, max]",
        "keywords": ["numeric range", "out of range", "value range check",
                     "min max check", "range validation"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req",  "prompt": "Column to check"},
            {"name": "min_val",     "type": "float_val","prompt": "Minimum value",
             "default": 0.0},
            {"name": "max_val",     "type": "float_val","prompt": "Maximum value",
             "default": 1000000.0},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "val_date_range",
        "module": "validator", "fn": "validate_date_range",
        "desc": "Flag rows where a date column is outside a given range",
        "keywords": ["date range", "date validation", "invalid dates",
                     "date out of range"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Date column"},
            {"name": "min_date",    "type": "string",  "prompt": "Min date (YYYY-MM-DD)",
             "default": "2000-01-01"},
            {"name": "max_date",    "type": "string",  "prompt": "Max date (YYYY-MM-DD)",
             "default": "2030-12-31"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "val_referential",
        "module": "validator", "fn": "referential_integrity",
        "desc": "Check that values in a column exist in a reference file",
        "keywords": ["referential integrity", "foreign key", "reference check",
                     "lookup validation", "valid values"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "col",         "type": "col_req",  "prompt": "Column to validate"},
            {"name": "ref_file",    "type": "ref_file"},
            {"name": "ref_col",     "type": "string",   "prompt": "Reference column name"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "val_quality",
        "module": "validator", "fn": "data_quality_report",
        "desc": "Comprehensive data quality report with score 0-100",
        "keywords": ["data quality", "quality report", "quality score",
                     "dq report", "data health"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "val_pii",
        "module": "validator", "fn": "detect_pii",
        "desc": "Detect columns likely containing PII (email, phone, SSN, names)",
        "keywords": ["pii", "personally identifiable", "sensitive data",
                     "detect pii", "privacy check", "gdpr check"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },

    # ── ANALYTICS ─────────────────────────────────────────────────────────────
    {
        "id": "analytics_correlation",
        "module": "analytics", "fn": "correlation_matrix",
        "desc": "Pairwise Pearson correlation matrix for numeric columns",
        "keywords": ["correlation", "correlation matrix", "pearson",
                     "related columns", "correlation analysis"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "columns",     "type": "cols_opt", "prompt": "Columns (Enter=all numeric)"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "analytics_pareto",
        "module": "analytics", "fn": "pareto_analysis",
        "desc": "Pareto (80/20) analysis",
        "keywords": ["pareto", "80 20", "pareto analysis", "pareto chart",
                     "80 percent 20 percent"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "category_col", "type": "col_req", "prompt": "Category column"},
            {"name": "value_col",    "type": "col_req", "prompt": "Value column"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "analytics_regression",
        "module": "analytics", "fn": "linear_regression",
        "desc": "Simple OLS linear regression",
        "keywords": ["regression", "linear regression", "ols",
                     "linear model", "fit line"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "x_col",       "type": "col_req", "prompt": "Independent variable (X)"},
            {"name": "y_col",       "type": "col_req", "prompt": "Dependent variable (Y)"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "analytics_forecast",
        "module": "analytics", "fn": "trend_forecast",
        "desc": "Linear trend forecast N periods ahead",
        "keywords": ["forecast", "trend forecast", "predict future",
                     "time series forecast", "projection"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "date_col",    "type": "col_req", "prompt": "Date column"},
            {"name": "value_col",   "type": "col_req", "prompt": "Value column to forecast"},
            {"name": "output_path", "type": "output"},
            {"name": "periods",     "type": "number",  "prompt": "Periods to forecast ahead",
             "default": 12},
        ]
    },
    {
        "id": "analytics_freq_dist",
        "module": "analytics", "fn": "frequency_distribution",
        "desc": "Histogram / frequency distribution for a numeric column",
        "keywords": ["frequency distribution", "histogram", "distribution",
                     "bin data", "value distribution"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column to analyse"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "analytics_zscore",
        "module": "analytics", "fn": "z_score_analysis",
        "desc": "Z-score standardization for a numeric column",
        "keywords": ["z score", "z-score", "standard score", "standardize",
                     "zscore analysis", "standard deviation score"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "column",      "type": "col_req", "prompt": "Column to analyse"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "analytics_cohort",
        "module": "analytics", "fn": "cohort_retention",
        "desc": "Monthly cohort retention matrix",
        "keywords": ["cohort", "cohort retention", "retention matrix",
                     "customer retention", "cohort analysis"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "customer_col", "type": "col_req", "prompt": "Customer ID column"},
            {"name": "date_col",     "type": "col_req", "prompt": "Transaction date column"},
            {"name": "output_path",  "type": "output"},
        ]
    },

    # ── CONVERTER ─────────────────────────────────────────────────────────────
    {
        "id": "conv_to_csv",
        "module": "converter", "fn": "excel_to_csv",
        "desc": "Export every sheet in an Excel file to separate CSVs",
        "keywords": ["excel to csv", "convert to csv", "export csv",
                     "save as csv", "to csv"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",       "type": "file"},
            {"name": "output_dir", "type": "output_dir"},
        ]
    },
    {
        "id": "conv_from_csv",
        "module": "converter", "fn": "csv_to_excel",
        "desc": "Merge multiple CSV files into one Excel workbook",
        "keywords": ["csv to excel", "import csv", "csv to xlsx",
                     "combine csv to excel"],
        "anti": [], "multi": True,
        "params": [
            {"name": "csv_files",   "type": "files"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "conv_to_json",
        "module": "converter", "fn": "excel_to_json",
        "desc": "Export all sheets from Excel to a JSON file",
        "keywords": ["excel to json", "convert to json", "export json",
                     "to json"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "output_path", "type": "output_json"},
        ]
    },
    {
        "id": "conv_from_json",
        "module": "converter", "fn": "json_to_excel",
        "desc": "Convert a JSON file to Excel",
        "keywords": ["json to excel", "import json", "json to xlsx",
                     "from json"],
        "anti": [], "multi": False,
        "params": [
            {"name": "json_file",   "type": "file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "conv_xls_to_xlsx",
        "module": "converter", "fn": "xls_to_xlsx_batch",
        "desc": "Batch convert old .xls files to modern .xlsx format",
        "keywords": ["xls to xlsx", "upgrade excel", "convert xls",
                     "old excel format", ".xls to .xlsx"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",      "type": "files"},
            {"name": "output_dir", "type": "output_dir"},
        ]
    },
    {
        "id": "conv_to_text",
        "module": "converter", "fn": "excel_to_text",
        "desc": "Export each sheet to a delimited text file",
        "keywords": ["excel to text", "export text", "to text file",
                     "text delimited"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",       "type": "file"},
            {"name": "output_dir", "type": "output_dir"},
        ]
    },
    {
        "id": "conv_merge_csv",
        "module": "converter", "fn": "merge_csv_files",
        "desc": "Stack multiple CSV files into one Excel sheet",
        "keywords": ["merge csv", "combine csv", "stack csv",
                     "consolidate csv"],
        "anti": [], "multi": True,
        "params": [
            {"name": "csv_files",   "type": "files"},
            {"name": "output_path", "type": "output"},
        ]
    },

    # ── LOOKUP ────────────────────────────────────────────────────────────────
    {
        "id": "lookup_vlookup",
        "module": "lookup", "fn": "vlookup",
        "desc": "VLOOKUP: bring columns from a reference file",
        "keywords": ["vlookup", "lookup", "match values", "bring columns",
                     "lookup join", "reference join"],
        "anti": ["fuzzy"], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "lookup_col",  "type": "col_req",  "prompt": "Lookup column (in main file)"},
            {"name": "ref_file",    "type": "ref_file"},
            {"name": "ref_col",     "type": "string",   "prompt": "Matching column name in reference file"},
            {"name": "return_cols", "type": "cols_req", "prompt": "Columns to return from reference"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "lookup_fuzzy",
        "module": "lookup", "fn": "fuzzy_match",
        "desc": "Fuzzy string matching (approximate match, handles typos)",
        "keywords": ["fuzzy match", "approximate match", "fuzzy lookup",
                     "typo match", "similar match", "near match"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "col",         "type": "col_req",  "prompt": "Column to match from"},
            {"name": "ref_file",    "type": "ref_file"},
            {"name": "ref_col",     "type": "string",   "prompt": "Column to match against in reference"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "lookup_multi_key",
        "module": "lookup", "fn": "multi_key_lookup",
        "desc": "Multi-column JOIN on all key columns simultaneously",
        "keywords": ["multi key lookup", "multi column join", "compound key",
                     "join on multiple keys", "multi key"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "lookup_cols", "type": "cols_req", "prompt": "Key columns (must exist in both files)"},
            {"name": "ref_file",    "type": "ref_file"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "lookup_reverse",
        "module": "lookup", "fn": "reverse_lookup",
        "desc": "Given a value, find its corresponding key",
        "keywords": ["reverse lookup", "find key from value", "backward lookup"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "value_col",    "type": "col_req", "prompt": "Values column"},
            {"name": "key_col",      "type": "col_req", "prompt": "Keys column"},
            {"name": "lookup_value", "type": "string",  "prompt": "Value to look up"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "lookup_enrich",
        "module": "lookup", "fn": "enrich_from_lookup",
        "desc": "Enrich main file with columns from a reference file",
        "keywords": ["enrich", "add columns from reference", "enrich data",
                     "augment data", "bring in extra columns"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "join_col",     "type": "col_req", "prompt": "Join column (main file)"},
            {"name": "ref_file",     "type": "ref_file"},
            {"name": "ref_join_col", "type": "string",  "prompt": "Join column in reference file"},
            {"name": "output_path",  "type": "output"},
        ]
    },

    # ── PROJECT MANAGEMENT ────────────────────────────────────────────────────
    {
        "id": "pm_team_consolidate",
        "module": "project_mgmt", "fn": "team_consolidator",
        "desc": "Merge team member data from multiple files (dedup by ID)",
        "keywords": ["team consolidate", "merge team data", "team members merge",
                     "consolidate team"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",       "type": "files"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "pm_split_team",
        "module": "project_mgmt", "fn": "split_by_team",
        "desc": "Split a master sheet into one file per department/team",
        "keywords": ["split by team", "split by department", "one file per team",
                     "separate by department"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",       "type": "file"},
            {"name": "split_col",  "type": "col_req",  "prompt": "Team/department column"},
            {"name": "output_dir", "type": "output_dir"},
        ]
    },
    {
        "id": "pm_timesheet",
        "module": "project_mgmt", "fn": "timesheet_rollup",
        "desc": "Consolidate timesheet files → Detail + By Person + By Project + pivot",
        "keywords": ["timesheet", "timesheet rollup", "hours consolidate",
                     "time tracking", "working hours", "quarterly sheets",
                     "consolidate hours", "timesheet consolidation"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",       "type": "files"},
            {"name": "person_col",  "type": "col_req", "prompt": "Person/name column"},
            {"name": "project_col", "type": "col_req", "prompt": "Project column"},
            {"name": "hours_col",   "type": "col_req", "prompt": "Hours column"},
            {"name": "date_col",    "type": "col_req", "prompt": "Date column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "pm_resource_alloc",
        "module": "project_mgmt", "fn": "resource_allocation",
        "desc": "Resource allocation % per resource per project with over-allocation flag",
        "keywords": ["resource allocation", "allocation analysis", "resource planning",
                     "over allocation", "resource utilization"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",           "type": "file"},
            {"name": "resource_col",   "type": "col_req", "prompt": "Resource/person column"},
            {"name": "project_col",    "type": "col_req", "prompt": "Project column"},
            {"name": "allocation_col", "type": "col_req", "prompt": "Allocation % column"},
            {"name": "output_path",    "type": "output"},
        ]
    },
    {
        "id": "pm_milestone",
        "module": "project_mgmt", "fn": "milestone_tracker",
        "desc": "Milestone tracker with RAG status, slippage, and owner summary",
        "keywords": ["milestone", "milestone tracker", "delivery tracker",
                     "deadline tracker", "project milestones", "slippage"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",         "type": "file"},
            {"name": "task_col",     "type": "col_req", "prompt": "Task/milestone column"},
            {"name": "owner_col",    "type": "col_req", "prompt": "Owner column"},
            {"name": "planned_col",  "type": "col_req", "prompt": "Planned date column"},
            {"name": "actual_col",   "type": "col_req", "prompt": "Actual date column"},
            {"name": "output_path",  "type": "output"},
        ]
    },
    {
        "id": "pm_raci",
        "module": "project_mgmt", "fn": "raci_matrix",
        "desc": "Build and validate a RACI matrix (Responsible/Accountable/Consulted/Informed)",
        "keywords": ["raci", "raci matrix", "responsibility matrix",
                     "responsible accountable", "raci chart"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",      "type": "file"},
            {"name": "task_col",  "type": "col_req",  "prompt": "Task column"},
            {"name": "role_cols", "type": "cols_req", "prompt": "Role columns (R, A, C, I)"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "pm_risk",
        "module": "project_mgmt", "fn": "risk_register",
        "desc": "Risk register: Prob×Impact scoring (1-5) with heat map",
        "keywords": ["risk register", "risk analysis", "risk assessment",
                     "probability impact", "risk scoring", "risk matrix"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",       "type": "file"},
            {"name": "desc_col",   "type": "col_req", "prompt": "Risk description column"},
            {"name": "prob_col",   "type": "col_req", "prompt": "Probability column (1-5)"},
            {"name": "impact_col", "type": "col_req", "prompt": "Impact column (1-5)"},
            {"name": "output_path","type": "output"},
        ]
    },
    {
        "id": "pm_actions",
        "module": "project_mgmt", "fn": "action_tracker",
        "desc": "Consolidate meeting actions with Days_Overdue and priority",
        "keywords": ["action tracker", "action items", "meeting actions",
                     "follow up", "action log"],
        "anti": [], "multi": True,
        "params": [
            {"name": "files",       "type": "files"},
            {"name": "action_col",  "type": "col_req", "prompt": "Action description column"},
            {"name": "owner_col",   "type": "col_req", "prompt": "Owner column"},
            {"name": "due_date_col","type": "col_req", "prompt": "Due date column"},
            {"name": "output_path", "type": "output"},
        ]
    },
    {
        "id": "pm_capacity",
        "module": "project_mgmt", "fn": "capacity_planner",
        "desc": "Capacity vs demand: available vs allocated hours, utilisation %",
        "keywords": ["capacity", "capacity planning", "capacity planner",
                     "available hours", "capacity vs demand"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",                  "type": "file"},
            {"name": "resource_col",          "type": "col_req", "prompt": "Resource column"},
            {"name": "role_col",              "type": "col_req", "prompt": "Role column"},
            {"name": "available_hours_col",   "type": "col_req", "prompt": "Available hours column"},
            {"name": "allocated_hours_col",   "type": "col_req", "prompt": "Allocated hours column"},
            {"name": "output_path",           "type": "output"},
        ]
    },
    {
        "id": "pm_sprint",
        "module": "project_mgmt", "fn": "sprint_tracker",
        "desc": "Sprint tracker: velocity, completion %, backlog health",
        "keywords": ["sprint", "sprint tracker", "velocity", "backlog",
                     "agile sprint", "scrum", "story points"],
        "anti": [], "multi": False,
        "params": [
            {"name": "file",        "type": "file"},
            {"name": "story_col",   "type": "col_req", "prompt": "Story/task column"},
            {"name": "points_col",  "type": "col_req", "prompt": "Story points column"},
            {"name": "sprint_col",  "type": "col_req", "prompt": "Sprint column"},
            {"name": "status_col",  "type": "col_req", "prompt": "Status column"},
            {"name": "output_path", "type": "output"},
        ]
    },
]

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def extract_number_from_query(query: str, default: Optional[int] = None) -> Optional[int]:
    """Extract first integer from query string. 'top 10' → 10."""
    # Numeric words first
    q_lower = query.lower()
    for word, num in QUANTITY_WORDS.items():
        if re.search(r'\b' + word + r'\b', q_lower):
            return num
    # Digit patterns
    match = re.search(r'\b(\d+)\b', query)
    if match:
        return int(match.group(1))
    return default


def get_columns_from_file(file_path: str) -> List[str]:
    """Return column names from first sheet of an Excel/CSV file."""
    try:
        p = Path(file_path)
        if p.suffix.lower() == ".csv":
            import pandas as pd
            df = pd.read_csv(file_path, nrows=0)
        else:
            import pandas as pd
            df = pd.read_excel(file_path, nrows=0)
        return [str(c) for c in df.columns if c is not None and str(c) != "nan"]
    except Exception:
        return []


def inspect_file(file_path: str) -> Dict[str, Any]:
    """
    Read file structure and return a FileInfo dict:
    {
      sheets, sheet_count, columns, row_counts,
      col_types, domain_hint, file_count
    }
    """
    info: Dict[str, Any] = {
        "sheets": [],
        "sheet_count": 1,
        "columns": {},
        "row_counts": {},
        "col_types": {},
        "domain_hint": None,
        "file_count": 1,
    }
    try:
        p = Path(file_path)
        if p.suffix.lower() == ".csv":
            import pandas as pd
            df = pd.read_csv(file_path, nrows=5)
            info["sheets"] = ["Sheet1"]
            info["sheet_count"] = 1
            info["columns"] = {"Sheet1": list(df.columns)}
            info["row_counts"] = {"Sheet1": len(df)}
        else:
            import pandas as pd, openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            info["sheets"] = wb.sheetnames
            info["sheet_count"] = len(wb.sheetnames)
            for sheet in wb.sheetnames[:5]:  # cap at 5 sheets
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet, nrows=3)
                    info["columns"][sheet] = list(df.columns)
                    ws = wb[sheet]
                    info["row_counts"][sheet] = ws.max_row - 1
                except Exception:
                    info["columns"][sheet] = []
            wb.close()

        # Domain detection from column names
        all_cols_lower = " ".join(
            str(c).lower()
            for cols in info["columns"].values()
            for c in cols
            if c is not None
        )
        for domain, hints in DOMAIN_HINTS.items():
            if any(h in all_cols_lower for h in hints):
                info["domain_hint"] = domain
                break

    except Exception:
        pass
    return info


def extract_context(query: str) -> Dict[str, Any]:
    """
    Parse query for structural and intent signals.
    Returns ContextClues dict.
    """
    q_lower = query.lower()
    context: Dict[str, Any] = {
        "quantities": {},
        "time_refs": [],
        "domain_hints": [],
        "goal_hints": [],
        "multi_file": False,
        "numbers": [],
    }

    # Extract quantities
    for word, num in QUANTITY_WORDS.items():
        if re.search(r'\b' + word + r'\b', q_lower):
            context["quantities"][word] = num

    # Extract raw numbers
    context["numbers"] = [int(n) for n in re.findall(r'\b(\d+)\b', query)]

    # Time references
    time_words = ["quarter", "q1", "q2", "q3", "q4", "monthly", "weekly",
                  "annual", "yearly", "period", "season"]
    context["time_refs"] = [w for w in time_words if w in q_lower]

    # Domain hints from query text
    domain_kws = {
        "hr_project": ["team members", "team member", "employee", "staff",
                       "person", "people"],
        "finance": ["invoice", "payment", "budget", "cost", "expense"],
        "sales": ["customer", "sale", "revenue", "deal", "prospect"],
        "inventory": ["stock", "item", "sku", "warehouse", "goods"],
    }
    for domain, kws in domain_kws.items():
        if any(kw in q_lower for kw in kws):
            context["domain_hints"].append(domain)

    # Goal hints
    goal_kws = {
        "consolidate": ["consolidate", "consolidated", "combine", "merge", "roll up"],
        "split":       ["split", "separate", "divide"],
        "analyze":     ["analyze", "analyse", "analysis", "insight"],
        "clean":       ["clean", "fix", "remove", "clear"],
        "report":      ["report", "summary", "overview"],
        "compare":     ["compare", "diff", "difference"],
    }
    for goal, kws in goal_kws.items():
        if any(kw in q_lower for kw in kws):
            context["goal_hints"].append(goal)

    # Multi-file detection
    context["multi_file"] = bool(
        re.search(r'\b(files?|multiple files|several files)\b', q_lower)
    )

    return context


# =============================================================================
# SCORING ENGINE
# =============================================================================

def _expand_query(query: str) -> str:
    """Expand query with synonyms to improve matching.
    Uses word-boundary matching to avoid substring false positives
    (e.g. 'ar' in 'department' should NOT trigger the 'aging' synonym).
    """
    q = query.lower()
    for canonical, synonyms in SYNONYMS.items():
        if canonical not in q:
            if any(re.search(r'\b' + re.escape(s) + r'\b', q) for s in synonyms):
                q += f" {canonical}"
    return q


def score_intent(query: str, intent: Dict) -> float:
    """
    Score a single intent against the query.
    Returns normalized score 0.0–1.0.

    Normalization: divide by top-3 keyword scores (sum).
    This means matching the 3 strongest keywords ≈ 100% confidence.
    A typical short query matches 1-3 keywords — this keeps scores intuitive.
    """
    q = _expand_query(query)
    raw = 0.0
    all_kw_scores = []

    for kw in intent["keywords"]:
        kw_score = 2.0 if " " in kw else 1.0
        all_kw_scores.append(kw_score)
        if kw in q:
            raw += kw_score

    # Anti-keyword penalties
    for anti in intent.get("anti", []):
        if anti in q:
            raw -= 1.5

    if not all_kw_scores:
        return 0.0

    # Normalize by a fixed divisor of 3.0:
    #   phrase match (2.0) alone  → 67%  (high)
    #   phrase + word (3.0)       → 100% (high)
    #   word match only (1.0)     → 33%  (medium)
    # This keeps scores intuitive for short natural-language queries.
    DIVISOR = 3.0
    return max(0.0, min(1.0, raw / DIVISOR))


def match_scenario(
    query: str,
    file_info: Optional[Dict],
    context: Dict,
) -> Dict[str, float]:
    """
    Scenario-aware score boosts.
    Returns {intent_id: boost_score} — these are added to keyword scores.
    """
    boosts: Dict[str, float] = {}
    q = query.lower()

    if not file_info:
        return boosts

    sc = file_info.get("sheet_count", 1)
    domain = file_info.get("domain_hint")
    all_cols_lower = " ".join(
        str(c).lower()
        for cols in file_info.get("columns", {}).values()
        for c in cols
        if c is not None
    )
    time_refs = context.get("time_refs", [])
    goal_hints = context.get("goal_hints", [])
    qty = context.get("quantities", {})

    # Multi-sheet + HR/project + quarterly → timesheet rollup
    if (sc > 1 and domain in ("hr_project", "hr")
            and ("consolidate" in goal_hints or "consolidate" in q)
            and (time_refs or any(d in q for d in ["quarter", "monthly"]))):
        boosts["pm_timesheet"] = 0.45
        boosts["consolidate_sheets"] = 0.15

    # Multi-sheet + consolidate → merge sheets
    if sc > 1 and "consolidate" in goal_hints:
        boosts["consolidate_sheets"] = boosts.get("consolidate_sheets", 0) + 0.20

    # Sheet count matches quantity word
    for word, num in qty.items():
        if num == sc and sc > 1:
            boosts["pm_timesheet"] = boosts.get("pm_timesheet", 0) + 0.20
            boosts["consolidate_sheets"] = boosts.get("consolidate_sheets", 0) + 0.10

    # Two files + compare
    if file_info.get("file_count", 1) == 2 and "compare" in goal_hints:
        boosts["compare_files"] = 0.35

    # Date column + forecast → trend forecast
    if any("date" in c for c in all_cols_lower.split()) and "analyze" in goal_hints:
        boosts["analytics_forecast"] = 0.15

    # Hours column → timesheet or capacity
    if "hours" in all_cols_lower or "hrs" in all_cols_lower:
        boosts["pm_timesheet"] = boosts.get("pm_timesheet", 0) + 0.10
        boosts["pm_capacity"]  = boosts.get("pm_capacity",  0) + 0.08

    # Many rows + profile/analyze → data_profile
    if "analyze" in goal_hints and "report" in goal_hints:
        boosts["report_profile"] = 0.20

    return boosts


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def parse_intent(
    query: str,
    files: Optional[List[str]] = None,
    top_n: int = 3,
) -> List[Dict]:
    """
    Main NL routing function.

    Args:
        query: Natural language description of what to do
        files : List of provided file paths (used for file inspection + param resolution)
        top_n : How many results to return

    Returns:
        List of dicts sorted by score descending:
        [
          {
            "intent": {...},       # full INTENT_MAP entry
            "score": 0.92,
            "confidence": "high",  # high / medium / low
            "file_info": {...},    # from inspect_file (if file provided)
            "context":  {...},     # from extract_context
          },
          ...
        ]
    """
    context = extract_context(query)

    # Inspect first file if available
    file_info = None
    if files:
        primary = files[0]
        try:
            file_info = inspect_file(primary)
            if len(files) > 1:
                file_info["file_count"] = len(files)
        except Exception:
            pass

    # Scenario boosts
    boosts = match_scenario(query, file_info, context)

    # Score all intents
    scored = []
    for intent in INTENT_MAP:
        base = score_intent(query, intent)
        boost = boosts.get(intent["id"], 0.0)
        final = min(1.0, base + boost)
        scored.append((final, intent))

    # Sort descending
    scored.sort(key=lambda x: x[0], reverse=True)

    results = []
    for score, intent in scored[:top_n]:
        if score < 0.05:
            break
        confidence = (
            "high"   if score >= 0.60 else
            "medium" if score >= 0.30 else
            "low"
        )
        results.append({
            "intent":     intent,
            "score":      round(score, 3),
            "confidence": confidence,
            "file_info":  file_info,
            "context":    context,
        })

    return results


def get_output_path(fn_name: str, ext: str = ".xlsx") -> str:
    """Auto-generate a timestamped output path in the output folder."""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(__file__).parent / "output"
    out_dir.mkdir(exist_ok=True)
    return str(out_dir / f"{fn_name}_{ts}{ext}")


def get_output_dir(fn_name: str) -> str:
    """Auto-generate a timestamped output directory path."""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(__file__).parent / "output" / f"{fn_name}_{ts}"
    out_dir.mkdir(parents=True, exist_ok=True)
    return str(out_dir)
