"""
=============================================================================
REPORTER MODULE
=============================================================================
Generate formatted Excel reports and data profiles.

Functions:
  generate_summary_report    - Stats for multiple files in one report
  data_profile               - Detailed column-by-column profiling
  generate_kpi_report        - Formatted KPI dashboard
  top_n_report               - Top / Bottom N records by a column
  frequency_report           - Value frequency counts for text columns
  monthly_summary_report     - Aggregate data by month
  generate_multi_sheet_report- Write any dict of DataFrames to sheets
=============================================================================
"""
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Optional, Dict
import datetime


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows × {len(df.columns)} cols)")
    return df


def _save_multi(sheets: Dict[str, pd.DataFrame], output_path: str) -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    print(f"    Saved   : {output_path}  ({len(sheets)} sheet(s))")
    return output_path


# ── public API ────────────────────────────────────────────────────────────────

def generate_summary_report(files: List[str], output_path: str) -> str:
    """
    Generate a summary statistics report across multiple Excel files.
    Each file gets its own sheet; a cross-file overview is also included.
    """
    sheets = {}
    overview_rows = []

    for f in files:
        df = _load(f)
        name = Path(f).stem[:31]

        num_cols = df.select_dtypes(include="number").columns.tolist()
        stats_rows = []
        for col in num_cols:
            s = df[col].dropna()
            stats_rows.append({
                "Column"  : col,
                "Count"   : int(s.count()),
                "Sum"     : round(s.sum(), 2),
                "Mean"    : round(s.mean(), 2),
                "Median"  : round(s.median(), 2),
                "Std_Dev" : round(s.std(), 2),
                "Min"     : round(s.min(), 2),
                "Max"     : round(s.max(), 2),
            })
        if stats_rows:
            sheets[name] = pd.DataFrame(stats_rows)

        overview_rows.append({
            "File"          : Path(f).name,
            "Rows"          : len(df),
            "Columns"       : len(df.columns),
            "Numeric_Cols"  : len(num_cols),
            "Missing_Values": int(df.isna().sum().sum()),
            "Duplicates"    : int(df.duplicated().sum()),
        })

    overview_df = pd.DataFrame(overview_rows)
    sheets = {"Overview": overview_df, **sheets}
    return _save_multi(sheets, output_path)


def data_profile(file: str, output_path: str) -> str:
    """
    Detailed column-by-column data profiling report.
    Includes: dtype, count, nulls, unique values, min/max/mean (for numerics),
    top values (for categoricals).
    """
    df = _load(file)
    rows = []

    for col in df.columns:
        col_data = df[col]
        null_count = int(col_data.isna().sum())
        unique_count = int(col_data.nunique())
        total = len(df)

        row = {
            "Column"          : col,
            "Data_Type"       : str(col_data.dtype),
            "Total_Values"    : total,
            "Non_Null_Count"  : int(col_data.notna().sum()),
            "Null_Count"      : null_count,
            "Null_Pct"        : round(null_count / total * 100, 1),
            "Unique_Values"   : unique_count,
            "Unique_Pct"      : round(unique_count / total * 100, 1),
            "Top_Value"       : str(col_data.value_counts().index[0]) if unique_count > 0 else "",
            "Top_Value_Freq"  : int(col_data.value_counts().iloc[0]) if unique_count > 0 else 0,
        }

        if pd.api.types.is_numeric_dtype(col_data):
            s = col_data.dropna()
            row.update({
                "Min"   : round(float(s.min()), 4) if len(s) else "",
                "Max"   : round(float(s.max()), 4) if len(s) else "",
                "Mean"  : round(float(s.mean()), 4) if len(s) else "",
                "Median": round(float(s.median()), 4) if len(s) else "",
                "Std"   : round(float(s.std()), 4) if len(s) else "",
            })
        else:
            row.update({"Min": "", "Max": "", "Mean": "", "Median": "", "Std": ""})

        rows.append(row)

    profile_df = pd.DataFrame(rows)

    meta_df = pd.DataFrame([{
        "File"              : Path(file).name,
        "Total_Rows"        : len(df),
        "Total_Columns"     : len(df.columns),
        "Total_Cells"       : len(df) * len(df.columns),
        "Missing_Cells"     : int(df.isna().sum().sum()),
        "Missing_Pct"       : round(df.isna().sum().sum() / (len(df) * len(df.columns)) * 100, 1),
        "Duplicate_Rows"    : int(df.duplicated().sum()),
        "Profile_Generated" : datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }])

    return _save_multi({"Metadata": meta_df, "Column_Profile": profile_df}, output_path)


def generate_kpi_report(file: str, kpi_columns: List[str],
                         label_column: Optional[str] = None,
                         output_path: str = "") -> str:
    """
    Generate a formatted KPI report with key metrics and rankings.

    Args:
        kpi_columns  : Numeric columns to include as KPIs.
        label_column : Category/name column for grouping.
    """
    df = _load(file)
    sheets = {}

    if label_column and label_column in df.columns:
        grouped = df.groupby(label_column)[kpi_columns].agg(["sum", "mean", "max", "min", "count"])
        grouped.columns = [f"{col}_{fn}" for col, fn in grouped.columns]
        grouped = grouped.reset_index()
        sheets["KPI_by_Group"] = grouped

    summary_rows = []
    for col in kpi_columns:
        if col not in df.columns:
            continue
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        summary_rows.append({
            "KPI"           : col,
            "Total"         : round(s.sum(), 2),
            "Average"       : round(s.mean(), 2),
            "Peak"          : round(s.max(), 2),
            "Lowest"        : round(s.min(), 2),
            "Std_Dev"       : round(s.std(), 2),
            "Count"         : int(s.count()),
            "Above_Avg_Rows": int((s > s.mean()).sum()),
            "Below_Avg_Rows": int((s < s.mean()).sum()),
        })

    sheets["KPI_Summary"] = pd.DataFrame(summary_rows)
    sheets["Raw_Data"]    = df

    return _save_multi(sheets, output_path)


def top_n_report(file: str, sort_column: str, n: int,
                  output_path: str, ascending: bool = False) -> str:
    """
    Generate a Top-N and Bottom-N report sorted by a numeric column.
    """
    df = _load(file)
    top_n    = df.nlargest(n, sort_column)  if not ascending else df.nsmallest(n, sort_column)
    bottom_n = df.nsmallest(n, sort_column) if not ascending else df.nlargest(n, sort_column)
    top_n.insert(0, "Rank", range(1, len(top_n) + 1))
    bottom_n.insert(0, "Rank", range(1, len(bottom_n) + 1))

    label_top    = f"Top_{n}"
    label_bottom = f"Bottom_{n}"

    print(f"    Top {n}    : {sort_column} max = {top_n[sort_column].iloc[0]:,.2f}")
    print(f"    Bottom {n} : {sort_column} min = {bottom_n[sort_column].iloc[0]:,.2f}")
    return _save_multi({label_top: top_n, label_bottom: bottom_n}, output_path)


def frequency_report(file: str, columns: List[str],
                      output_path: str, top_n: int = 20) -> str:
    """
    Value frequency count for categorical columns (like a pivot count).
    """
    df = _load(file)
    sheets = {}

    for col in columns:
        if col not in df.columns:
            continue
        freq = df[col].value_counts().head(top_n).reset_index()
        freq.columns = [col, "Count"]
        freq["Percentage_%"] = (freq["Count"] / len(df) * 100).round(2)
        sheets[col[:31]] = freq

    return _save_multi(sheets, output_path)


def monthly_summary_report(file: str, date_column: str,
                             value_columns: List[str],
                             output_path: str,
                             aggfunc: str = "sum") -> str:
    """
    Aggregate data by month from a date column.
    """
    df = _load(file)
    df[date_column] = pd.to_datetime(df[date_column], errors="coerce", infer_datetime_format=True)
    df = df.dropna(subset=[date_column])
    df["_Year"]  = df[date_column].dt.year
    df["_Month"] = df[date_column].dt.month
    df["_Month_Label"] = df[date_column].dt.to_period("M").astype(str)

    agg_cols = [c for c in value_columns if c in df.columns and pd.api.types.is_numeric_dtype(df[c])]
    grouped = df.groupby(["_Year", "_Month", "_Month_Label"])[agg_cols].agg(aggfunc).reset_index()
    grouped = grouped.sort_values(["_Year", "_Month"]).drop(columns=["_Year", "_Month"])
    grouped = grouped.rename(columns={"_Month_Label": "Month"})

    print(f"    Monthly aggregation  ({aggfunc})  →  {len(grouped):,} months")
    return _save_multi({"Monthly_Summary": grouped, "Raw_Data": df.drop(columns=["_Year","_Month","_Month_Label"])},
                        output_path)


def generate_multi_sheet_report(data_dict: Dict[str, pd.DataFrame],
                                  output_path: str) -> str:
    """
    Write any dict of DataFrames to a multi-sheet Excel file.
    Useful as a utility for building custom reports.
    """
    return _save_multi(data_dict, output_path)
