"""
=============================================================================
CALCULATOR MODULE
=============================================================================
Performs business calculations on Excel data.

Functions:
  calculate_efficiency       - (Actual / Target) × 100
  calculate_productivity     - Output per unit of input (e.g., per hour)
  calculate_utilization      - Used / Available × 100
  calculate_variance         - Actual vs. Budget (absolute + %)
  calculate_growth_rate      - Period-over-period growth %
  calculate_summary_stats    - Mean, Median, Std, Min, Max, Sum, Count
  calculate_percentage_of_total - Each row's % share of a column total
  calculate_moving_average   - Rolling/moving average for a column
  calculate_kpi_dashboard    - Multi-KPI report across selected columns
  calculate_weighted_average - Weighted mean
=============================================================================
"""
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Dict, Optional


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Calculated") -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, sheet_name=sheet_name, index=False)
    print(f"    Saved   : {output_path}  ({len(df):,} rows × {len(df.columns)} cols)")
    return output_path


def _numeric(df: pd.DataFrame, col: str) -> pd.Series:
    """Coerce column to numeric, warn on errors."""
    s = pd.to_numeric(df[col], errors="coerce")
    nulls = s.isna().sum()
    if nulls:
        print(f"    Warning : {nulls} non-numeric value(s) in '{col}' treated as NaN")
    return s


# ── public API ────────────────────────────────────────────────────────────────

def calculate_efficiency(file: str, actual_col: str, target_col: str,
                         output_path: str,
                         result_col: str = "Efficiency_%") -> str:
    """
    Efficiency = (Actual / Target) × 100

    Adds a new column with the efficiency percentage.
    """
    df = _load(file)
    actual = _numeric(df, actual_col)
    target = _numeric(df, target_col)
    df[result_col] = (actual / target * 100).round(2)
    df["Efficiency_Status"] = df[result_col].apply(
        lambda x: "Above Target" if x >= 100
        else ("Near Target" if x >= 90 else "Below Target")
    )
    print(f"    Avg Efficiency: {df[result_col].mean():.1f}%  |  "
          f"Max: {df[result_col].max():.1f}%  |  Min: {df[result_col].min():.1f}%")
    return _save(df, output_path)


def calculate_productivity(file: str, output_col: str, input_col: str,
                           output_path: str,
                           result_col: str = "Productivity") -> str:
    """
    Productivity = Output / Input  (e.g., units produced per hour)

    Adds a new column with the productivity ratio.
    """
    df = _load(file)
    out = _numeric(df, output_col)
    inp = _numeric(df, input_col)
    df[result_col] = (out / inp).round(4)
    print(f"    Avg Productivity: {df[result_col].mean():.2f}  |  "
          f"Max: {df[result_col].max():.2f}  |  Min: {df[result_col].min():.2f}")
    return _save(df, output_path)


def calculate_utilization(file: str, used_col: str, available_col: str,
                          output_path: str,
                          result_col: str = "Utilization_%") -> str:
    """
    Utilization = (Used / Available) × 100
    """
    df = _load(file)
    used = _numeric(df, used_col)
    avail = _numeric(df, available_col)
    df[result_col] = (used / avail * 100).round(2)
    df["Utilization_Status"] = df[result_col].apply(
        lambda x: "Over-utilized" if x > 100
        else ("High" if x >= 80 else ("Moderate" if x >= 50 else "Low"))
    )
    print(f"    Avg Utilization: {df[result_col].mean():.1f}%")
    return _save(df, output_path)


def calculate_variance(file: str, actual_col: str, budget_col: str,
                       output_path: str) -> str:
    """
    Variance = Actual − Budget
    Variance % = (Actual − Budget) / Budget × 100

    Adds both absolute and percentage variance columns.
    """
    df = _load(file)
    actual = _numeric(df, actual_col)
    budget = _numeric(df, budget_col)
    df["Variance"] = (actual - budget).round(2)
    df["Variance_%"] = ((actual - budget) / budget * 100).round(2)
    df["Variance_Type"] = df["Variance"].apply(
        lambda x: "Favourable" if x > 0 else ("On-Budget" if x == 0 else "Unfavourable")
    )
    print(f"    Total Variance: {df['Variance'].sum():,.2f}  |  "
          f"Avg Variance%: {df['Variance_%'].mean():.1f}%")
    return _save(df, output_path)


def calculate_growth_rate(file: str, value_col: str,
                          period_col: Optional[str] = None,
                          output_path: str = "",
                          result_col: str = "Growth_Rate_%") -> str:
    """
    Period-over-period growth rate.
    Growth % = (Current − Previous) / |Previous| × 100
    Rows are assumed to be in chronological order unless period_col is provided.
    """
    df = _load(file)
    if period_col and period_col in df.columns:
        df = df.sort_values(period_col).reset_index(drop=True)

    values = _numeric(df, value_col)
    df[result_col] = (values.pct_change() * 100).round(2)
    df["Growth_Direction"] = df[result_col].apply(
        lambda x: "Growth" if x > 0 else ("Decline" if x < 0 else "Flat")
    )
    avg_growth = df[result_col].dropna().mean()
    print(f"    Avg Growth Rate: {avg_growth:.1f}%")
    return _save(df, output_path)


def calculate_summary_stats(file: str, columns: List[str],
                             output_path: str) -> str:
    """
    Compute descriptive statistics (count, sum, mean, median, std, min, max)
    for selected numeric columns.
    Results are written as a stats summary sheet.
    """
    df = _load(file)
    cols = [c for c in columns if c in df.columns] if columns else df.select_dtypes(include="number").columns.tolist()

    if not cols:
        raise ValueError("No numeric columns found / selected.")

    rows = []
    for col in cols:
        s = _numeric(df, col).dropna()
        rows.append({
            "Column"  : col,
            "Count"   : int(s.count()),
            "Sum"     : round(s.sum(), 2),
            "Mean"    : round(s.mean(), 2),
            "Median"  : round(s.median(), 2),
            "Std_Dev" : round(s.std(), 2),
            "Min"     : round(s.min(), 2),
            "Max"     : round(s.max(), 2),
            "Range"   : round(s.max() - s.min(), 2),
        })
    stats_df = pd.DataFrame(rows)
    print(stats_df.to_string(index=False))
    return _save(stats_df, output_path, sheet_name="Summary_Stats")


def calculate_percentage_of_total(file: str, value_col: str,
                                   group_col: Optional[str] = None,
                                   output_path: str = "",
                                   result_col: str = "Pct_of_Total") -> str:
    """
    Each row's percentage share of the column total.
    If group_col is provided, percentage is within each group.
    """
    df = _load(file)
    values = _numeric(df, value_col)

    if group_col and group_col in df.columns:
        group_totals = df.groupby(group_col)[value_col].transform("sum")
        df[result_col] = (values / group_totals * 100).round(2)
        print(f"    Pct_of_Total computed within groups of '{group_col}'")
    else:
        total = values.sum()
        df[result_col] = (values / total * 100).round(2)
        print(f"    Grand Total: {total:,.2f}")

    return _save(df, output_path)


def calculate_moving_average(file: str, value_col: str, window: int,
                              output_path: str) -> str:
    """
    Compute a rolling/moving average for a numeric column.

    Args:
        window: Number of periods (rows) to average over.
    """
    df = _load(file)
    values = _numeric(df, value_col)
    col_name = f"MA_{window}_periods"
    df[col_name] = values.rolling(window=window, min_periods=1).mean().round(4)
    print(f"    Moving average ({window} periods) added as '{col_name}'")
    return _save(df, output_path)


def calculate_kpi_dashboard(file: str, kpi_columns: List[str],
                             output_path: str) -> str:
    """
    Generate a KPI dashboard with key metrics for multiple columns.
    Outputs two sheets: raw data + KPI summary.
    """
    df = _load(file)
    num_cols = [c for c in kpi_columns if c in df.columns]

    kpi_rows = []
    for col in num_cols:
        s = _numeric(df, col).dropna()
        total = s.sum()
        mean  = s.mean()
        kpi_rows.append({
            "KPI_Column"       : col,
            "Total"            : round(total, 2),
            "Average"          : round(mean, 2),
            "Peak"             : round(s.max(), 2),
            "Lowest"           : round(s.min(), 2),
            "Std_Deviation"    : round(s.std(), 2),
            "Records_Count"    : int(s.count()),
            "Above_Average_cnt": int((s > mean).sum()),
            "Below_Average_cnt": int((s < mean).sum()),
        })

    kpi_df = pd.DataFrame(kpi_rows)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Raw_Data", index=False)
        kpi_df.to_excel(writer, sheet_name="KPI_Dashboard", index=False)

    print(f"    Saved KPI dashboard: {output_path}")
    return output_path


def calculate_weighted_average(file: str, value_col: str, weight_col: str,
                                group_col: Optional[str] = None,
                                output_path: str = "") -> str:
    """
    Weighted average = sum(value × weight) / sum(weight)
    If group_col is provided, compute weighted average per group.
    """
    df = _load(file)
    values  = _numeric(df, value_col)
    weights = _numeric(df, weight_col)

    if group_col and group_col in df.columns:
        result = (
            df.assign(_v=values * weights)
              .groupby(group_col)
              .apply(lambda g: g["_v"].sum() / weights.loc[g.index].sum())
              .reset_index()
        )
        result.columns = [group_col, "Weighted_Average"]
    else:
        wa = (values * weights).sum() / weights.sum()
        result = pd.DataFrame([{"Weighted_Average": round(wa, 4)}])
        print(f"    Overall Weighted Average: {wa:.4f}")

    return _save(result, output_path, sheet_name="Weighted_Avg")
