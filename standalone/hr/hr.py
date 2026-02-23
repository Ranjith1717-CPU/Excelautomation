"""
=============================================================================
HR MODULE
=============================================================================
Human Resources analytics on Excel workforce data.

Functions:
  attrition_analysis          - Turnover rate by department
  headcount_summary           - Count/% by any group columns
  tenure_analysis             - Years-of-service bands + averages
  age_band_analysis           - Workforce age demographics
  salary_analysis             - Min/max/median/percentiles by group
  performance_distribution    - Rating distribution + forced ranking
  salary_increment_calculator - Apply % increment to salary column
=============================================================================
"""
import pandas as pd
import numpy as np
from pathlib import Path
import datetime
from typing import List, Optional, Union


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "HR") -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, sheet_name=sheet_name, index=False)
    print(f"    Saved   : {output_path}  ({len(df):,} rows × {len(df.columns)} cols)")
    return output_path


def _save_multi(sheets: dict, output_path: str) -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    print(f"    Saved   : {output_path}  ({len(sheets)} sheet(s))")
    return output_path


# ── public API ────────────────────────────────────────────────────────────────

def attrition_analysis(file: str, status_col: str,
                       dept_col: str, output_path: str,
                       active_value: str = "Active",
                       left_value: str = "Left") -> str:
    """
    Attrition / turnover rate by department.
    status_col values should include active_value and left_value.
    """
    df = _load(file)
    total = df.groupby(dept_col).size().rename("Total_Headcount")
    left = df[df[status_col].astype(str).str.strip() == left_value].groupby(dept_col).size().rename("Attritions")
    summary = pd.concat([total, left], axis=1).fillna(0)
    summary["Attrition_Rate_%"] = (summary["Attritions"] / summary["Total_Headcount"] * 100).round(2)
    summary = summary.reset_index()
    summary.loc[len(summary)] = [
        "TOTAL",
        summary["Total_Headcount"].sum(),
        summary["Attritions"].sum(),
        round(summary["Attritions"].sum() / summary["Total_Headcount"].sum() * 100, 2)
        if summary["Total_Headcount"].sum() else 0
    ]
    return _save_multi({"Detail": df, "Attrition_Summary": summary}, output_path)


def headcount_summary(file: str, group_cols: List[str],
                      output_path: str) -> str:
    """
    Headcount count and % share grouped by one or more columns.
    """
    df = _load(file)
    summary = df.groupby(group_cols).size().rename("Headcount").reset_index()
    summary["% of Total"] = (summary["Headcount"] / summary["Headcount"].sum() * 100).round(2)
    summary.sort_values("Headcount", ascending=False, inplace=True)
    return _save(summary, output_path, "Headcount_Summary")


def tenure_analysis(file: str, join_date_col: str,
                    output_path: str,
                    exit_date_col: Optional[str] = None,
                    as_of_date: Optional[str] = None) -> str:
    """
    Years-of-service distribution.
    Uses today (or as_of_date) as the end date unless exit_date_col given.
    """
    df = _load(file)
    as_of = pd.Timestamp(as_of_date) if as_of_date else pd.Timestamp.today()
    df["_join"] = pd.to_datetime(df[join_date_col], errors="coerce")

    if exit_date_col and exit_date_col in df.columns:
        df["_end"] = pd.to_datetime(df[exit_date_col], errors="coerce").fillna(as_of)
    else:
        df["_end"] = as_of

    df["Years_of_Service"] = ((df["_end"] - df["_join"]).dt.days / 365.25).round(2)

    def _band(y):
        if pd.isna(y):       return "Unknown"
        elif y < 1:          return "< 1 Year"
        elif y < 3:          return "1-3 Years"
        elif y < 5:          return "3-5 Years"
        elif y < 10:         return "5-10 Years"
        else:                return "10+ Years"

    df["Tenure_Band"] = df["Years_of_Service"].apply(_band)
    df.drop(columns=["_join", "_end"], inplace=True)

    band_order = ["< 1 Year", "1-3 Years", "3-5 Years", "5-10 Years", "10+ Years", "Unknown"]
    summary = df.groupby("Tenure_Band")["Years_of_Service"].agg(
        Count="count", Avg_Years="mean", Min_Years="min", Max_Years="max"
    ).reindex(band_order).fillna(0).round(2).reset_index()
    summary["% of Total"] = (summary["Count"] / summary["Count"].sum() * 100).round(2)

    return _save_multi({"Detail": df, "Tenure_Summary": summary}, output_path)


def age_band_analysis(file: str, dob_col: str,
                      output_path: str,
                      as_of_date: Optional[str] = None) -> str:
    """
    Workforce age demographics by 10-year bands.
    """
    df = _load(file)
    as_of = pd.Timestamp(as_of_date) if as_of_date else pd.Timestamp.today()
    df["_dob"] = pd.to_datetime(df[dob_col], errors="coerce")
    df["Age"] = ((as_of - df["_dob"]).dt.days / 365.25).round(1)
    df.drop(columns=["_dob"], inplace=True)

    def _band(a):
        if pd.isna(a):  return "Unknown"
        elif a < 25:    return "Under 25"
        elif a < 35:    return "25-34"
        elif a < 45:    return "35-44"
        elif a < 55:    return "45-54"
        else:           return "55+"

    df["Age_Band"] = df["Age"].apply(_band)
    band_order = ["Under 25", "25-34", "35-44", "45-54", "55+", "Unknown"]
    summary = df.groupby("Age_Band")["Age"].agg(
        Count="count", Avg_Age="mean"
    ).reindex(band_order).fillna(0).round(2).reset_index()
    summary["% of Total"] = (summary["Count"] / summary["Count"].sum() * 100).round(2)

    return _save_multi({"Detail": df, "Age_Band_Summary": summary}, output_path)


def salary_analysis(file: str, salary_col: str,
                    dept_col: str, output_path: str) -> str:
    """
    Salary statistics (min, max, mean, median, P25, P75) by department.
    """
    df = _load(file)
    df[salary_col] = pd.to_numeric(df[salary_col], errors="coerce")

    def _pct(g, p): return np.nanpercentile(g, p)

    summary = df.groupby(dept_col)[salary_col].agg(
        Count="count",
        Min="min",
        Max="max",
        Mean="mean",
        Median="median",
        P25=lambda x: _pct(x.dropna(), 25),
        P75=lambda x: _pct(x.dropna(), 75),
    ).round(2).reset_index()

    return _save_multi({"Detail": df, "Salary_Summary": summary}, output_path)


def performance_distribution(file: str, rating_col: str,
                              output_path: str) -> str:
    """
    Performance rating distribution with count, %, and forced ranking label.
    """
    df = _load(file)
    dist = df[rating_col].value_counts().sort_index().reset_index()
    dist.columns = ["Rating", "Count"]
    dist["% of Total"] = (dist["Count"] / dist["Count"].sum() * 100).round(2)
    dist["Cumulative_%"] = dist["% of Total"].cumsum().round(2)

    # Add forced ranking labels based on cumulative %
    def _rank(cum_pct):
        if cum_pct <= 10:    return "Top 10% — Star"
        elif cum_pct <= 35:  return "Top 35% — High Performer"
        elif cum_pct <= 75:  return "Mid 40% — Solid Performer"
        else:                return "Bottom 25% — Needs Improvement"

    dist["Forced_Rank_Label"] = dist["Cumulative_%"].apply(_rank)

    return _save_multi({"Detail": df, "Rating_Distribution": dist}, output_path)


def salary_increment_calculator(file: str, salary_col: str,
                                 pct_or_col: Union[str, float],
                                 output_path: str) -> str:
    """
    Apply salary increment.
    pct_or_col: float = flat % for all (e.g. 10.0 for 10%)
                str   = column name containing individual % values
    """
    df = _load(file)
    df[salary_col] = pd.to_numeric(df[salary_col], errors="coerce").fillna(0)

    if isinstance(pct_or_col, (int, float)):
        df["Increment_%"] = float(pct_or_col)
    else:
        df["Increment_%"] = pd.to_numeric(df[pct_or_col], errors="coerce").fillna(0)

    df["Increment_Amount"] = (df[salary_col] * df["Increment_%"] / 100).round(2)
    df["New_Salary"] = (df[salary_col] + df["Increment_Amount"]).round(2)

    return _save(df, output_path, "Salary_Increment")
