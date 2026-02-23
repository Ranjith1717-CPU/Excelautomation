"""
=============================================================================
INVENTORY MODULE
=============================================================================
Inventory management and operations analytics.

Functions:
  abc_analysis         - A/B/C classification by cumulative value %
  reorder_point        - ROP = avg_daily_usage × lead_time + safety_stock
  stock_aging          - Age-bucket inventory (like AR aging)
  inventory_turnover   - Turnover ratio + days on hand
  oee_calculator       - OEE = Availability × Performance × Quality
  dead_stock_analysis  - Items with no movement beyond N days
=============================================================================
"""
import pandas as pd
import numpy as np
from pathlib import Path
import datetime
from typing import Optional


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Inventory") -> str:
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

def abc_analysis(file: str, item_col: str,
                 value_col: str, output_path: str,
                 a_threshold: float = 70.0,
                 b_threshold: float = 90.0) -> str:
    """
    ABC inventory classification by cumulative value %.
    A = top items contributing to a_threshold% of total value.
    B = next up to b_threshold%.
    C = rest.
    """
    df = _load(file)
    df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0)

    df_sorted = df.sort_values(value_col, ascending=False).copy()
    df_sorted["Cumulative_Value"] = df_sorted[value_col].cumsum()
    total_value = df_sorted[value_col].sum()
    df_sorted["Cumulative_%"] = (df_sorted["Cumulative_Value"] / total_value * 100).round(2)
    df_sorted["Value_%"] = (df_sorted[value_col] / total_value * 100).round(2)

    def _abc(cum_pct):
        if cum_pct <= a_threshold: return "A"
        elif cum_pct <= b_threshold: return "B"
        else: return "C"

    df_sorted["ABC_Class"] = df_sorted["Cumulative_%"].apply(_abc)

    summary = df_sorted.groupby("ABC_Class").agg(
        Item_Count=(item_col, "count"),
        Total_Value=(value_col, "sum"),
    ).round(2).reset_index()
    summary["% of Items"] = (summary["Item_Count"] / len(df_sorted) * 100).round(2)
    summary["% of Value"] = (summary["Total_Value"] / total_value * 100).round(2)

    return _save_multi({"ABC_Detail": df_sorted, "ABC_Summary": summary}, output_path)


def reorder_point(file: str, output_path: str,
                  avg_daily_usage_col: str = "Avg_Daily_Usage",
                  lead_time_col: str = "Lead_Time_Days",
                  safety_stock_col: str = "Safety_Stock",
                  current_stock_col: Optional[str] = "Current_Stock") -> str:
    """
    Reorder Point = (Avg Daily Usage × Lead Time) + Safety Stock.
    If current_stock_col provided, flags items below ROP.
    """
    df = _load(file)
    for c in [avg_daily_usage_col, lead_time_col, safety_stock_col]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["Reorder_Point"] = (
        df[avg_daily_usage_col] * df[lead_time_col] + df[safety_stock_col]
    ).round(2)

    if current_stock_col and current_stock_col in df.columns:
        df[current_stock_col] = pd.to_numeric(df[current_stock_col], errors="coerce").fillna(0)
        df["Reorder_Required"] = df[current_stock_col] <= df["Reorder_Point"]
        df["Stock_vs_ROP"] = (df[current_stock_col] - df["Reorder_Point"]).round(2)

    return _save(df, output_path, "Reorder_Point")


def stock_aging(file: str, receipt_date_col: str,
                qty_col: str, output_path: str,
                as_of_date: Optional[str] = None) -> str:
    """
    Stock age analysis: bucket inventory by receipt date into aging bands.
    """
    df = _load(file)
    as_of = pd.Timestamp(as_of_date) if as_of_date else pd.Timestamp.today()
    df["_receipt"] = pd.to_datetime(df[receipt_date_col], errors="coerce")
    df["Age_Days"] = (as_of - df["_receipt"]).dt.days.fillna(0).astype(int)
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    df.drop(columns=["_receipt"], inplace=True)

    def _bucket(d):
        if d <= 30:   return "0-30 Days"
        elif d <= 60: return "31-60 Days"
        elif d <= 90: return "61-90 Days"
        elif d <= 180:return "91-180 Days"
        else:         return "180+ Days"

    df["Age_Bucket"] = df["Age_Days"].apply(_bucket)

    order = ["0-30 Days", "31-60 Days", "61-90 Days", "91-180 Days", "180+ Days"]
    summary = df.groupby("Age_Bucket")[qty_col].agg(
        Item_Count="count", Total_Qty="sum"
    ).reindex(order).fillna(0).reset_index()
    summary["% of Total_Qty"] = (summary["Total_Qty"] / summary["Total_Qty"].sum() * 100).round(2)

    return _save_multi({"Detail": df, "Aging_Summary": summary}, output_path)


def inventory_turnover(file: str, cogs_col: str,
                       inventory_col: str, output_path: str,
                       item_col: Optional[str] = None) -> str:
    """
    Inventory Turnover = COGS / Average Inventory.
    Days on Hand = 365 / Turnover.
    """
    df = _load(file)
    df[cogs_col] = pd.to_numeric(df[cogs_col], errors="coerce").fillna(0)
    df[inventory_col] = pd.to_numeric(df[inventory_col], errors="coerce").replace(0, np.nan)

    df["Inventory_Turnover"] = (df[cogs_col] / df[inventory_col]).round(4)
    df["Days_on_Hand"] = (365 / df["Inventory_Turnover"]).round(1)

    def _rating(doh):
        if pd.isna(doh): return "N/A"
        elif doh <= 30:  return "Fast Moving"
        elif doh <= 90:  return "Normal"
        elif doh <= 180: return "Slow Moving"
        else:            return "Dead Stock Risk"

    df["Movement_Rating"] = df["Days_on_Hand"].apply(_rating)

    if item_col:
        summary = df.groupby(item_col).agg(
            Avg_Turnover=("Inventory_Turnover", "mean"),
            Avg_Days_on_Hand=("Days_on_Hand", "mean"),
        ).round(2).reset_index()
        return _save_multi({"Detail": df, "Turnover_Summary": summary}, output_path)

    return _save(df, output_path, "Inventory_Turnover")


def oee_calculator(file: str, output_path: str,
                   planned_time_col: str = "Planned_Time",
                   downtime_col: str = "Downtime",
                   ideal_rate_col: str = "Ideal_Rate",
                   actual_rate_col: str = "Actual_Rate",
                   good_units_col: str = "Good_Units",
                   total_units_col: str = "Total_Units") -> str:
    """
    OEE = Availability × Performance × Quality.
    Availability = (Planned - Downtime) / Planned
    Performance  = Actual Rate / Ideal Rate
    Quality      = Good Units / Total Units
    """
    df = _load(file)
    for c in [planned_time_col, downtime_col, ideal_rate_col,
              actual_rate_col, good_units_col, total_units_col]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    run_time = df[planned_time_col] - df[downtime_col]
    df["Availability_%"] = (run_time / df[planned_time_col].replace(0, np.nan) * 100).round(2)
    df["Performance_%"] = (df[actual_rate_col] / df[ideal_rate_col].replace(0, np.nan) * 100).round(2)
    df["Quality_%"] = (df[good_units_col] / df[total_units_col].replace(0, np.nan) * 100).round(2)
    df["OEE_%"] = (
        df["Availability_%"] * df["Performance_%"] * df["Quality_%"] / 10000
    ).round(2)

    def _oee_grade(oee):
        if pd.isna(oee):  return "N/A"
        elif oee >= 85:   return "World Class"
        elif oee >= 65:   return "Good"
        elif oee >= 45:   return "Average"
        else:             return "Needs Improvement"

    df["OEE_Grade"] = df["OEE_%"].apply(_oee_grade)
    return _save(df, output_path, "OEE")


def dead_stock_analysis(file: str, last_movement_col: str,
                        qty_col: str, output_path: str,
                        days: int = 180,
                        as_of_date: Optional[str] = None) -> str:
    """
    Identify items with no movement for 'days' or more.
    """
    df = _load(file)
    as_of = pd.Timestamp(as_of_date) if as_of_date else pd.Timestamp.today()
    df["_last_move"] = pd.to_datetime(df[last_movement_col], errors="coerce")
    df["Days_Since_Movement"] = (as_of - df["_last_move"]).dt.days.fillna(9999).astype(int)
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    df.drop(columns=["_last_move"], inplace=True)

    df["Is_Dead_Stock"] = df["Days_Since_Movement"] >= days

    dead = df[df["Is_Dead_Stock"]].copy()
    print(f"    Found   : {len(dead)} dead stock items (no movement ≥ {days} days)")

    summary = {
        "Total_Items": len(df),
        "Dead_Stock_Items": len(dead),
        "Dead_Stock_%": round(len(dead) / len(df) * 100, 2) if len(df) else 0,
        "Dead_Stock_Qty": dead[qty_col].sum(),
        "Active_Items": len(df) - len(dead),
    }
    summary_df = pd.DataFrame([summary])

    return _save_multi({"All_Items": df, "Dead_Stock": dead, "Summary": summary_df}, output_path)
