"""
=============================================================================
SALES MODULE
=============================================================================
Sales analytics, commission, segmentation and pipeline calculations.

Functions:
  commission_calculator  - Flat % or tiered slab commission
  rfm_segmentation       - RFM scores + customer segments
  quota_attainment       - % attainment + Above/Near/Below labels
  pipeline_analysis      - Funnel by stage + conversion rates
  sales_by_territory     - Territory summary + rank
  customer_abc           - A/B/C by revenue contribution
  discount_analysis      - Discount % + revenue leakage
=============================================================================
"""
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Optional


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Sales") -> str:
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

def commission_calculator(file: str, sales_col: str,
                           output_path: str,
                           tiers: Optional[List[tuple]] = None,
                           flat_pct: float = 5.0) -> str:
    """
    Commission calculation.
    tiers: list of (max_sales, pct) e.g. [(50000, 3), (100000, 5), (float('inf'), 8)]
           If None, flat_pct is applied.
    """
    df = _load(file)
    df[sales_col] = pd.to_numeric(df[sales_col], errors="coerce").fillna(0)

    if tiers:
        def _apply_tier(amount):
            for limit, pct in sorted(tiers, key=lambda x: x[0]):
                if amount <= limit:
                    return round(amount * pct / 100, 2)
            return round(amount * tiers[-1][1] / 100, 2)
        df["Commission_Rate_%"] = df[sales_col].apply(
            lambda a: next((p for l, p in sorted(tiers, key=lambda x: x[0]) if a <= l), tiers[-1][1])
        )
        df["Commission_Amount"] = df[sales_col].apply(_apply_tier)
    else:
        df["Commission_Rate_%"] = flat_pct
        df["Commission_Amount"] = (df[sales_col] * flat_pct / 100).round(2)

    df["Net_Payout"] = (df[sales_col] - df["Commission_Amount"]).round(2)
    return _save(df, output_path, "Commission")


def rfm_segmentation(file: str, customer_col: str,
                     date_col: str, amount_col: str,
                     output_path: str,
                     as_of_date: Optional[str] = None) -> str:
    """
    RFM (Recency, Frequency, Monetary) segmentation.
    Scores 1-5 each dimension; assigns segment label.
    """
    df = _load(file)
    as_of = pd.Timestamp(as_of_date) if as_of_date else pd.Timestamp.today()
    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)

    rfm = df.groupby(customer_col).agg(
        Recency=("_date", lambda x: (as_of - x.max()).days),
        Frequency=(date_col, "count"),
        Monetary=(amount_col, "sum"),
    ).reset_index()

    # Score 1-5 (5 = best): Recency is inverse
    rfm["R_Score"] = pd.qcut(rfm["Recency"].rank(method="first"), 5, labels=[5, 4, 3, 2, 1]).astype(int)
    rfm["F_Score"] = pd.qcut(rfm["Frequency"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
    rfm["M_Score"] = pd.qcut(rfm["Monetary"].rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
    rfm["RFM_Score"] = rfm["R_Score"] + rfm["F_Score"] + rfm["M_Score"]

    def _segment(row):
        r, f, m = row["R_Score"], row["F_Score"], row["M_Score"]
        if r >= 4 and f >= 4 and m >= 4: return "Champions"
        elif r >= 3 and f >= 3:          return "Loyal Customers"
        elif r >= 4 and f <= 2:          return "Recent Customers"
        elif r <= 2 and f >= 4:          return "At Risk"
        elif r <= 2 and f <= 2:          return "Lost Customers"
        else:                             return "Potential Loyalists"

    rfm["Segment"] = rfm.apply(_segment, axis=1)
    rfm["Recency"] = rfm["Recency"].astype(int)
    rfm["Monetary"] = rfm["Monetary"].round(2)

    seg_summary = rfm.groupby("Segment").agg(
        Count=("Segment", "count"),
        Avg_Recency=("Recency", "mean"),
        Avg_Frequency=("Frequency", "mean"),
        Avg_Monetary=("Monetary", "mean"),
    ).round(2).reset_index()

    return _save_multi({"RFM_Detail": rfm, "Segment_Summary": seg_summary}, output_path)


def quota_attainment(file: str, actual_col: str,
                     quota_col: str, output_path: str) -> str:
    """
    % quota attainment with Above/Near/Below labels.
    Above ≥ 100%, Near = 80-99%, Below < 80%.
    """
    df = _load(file)
    df[actual_col] = pd.to_numeric(df[actual_col], errors="coerce").fillna(0)
    df[quota_col] = pd.to_numeric(df[quota_col], errors="coerce").replace(0, np.nan)

    df["Attainment_%"] = (df[actual_col] / df[quota_col] * 100).round(2)

    def _label(pct):
        if pd.isna(pct):   return "No Quota"
        elif pct >= 100:   return "Above Quota"
        elif pct >= 80:    return "Near Quota"
        else:              return "Below Quota"

    df["Attainment_Status"] = df["Attainment_%"].apply(_label)
    summary = df["Attainment_Status"].value_counts().reset_index()
    summary.columns = ["Status", "Count"]

    return _save_multi({"Detail": df, "Attainment_Summary": summary}, output_path)


def pipeline_analysis(file: str, stage_col: str,
                      value_col: str, output_path: str) -> str:
    """
    Sales pipeline funnel: count and total value by stage.
    Computes stage-to-stage conversion rates.
    """
    df = _load(file)
    df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0)

    funnel = df.groupby(stage_col).agg(
        Count=(stage_col, "count"),
        Total_Value=(value_col, "sum"),
        Avg_Deal_Size=(value_col, "mean"),
    ).round(2).reset_index()

    funnel.sort_values("Total_Value", ascending=False, inplace=True)
    total = funnel["Count"].sum()
    funnel["% of Pipeline"] = (funnel["Count"] / total * 100).round(2)

    return _save_multi({"Detail": df, "Pipeline_Funnel": funnel}, output_path)


def sales_by_territory(file: str, territory_col: str,
                       sales_col: str, output_path: str) -> str:
    """
    Territory-level sales summary with rank.
    """
    df = _load(file)
    df[sales_col] = pd.to_numeric(df[sales_col], errors="coerce").fillna(0)

    summary = df.groupby(territory_col)[sales_col].agg(
        Transactions="count",
        Total_Sales="sum",
        Avg_Sale="mean",
        Min_Sale="min",
        Max_Sale="max",
    ).round(2).reset_index()

    summary.sort_values("Total_Sales", ascending=False, inplace=True)
    summary["Rank"] = range(1, len(summary) + 1)
    summary["% of Total"] = (summary["Total_Sales"] / summary["Total_Sales"].sum() * 100).round(2)

    return _save_multi({"Detail": df, "Territory_Summary": summary}, output_path)


def customer_abc(file: str, customer_col: str,
                 revenue_col: str, output_path: str,
                 a_threshold: float = 70.0,
                 b_threshold: float = 90.0) -> str:
    """
    A/B/C customer classification by cumulative revenue contribution.
    A = top customers up to a_threshold% of revenue
    B = next customers up to b_threshold%
    C = rest
    """
    df = _load(file)
    df[revenue_col] = pd.to_numeric(df[revenue_col], errors="coerce").fillna(0)

    cust = df.groupby(customer_col)[revenue_col].sum().reset_index()
    cust.columns = [customer_col, "Total_Revenue"]
    cust.sort_values("Total_Revenue", ascending=False, inplace=True)
    cust["Cumulative_Revenue"] = cust["Total_Revenue"].cumsum()
    total_rev = cust["Total_Revenue"].sum()
    cust["Cumulative_%"] = (cust["Cumulative_Revenue"] / total_rev * 100).round(2)
    cust["Revenue_%"] = (cust["Total_Revenue"] / total_rev * 100).round(2)

    def _abc(cum_pct):
        if cum_pct <= a_threshold: return "A"
        elif cum_pct <= b_threshold: return "B"
        else: return "C"

    cust["ABC_Class"] = cust["Cumulative_%"].apply(_abc)

    abc_summary = cust.groupby("ABC_Class").agg(
        Customers=(customer_col, "count"),
        Total_Revenue=("Total_Revenue", "sum"),
    ).round(2).reset_index()
    abc_summary["Revenue_%"] = (abc_summary["Total_Revenue"] / total_rev * 100).round(2)

    return _save_multi({"Customer_ABC": cust, "ABC_Summary": abc_summary}, output_path)


def discount_analysis(file: str, list_price_col: str,
                      sell_price_col: str, output_path: str) -> str:
    """
    Discount % per row and total revenue leakage analysis.
    """
    df = _load(file)
    df[list_price_col] = pd.to_numeric(df[list_price_col], errors="coerce").fillna(0)
    df[sell_price_col] = pd.to_numeric(df[sell_price_col], errors="coerce").fillna(0)

    df["Discount_Amount"] = (df[list_price_col] - df[sell_price_col]).round(2)
    lp = df[list_price_col].replace(0, np.nan)
    df["Discount_%"] = (df["Discount_Amount"] / lp * 100).round(2)

    def _band(pct):
        if pd.isna(pct) or pct <= 0: return "No Discount"
        elif pct <= 5:  return "0-5%"
        elif pct <= 10: return "5-10%"
        elif pct <= 20: return "10-20%"
        else:           return "20%+"

    df["Discount_Band"] = df["Discount_%"].apply(_band)

    total_list = df[list_price_col].sum()
    total_sell = df[sell_price_col].sum()
    total_leakage = total_list - total_sell

    summary_data = {
        "Total_List_Price": [round(total_list, 2)],
        "Total_Sell_Price": [round(total_sell, 2)],
        "Total_Discount_Leakage": [round(total_leakage, 2)],
        "Overall_Discount_%": [round(total_leakage / total_list * 100, 2) if total_list else 0],
    }
    summary = pd.DataFrame(summary_data)

    band_summary = df.groupby("Discount_Band").agg(
        Count=(list_price_col, "count"),
        Total_Discount=("Discount_Amount", "sum"),
    ).round(2).reset_index()

    return _save_multi({"Detail": df, "Discount_Overview": summary, "Band_Summary": band_summary}, output_path)
