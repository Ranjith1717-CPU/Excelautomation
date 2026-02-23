"""
=============================================================================
FINANCE MODULE
=============================================================================
Domain-specific financial calculations on Excel data.

Functions:
  aging_analysis              - AR/AP aging buckets (0-30, 31-60, 61-90, 90+)
  loan_amortization           - Full EMI amortization schedule
  depreciation_schedule       - Straight-line + declining balance
  financial_ratios            - Gross margin, ROI, current ratio, etc.
  payroll_calculator          - Gross→net with HRA/PF/ESI/TDS
  budget_vs_actual            - Variance + % variance report
  compound_interest_schedule  - FV growth table
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


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Finance") -> str:
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

def aging_analysis(file: str, date_col: str, amount_col: str,
                   output_path: str,
                   as_of_date: Optional[str] = None) -> str:
    """
    AR/AP Aging Analysis.
    Buckets outstanding amounts into: 0-30, 31-60, 61-90, 90+ days.
    as_of_date: 'YYYY-MM-DD' string, defaults to today.
    """
    df = _load(file)
    as_of = pd.Timestamp(as_of_date) if as_of_date else pd.Timestamp.today()
    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    df["Days_Outstanding"] = (as_of - df["_date"]).dt.days.fillna(0).astype(int)
    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)

    def _bucket(d):
        if d <= 30:   return "0-30 Days"
        elif d <= 60: return "31-60 Days"
        elif d <= 90: return "61-90 Days"
        else:         return "90+ Days"

    df["Aging_Bucket"] = df["Days_Outstanding"].apply(_bucket)
    df.drop(columns=["_date"], inplace=True)

    summary = df.groupby("Aging_Bucket")[amount_col].agg(
        Count="count", Total="sum"
    ).reindex(["0-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]).fillna(0)
    summary["% of Total"] = (summary["Total"] / summary["Total"].sum() * 100).round(2)
    summary = summary.reset_index()

    return _save_multi({"Detail": df, "Aging_Summary": summary}, output_path)


def loan_amortization(principal: float, annual_rate: float,
                      months: int, output_path: str) -> str:
    """
    Full EMI amortization schedule.
    annual_rate: percentage (e.g. 12 for 12%).
    """
    r = annual_rate / 100 / 12
    if r == 0:
        emi = principal / months
    else:
        emi = principal * r * (1 + r) ** months / ((1 + r) ** months - 1)

    rows = []
    balance = principal
    for m in range(1, months + 1):
        interest = balance * r
        principal_part = emi - interest
        balance -= principal_part
        rows.append({
            "Month": m,
            "EMI": round(emi, 2),
            "Principal": round(principal_part, 2),
            "Interest": round(interest, 2),
            "Balance": round(max(balance, 0), 2),
        })

    df = pd.DataFrame(rows)
    print(f"    Computed: {months}-month amortization  EMI={emi:.2f}")
    return _save(df, output_path, "Amortization")


def depreciation_schedule(file: str, output_path: str,
                          asset_col: str = "Asset",
                          cost_col: str = "Cost",
                          salvage_col: str = "Salvage_Value",
                          life_col: str = "Useful_Life_Years") -> str:
    """
    Straight-line AND declining balance depreciation for each asset row.
    Expected columns: Asset, Cost, Salvage_Value, Useful_Life_Years
    """
    df = _load(file)
    for c in [cost_col, salvage_col, life_col]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["SL_Annual_Dep"] = ((df[cost_col] - df[salvage_col]) / df[life_col]).round(2)
    df["SL_Dep_Rate_%"] = (df["SL_Annual_Dep"] / df[cost_col] * 100).round(2)
    df["DB_Rate_%"] = (2 / df[life_col] * 100).round(2)  # double-declining rate
    df["DB_Year1_Dep"] = (df[cost_col] * df["DB_Rate_%"] / 100).round(2)
    df["DB_Year1_Book_Value"] = (df[cost_col] - df["DB_Year1_Dep"]).round(2)

    return _save(df, output_path, "Depreciation")


def financial_ratios(file: str, output_path: str,
                     revenue_col: str = "Revenue",
                     cogs_col: str = "COGS",
                     net_profit_col: str = "Net_Profit",
                     total_assets_col: str = "Total_Assets",
                     current_assets_col: str = "Current_Assets",
                     current_liab_col: str = "Current_Liabilities",
                     total_debt_col: str = "Total_Debt",
                     equity_col: str = "Equity") -> str:
    """
    Compute common financial ratios per row.
    Missing columns are skipped gracefully.
    """
    df = _load(file)
    cols = df.columns.tolist()

    def _col(c):
        return pd.to_numeric(df[c], errors="coerce") if c in cols else pd.Series([np.nan] * len(df))

    rev = _col(revenue_col)
    cogs = _col(cogs_col)
    np_ = _col(net_profit_col)
    ta = _col(total_assets_col)
    ca = _col(current_assets_col)
    cl = _col(current_liab_col)
    td = _col(total_debt_col)
    eq = _col(equity_col)

    gross_profit = rev - cogs
    df["Gross_Profit"] = gross_profit.round(2)
    df["Gross_Margin_%"] = (gross_profit / rev * 100).round(2)
    df["Net_Margin_%"] = (np_ / rev * 100).round(2)
    df["ROI_%"] = (np_ / ta * 100).round(2)
    df["Current_Ratio"] = (ca / cl).round(4)
    df["Debt_to_Equity"] = (td / eq).round(4)

    return _save(df, output_path, "Financial_Ratios")


def payroll_calculator(file: str, basic_col: str,
                       output_path: str,
                       hra_pct: float = 40.0,
                       pf_pct: float = 12.0,
                       esi_pct: float = 0.75,
                       tds_pct: float = 10.0) -> str:
    """
    Payroll: Gross → Net after standard Indian deductions.
    HRA = 40% of basic, PF = 12%, ESI = 0.75%, TDS = 10%.
    """
    df = _load(file)
    df["Basic"] = pd.to_numeric(df[basic_col], errors="coerce").fillna(0)
    df["HRA"] = (df["Basic"] * hra_pct / 100).round(2)
    df["Gross_Salary"] = (df["Basic"] + df["HRA"]).round(2)
    df["PF_Deduction"] = (df["Basic"] * pf_pct / 100).round(2)
    df["ESI_Deduction"] = (df["Gross_Salary"] * esi_pct / 100).round(2)
    df["TDS_Deduction"] = (df["Gross_Salary"] * tds_pct / 100).round(2)
    df["Total_Deductions"] = (df["PF_Deduction"] + df["ESI_Deduction"] + df["TDS_Deduction"]).round(2)
    df["Net_Salary"] = (df["Gross_Salary"] - df["Total_Deductions"]).round(2)

    return _save(df, output_path, "Payroll")


def budget_vs_actual(budget_file: str, actual_file: str,
                     key_col: str, output_path: str) -> str:
    """
    Merge budget and actual files on key_col.
    Compute variance and % variance for all numeric columns.
    """
    dfb = _load(budget_file)
    dfa = _load(actual_file)

    merged = pd.merge(dfb, dfa, on=key_col, suffixes=("_Budget", "_Actual"), how="outer")

    numeric_budget = [c for c in dfb.columns if c != key_col
                      and pd.api.types.is_numeric_dtype(dfb[c])]

    for col in numeric_budget:
        bc = f"{col}_Budget"
        ac = f"{col}_Actual"
        if bc in merged.columns and ac in merged.columns:
            merged[f"{col}_Variance"] = (
                pd.to_numeric(merged[ac], errors="coerce") -
                pd.to_numeric(merged[bc], errors="coerce")
            ).round(2)
            bvals = pd.to_numeric(merged[bc], errors="coerce")
            merged[f"{col}_Var_%"] = (merged[f"{col}_Variance"] / bvals.replace(0, np.nan) * 100).round(2)

    return _save(merged, output_path, "Budget_vs_Actual")


def compound_interest_schedule(principal: float, annual_rate: float,
                                periods: int, output_path: str,
                                frequency: str = "annual") -> str:
    """
    Future value compounding growth table.
    frequency: 'annual' | 'semi-annual' | 'quarterly' | 'monthly'
    """
    freq_map = {"annual": 1, "semi-annual": 2, "quarterly": 4, "monthly": 12}
    n = freq_map.get(frequency.lower(), 1)
    r = annual_rate / 100 / n
    total_periods = periods * n

    rows = []
    amount = principal
    for t in range(1, total_periods + 1):
        interest = amount * r
        amount += interest
        rows.append({
            "Period": t,
            "Opening_Balance": round(amount - interest, 2),
            "Interest_Earned": round(interest, 2),
            "Closing_Balance": round(amount, 2),
            "Total_Return_%": round((amount - principal) / principal * 100, 2),
        })

    df = pd.DataFrame(rows)
    print(f"    Computed: {total_periods} compounding periods  FV={amount:.2f}")
    return _save(df, output_path, "Compound_Interest")
