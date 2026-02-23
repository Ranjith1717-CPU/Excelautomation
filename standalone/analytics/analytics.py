"""
=============================================================================
ANALYTICS MODULE
=============================================================================
Statistical analytics and predictive calculations on Excel data.

Functions:
  correlation_matrix    - Pairwise Pearson correlation
  pareto_analysis       - 80/20 with cumulative %
  linear_regression     - OLS + R², predicted values (numpy only)
  trend_forecast        - Extrapolate trend N periods ahead
  frequency_distribution- Histogram bin data
  z_score_analysis      - Standardized scores + outlier flag
  cohort_retention      - Monthly retention matrix
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


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Analytics") -> str:
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

def correlation_matrix(file: str, columns: List[str],
                       output_path: str) -> str:
    """
    Pairwise Pearson correlation matrix for selected numeric columns.
    Also outputs a heatmap-style flat table (col_A, col_B, correlation).
    """
    df = _load(file)
    if not columns:
        columns = df.select_dtypes(include=[np.number]).columns.tolist()

    num_df = df[columns].apply(pd.to_numeric, errors="coerce")
    corr_matrix = num_df.corr(method="pearson").round(4)

    # Flat pairwise table
    pairs = []
    for i, c1 in enumerate(corr_matrix.columns):
        for j, c2 in enumerate(corr_matrix.columns):
            if i < j:
                val = corr_matrix.loc[c1, c2]
                strength = "Strong" if abs(val) >= 0.7 else "Moderate" if abs(val) >= 0.4 else "Weak"
                direction = "Positive" if val > 0 else "Negative"
                pairs.append({"Column_A": c1, "Column_B": c2,
                               "Pearson_r": val, "Strength": strength, "Direction": direction})

    pairs_df = pd.DataFrame(pairs).sort_values("Pearson_r", key=abs, ascending=False)
    corr_reset = corr_matrix.reset_index().rename(columns={"index": "Column"})

    return _save_multi({"Correlation_Matrix": corr_reset, "Pairwise_Table": pairs_df}, output_path)


def pareto_analysis(file: str, category_col: str,
                    value_col: str, output_path: str) -> str:
    """
    Pareto (80/20) analysis.
    Sorts by value descending, computes cumulative % contribution.
    """
    df = _load(file)
    df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0)

    pareto = df.groupby(category_col)[value_col].sum().reset_index()
    pareto.columns = [category_col, "Total_Value"]
    pareto.sort_values("Total_Value", ascending=False, inplace=True)

    total = pareto["Total_Value"].sum()
    pareto["Value_%"] = (pareto["Total_Value"] / total * 100).round(2)
    pareto["Cumulative_%"] = pareto["Value_%"].cumsum().round(2)

    def _zone(cum):
        if cum <= 80: return "Vital Few (80%)"
        else:          return "Trivial Many (20%)"

    pareto["Pareto_Zone"] = pareto["Cumulative_%"].apply(_zone)
    pareto["Rank"] = range(1, len(pareto) + 1)

    vital = pareto[pareto["Pareto_Zone"] == "Vital Few (80%)"]
    print(f"    80/20   : {len(vital)} items drive 80% of {value_col}")

    return _save(pareto, output_path, "Pareto")


def linear_regression(file: str, x_col: str, y_col: str,
                       output_path: str) -> str:
    """
    Simple OLS linear regression using numpy.
    Adds Predicted, Residual columns. Reports R², slope, intercept.
    """
    df = _load(file)
    df[x_col] = pd.to_numeric(df[x_col], errors="coerce")
    df[y_col] = pd.to_numeric(df[y_col], errors="coerce")

    clean = df[[x_col, y_col]].dropna()
    x = clean[x_col].values
    y = clean[y_col].values

    # OLS: [slope, intercept] via numpy polyfit
    coeffs = np.polyfit(x, y, 1)
    slope, intercept = coeffs[0], coeffs[1]
    y_pred = slope * x + intercept

    ss_res = np.sum((y - y_pred) ** 2)
    ss_tot = np.sum((y - y.mean()) ** 2)
    r_squared = 1 - ss_res / ss_tot if ss_tot != 0 else 0.0

    print(f"    Regression: y = {slope:.4f}x + {intercept:.4f}  R²={r_squared:.4f}")

    df["Predicted"] = slope * pd.to_numeric(df[x_col], errors="coerce") + intercept
    df["Residual"] = pd.to_numeric(df[y_col], errors="coerce") - df["Predicted"]
    df["Predicted"] = df["Predicted"].round(4)
    df["Residual"] = df["Residual"].round(4)

    stats = pd.DataFrame([{
        "Slope": round(slope, 6),
        "Intercept": round(intercept, 6),
        "R_Squared": round(r_squared, 6),
        "RMSE": round(np.sqrt(ss_res / len(y)), 4),
        "N": len(clean),
        "Equation": f"y = {slope:.4f}×{x_col} + {intercept:.4f}",
    }])

    return _save_multi({"Regression_Data": df, "Statistics": stats}, output_path)


def trend_forecast(file: str, date_col: str, value_col: str,
                   periods: int, output_path: str) -> str:
    """
    Linear trend forecast N periods ahead from historical data.
    Uses numeric index regression then extrapolates.
    """
    df = _load(file)
    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["_date"]).sort_values("_date")
    df[value_col] = pd.to_numeric(df[value_col], errors="coerce")

    # Use ordinal day numbers as X
    min_date = df["_date"].min()
    df["_x"] = (df["_date"] - min_date).dt.days

    clean = df[["_x", value_col]].dropna()
    x = clean["_x"].values
    y = clean[value_col].values
    coeffs = np.polyfit(x, y, 1)
    slope, intercept = coeffs[0], coeffs[1]

    df["Trend_Predicted"] = (slope * df["_x"] + intercept).round(4)

    # Forecast
    last_date = df["_date"].max()
    try:
        freq = (df["_date"].diff().median()).days
        freq = max(1, int(freq))
    except Exception:
        freq = 30

    future_dates = [last_date + pd.Timedelta(days=freq * i) for i in range(1, periods + 1)]
    future_x = [(d - min_date).days for d in future_dates]
    future_vals = [round(slope * x_ + intercept, 4) for x_ in future_x]

    forecast_df = pd.DataFrame({
        date_col: future_dates,
        f"{value_col}_Forecast": future_vals,
        "Type": "Forecast",
    })

    df_out = df.drop(columns=["_x", "_date"]).copy()
    df_out["Type"] = "Historical"

    print(f"    Forecast: {periods} periods ahead  slope={slope:.4f}")
    return _save_multi({"Historical": df_out, "Forecast": forecast_df}, output_path)


def frequency_distribution(file: str, column: str,
                            bins: int, output_path: str) -> str:
    """
    Histogram bin data for a numeric column.
    Returns bin ranges with count and % frequency.
    """
    df = _load(file)
    df[column] = pd.to_numeric(df[column], errors="coerce")
    clean = df[column].dropna()

    counts, bin_edges = np.histogram(clean, bins=bins)
    bin_labels = [f"{bin_edges[i]:.2f} – {bin_edges[i+1]:.2f}" for i in range(len(bin_edges) - 1)]

    freq_df = pd.DataFrame({
        "Bin": bin_labels,
        "Bin_Start": [round(bin_edges[i], 4) for i in range(len(bin_edges) - 1)],
        "Bin_End":   [round(bin_edges[i+1], 4) for i in range(len(bin_edges) - 1)],
        "Count": counts,
        "Frequency_%": (counts / counts.sum() * 100).round(2),
    })

    stats_df = pd.DataFrame([{
        "Count": len(clean),
        "Mean": round(clean.mean(), 4),
        "Median": round(clean.median(), 4),
        "Std_Dev": round(clean.std(), 4),
        "Min": round(clean.min(), 4),
        "Max": round(clean.max(), 4),
        "Skewness": round(clean.skew(), 4),
        "Kurtosis": round(clean.kurtosis(), 4),
    }])

    return _save_multi({"Frequency_Distribution": freq_df, "Statistics": stats_df}, output_path)


def z_score_analysis(file: str, column: str, output_path: str,
                     threshold: float = 3.0) -> str:
    """
    Z-score standardization for a numeric column.
    Flags values beyond ±threshold as outliers.
    """
    df = _load(file)
    df[column] = pd.to_numeric(df[column], errors="coerce")

    mean_val = df[column].mean()
    std_val  = df[column].std()

    df["Z_Score"] = ((df[column] - mean_val) / std_val).round(4)
    df["Is_Outlier"] = df["Z_Score"].abs() > threshold

    outliers = df[df["Is_Outlier"]].copy()
    print(f"    Outliers: {len(outliers)} value(s) beyond ±{threshold} σ")

    stats_df = pd.DataFrame([{
        "Column": column,
        "Mean": round(mean_val, 4),
        "Std_Dev": round(std_val, 4),
        "Threshold": threshold,
        "Outlier_Count": len(outliers),
        "Outlier_%": round(len(outliers) / len(df) * 100, 2),
    }])

    return _save_multi({"Z_Score_Data": df, "Outliers": outliers, "Statistics": stats_df}, output_path)


def cohort_retention(file: str, customer_col: str,
                     date_col: str, output_path: str) -> str:
    """
    Monthly cohort retention matrix.
    Rows = acquisition month cohort.
    Columns = Month 0, 1, 2, ... (periods since acquisition).
    Values = # customers retained / cohort size (%).
    """
    df = _load(file)
    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["_date", customer_col])

    df["_period"] = df["_date"].dt.to_period("M")

    # First transaction date per customer
    first_purchase = df.groupby(customer_col)["_period"].min().rename("Cohort")
    df = df.join(first_purchase, on=customer_col)
    df["Period_Number"] = (df["_period"] - df["Cohort"]).apply(lambda x: x.n)

    cohort_data = df.groupby(["Cohort", "Period_Number"])[customer_col].nunique().reset_index()
    cohort_pivot = cohort_data.pivot_table(
        index="Cohort", columns="Period_Number", values=customer_col
    )

    # Convert to retention %
    cohort_sizes = cohort_pivot[0]
    retention_pct = cohort_pivot.divide(cohort_sizes, axis=0).multiply(100).round(2)

    # Rename columns
    retention_pct.columns = [f"Month_{c}" for c in retention_pct.columns]
    cohort_pivot.columns = [f"Month_{c}" for c in cohort_pivot.columns]

    retention_reset = retention_pct.reset_index()
    retention_reset["Cohort"] = retention_reset["Cohort"].astype(str)

    counts_reset = cohort_pivot.reset_index()
    counts_reset["Cohort"] = counts_reset["Cohort"].astype(str)

    print(f"    Cohorts : {len(retention_reset)} monthly cohorts")
    return _save_multi({"Retention_%": retention_reset, "Retention_Counts": counts_reset}, output_path)
