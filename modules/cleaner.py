"""
=============================================================================
CLEANER MODULE
=============================================================================
Data cleaning operations for Excel files.

Functions:
  remove_duplicates          - Drop duplicate rows
  remove_empty_rows_cols     - Drop fully empty rows and/or columns
  trim_whitespace            - Strip leading/trailing spaces in string columns
  standardize_dates          - Parse and reformat date columns
  fill_missing_values        - Fill NaN with a chosen strategy
  fix_data_types             - Auto-detect and coerce column data types
  normalize_text_case        - upper / lower / title / sentence case
  remove_special_characters  - Strip non-alphanumeric chars from columns
  remove_outliers            - Drop rows where value is N std devs from mean
  full_clean                 - Run all cleaning steps in one shot
=============================================================================
"""
import pandas as pd
import numpy as np
import re
from pathlib import Path
from typing import List, Optional, Dict


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows × {len(df.columns)} cols)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Cleaned") -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, sheet_name=sheet_name, index=False)
    print(f"    Saved   : {output_path}  ({len(df):,} rows × {len(df.columns)} cols)")
    return output_path


# ── public API ────────────────────────────────────────────────────────────────

def remove_duplicates(file: str, output_path: str,
                      subset: Optional[List[str]] = None,
                      keep: str = "first") -> str:
    """
    Remove duplicate rows.

    Args:
        subset: Column(s) to consider for duplication check. None = all columns.
        keep  : 'first' | 'last' | False (drop all duplicates)
    """
    df = _load(file)
    before = len(df)
    df = df.drop_duplicates(subset=subset, keep=keep).reset_index(drop=True)
    removed = before - len(df)
    print(f"    Removed : {removed:,} duplicate row(s)  ({len(df):,} remaining)")
    return _save(df, output_path)


def remove_empty_rows_cols(file: str, output_path: str,
                            remove_rows: bool = True,
                            remove_cols: bool = True,
                            threshold: float = 1.0) -> str:
    """
    Drop rows/columns that are entirely (or mostly) empty.

    Args:
        threshold: Fraction of NaN required to drop (1.0 = all NaN, 0.5 = 50% NaN)
    """
    df = _load(file)

    if remove_rows:
        before = len(df)
        df = df.dropna(axis=0, how="all" if threshold == 1.0 else "any",
                       thresh=None if threshold == 1.0 else int(len(df.columns) * (1 - threshold)))
        print(f"    Removed : {before - len(df):,} empty row(s)")

    if remove_cols:
        before_cols = len(df.columns)
        df = df.dropna(axis=1, how="all" if threshold == 1.0 else "any",
                       thresh=None if threshold == 1.0 else int(len(df) * (1 - threshold)))
        print(f"    Removed : {before_cols - len(df.columns):,} empty column(s)")

    return _save(df, output_path)


def trim_whitespace(file: str, output_path: str,
                    columns: Optional[List[str]] = None) -> str:
    """
    Strip leading/trailing whitespace from all (or selected) string columns.
    Also collapses multiple internal spaces into one.
    """
    df = _load(file)
    cols = columns if columns else df.select_dtypes(include="object").columns.tolist()
    fixed = 0
    for col in cols:
        if col in df.columns:
            original = df[col].astype(str)
            df[col] = df[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
            fixed += (original != df[col]).sum()
    print(f"    Fixed   : {fixed:,} cell(s) with extra whitespace")
    return _save(df, output_path)


def standardize_dates(file: str, date_columns: List[str],
                      output_format: str = "%Y-%m-%d",
                      output_path: str = "") -> str:
    """
    Parse date columns and reformat to a uniform output format.

    Args:
        date_columns : List of column names to treat as dates.
        output_format: strftime format string (default ISO 8601: %Y-%m-%d)
    """
    df = _load(file)
    for col in date_columns:
        if col not in df.columns:
            print(f"    Warning : Column '{col}' not found — skipped")
            continue
        parsed = pd.to_datetime(df[col], errors="coerce", infer_datetime_format=True)
        failures = parsed.isna().sum()
        df[col] = parsed.dt.strftime(output_format)
        print(f"    Column '{col}': reformatted  ({failures} parse failures → NaT)")
    return _save(df, output_path)


def fill_missing_values(file: str, output_path: str,
                         strategy: str = "mean",
                         fill_value=None,
                         columns: Optional[List[str]] = None) -> str:
    """
    Fill NaN values using a chosen strategy.

    Args:
        strategy  : 'mean' | 'median' | 'mode' | 'ffill' | 'bfill' | 'value'
        fill_value: Used when strategy='value'
        columns   : Columns to fill. None = all.
    """
    df = _load(file)
    cols = columns if columns else df.columns.tolist()
    total_filled = 0

    for col in cols:
        if col not in df.columns:
            continue
        nulls_before = df[col].isna().sum()
        if nulls_before == 0:
            continue

        if strategy == "mean" and pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(df[col].mean())
        elif strategy == "median" and pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(df[col].median())
        elif strategy == "mode":
            mode_val = df[col].mode()
            if not mode_val.empty:
                df[col] = df[col].fillna(mode_val[0])
        elif strategy == "ffill":
            df[col] = df[col].ffill()
        elif strategy == "bfill":
            df[col] = df[col].bfill()
        elif strategy == "value" and fill_value is not None:
            df[col] = df[col].fillna(fill_value)

        filled = nulls_before - df[col].isna().sum()
        total_filled += filled

    print(f"    Filled  : {total_filled:,} missing value(s) using strategy='{strategy}'")
    return _save(df, output_path)


def fix_data_types(file: str, output_path: str) -> str:
    """
    Auto-detect and coerce column data types:
    - Columns that look numeric → float/int
    - Columns that look like dates → datetime → string ISO
    - Strip 'object' columns that are actually numbers stored as text
    """
    df = _load(file)
    conversions = []

    for col in df.columns:
        original_dtype = str(df[col].dtype)

        # Try numeric
        if df[col].dtype == object:
            converted = pd.to_numeric(df[col].astype(str).str.replace(",", ""), errors="coerce")
            if converted.notna().mean() > 0.8:
                df[col] = converted
                conversions.append(f"'{col}': object → numeric")
                continue

        # Try datetime
        if df[col].dtype == object:
            try:
                converted = pd.to_datetime(df[col], errors="coerce", infer_datetime_format=True)
                if converted.notna().mean() > 0.8:
                    df[col] = converted.dt.strftime("%Y-%m-%d")
                    conversions.append(f"'{col}': object → date")
                    continue
            except Exception:
                pass

    if conversions:
        for c in conversions:
            print(f"    Converted: {c}")
    else:
        print("    No type conversions needed.")

    return _save(df, output_path)


def normalize_text_case(file: str, output_path: str,
                         columns: Optional[List[str]] = None,
                         case: str = "title") -> str:
    """
    Normalize text case in string columns.

    Args:
        case: 'upper' | 'lower' | 'title' | 'sentence'
    """
    df = _load(file)
    cols = columns if columns else df.select_dtypes(include="object").columns.tolist()

    for col in cols:
        if col not in df.columns:
            continue
        if case == "upper":
            df[col] = df[col].astype(str).str.upper()
        elif case == "lower":
            df[col] = df[col].astype(str).str.lower()
        elif case == "title":
            df[col] = df[col].astype(str).str.title()
        elif case == "sentence":
            df[col] = df[col].astype(str).str.capitalize()

    print(f"    Applied '{case}' case to {len(cols)} column(s)")
    return _save(df, output_path)


def remove_special_characters(file: str, output_path: str,
                               columns: Optional[List[str]] = None,
                               keep_pattern: str = r"[^\w\s.,\-]") -> str:
    """
    Strip special/unwanted characters from string columns.

    Args:
        keep_pattern: Regex of characters to REMOVE (default removes most special chars
                      but keeps word chars, spaces, commas, periods, hyphens).
    """
    df = _load(file)
    cols = columns if columns else df.select_dtypes(include="object").columns.tolist()
    total = 0

    for col in cols:
        if col not in df.columns:
            continue
        before = df[col].astype(str).copy()
        df[col] = df[col].astype(str).str.replace(keep_pattern, "", regex=True)
        changed = (before != df[col]).sum()
        total += changed

    print(f"    Cleaned : {total:,} cell(s) containing special characters")
    return _save(df, output_path)


def remove_outliers(file: str, value_column: str, output_path: str,
                    std_threshold: float = 3.0) -> str:
    """
    Remove rows where the value in value_column is more than
    std_threshold standard deviations from the mean.
    """
    df = _load(file)
    values = pd.to_numeric(df[value_column], errors="coerce")
    mean   = values.mean()
    std    = values.std()
    mask   = (values - mean).abs() <= std_threshold * std
    removed = (~mask).sum()
    df = df[mask].reset_index(drop=True)
    print(f"    Removed : {removed:,} outlier row(s) (±{std_threshold}σ from mean {mean:.2f})")
    return _save(df, output_path)


def full_clean(file: str, output_path: str) -> str:
    """
    Run all cleaning steps in sequence:
    1. Remove fully empty rows/cols
    2. Remove duplicates
    3. Trim whitespace
    4. Fix data types
    5. Fill missing numerics with median
    """
    df = _load(file)

    # 1. Empty rows/cols
    before = df.shape
    df = df.dropna(how="all").dropna(axis=1, how="all")
    print(f"    Step 1 — Empty rows/cols removed: {before} → {df.shape}")

    # 2. Duplicates
    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    print(f"    Step 2 — Duplicates removed: {before} → {len(df)}")

    # 3. Trim whitespace
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    print("    Step 3 — Whitespace trimmed")

    # 4. Fix types
    for col in df.columns:
        if df[col].dtype == object:
            converted = pd.to_numeric(df[col].astype(str).str.replace(",", ""), errors="coerce")
            if converted.notna().mean() > 0.8:
                df[col] = converted

    print("    Step 4 — Data types fixed")

    # 5. Fill missing numerics
    for col in df.select_dtypes(include="number").columns:
        if df[col].isna().any():
            df[col] = df[col].fillna(df[col].median())
    print("    Step 5 — Missing numeric values filled with median")

    return _save(df, output_path)
