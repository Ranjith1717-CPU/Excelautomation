"""
=============================================================================
VALIDATOR MODULE
=============================================================================
Data quality validation and flagging for Excel files.

Functions:
  check_mandatory_fields  - Flag rows missing required values
  validate_email          - Regex email validation
  validate_phone          - Phone format check
  validate_numeric_range  - Out-of-range flag
  validate_date_range     - Date boundary check
  referential_integrity   - Values must exist in a lookup
  data_quality_report     - Comprehensive quality score 0-100
  detect_pii              - Flag columns likely containing PII
=============================================================================
"""
import pandas as pd
import numpy as np
import re
from pathlib import Path
from typing import List, Optional, Union


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Validated") -> str:
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

def check_mandatory_fields(file: str, required_cols: List[str],
                            output_path: str) -> str:
    """
    Flag rows with missing values in any of the required columns.
    Adds 'Missing_Fields' column listing which fields are blank.
    Adds 'Has_Missing' boolean column.
    """
    df = _load(file)
    present_cols = [c for c in required_cols if c in df.columns]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        print(f"    Warning : Columns not found in file: {missing_cols}")

    def _missing_list(row):
        return ", ".join([c for c in present_cols if pd.isna(row[c]) or str(row[c]).strip() == ""])

    df["Missing_Fields"] = df.apply(_missing_list, axis=1)
    df["Has_Missing"] = df["Missing_Fields"] != ""

    flagged = df[df["Has_Missing"]].copy()
    print(f"    Found   : {len(flagged)} rows with missing mandatory fields")

    summary = {
        "Total_Rows": len(df),
        "Rows_With_Missing": len(flagged),
        "Missing_%": round(len(flagged) / len(df) * 100, 2) if len(df) else 0,
    }
    for c in present_cols:
        count = (df[c].isna() | (df[c].astype(str).str.strip() == "")).sum()
        summary[f"{c}_Missing_Count"] = int(count)

    summary_df = pd.DataFrame([summary])
    return _save_multi({"All_Data": df, "Missing_Rows": flagged, "Summary": summary_df}, output_path)


def validate_email(file: str, email_col: str, output_path: str) -> str:
    """
    Validate email addresses using regex.
    Adds 'Email_Valid' boolean column and 'Email_Error' column.
    """
    EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")
    df = _load(file)

    def _check(val):
        if pd.isna(val) or str(val).strip() == "":
            return False, "Empty"
        if EMAIL_RE.match(str(val).strip()):
            return True, ""
        return False, "Invalid format"

    results = df[email_col].apply(_check)
    df["Email_Valid"] = results.apply(lambda x: x[0])
    df["Email_Error"] = results.apply(lambda x: x[1])

    invalid = df[~df["Email_Valid"]].copy()
    print(f"    Found   : {len(invalid)} invalid email(s)")
    return _save_multi({"All_Data": df, "Invalid_Emails": invalid}, output_path)


def validate_phone(file: str, phone_col: str, output_path: str,
                   pattern: str = r"^[\+]?[\d\s\-\(\)]{7,15}$") -> str:
    """
    Phone number format validation.
    Default pattern accepts international formats with 7-15 digits.
    """
    PHONE_RE = re.compile(pattern)
    df = _load(file)

    def _check(val):
        if pd.isna(val) or str(val).strip() == "":
            return False, "Empty"
        cleaned = re.sub(r"[\s\-\(\)]", "", str(val).strip())
        if PHONE_RE.match(str(val).strip()) and 7 <= len(cleaned.lstrip("+")) <= 15:
            return True, ""
        return False, "Invalid format"

    results = df[phone_col].apply(_check)
    df["Phone_Valid"] = results.apply(lambda x: x[0])
    df["Phone_Error"] = results.apply(lambda x: x[1])

    invalid = df[~df["Phone_Valid"]].copy()
    print(f"    Found   : {len(invalid)} invalid phone number(s)")
    return _save_multi({"All_Data": df, "Invalid_Phones": invalid}, output_path)


def validate_numeric_range(file: str, col: str,
                            min_val: float, max_val: float,
                            output_path: str) -> str:
    """
    Flag rows where a numeric column is outside [min_val, max_val].
    """
    df = _load(file)
    df[col] = pd.to_numeric(df[col], errors="coerce")
    df["In_Range"] = df[col].between(min_val, max_val, inclusive="both")
    df["Range_Error"] = df.apply(
        lambda r: f"Value {r[col]} outside [{min_val}, {max_val}]" if not r["In_Range"] else "",
        axis=1
    )
    out_of_range = df[~df["In_Range"]].copy()
    print(f"    Found   : {len(out_of_range)} out-of-range value(s)")
    return _save_multi({"All_Data": df, "Out_of_Range": out_of_range}, output_path)


def validate_date_range(file: str, date_col: str,
                        min_date: str, max_date: str,
                        output_path: str) -> str:
    """
    Flag rows where a date column is outside [min_date, max_date].
    min_date / max_date: 'YYYY-MM-DD' strings.
    """
    df = _load(file)
    df["_date_parsed"] = pd.to_datetime(df[date_col], errors="coerce")
    min_dt = pd.Timestamp(min_date)
    max_dt = pd.Timestamp(max_date)

    df["Date_Valid"] = df["_date_parsed"].between(min_dt, max_dt, inclusive="both")
    df.loc[df["_date_parsed"].isna(), "Date_Valid"] = False
    df["Date_Error"] = df.apply(
        lambda r: "Cannot parse date" if pd.isna(r["_date_parsed"])
        else (f"Date {r['_date_parsed'].date()} outside [{min_date}, {max_date}]" if not r["Date_Valid"] else ""),
        axis=1
    )
    df.drop(columns=["_date_parsed"], inplace=True)

    invalid = df[~df["Date_Valid"]].copy()
    print(f"    Found   : {len(invalid)} date validation failure(s)")
    return _save_multi({"All_Data": df, "Invalid_Dates": invalid}, output_path)


def referential_integrity(file: str, col: str,
                           ref_file: str, ref_col: str,
                           output_path: str) -> str:
    """
    Check that values in col exist in ref_col of ref_file.
    Flags rows where the value is not found.
    """
    df = _load(file)
    ref_df = pd.read_excel(ref_file)
    print(f"    Loaded  : {Path(ref_file).name}  ({len(ref_df):,} rows) [reference]")

    valid_values = set(ref_df[ref_col].dropna().astype(str).str.strip())
    df["RI_Valid"] = df[col].astype(str).str.strip().isin(valid_values)
    df["RI_Error"] = df.apply(
        lambda r: f"'{r[col]}' not found in {ref_col}" if not r["RI_Valid"] else "",
        axis=1
    )
    invalid = df[~df["RI_Valid"]].copy()
    print(f"    Found   : {len(invalid)} referential integrity violation(s)")
    return _save_multi({"All_Data": df, "RI_Violations": invalid}, output_path)


def data_quality_report(file: str, output_path: str) -> str:
    """
    Comprehensive data quality report with score 0-100.
    Checks: completeness, uniqueness, format consistency, numeric validity.
    """
    df = _load(file)
    total_cells = len(df) * len(df.columns)
    report_rows = []

    for col in df.columns:
        series = df[col]
        total = len(series)
        null_count = series.isna().sum()
        empty_str_count = (series.astype(str).str.strip() == "").sum() - null_count
        blank_count = null_count + max(0, empty_str_count)
        completeness = round((1 - blank_count / total) * 100, 2) if total else 0

        unique_count = series.nunique(dropna=True)
        uniqueness_pct = round(unique_count / total * 100, 2) if total else 0

        dup_count = series.duplicated(keep=False).sum()

        # Numeric validity
        numeric_ratio = 0.0
        try:
            converted = pd.to_numeric(series.dropna(), errors="coerce")
            numeric_ratio = round(converted.notna().sum() / len(series.dropna()) * 100, 2) if len(series.dropna()) else 0.0
        except Exception:
            pass

        quality_score = round(completeness * 0.5 + min(uniqueness_pct, 100) * 0.3 + 20, 2)
        quality_score = min(quality_score, 100)

        report_rows.append({
            "Column": col,
            "Data_Type": str(series.dtype),
            "Total_Values": total,
            "Null_Count": int(null_count),
            "Blank_%": round(blank_count / total * 100, 2) if total else 0,
            "Completeness_%": completeness,
            "Unique_Values": unique_count,
            "Uniqueness_%": uniqueness_pct,
            "Duplicate_Count": int(dup_count),
            "Numeric_Validity_%": numeric_ratio,
            "Quality_Score": quality_score,
        })

    report_df = pd.DataFrame(report_rows)
    overall_score = round(report_df["Quality_Score"].mean(), 2)
    print(f"    Quality : Overall score = {overall_score}/100")

    overview = pd.DataFrame([{
        "Total_Rows": len(df),
        "Total_Columns": len(df.columns),
        "Total_Cells": total_cells,
        "Overall_Quality_Score": overall_score,
    }])

    return _save_multi({"Column_Report": report_df, "Overview": overview}, output_path)


def detect_pii(file: str, output_path: str) -> str:
    """
    Detect columns likely containing PII (email, phone, SSN, credit card, name patterns).
    Returns a report of suspected PII columns with sample values.
    """
    df = _load(file)
    EMAIL_RE   = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
    PHONE_RE   = re.compile(r"\b[\+]?[\d\s\-\(\)]{7,15}\b")
    SSN_RE     = re.compile(r"\b\d{3}[-\s]?\d{2}[-\s]?\d{4}\b")
    CC_RE      = re.compile(r"\b(?:\d{4}[-\s]?){3}\d{4}\b")
    NAME_KEYWORDS = {"name", "first", "last", "fname", "lname", "fullname", "surname"}

    pii_rows = []
    for col in df.columns:
        pii_types = []
        sample_vals = df[col].dropna().astype(str).head(20)

        # Check column name
        col_lower = col.lower().replace("_", "").replace(" ", "")
        if any(k in col_lower for k in NAME_KEYWORDS):
            pii_types.append("Possible Name")

        # Check values
        email_hits = sum(1 for v in sample_vals if EMAIL_RE.search(v))
        phone_hits = sum(1 for v in sample_vals if PHONE_RE.search(v))
        ssn_hits   = sum(1 for v in sample_vals if SSN_RE.search(v))
        cc_hits    = sum(1 for v in sample_vals if CC_RE.search(v))

        if email_hits >= 2: pii_types.append("Email")
        if phone_hits >= 2: pii_types.append("Phone")
        if ssn_hits >= 1:   pii_types.append("SSN")
        if cc_hits >= 1:    pii_types.append("Credit Card")

        if pii_types:
            pii_rows.append({
                "Column": col,
                "PII_Types_Detected": ", ".join(pii_types),
                "Sample_Values": " | ".join(list(sample_vals[:3])),
                "Risk_Level": "HIGH" if any(t in ("SSN", "Credit Card") for t in pii_types) else "MEDIUM",
            })

    pii_df = pd.DataFrame(pii_rows) if pii_rows else pd.DataFrame(
        columns=["Column", "PII_Types_Detected", "Sample_Values", "Risk_Level"]
    )
    print(f"    Detected: {len(pii_df)} column(s) with potential PII")
    return _save_multi({"PII_Report": pii_df, "Full_Data": df}, output_path)
