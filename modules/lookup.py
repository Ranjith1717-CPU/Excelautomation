"""
=============================================================================
LOOKUP MODULE
=============================================================================
Excel VLOOKUP-style and fuzzy matching operations.

Functions:
  vlookup            - pandas-based VLOOKUP
  fuzzy_match        - difflib approximate string matching
  multi_key_lookup   - Multi-column JOIN lookup
  reverse_lookup     - Find key by value
  enrich_from_lookup - Enrich master with reference columns
=============================================================================
"""
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Optional
import difflib


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Lookup") -> str:
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

def vlookup(file: str, lookup_col: str,
            ref_file: str, ref_col: str,
            return_cols: List[str],
            output_path: str,
            how: str = "left") -> str:
    """
    pandas-based VLOOKUP: join return_cols from ref_file into main file.
    how: 'left' (keep all from main), 'inner' (only matches), 'outer'.
    """
    df = _load(file)
    ref_df = _load(ref_file)
    print(f"    Loaded  : {Path(ref_file).name}  ({len(ref_df):,} rows) [reference]")

    # Only keep ref_col + return_cols from reference
    keep_cols = [ref_col] + [c for c in return_cols if c in ref_df.columns and c != ref_col]
    ref_sub = ref_df[keep_cols].drop_duplicates(subset=[ref_col])

    merged = pd.merge(df, ref_sub, left_on=lookup_col, right_on=ref_col, how=how, suffixes=("", "_ref"))

    matched = merged[ref_col].notna().sum() if ref_col in merged.columns else len(merged)
    print(f"    Matched : {matched}/{len(df)} rows")

    return _save(merged, output_path, "VLOOKUP_Result")


def fuzzy_match(file: str, col: str,
                ref_file: str, ref_col: str,
                output_path: str,
                threshold: float = 0.75) -> str:
    """
    Fuzzy string matching using difflib SequenceMatcher.
    threshold: 0.0–1.0 (e.g. 0.75 = 75% similarity).
    Adds 'Best_Match', 'Match_Score', 'Is_Match' columns.

    Note: For large datasets, install rapidfuzz for much better performance:
          pip install rapidfuzz
    """
    df = _load(file)
    ref_df = pd.read_excel(ref_file)
    print(f"    Loaded  : {Path(ref_file).name}  ({len(ref_df):,} rows) [reference]")

    ref_values = ref_df[ref_col].dropna().astype(str).tolist()

    # Try to use rapidfuzz if available, fall back to difflib
    try:
        from rapidfuzz import process as rf_process, fuzz as rf_fuzz
        def _best_match(val):
            if pd.isna(val) or str(val).strip() == "":
                return "", 0.0
            result = rf_process.extractOne(str(val), ref_values, scorer=rf_fuzz.ratio)
            if result:
                match, score, _ = result
                return match, round(score / 100, 4)
            return "", 0.0
        print("    Engine  : rapidfuzz (fast)")
    except ImportError:
        def _best_match(val):
            if pd.isna(val) or str(val).strip() == "":
                return "", 0.0
            best = max(ref_values,
                       key=lambda r: difflib.SequenceMatcher(None, str(val).lower(), r.lower()).ratio())
            score = difflib.SequenceMatcher(None, str(val).lower(), best.lower()).ratio()
            return best, round(score, 4)
        print("    Engine  : difflib (install rapidfuzz for speed boost)")

    print(f"    Matching: {len(df)} values against {len(ref_values)} reference entries...")
    results = df[col].apply(_best_match)
    df["Best_Match"] = results.apply(lambda x: x[0])
    df["Match_Score"] = results.apply(lambda x: x[1])
    df["Is_Match"] = df["Match_Score"] >= threshold

    matched = df["Is_Match"].sum()
    print(f"    Matched : {matched}/{len(df)} rows above threshold {threshold}")

    no_match = df[~df["Is_Match"]].copy()
    return _save_multi({"All_Results": df, "No_Match": no_match}, output_path)


def multi_key_lookup(file: str, lookup_cols: List[str],
                     ref_file: str, output_path: str,
                     how: str = "left") -> str:
    """
    Multi-column JOIN: match on all lookup_cols simultaneously (same columns must exist in ref_file).
    Returns enriched dataframe.
    """
    df = _load(file)
    ref_df = pd.read_excel(ref_file)
    print(f"    Loaded  : {Path(ref_file).name}  ({len(ref_df):,} rows) [reference]")

    merged = pd.merge(df, ref_df, on=lookup_cols, how=how, suffixes=("", "_ref"))
    print(f"    Merged  : {len(merged)} rows  (how='{how}')")
    return _save(merged, output_path, "MultiKey_Lookup")


def reverse_lookup(file: str, value_col: str,
                   ref_file: str, key_col: str, val_col: str,
                   output_path: str) -> str:
    """
    Reverse lookup: given values, find their corresponding keys.
    Looks up value_col in val_col of ref_file, returns key_col.
    """
    df = _load(file)
    ref_df = pd.read_excel(ref_file)
    print(f"    Loaded  : {Path(ref_file).name}  ({len(ref_df):,} rows) [reference]")

    lookup_map = ref_df.set_index(val_col)[key_col].to_dict()
    df[f"Lookup_{key_col}"] = df[value_col].map(lookup_map)

    not_found = df[f"Lookup_{key_col}"].isna().sum()
    print(f"    Matched : {len(df) - not_found}/{len(df)} rows  ({not_found} not found)")
    return _save(df, output_path, "Reverse_Lookup")


def enrich_from_lookup(file: str, join_col: str,
                       ref_file: str, ref_join_col: str,
                       enrich_cols: List[str],
                       output_path: str) -> str:
    """
    Enrich main file with specific columns from a reference file.
    join_col (main) is matched to ref_join_col (reference).
    enrich_cols: columns from ref_file to bring into main file.
    """
    df = _load(file)
    ref_df = pd.read_excel(ref_file)
    print(f"    Loaded  : {Path(ref_file).name}  ({len(ref_df):,} rows) [reference]")

    available = [c for c in enrich_cols if c in ref_df.columns]
    missing_e = [c for c in enrich_cols if c not in ref_df.columns]
    if missing_e:
        print(f"    Warning : Enrich columns not in ref: {missing_e}")

    keep_cols = [ref_join_col] + available
    ref_sub = ref_df[keep_cols].drop_duplicates(subset=[ref_join_col])

    merged = pd.merge(df, ref_sub, left_on=join_col, right_on=ref_join_col,
                      how="left", suffixes=("", "_enriched"))

    # Drop duplicate ref join column if it was added
    if ref_join_col != join_col and ref_join_col in merged.columns:
        merged.drop(columns=[ref_join_col], inplace=True)

    enriched = merged[list(df.columns) + [c for c in available if c not in df.columns]]
    matched = enriched[available[0]].notna().sum() if available else 0
    print(f"    Enriched: {matched}/{len(df)} rows enriched with {len(available)} column(s)")
    return _save(enriched, output_path, "Enriched")
