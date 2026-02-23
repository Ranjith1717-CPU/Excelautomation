"""
=============================================================================
COMPARATOR MODULE
=============================================================================
Compare Excel files and identify differences.

Functions:
  compare_two_files          - Side-by-side diff with highlighted changes
  find_new_rows              - Rows present in file2 but not in file1
  find_deleted_rows          - Rows present in file1 but not in file2
  find_changed_values        - Same key, different cell values
  find_duplicates_in_file    - Within-file duplicate finder
  find_common_rows           - Rows that exist in both files
  cross_file_duplicate_check - Find rows that appear in multiple files
=============================================================================
"""
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Optional


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows × {len(df.columns)} cols)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Result") -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, sheet_name=sheet_name, index=False)
    print(f"    Saved   : {output_path}  ({len(df):,} rows)")
    return output_path


def _save_multi(sheets: dict, output_path: str) -> str:
    """Save multiple DataFrames to multiple sheets in one file."""
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    print(f"    Saved   : {output_path}  ({len(sheets)} sheet(s))")
    return output_path


# ── public API ────────────────────────────────────────────────────────────────

def compare_two_files(file1: str, file2: str, output_path: str,
                       key_column: Optional[str] = None) -> str:
    """
    Full side-by-side comparison of two Excel files.
    Produces a multi-sheet report:
      - Summary     : high-level diff stats
      - New_Rows    : in file2 but not file1
      - Deleted_Rows: in file1 but not file2
      - Changed     : same key, different values
      - Unchanged   : identical rows

    Args:
        key_column: Column to use as unique identifier for matching rows.
                    If None, uses row position.
    """
    df1 = _load(file1)
    df2 = _load(file2)

    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    sheets = {}

    if key_column and key_column in df1.columns and key_column in df2.columns:
        df1_indexed = df1.set_index(key_column)
        df2_indexed = df2.set_index(key_column)

        keys1 = set(df1_indexed.index)
        keys2 = set(df2_indexed.index)

        new_keys = keys2 - keys1
        del_keys = keys1 - keys2
        common_keys = keys1 & keys2

        new_rows = df2[df2[key_column].isin(new_keys)]
        del_rows = df1[df1[key_column].isin(del_keys)]

        # Changed rows
        common_cols = [c for c in df1.columns if c in df2.columns and c != key_column]
        changed_records = []
        unchanged_count = 0

        for key in common_keys:
            row1 = df1_indexed.loc[key]
            row2 = df2_indexed.loc[key]
            diffs = {}
            for col in common_cols:
                v1 = row1.get(col, None)
                v2 = row2.get(col, None)
                if str(v1) != str(v2):
                    diffs[col] = {"old": v1, "new": v2}

            if diffs:
                record = {key_column: key}
                for col, vals in diffs.items():
                    record[f"{col} (OLD)"] = vals["old"]
                    record[f"{col} (NEW)"] = vals["new"]
                changed_records.append(record)
            else:
                unchanged_count += 1

        changed_df = pd.DataFrame(changed_records)
    else:
        # Position-based comparison
        max_rows = max(len(df1), len(df2))
        df1_padded = df1.reindex(range(max_rows))
        df2_padded = df2.reindex(range(max_rows))

        diff_mask = df1_padded.astype(str) != df2_padded.astype(str)
        changed_positions = diff_mask.any(axis=1)
        new_rows = df2.iloc[len(df1):] if len(df2) > len(df1) else pd.DataFrame()
        del_rows = df1.iloc[len(df2):] if len(df1) > len(df2) else pd.DataFrame()
        changed_df = pd.DataFrame({"Row_Index": diff_mask[changed_positions].index.tolist(),
                                   "Changed_Columns": [
                                       ", ".join(diff_mask.columns[diff_mask.loc[i]].tolist())
                                       for i in diff_mask[changed_positions].index
                                   ]})
        unchanged_count = (~changed_positions).sum()

    # Summary
    summary = pd.DataFrame([{
        "File_1"              : Path(file1).name,
        "File_2"              : Path(file2).name,
        "File1_Rows"          : len(df1),
        "File2_Rows"          : len(df2),
        "New_Rows_in_File2"   : len(new_rows),
        "Deleted_from_File1"  : len(del_rows),
        "Changed_Rows"        : len(changed_df),
        "Unchanged_Rows"      : unchanged_count,
    }])

    sheets["Summary"]      = summary
    sheets["New_Rows"]     = new_rows if len(new_rows) > 0 else pd.DataFrame({"Note": ["No new rows"]})
    sheets["Deleted_Rows"] = del_rows if len(del_rows) > 0 else pd.DataFrame({"Note": ["No deleted rows"]})
    sheets["Changed"]      = changed_df if len(changed_df) > 0 else pd.DataFrame({"Note": ["No changed values"]})

    print(f"    New     : {len(new_rows):,} rows")
    print(f"    Deleted : {len(del_rows):,} rows")
    print(f"    Changed : {len(changed_df):,} rows")
    print(f"    Same    : {unchanged_count:,} rows")

    return _save_multi(sheets, output_path)


def find_new_rows(file1: str, file2: str, output_path: str,
                   key_columns: Optional[List[str]] = None) -> str:
    """
    Return rows that exist in file2 but NOT in file1 (new additions).

    Args:
        key_columns: Columns to use as the unique identifier.
                     None = uses ALL columns (exact row match).
    """
    df1 = _load(file1)
    df2 = _load(file2)

    keys = key_columns if key_columns else df1.columns.tolist()
    keys = [k for k in keys if k in df1.columns and k in df2.columns]

    merged = df2.merge(df1[keys].drop_duplicates(), on=keys, how="left", indicator=True)
    new_rows = merged[merged["_merge"] == "left_only"].drop("_merge", axis=1)

    print(f"    New rows in {Path(file2).name}: {len(new_rows):,}")
    return _save(new_rows, output_path, sheet_name="New_Rows")


def find_deleted_rows(file1: str, file2: str, output_path: str,
                       key_columns: Optional[List[str]] = None) -> str:
    """
    Return rows that exist in file1 but NOT in file2 (deleted/removed rows).
    """
    df1 = _load(file1)
    df2 = _load(file2)

    keys = key_columns if key_columns else df1.columns.tolist()
    keys = [k for k in keys if k in df1.columns and k in df2.columns]

    merged = df1.merge(df2[keys].drop_duplicates(), on=keys, how="left", indicator=True)
    del_rows = merged[merged["_merge"] == "left_only"].drop("_merge", axis=1)

    print(f"    Deleted rows (in {Path(file1).name} but not {Path(file2).name}): {len(del_rows):,}")
    return _save(del_rows, output_path, sheet_name="Deleted_Rows")


def find_changed_values(file1: str, file2: str, key_column: str,
                         output_path: str) -> str:
    """
    Find rows with the same key but different cell values between two files.
    Output has columns: key, field_name, old_value, new_value.
    """
    df1 = _load(file1)
    df2 = _load(file2)

    df1 = df1.set_index(key_column)
    df2 = df2.set_index(key_column)

    common_keys = df1.index.intersection(df2.index)
    common_cols = [c for c in df1.columns if c in df2.columns]

    records = []
    for key in common_keys:
        for col in common_cols:
            v1 = str(df1.at[key, col])
            v2 = str(df2.at[key, col])
            if v1 != v2:
                records.append({
                    key_column  : key,
                    "Column"    : col,
                    "Old_Value" : df1.at[key, col],
                    "New_Value" : df2.at[key, col],
                })

    result = pd.DataFrame(records)
    print(f"    Changed cell(s): {len(result):,}")
    return _save(result, output_path, sheet_name="Changed_Values")


def find_duplicates_in_file(file: str, output_path: str,
                             subset: Optional[List[str]] = None,
                             keep: str = "first") -> str:
    """
    Find and report duplicate rows within a single file.

    Args:
        subset: Columns to check for duplicates. None = all columns.
        keep  : 'first' | 'last' | False (mark all duplicates)
    """
    df = _load(file)
    dup_mask = df.duplicated(subset=subset, keep=keep)
    dups = df[dup_mask].copy()
    dups.insert(0, "_Duplicate_of_Row", df[~dup_mask].reset_index().index.tolist()[:len(dups)] if keep else "")

    print(f"    Duplicates found: {len(dups):,}")
    return _save(dups, output_path, sheet_name="Duplicates")


def find_common_rows(file1: str, file2: str, output_path: str,
                      key_columns: Optional[List[str]] = None) -> str:
    """
    Return rows that exist in BOTH files (intersection).
    """
    df1 = _load(file1)
    df2 = _load(file2)

    keys = key_columns if key_columns else df1.columns.tolist()
    keys = [k for k in keys if k in df1.columns and k in df2.columns]

    common = df1.merge(df2[keys].drop_duplicates(), on=keys, how="inner")
    print(f"    Common rows (in both files): {len(common):,}")
    return _save(common, output_path, sheet_name="Common_Rows")


def cross_file_duplicate_check(files: List[str], key_columns: List[str],
                                 output_path: str) -> str:
    """
    Find rows (by key columns) that appear in more than one file.
    Useful for detecting data entry duplicates across multiple sources.
    """
    dfs = []
    for f in files:
        df = _load(f)
        df["_Source_File"] = Path(f).name
        dfs.append(df)

    combined = pd.concat(dfs, ignore_index=True)
    dup_keys = combined.duplicated(subset=key_columns, keep=False)
    duplicates = combined[dup_keys].sort_values(key_columns)

    print(f"    Cross-file duplicates found: {len(duplicates):,} rows across {len(files)} files")
    return _save(duplicates, output_path, sheet_name="Cross_File_Dups")
