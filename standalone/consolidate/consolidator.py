"""
=============================================================================
CONSOLIDATOR MODULE
=============================================================================
Handles merging and combining multiple Excel files in various ways.

Functions:
  merge_files_stack          - Stack N files vertically (append rows)
  merge_files_by_key         - Join N files horizontally on a key column
  merge_specific_columns     - Pull selected columns from multiple files
  merge_sheets_in_file       - Consolidate all sheets of one file into one
  merge_same_sheet_cross_files - Merge same-named sheet across N files
=============================================================================
"""
import pandas as pd
from pathlib import Path
from typing import List, Optional


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    """Read an Excel file and return a DataFrame."""
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows × {len(df.columns)} cols)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Consolidated") -> str:
    """Save DataFrame to Excel and report the result."""
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, sheet_name=sheet_name, index=False)
    print(f"\n    Saved   : {output_path}  ({len(df):,} rows × {len(df.columns)} cols)")
    return output_path


# ── public API ────────────────────────────────────────────────────────────────

def merge_files_stack(files: List[str], output_path: str,
                      add_source_column: bool = True) -> str:
    """
    Stack multiple Excel files vertically (append rows).
    Files may have different columns – missing values become NaN.

    Args:
        files           : List of Excel file paths.
        output_path     : Where to save the result.
        add_source_column: Add a '_Source_File' column with filename.

    Returns:
        output_path
    """
    dfs = []
    for f in files:
        try:
            df = _load(f)
            if add_source_column:
                df["_Source_File"] = Path(f).name
            dfs.append(df)
        except Exception as e:
            print(f"    SKIP    : {Path(f).name} — {e}")

    if not dfs:
        raise ValueError("No files could be loaded.")

    combined = pd.concat(dfs, ignore_index=True, sort=False)
    return _save(combined, output_path)


def merge_files_by_key(files: List[str], key_column: str,
                       join_type: str = "outer",
                       output_path: str = "") -> str:
    """
    Join multiple Excel files horizontally on a shared key column (SQL-style JOIN).

    Args:
        files      : List of Excel file paths.
        key_column : Column name to join on (must exist in all files).
        join_type  : 'inner' | 'outer' | 'left' | 'right'
        output_path: Where to save the result.
    """
    if not files:
        raise ValueError("No files provided.")

    base = _load(files[0])

    for f in files[1:]:
        df = _load(f)
        suffix = f"_{Path(f).stem}"
        base = base.merge(df, on=key_column, how=join_type,
                          suffixes=("", suffix))

    return _save(base, output_path)


def merge_specific_columns(files: List[str], columns: List[str],
                           output_path: str,
                           add_source: bool = True) -> str:
    """
    Extract a specific set of columns from multiple files and stack them.

    Args:
        files      : List of Excel file paths.
        columns    : Column names to extract (case-sensitive).
        output_path: Where to save the result.
        add_source : Append a '_Source_File' column.
    """
    dfs = []
    for f in files:
        try:
            df = _load(f)
            available = [c for c in columns if c in df.columns]
            if not available:
                print(f"    SKIP    : {Path(f).name} — none of the columns found")
                continue
            sub = df[available].copy()
            if add_source:
                sub["_Source_File"] = Path(f).name
            dfs.append(sub)
        except Exception as e:
            print(f"    SKIP    : {Path(f).name} — {e}")

    if not dfs:
        raise ValueError("No matching columns found in any file.")

    combined = pd.concat(dfs, ignore_index=True, sort=False)
    return _save(combined, output_path)


def merge_sheets_in_file(file_path: str, output_path: str,
                          add_sheet_column: bool = True) -> str:
    """
    Consolidate ALL sheets within a single Excel file into one sheet.

    Args:
        file_path       : Path to the Excel file.
        output_path     : Where to save the result.
        add_sheet_column: Add a '_Sheet_Name' column.
    """
    xl = pd.ExcelFile(file_path)
    dfs = []

    for sheet in xl.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet)
        if add_sheet_column:
            df["_Sheet_Name"] = sheet
        dfs.append(df)
        print(f"    Sheet   : '{sheet}'  ({len(df):,} rows)")

    combined = pd.concat(dfs, ignore_index=True, sort=False)
    return _save(combined, output_path)


def merge_same_sheet_cross_files(files: List[str], sheet_name: str,
                                  output_path: str,
                                  add_source: bool = True) -> str:
    """
    Extract and stack the same sheet from multiple Excel files.

    Args:
        files      : List of Excel file paths.
        sheet_name : Sheet tab name to extract from each file.
        output_path: Where to save the result.
        add_source : Append a '_Source_File' column.
    """
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f, sheet_name=sheet_name)
            if add_source:
                df["_Source_File"] = Path(f).name
            dfs.append(df)
            print(f"    Loaded  : {Path(f).name} → '{sheet_name}'  ({len(df):,} rows)")
        except Exception as e:
            print(f"    SKIP    : {Path(f).name} — {e}")

    if not dfs:
        raise ValueError(f"Sheet '{sheet_name}' not found in any file.")

    combined = pd.concat(dfs, ignore_index=True, sort=False)
    return _save(combined, output_path)
