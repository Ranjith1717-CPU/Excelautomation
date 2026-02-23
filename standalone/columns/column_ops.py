"""
=============================================================================
COLUMN OPERATIONS MODULE
=============================================================================
Manipulate columns in Excel files.

Functions:
  rename_columns             - Rename columns via a mapping dict or interactive
  merge_columns              - Concatenate two or more columns into one
  split_column               - Split one column into multiple by delimiter
  reorder_columns            - Specify the desired column order
  drop_columns               - Remove unwanted columns
  add_calculated_column      - Add a new column from a formula/expression
  extract_from_column        - Regex extraction from a column
  map_column_values          - Replace values using a mapping dict
  pivot_column_to_rows       - Expand a multi-value cell column to rows
  normalize_column_names     - Standardize all header names
=============================================================================
"""
import pandas as pd
import numpy as np
import re
from pathlib import Path
from typing import List, Optional, Dict, Union


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows × {len(df.columns)} cols)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Result") -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, sheet_name=sheet_name, index=False)
    print(f"    Saved   : {output_path}  ({len(df):,} rows × {len(df.columns)} cols)")
    return output_path


# ── public API ────────────────────────────────────────────────────────────────

def rename_columns(file: str, mapping: Dict[str, str], output_path: str) -> str:
    """
    Rename columns using a dict mapping {old_name: new_name}.

    Args:
        mapping: e.g. {"Emp Name": "Employee_Name", "Sal": "Salary"}
    """
    df = _load(file)
    actual_mapping = {k: v for k, v in mapping.items() if k in df.columns}
    missing = [k for k in mapping if k not in df.columns]
    if missing:
        print(f"    Warning : Column(s) not found: {missing}")
    df = df.rename(columns=actual_mapping)
    print(f"    Renamed : {len(actual_mapping)} column(s)")
    return _save(df, output_path)


def merge_columns(file: str, columns: List[str],
                   new_column_name: str, output_path: str,
                   separator: str = " ",
                   drop_originals: bool = False) -> str:
    """
    Concatenate multiple columns into one new column.

    Args:
        columns          : List of column names to concatenate.
        separator        : String placed between values (default: space).
        drop_originals   : Remove the source columns after merging.
    """
    df = _load(file)
    available = [c for c in columns if c in df.columns]
    if not available:
        raise ValueError(f"None of the specified columns found: {columns}")

    df[new_column_name] = df[available].astype(str).apply(
        lambda row: separator.join(v for v in row if v not in ("nan", "None", "")),
        axis=1
    )

    if drop_originals:
        df = df.drop(columns=available)

    print(f"    Merged  : {available} → '{new_column_name}'")
    return _save(df, output_path)


def split_column(file: str, column: str, delimiter: str,
                  new_column_names: Optional[List[str]] = None,
                  output_path: str = "",
                  drop_original: bool = False,
                  expand_all: bool = True) -> str:
    """
    Split a column by a delimiter into multiple columns.

    Args:
        delimiter        : Character(s) to split on.
        new_column_names : Names for the new columns. Auto-generated if None.
        expand_all       : If True, create one column per split part.
    """
    df = _load(file)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found")

    split_data = df[column].astype(str).str.split(delimiter, expand=expand_all)

    if expand_all:
        n_parts = split_data.shape[1]
        if new_column_names and len(new_column_names) == n_parts:
            split_data.columns = new_column_names
        else:
            split_data.columns = [f"{column}_part{i+1}" for i in range(n_parts)]

        for col in split_data.columns:
            df[col] = split_data[col].str.strip()

        print(f"    Split   : '{column}' → {list(split_data.columns)}")
    else:
        df[f"{column}_split"] = split_data

    if drop_original:
        df = df.drop(columns=[column])

    return _save(df, output_path)


def reorder_columns(file: str, column_order: List[str],
                     output_path: str,
                     put_rest_at_end: bool = True) -> str:
    """
    Reorder columns to a specified order.

    Args:
        column_order  : List of column names in desired order.
        put_rest_at_end: Append any columns not listed to the end.
    """
    df = _load(file)
    ordered = [c for c in column_order if c in df.columns]
    missing = [c for c in column_order if c not in df.columns]
    if missing:
        print(f"    Warning : Column(s) not found: {missing}")

    if put_rest_at_end:
        rest = [c for c in df.columns if c not in ordered]
        final_order = ordered + rest
    else:
        final_order = ordered

    df = df[final_order]
    print(f"    Reordered to: {final_order}")
    return _save(df, output_path)


def drop_columns(file: str, columns: List[str], output_path: str) -> str:
    """
    Remove specified columns from the file.
    """
    df = _load(file)
    existing = [c for c in columns if c in df.columns]
    not_found = [c for c in columns if c not in df.columns]
    if not_found:
        print(f"    Warning : Column(s) not found: {not_found}")
    df = df.drop(columns=existing)
    print(f"    Dropped : {existing}")
    return _save(df, output_path)


def add_calculated_column(file: str, new_column_name: str,
                            expression: str, output_path: str) -> str:
    """
    Add a new column using a Python/pandas expression.

    Args:
        expression: A pandas eval() expression.
                    Use column names directly. Arithmetic and comparison operators work.
                    Examples:
                      "Salary * 1.10"
                      "Revenue - Cost"
                      "Units * Unit_Price"
                      "Score / Max_Score * 100"

    The expression is evaluated using df.eval() with the dataframe columns in scope.
    """
    df = _load(file)
    try:
        df[new_column_name] = df.eval(expression)
        print(f"    Added   : '{new_column_name}' = {expression}")
    except Exception as e:
        raise ValueError(f"Expression evaluation failed: {e}\n"
                         f"Available columns: {list(df.columns)}")
    return _save(df, output_path)


def extract_from_column(file: str, source_column: str,
                         pattern: str, output_path: str,
                         new_column_name: Optional[str] = None,
                         group: int = 0) -> str:
    """
    Extract text from a column using a regex pattern.

    Args:
        pattern         : Python regex pattern.
        new_column_name : Name for the extracted column. Defaults to source_column + '_extracted'.
        group           : Capture group number (0 = whole match, 1+ = group).
    """
    df = _load(file)
    if source_column not in df.columns:
        raise ValueError(f"Column '{source_column}' not found")

    out_col = new_column_name or f"{source_column}_extracted"

    if group == 0:
        df[out_col] = df[source_column].astype(str).str.extract(f"({pattern})", expand=False)
    else:
        extracted = df[source_column].astype(str).str.extract(pattern, expand=True)
        df[out_col] = extracted.iloc[:, group - 1] if group <= extracted.shape[1] else None

    matched = df[out_col].notna().sum()
    print(f"    Extracted '{out_col}' from '{source_column}'  ({matched:,} matches)")
    return _save(df, output_path)


def map_column_values(file: str, column: str,
                       mapping: Dict, output_path: str,
                       unmapped_strategy: str = "keep") -> str:
    """
    Replace values in a column using a lookup dict.

    Args:
        mapping           : {old_value: new_value}
        unmapped_strategy : 'keep' (original) | 'null' (NaN) | 'other' (literal "Other")
    """
    df = _load(file)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found")

    before = df[column].copy()
    df[column] = df[column].map(mapping)

    if unmapped_strategy == "keep":
        null_mask = df[column].isna() & before.notna()
        df.loc[null_mask, column] = before[null_mask]
    elif unmapped_strategy == "other":
        df[column] = df[column].fillna("Other")
    # 'null' → leave as NaN

    changed = (before != df[column]).sum()
    print(f"    Mapped  : {changed:,} value(s) in '{column}'")
    return _save(df, output_path)


def pivot_column_to_rows(file: str, column: str, output_path: str,
                          delimiter: str = ",") -> str:
    """
    Expand a column where cells contain multiple values separated by a delimiter
    into separate rows (one value per row).

    Example:
        Row with Tags = "sales,marketing,hr"  →  3 rows: sales | marketing | hr
    """
    df = _load(file)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found")

    expanded = (
        df.assign(**{column: df[column].astype(str).str.split(delimiter)})
          .explode(column)
    )
    expanded[column] = expanded[column].str.strip()
    expanded = expanded.reset_index(drop=True)

    print(f"    Expanded: {len(df):,} rows → {len(expanded):,} rows (column '{column}')")
    return _save(expanded, output_path)


def normalize_column_names(file: str, output_path: str,
                            style: str = "snake_case") -> str:
    """
    Standardize all column names to a consistent format.

    Args:
        style: 'snake_case' (my_column) | 'title_case' (My Column) |
               'upper'  (MY_COLUMN)    | 'lower' (my column)
    """
    df = _load(file)
    original = list(df.columns)

    def to_snake(name: str) -> str:
        name = re.sub(r"[^\w\s]", "", str(name))
        name = re.sub(r"\s+", "_", name.strip())
        name = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1_\2", name)
        name = re.sub(r"([a-z\d])([A-Z])", r"\1_\2", name)
        return name.lower()

    if style == "snake_case":
        df.columns = [to_snake(c) for c in df.columns]
    elif style == "title_case":
        df.columns = [re.sub(r"[_]+", " ", str(c)).title() for c in df.columns]
    elif style == "upper":
        df.columns = [to_snake(c).upper() for c in df.columns]
    elif style == "lower":
        df.columns = [str(c).lower().replace(" ", "_") for c in df.columns]

    renamed = sum(o != n for o, n in zip(original, df.columns))
    print(f"    Renamed : {renamed} column(s) to '{style}' format")
    return _save(df, output_path)
