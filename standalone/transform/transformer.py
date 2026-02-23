"""
=============================================================================
TRANSFORMER MODULE
=============================================================================
Reshape and restructure Excel data.

Functions:
  create_pivot_table         - Excel-style pivot table
  unpivot_data               - Melt wide → long format
  transpose_data             - Flip rows and columns
  split_by_column_value      - One file per unique value in a column
  split_sheets_to_files      - Each sheet → separate Excel file
  split_file_by_rows         - Break large file into N-row chunks
  reshape_wide_to_long       - Wide (repeated columns) → long format
  reshape_long_to_wide       - Long → wide (crosstab)
  add_running_total          - Cumulative sum column
  rank_column                - Rank rows by a numeric column
=============================================================================
"""
import pandas as pd
from pathlib import Path
from typing import List, Optional, Dict, Union


# ── helpers ──────────────────────────────────────────────────────────────────

def _load(file: str, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet)
    print(f"    Loaded  : {Path(file).name}  ({len(df):,} rows × {len(df.columns)} cols)")
    return df


def _save(df: pd.DataFrame, output_path: str, sheet_name: str = "Transformed") -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, sheet_name=sheet_name, index=False)
    print(f"    Saved   : {output_path}  ({len(df):,} rows × {len(df.columns)} cols)")
    return output_path


# ── public API ────────────────────────────────────────────────────────────────

def create_pivot_table(file: str,
                        index_cols: List[str],
                        value_cols: List[str],
                        output_path: str,
                        columns_col: Optional[str] = None,
                        aggfunc: str = "sum") -> str:
    """
    Create an Excel-style pivot table.

    Args:
        index_cols : Row grouping columns.
        value_cols : Columns to aggregate.
        columns_col: Optional column to spread as new columns (like Excel column area).
        aggfunc    : 'sum' | 'mean' | 'count' | 'min' | 'max'
    """
    df = _load(file)
    func_map = {"sum": "sum", "mean": "mean", "count": "count",
                "min": "min", "max": "max"}
    fn = func_map.get(aggfunc, "sum")

    pivot = df.pivot_table(
        index=index_cols,
        columns=columns_col,
        values=value_cols,
        aggfunc=fn,
        fill_value=0,
    ).reset_index()

    # Flatten multi-level columns
    pivot.columns = [" - ".join([str(c) for c in col]).strip(" - ")
                     if isinstance(col, tuple) else col
                     for col in pivot.columns]

    return _save(pivot, output_path, sheet_name="Pivot")


def unpivot_data(file: str,
                  id_columns: List[str],
                  value_columns: List[str],
                  output_path: str,
                  var_name: str = "Variable",
                  value_name: str = "Value") -> str:
    """
    Melt (unpivot) from wide to long format.

    Args:
        id_columns   : Columns to keep as-is (identifier columns).
        value_columns: Columns to unpivot into rows.
    """
    df = _load(file)
    melted = df.melt(id_vars=id_columns,
                     value_vars=value_columns,
                     var_name=var_name,
                     value_name=value_name)
    print(f"    Unpivoted {len(value_columns)} columns → {len(melted):,} rows")
    return _save(melted, output_path, sheet_name="Unpivoted")


def transpose_data(file: str, output_path: str,
                   header_col: Optional[str] = None) -> str:
    """
    Flip rows and columns (transpose).

    Args:
        header_col: If set, use this column as the new header row after transposing.
    """
    df = _load(file)
    if header_col and header_col in df.columns:
        df = df.set_index(header_col)

    transposed = df.T.reset_index()
    transposed.columns = ["Field"] + list(transposed.columns[1:])
    print(f"    Transposed: {df.shape} → {transposed.shape}")
    return _save(transposed, output_path, sheet_name="Transposed")


def split_by_column_value(file: str, split_column: str,
                           output_dir: str) -> List[str]:
    """
    Split a file into separate Excel files, one per unique value in split_column.

    Returns:
        List of created file paths.
    """
    df = _load(file)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    unique_vals = df[split_column].dropna().unique()
    created = []
    for val in unique_vals:
        subset = df[df[split_column] == val].reset_index(drop=True)
        safe_val = str(val).replace("/", "-").replace("\\", "-").replace(":", "-")
        out_path = str(out_dir / f"{safe_val}.xlsx")
        subset.to_excel(out_path, index=False)
        created.append(out_path)
        print(f"    Created : {safe_val}.xlsx  ({len(subset):,} rows)")

    print(f"\n    Total   : {len(created)} file(s) created in {output_dir}")
    return created


def split_sheets_to_files(file: str, output_dir: str) -> List[str]:
    """
    Extract each sheet from an Excel file into its own separate file.

    Returns:
        List of created file paths.
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    xl = pd.ExcelFile(file)
    created = []

    for sheet in xl.sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        safe_name = sheet.replace("/", "-").replace("\\", "-").replace(":", "-")
        out_path = str(out_dir / f"{safe_name}.xlsx")
        df.to_excel(out_path, index=False)
        created.append(out_path)
        print(f"    Created : {safe_name}.xlsx  ({len(df):,} rows)")

    print(f"\n    Total   : {len(created)} file(s) created in {output_dir}")
    return created


def split_file_by_rows(file: str, chunk_size: int,
                        output_dir: str) -> List[str]:
    """
    Break a large Excel file into smaller chunks of chunk_size rows each.

    Returns:
        List of created file paths.
    """
    df = _load(file)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    stem = Path(file).stem
    created = []
    total_chunks = (len(df) + chunk_size - 1) // chunk_size

    for i, start in enumerate(range(0, len(df), chunk_size), 1):
        chunk = df.iloc[start:start + chunk_size].reset_index(drop=True)
        out_path = str(out_dir / f"{stem}_part{i:03d}_of_{total_chunks:03d}.xlsx")
        chunk.to_excel(out_path, index=False)
        created.append(out_path)
        print(f"    Part {i:03d} : {out_path}  ({len(chunk):,} rows)")

    print(f"\n    Total   : {len(created)} file(s) created")
    return created


def reshape_wide_to_long(file: str,
                          stub_names: List[str],
                          output_path: str) -> str:
    """
    Convert wide-format data with repeated column patterns to long format.
    E.g. Q1_Sales, Q2_Sales, Q3_Sales → Quarter, Sales columns.

    Args:
        stub_names: Base column name stubs (e.g., ['Sales', 'Cost'])
    """
    df = _load(file)
    long_df = pd.wide_to_long(df, stubnames=stub_names,
                               i=df.index.tolist(),
                               j="Period",
                               sep="_").reset_index()
    print(f"    Reshaped: {df.shape} → {long_df.shape}")
    return _save(long_df, output_path, sheet_name="Long_Format")


def reshape_long_to_wide(file: str, index_cols: List[str],
                          columns_col: str, values_col: str,
                          output_path: str,
                          aggfunc: str = "sum") -> str:
    """
    Convert long-format data to wide format (crosstab / unstack).

    Args:
        index_cols : Row identifier columns.
        columns_col: Column whose values become new column headers.
        values_col : Column whose values fill the cells.
    """
    df = _load(file)
    wide = df.pivot_table(index=index_cols,
                           columns=columns_col,
                           values=values_col,
                           aggfunc=aggfunc,
                           fill_value=0).reset_index()
    wide.columns.name = None
    print(f"    Reshaped: {df.shape} → {wide.shape}")
    return _save(wide, output_path, sheet_name="Wide_Format")


def add_running_total(file: str, value_col: str,
                       output_path: str,
                       group_col: Optional[str] = None) -> str:
    """
    Add a cumulative sum (running total) column.

    Args:
        group_col: If provided, reset running total for each group.
    """
    df = _load(file)
    values = pd.to_numeric(df[value_col], errors="coerce")
    col_name = f"Running_Total_{value_col}"

    if group_col and group_col in df.columns:
        df[col_name] = df.groupby(group_col)[value_col].cumsum()
    else:
        df[col_name] = values.cumsum()

    print(f"    Running total added as '{col_name}'  (grand total: {df[col_name].iloc[-1]:,.2f})")
    return _save(df, output_path)


def rank_column(file: str, value_col: str, output_path: str,
                ascending: bool = False,
                group_col: Optional[str] = None) -> str:
    """
    Add a rank column based on a numeric column.

    Args:
        ascending: False = highest value = rank 1 (default for sales/performance).
        group_col : If set, rank within each group.
    """
    df = _load(file)
    rank_col = f"Rank_{value_col}"

    if group_col and group_col in df.columns:
        df[rank_col] = df.groupby(group_col)[value_col].rank(
            ascending=ascending, method="min"
        ).astype(int)
    else:
        df[rank_col] = df[value_col].rank(ascending=ascending, method="min").astype(int)

    print(f"    Rank column '{rank_col}' added  (ascending={ascending})")
    return _save(df, output_path)
