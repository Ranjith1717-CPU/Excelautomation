"""
=============================================================================
CONVERTER MODULE
=============================================================================
File format conversion utilities for Excel data.

Functions:
  excel_to_csv        - Each sheet → separate CSV file
  csv_to_excel        - Multiple CSVs → one Excel (each = one sheet)
  excel_to_json       - All sheets → JSON (records format)
  json_to_excel       - JSON array → Excel
  xls_to_xlsx_batch   - Batch convert .xls → .xlsx
  excel_to_text       - Tab/pipe/custom delimited text export
  merge_csv_files     - Stack multiple CSVs → one Excel
=============================================================================
"""
import pandas as pd
import json
from pathlib import Path
from typing import List, Optional


# ── helpers ──────────────────────────────────────────────────────────────────

def _ensure_dir(path: str) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


# ── public API ────────────────────────────────────────────────────────────────

def excel_to_csv(file: str, output_dir: str,
                 encoding: str = "utf-8-sig") -> List[str]:
    """
    Export every sheet in an Excel file to a separate CSV.
    Returns list of created CSV paths.
    """
    out_dir = _ensure_dir(output_dir)
    xl = pd.ExcelFile(file)
    stem = Path(file).stem
    created = []
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" for c in sheet)
        csv_path = str(out_dir / f"{stem}_{safe_name}.csv")
        df.to_csv(csv_path, index=False, encoding=encoding)
        print(f"    Saved   : {csv_path}  ({len(df):,} rows)")
        created.append(csv_path)
    return created


def csv_to_excel(csv_files: List[str], output_path: str,
                 encoding: str = "utf-8-sig") -> str:
    """
    Merge multiple CSV files into one Excel workbook.
    Each CSV becomes one sheet named after the CSV filename.
    """
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for csv_file in csv_files:
            try:
                df = pd.read_csv(csv_file, encoding=encoding)
            except UnicodeDecodeError:
                df = pd.read_csv(csv_file, encoding="latin-1")
            sheet_name = Path(csv_file).stem[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"    Added   : sheet '{sheet_name}'  ({len(df):,} rows)")
    print(f"    Saved   : {output_path}")
    return output_path


def excel_to_json(file: str, output_path: str,
                  orient: str = "records") -> str:
    """
    Export all sheets from an Excel file to a JSON file.
    Output format: { "Sheet1": [...records...], "Sheet2": [...] }
    """
    xl = pd.ExcelFile(file)
    result = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        # Convert dates to strings for JSON serialization
        for col in df.select_dtypes(include=["datetime64", "datetimetz"]).columns:
            df[col] = df[col].astype(str)
        result[sheet] = json.loads(df.to_json(orient=orient, date_format="iso"))

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    total_rows = sum(len(v) for v in result.values())
    print(f"    Saved   : {output_path}  ({len(xl.sheet_names)} sheet(s), {total_rows:,} total records)")
    return output_path


def json_to_excel(json_file: str, output_path: str) -> str:
    """
    Convert a JSON file to Excel.
    Supports:
      - JSON array of objects → single sheet 'Data'
      - JSON object of arrays → each key becomes a sheet
    """
    with open(json_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    if isinstance(data, list):
        df = pd.DataFrame(data)
        df.to_excel(output_path, sheet_name="Data", index=False)
        print(f"    Saved   : {output_path}  ({len(df):,} rows)")
    elif isinstance(data, dict):
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for key, rows in data.items():
                df = pd.DataFrame(rows) if isinstance(rows, list) else pd.DataFrame([rows])
                df.to_excel(writer, sheet_name=str(key)[:31], index=False)
                print(f"    Added   : sheet '{key}'  ({len(df):,} rows)")
        print(f"    Saved   : {output_path}")
    else:
        raise ValueError("JSON must be a list (array) or a dict of arrays")

    return output_path


def xls_to_xlsx_batch(files: List[str], output_dir: str) -> List[str]:
    """
    Batch convert .xls files to .xlsx format.
    Returns list of created .xlsx paths.
    """
    out_dir = _ensure_dir(output_dir)
    created = []
    for file in files:
        try:
            df_dict = pd.read_excel(file, sheet_name=None, engine="xlrd")
            out_path = str(out_dir / (Path(file).stem + ".xlsx"))
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                for sheet_name, df in df_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            print(f"    Converted: {Path(file).name} → {Path(out_path).name}")
            created.append(out_path)
        except Exception as e:
            print(f"    Error    : {Path(file).name}: {e}")
    return created


def excel_to_text(file: str, output_dir: str,
                  delimiter: str = "\t",
                  encoding: str = "utf-8-sig") -> List[str]:
    """
    Export each sheet to a delimited text file.
    delimiter: '\\t' (tab), '|' (pipe), ',' (csv), etc.
    """
    out_dir = _ensure_dir(output_dir)
    xl = pd.ExcelFile(file)
    stem = Path(file).stem
    ext_map = {"\t": "txt", "|": "txt", ",": "csv"}
    ext = ext_map.get(delimiter, "txt")
    created = []

    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        safe_name = "".join(c if c.isalnum() or c in "._- " else "_" for c in sheet)
        out_path = str(out_dir / f"{stem}_{safe_name}.{ext}")
        df.to_csv(out_path, sep=delimiter, index=False, encoding=encoding)
        print(f"    Saved   : {out_path}  ({len(df):,} rows)")
        created.append(out_path)

    return created


def merge_csv_files(csv_files: List[str], output_path: str,
                    encoding: str = "utf-8-sig",
                    add_source_col: bool = True) -> str:
    """
    Stack multiple CSV files vertically into one Excel sheet.
    Optionally adds a 'Source_File' column.
    """
    dfs = []
    for csv_file in csv_files:
        try:
            df = pd.read_csv(csv_file, encoding=encoding)
        except UnicodeDecodeError:
            df = pd.read_csv(csv_file, encoding="latin-1")
        if add_source_col:
            df["Source_File"] = Path(csv_file).name
        dfs.append(df)
        print(f"    Loaded  : {Path(csv_file).name}  ({len(df):,} rows)")

    combined = pd.concat(dfs, ignore_index=True)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    combined.to_excel(output_path, sheet_name="Merged", index=False)
    print(f"    Saved   : {output_path}  ({len(combined):,} total rows)")
    return output_path
