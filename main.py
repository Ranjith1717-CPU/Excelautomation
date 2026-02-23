"""
=============================================================================
  EXCEL AUTOMATION TOOLKIT  v2.0
  All-in-one Excel automation powered by Python + pandas
=============================================================================
  Run this file directly:   python main.py
  Or use the launcher:      run.bat  (Windows)
  Jump to a module:         python main.py finance
                            python main.py hr
                            python main.py sales
                            python main.py inventory
                            python main.py format
                            python main.py validate
                            python main.py analytics
                            python main.py convert
                            python main.py lookup
=============================================================================
"""
import os
import sys
import glob
import datetime
from pathlib import Path

# ── Bootstrap: install missing deps silently ─────────────────────────────────
def _check_deps():
    missing = []
    for pkg in ["pandas", "openpyxl", "colorama", "tabulate", "numpy"]:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"[INFO] Installing missing packages: {missing}")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing + ["-q"])

_check_deps()

# ── Imports ───────────────────────────────────────────────────────────────────
import pandas as pd
from colorama import init, Fore, Style
init(autoreset=True)

# Add local modules path
sys.path.insert(0, str(Path(__file__).parent))
from modules import (consolidator, calculator, cleaner, transformer, comparator, reporter, column_ops,
                     finance, hr, sales, inventory, formatter, validator, analytics, converter, lookup)

# ── Output directory ──────────────────────────────────────────────────────────
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


# =============================================================================
# UI HELPERS
# =============================================================================

def banner():
    os.system("cls" if os.name == "nt" else "clear")
    print(Fore.CYAN + Style.BRIGHT + """
╔══════════════════════════════════════════════════════════╗
║          EXCEL AUTOMATION TOOLKIT  v2.0                  ║
║          Automate Everything Excel — Powered by Python   ║
╚══════════════════════════════════════════════════════════╝""")
    print(Fore.YELLOW + f"  Output folder: {OUTPUT_DIR}\n")


def section(title: str):
    print(Fore.CYAN + Style.BRIGHT + f"\n{'─'*56}")
    print(Fore.CYAN + Style.BRIGHT + f"  {title}")
    print(Fore.CYAN + Style.BRIGHT + f"{'─'*56}")


def success(msg: str):
    print(Fore.GREEN + Style.BRIGHT + f"\n  ✓  {msg}")


def error(msg: str):
    print(Fore.RED + Style.BRIGHT + f"\n  ✗  ERROR: {msg}")


def info(msg: str):
    print(Fore.YELLOW + f"  →  {msg}")


def prompt(msg: str) -> str:
    return input(Fore.WHITE + Style.BRIGHT + f"\n  {msg}: ").strip()


def menu_choice(options: list, title: str = "Choose an option") -> str:
    """Display a numbered menu and return the chosen key."""
    print()
    for key, label in options:
        print(f"    {Fore.CYAN}{key}{Style.RESET_ALL}  {label}")
    print(f"    {Fore.CYAN}0{Style.RESET_ALL}  Back / Exit")
    return prompt(title)


def pause():
    input(Fore.YELLOW + "\n  Press Enter to continue...")


def get_output_path(prefix: str, ext: str = ".xlsx") -> str:
    """Generate a timestamped output path in the output folder."""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return str(OUTPUT_DIR / f"{prefix}_{ts}{ext}")


# =============================================================================
# FILE SELECTION HELPERS
# =============================================================================

def pick_files(prompt_text: str = "Enter file path(s)",
               allow_multiple: bool = True,
               allow_folder: bool = True) -> list:
    """
    Prompt user for one or more Excel file paths.
    Accepts:
      - A single path
      - Comma-separated paths
      - A folder path (loads all .xlsx files in it)
      - A glob pattern  (e.g. C:/data/*.xlsx)
    """
    raw = prompt(prompt_text + (" [comma-sep or folder]" if allow_multiple else ""))
    if not raw:
        return []

    paths = []
    for part in raw.split(","):
        part = part.strip().strip('"').strip("'")

        if not part:
            continue

        p = Path(part)

        # Glob pattern
        if "*" in part or "?" in part:
            matched = glob.glob(part)
            paths.extend([m for m in matched if m.lower().endswith((".xlsx", ".xls", ".xlsm"))])
            continue

        # Folder
        if p.is_dir() and allow_folder:
            for ext in ["*.xlsx", "*.xls", "*.xlsm"]:
                paths.extend([str(f) for f in p.glob(ext)])
            continue

        # Single file
        if p.is_file():
            paths.append(str(p))
        else:
            error(f"Not found: {part}")

    if paths:
        info(f"Files selected ({len(paths)}):")
        for f in paths:
            print(f"      {Path(f).name}")
    return paths


def pick_single_file(prompt_text: str = "Enter file path") -> str:
    files = pick_files(prompt_text, allow_multiple=False, allow_folder=False)
    return files[0] if files else ""


def pick_columns(df_columns: list, prompt_text: str = "Enter column names") -> list:
    """Show available columns and let user pick some."""
    print(f"\n  Available columns:")
    for i, c in enumerate(df_columns, 1):
        print(f"    {i:3}. {c}")
    raw = prompt(prompt_text + " [comma-sep, or numbers, or ALL]")
    if raw.strip().upper() == "ALL":
        return list(df_columns)
    cols = []
    for part in raw.split(","):
        part = part.strip()
        if part.isdigit():
            idx = int(part) - 1
            if 0 <= idx < len(df_columns):
                cols.append(df_columns[idx])
        elif part in df_columns:
            cols.append(part)
    return cols


def preview_file(file_path: str, n: int = 5):
    """Print the first N rows of an Excel file."""
    try:
        df = pd.read_excel(file_path, nrows=n)
        print(f"\n  Preview — {Path(file_path).name}  (first {n} rows):")
        print(df.to_string(index=False))
        print(f"\n  Shape: {pd.read_excel(file_path).shape}")
    except Exception as e:
        error(str(e))


# =============================================================================
# MENU 1 — CONSOLIDATE FILES
# =============================================================================

def menu_consolidate():
    while True:
        banner()
        section("1. CONSOLIDATE FILES")
        choice = menu_choice([
            ("1", "Stack N files vertically (append all rows)"),
            ("2", "Join N files by a key column (SQL-style JOIN)"),
            ("3", "Extract specific columns from N files"),
            ("4", "Consolidate all sheets within ONE file"),
            ("5", "Merge same sheet name across multiple files"),
        ], "Select operation")

        if choice == "0":
            break

        elif choice == "1":
            section("Stack files vertically")
            files = pick_files("Enter Excel file paths (or folder)")
            if not files: continue
            add_src = prompt("Add source filename column? [y/n]").lower() != "n"
            out = get_output_path("consolidated_stack")
            try:
                result = consolidator.merge_files_stack(files, out, add_source_column=add_src)
                success(f"Saved → {result}")
            except Exception as e:
                error(str(e))
            pause()

        elif choice == "2":
            section("Join files by key column")
            files = pick_files("Enter Excel file paths")
            if not files: continue
            key = prompt("Enter the key column name (must exist in all files)")
            join = prompt("Join type [inner/outer/left/right] (default: outer)") or "outer"
            out = get_output_path("consolidated_join")
            try:
                result = consolidator.merge_files_by_key(files, key, join, out)
                success(f"Saved → {result}")
            except Exception as e:
                error(str(e))
            pause()

        elif choice == "3":
            section("Extract specific columns from multiple files")
            files = pick_files("Enter Excel file paths")
            if not files: continue
            # Preview columns from first file
            try:
                sample_cols = pd.read_excel(files[0], nrows=0).columns.tolist()
                cols = pick_columns(sample_cols, "Enter column names to extract")
            except Exception as e:
                error(str(e)); continue
            out = get_output_path("extracted_columns")
            try:
                result = consolidator.merge_specific_columns(files, cols, out)
                success(f"Saved → {result}")
            except Exception as e:
                error(str(e))
            pause()

        elif choice == "4":
            section("Consolidate all sheets in one file")
            file = pick_single_file("Enter Excel file path")
            if not file: continue
            add_sheet = prompt("Add sheet name column? [y/n]").lower() != "n"
            out = get_output_path("sheets_consolidated")
            try:
                result = consolidator.merge_sheets_in_file(file, out, add_sheet_column=add_sheet)
                success(f"Saved → {result}")
            except Exception as e:
                error(str(e))
            pause()

        elif choice == "5":
            section("Merge same sheet across multiple files")
            files = pick_files("Enter Excel file paths")
            if not files: continue
            sheet = prompt("Enter the sheet tab name to merge")
            out = get_output_path("cross_file_sheet")
            try:
                result = consolidator.merge_same_sheet_cross_files(files, sheet, out)
                success(f"Saved → {result}")
            except Exception as e:
                error(str(e))
            pause()


# =============================================================================
# MENU 2 — CALCULATE & ANALYZE
# =============================================================================

def menu_calculate():
    while True:
        banner()
        section("2. CALCULATE & ANALYZE")
        choice = menu_choice([
            ("1",  "Efficiency  (Actual / Target × 100)"),
            ("2",  "Productivity  (Output / Input)"),
            ("3",  "Utilization  (Used / Available × 100)"),
            ("4",  "Variance  (Actual vs Budget)"),
            ("5",  "Growth Rate  (period-over-period %)"),
            ("6",  "Summary Statistics  (count, sum, mean, median, min, max)"),
            ("7",  "Percentage of Total  (each row's % share)"),
            ("8",  "Moving / Rolling Average"),
            ("9",  "KPI Dashboard  (multi-column key metrics)"),
            ("10", "Weighted Average"),
        ], "Select calculation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                actual = prompt(f"Actual column {cols}"); target = prompt(f"Target column {cols}")
                out = get_output_path("efficiency")
                result = calculator.calculate_efficiency(file, actual, target, out)

            elif choice == "2":
                out_col = prompt(f"Output column {cols}"); in_col = prompt(f"Input column {cols}")
                out = get_output_path("productivity")
                result = calculator.calculate_productivity(file, out_col, in_col, out)

            elif choice == "3":
                used = prompt(f"Used column {cols}"); avail = prompt(f"Available column {cols}")
                out = get_output_path("utilization")
                result = calculator.calculate_utilization(file, used, avail, out)

            elif choice == "4":
                actual = prompt(f"Actual column {cols}"); budget = prompt(f"Budget column {cols}")
                out = get_output_path("variance")
                result = calculator.calculate_variance(file, actual, budget, out)

            elif choice == "5":
                val_col = prompt(f"Value column {cols}")
                period_col = prompt(f"Period/Date column (or Enter to skip) {cols}") or None
                out = get_output_path("growth_rate")
                result = calculator.calculate_growth_rate(file, val_col, period_col, out)

            elif choice == "6":
                selected = pick_columns(cols, "Columns to summarize (or ALL for all numeric)")
                out = get_output_path("summary_stats")
                result = calculator.calculate_summary_stats(file, selected, out)

            elif choice == "7":
                val_col = prompt(f"Value column {cols}")
                grp_col = prompt(f"Group column for sub-totals (or Enter to skip) {cols}") or None
                out = get_output_path("pct_of_total")
                result = calculator.calculate_percentage_of_total(file, val_col, grp_col, out)

            elif choice == "8":
                val_col = prompt(f"Value column {cols}")
                window = int(prompt("Window size (number of periods)") or "3")
                out = get_output_path("moving_avg")
                result = calculator.calculate_moving_average(file, val_col, window, out)

            elif choice == "9":
                kpi_cols = pick_columns(cols, "Select KPI columns")
                out = get_output_path("kpi_dashboard")
                result = calculator.calculate_kpi_dashboard(file, kpi_cols, out)

            elif choice == "10":
                val_col = prompt(f"Value column {cols}"); wt_col = prompt(f"Weight column {cols}")
                grp_col = prompt(f"Group column (or Enter for overall) {cols}") or None
                out = get_output_path("weighted_avg")
                result = calculator.calculate_weighted_average(file, val_col, wt_col, grp_col, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 3 — CLEAN DATA
# =============================================================================

def menu_clean():
    while True:
        banner()
        section("3. CLEAN DATA")
        choice = menu_choice([
            ("1",  "Remove duplicate rows"),
            ("2",  "Remove empty rows and/or columns"),
            ("3",  "Trim whitespace in text columns"),
            ("4",  "Standardize date columns to one format"),
            ("5",  "Fill missing / blank values"),
            ("6",  "Auto-fix column data types"),
            ("7",  "Normalize text case (upper/lower/title)"),
            ("8",  "Remove special characters"),
            ("9",  "Remove statistical outliers"),
            ("10", "Full Auto-Clean  (runs all steps)"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                subset_input = prompt("Columns to check for duplicates (or Enter for ALL columns)")
                subset = [c.strip() for c in subset_input.split(",") if c.strip()] or None
                keep = prompt("Keep: first / last / False") or "first"
                out = get_output_path("deduped")
                result = cleaner.remove_duplicates(file, out, subset=subset, keep=keep)

            elif choice == "2":
                rm_rows = prompt("Remove empty rows? [y/n]").lower() != "n"
                rm_cols = prompt("Remove empty columns? [y/n]").lower() != "n"
                out = get_output_path("empty_removed")
                result = cleaner.remove_empty_rows_cols(file, out, remove_rows=rm_rows, remove_cols=rm_cols)

            elif choice == "3":
                sel = prompt(f"Columns to trim (or Enter for ALL text columns) {cols}")
                columns = [c.strip() for c in sel.split(",") if c.strip()] if sel else None
                out = get_output_path("trimmed")
                result = cleaner.trim_whitespace(file, out, columns=columns)

            elif choice == "4":
                date_cols = pick_columns(cols, "Select date columns")
                fmt = prompt("Output date format (default: %Y-%m-%d)") or "%Y-%m-%d"
                out = get_output_path("dates_fixed")
                result = cleaner.standardize_dates(file, date_cols, output_format=fmt, output_path=out)

            elif choice == "5":
                strategy = prompt("Strategy [mean / median / mode / ffill / bfill / value]") or "mean"
                fill_val = prompt("Fill value (only used if strategy=value)") if strategy == "value" else None
                sel = prompt(f"Columns to fill (or Enter for ALL) {cols}")
                columns = [c.strip() for c in sel.split(",") if c.strip()] if sel else None
                out = get_output_path("missing_filled")
                result = cleaner.fill_missing_values(file, out, strategy=strategy,
                                                      fill_value=fill_val, columns=columns)

            elif choice == "6":
                out = get_output_path("types_fixed")
                result = cleaner.fix_data_types(file, out)

            elif choice == "7":
                case = prompt("Case [upper / lower / title / sentence]") or "title"
                sel = prompt(f"Columns (or Enter for ALL text) {cols}")
                columns = [c.strip() for c in sel.split(",") if c.strip()] if sel else None
                out = get_output_path("case_normalized")
                result = cleaner.normalize_text_case(file, out, columns=columns, case=case)

            elif choice == "8":
                sel = prompt(f"Columns (or Enter for ALL text) {cols}")
                columns = [c.strip() for c in sel.split(",") if c.strip()] if sel else None
                out = get_output_path("special_chars_removed")
                result = cleaner.remove_special_characters(file, out, columns=columns)

            elif choice == "9":
                val_col = prompt(f"Numeric column to check for outliers {cols}")
                thresh = float(prompt("Std-dev threshold (default 3.0)") or "3.0")
                out = get_output_path("outliers_removed")
                result = cleaner.remove_outliers(file, val_col, out, std_threshold=thresh)

            elif choice == "10":
                out = get_output_path("full_cleaned")
                result = cleaner.full_clean(file, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 4 — TRANSFORM DATA
# =============================================================================

def menu_transform():
    while True:
        banner()
        section("4. TRANSFORM DATA")
        choice = menu_choice([
            ("1", "Create Pivot Table"),
            ("2", "Unpivot / Melt  (wide → long)"),
            ("3", "Transpose  (flip rows ↔ columns)"),
            ("4", "Split file by column value  (one file per category)"),
            ("5", "Split sheets to separate files"),
            ("6", "Split large file into row-chunks"),
            ("7", "Wide to Long  (repeated column stubs)"),
            ("8", "Long to Wide  (crosstab / unstack)"),
            ("9", "Add Running / Cumulative Total column"),
            ("10","Rank rows by a column"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                index_cols  = pick_columns(cols, "Row (index) columns")
                value_cols  = pick_columns(cols, "Value columns to aggregate")
                col_col_raw = prompt(f"Column-spread column (or Enter to skip) {cols}")
                col_col     = col_col_raw if col_col_raw else None
                aggfunc     = prompt("Aggregation [sum/mean/count/min/max] (default: sum)") or "sum"
                out = get_output_path("pivot")
                result = transformer.create_pivot_table(file, index_cols, value_cols, out, col_col, aggfunc)

            elif choice == "2":
                id_cols  = pick_columns(cols, "ID / identifier columns (keep as-is)")
                val_cols = pick_columns(cols, "Value columns to unpivot into rows")
                var_name = prompt("Name for the variable column (default: Variable)") or "Variable"
                val_name = prompt("Name for the value column (default: Value)") or "Value"
                out = get_output_path("unpivoted")
                result = transformer.unpivot_data(file, id_cols, val_cols, out, var_name, val_name)

            elif choice == "3":
                header_col = prompt(f"Column to use as header after transpose (or Enter) {cols}") or None
                out = get_output_path("transposed")
                result = transformer.transpose_data(file, out, header_col=header_col)

            elif choice == "4":
                split_col = prompt(f"Column to split on {cols}")
                out_dir = str(OUTPUT_DIR / f"split_{datetime.datetime.now().strftime('%H%M%S')}")
                created = transformer.split_by_column_value(file, split_col, out_dir)
                success(f"Created {len(created)} file(s) in {out_dir}")
                pause(); continue

            elif choice == "5":
                out_dir = str(OUTPUT_DIR / f"sheets_{datetime.datetime.now().strftime('%H%M%S')}")
                created = transformer.split_sheets_to_files(file, out_dir)
                success(f"Created {len(created)} file(s) in {out_dir}")
                pause(); continue

            elif choice == "6":
                chunk = int(prompt("Rows per chunk (e.g. 1000)") or "1000")
                out_dir = str(OUTPUT_DIR / f"chunks_{datetime.datetime.now().strftime('%H%M%S')}")
                created = transformer.split_file_by_rows(file, chunk, out_dir)
                success(f"Created {len(created)} file(s) in {out_dir}")
                pause(); continue

            elif choice == "7":
                stub_raw = prompt("Column stub names to convert (comma-separated, e.g. Sales,Cost)")
                stubs = [s.strip() for s in stub_raw.split(",") if s.strip()]
                out = get_output_path("wide_to_long")
                result = transformer.reshape_wide_to_long(file, stubs, out)

            elif choice == "8":
                index_cols = pick_columns(cols, "Index / row identifier columns")
                col_col    = prompt(f"Column whose values become new headers {cols}")
                val_col    = prompt(f"Values column {cols}")
                aggfunc    = prompt("Aggregation [sum/mean/count] (default: sum)") or "sum"
                out = get_output_path("long_to_wide")
                result = transformer.reshape_long_to_wide(file, index_cols, col_col, val_col, out, aggfunc)

            elif choice == "9":
                val_col = prompt(f"Numeric column {cols}")
                grp_col = prompt(f"Group column (or Enter for no grouping) {cols}") or None
                out = get_output_path("running_total")
                result = transformer.add_running_total(file, val_col, out, group_col=grp_col)

            elif choice == "10":
                val_col   = prompt(f"Column to rank by {cols}")
                asc       = prompt("Ascending? [y/n] (n = highest = rank 1)").lower() == "y"
                grp_col   = prompt(f"Rank within group (or Enter for overall) {cols}") or None
                out = get_output_path("ranked")
                result = transformer.rank_column(file, val_col, out, ascending=asc, group_col=grp_col)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 5 — COMPARE FILES
# =============================================================================

def menu_compare():
    while True:
        banner()
        section("5. COMPARE FILES")
        choice = menu_choice([
            ("1", "Full comparison report  (2 files, all differences)"),
            ("2", "Find NEW rows  (rows in File2 but not File1)"),
            ("3", "Find DELETED rows  (rows in File1 but not File2)"),
            ("4", "Find CHANGED cell values  (same key, different data)"),
            ("5", "Find duplicates WITHIN a single file"),
            ("6", "Find COMMON rows  (exist in both files)"),
            ("7", "Cross-file duplicate check  (N files, same key)"),
        ], "Select operation")

        if choice == "0":
            break

        try:
            if choice in ("1", "2", "3", "4", "6"):
                f1 = pick_single_file("Enter File 1 path")
                if not f1: continue
                f2 = pick_single_file("Enter File 2 path")
                if not f2: continue

                if choice == "1":
                    cols = pd.read_excel(f1, nrows=0).columns.tolist()
                    key = prompt(f"Key column for matching rows (or Enter for row-position) {cols}") or None
                    out = get_output_path("comparison_report")
                    result = comparator.compare_two_files(f1, f2, out, key_column=key)

                elif choice == "2":
                    cols = pd.read_excel(f1, nrows=0).columns.tolist()
                    key_raw = prompt(f"Key column(s) — comma-sep (or Enter for all cols) {cols}")
                    keys = [k.strip() for k in key_raw.split(",") if k.strip()] or None
                    out = get_output_path("new_rows")
                    result = comparator.find_new_rows(f1, f2, out, key_columns=keys)

                elif choice == "3":
                    cols = pd.read_excel(f1, nrows=0).columns.tolist()
                    key_raw = prompt(f"Key column(s) — comma-sep (or Enter for all cols) {cols}")
                    keys = [k.strip() for k in key_raw.split(",") if k.strip()] or None
                    out = get_output_path("deleted_rows")
                    result = comparator.find_deleted_rows(f1, f2, out, key_columns=keys)

                elif choice == "4":
                    cols = pd.read_excel(f1, nrows=0).columns.tolist()
                    key = prompt(f"Key column {cols}")
                    out = get_output_path("changed_values")
                    result = comparator.find_changed_values(f1, f2, key, out)

                elif choice == "6":
                    cols = pd.read_excel(f1, nrows=0).columns.tolist()
                    key_raw = prompt(f"Key column(s) — comma-sep (or Enter for all cols) {cols}")
                    keys = [k.strip() for k in key_raw.split(",") if k.strip()] or None
                    out = get_output_path("common_rows")
                    result = comparator.find_common_rows(f1, f2, out, key_columns=keys)

                success(f"Saved → {result}")

            elif choice == "5":
                f1 = pick_single_file("Enter Excel file path")
                if not f1: continue
                cols = pd.read_excel(f1, nrows=0).columns.tolist()
                key_raw = prompt(f"Columns to check (comma-sep, or Enter for all) {cols}")
                keys = [k.strip() for k in key_raw.split(",") if k.strip()] or None
                out = get_output_path("duplicates")
                result = comparator.find_duplicates_in_file(f1, out, subset=keys)
                success(f"Saved → {result}")

            elif choice == "7":
                files = pick_files("Enter Excel file paths (or folder)")
                if not files: continue
                cols = pd.read_excel(files[0], nrows=0).columns.tolist()
                key_raw = prompt(f"Key column(s) — comma-sep {cols}")
                keys = [k.strip() for k in key_raw.split(",") if k.strip()]
                out = get_output_path("cross_file_dups")
                result = comparator.cross_file_duplicate_check(files, keys, out)
                success(f"Saved → {result}")

        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 6 — COLUMN OPERATIONS
# =============================================================================

def menu_columns():
    while True:
        banner()
        section("6. COLUMN OPERATIONS")
        choice = menu_choice([
            ("1",  "Rename columns"),
            ("2",  "Merge / Concatenate columns"),
            ("3",  "Split column by delimiter"),
            ("4",  "Reorder columns"),
            ("5",  "Drop / Delete columns"),
            ("6",  "Add calculated column  (formula)"),
            ("7",  "Extract from column  (regex)"),
            ("8",  "Map / Replace column values"),
            ("9",  "Expand multi-value cell to rows  (un-nest)"),
            ("10", "Normalize all column header names"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                print(f"\n  Current columns: {cols}")
                mapping = {}
                info("Enter renames one by one. Press Enter when done.")
                while True:
                    old = prompt("Old column name (or Enter to finish)")
                    if not old: break
                    new = prompt(f"New name for '{old}'")
                    mapping[old] = new
                out = get_output_path("renamed")
                result = column_ops.rename_columns(file, mapping, out)

            elif choice == "2":
                selected = pick_columns(cols, "Columns to merge")
                new_name = prompt("New column name")
                sep = prompt("Separator (default: space)") or " "
                drop = prompt("Drop original columns? [y/n]").lower() == "y"
                out = get_output_path("merged_cols")
                result = column_ops.merge_columns(file, selected, new_name, out, separator=sep, drop_originals=drop)

            elif choice == "3":
                col = prompt(f"Column to split {cols}")
                delim = prompt("Delimiter (e.g. comma, space, |)")
                names_raw = prompt("New column names (comma-sep, or Enter to auto-name)")
                names = [n.strip() for n in names_raw.split(",") if n.strip()] or None
                drop = prompt("Drop original column? [y/n]").lower() == "y"
                out = get_output_path("split_col")
                result = column_ops.split_column(file, col, delim, names, out, drop_original=drop)

            elif choice == "4":
                print(f"\n  Current columns: {cols}")
                order_raw = prompt("Enter column names in desired order (comma-sep)")
                order = [c.strip() for c in order_raw.split(",") if c.strip()]
                out = get_output_path("reordered")
                result = column_ops.reorder_columns(file, order, out)

            elif choice == "5":
                selected = pick_columns(cols, "Columns to DROP / DELETE")
                out = get_output_path("dropped_cols")
                result = column_ops.drop_columns(file, selected, out)

            elif choice == "6":
                info("Available columns: " + str(cols))
                info("Example expressions: 'Salary * 1.1'  |  'Revenue - Cost'  |  'Units * Price'")
                new_name = prompt("New column name")
                expr = prompt("Expression (use column names directly)")
                out = get_output_path("calculated_col")
                result = column_ops.add_calculated_column(file, new_name, expr, out)

            elif choice == "7":
                src_col = prompt(f"Source column {cols}")
                pattern = prompt("Regex pattern to extract (e.g. r'\\d+' for numbers)")
                new_name = prompt("New column name (or Enter for auto)") or None
                out = get_output_path("extracted_col")
                result = column_ops.extract_from_column(file, src_col, pattern, out, new_column_name=new_name)

            elif choice == "8":
                col = prompt(f"Column to map {cols}")
                mapping = {}
                info("Enter value mappings. Press Enter when done.")
                while True:
                    old = prompt("Old value (or Enter to finish)")
                    if not old: break
                    new = prompt(f"New value for '{old}'")
                    mapping[old] = new
                strategy = prompt("Unmapped values: keep / null / other") or "keep"
                out = get_output_path("mapped_values")
                result = column_ops.map_column_values(file, col, mapping, out, unmapped_strategy=strategy)

            elif choice == "9":
                col = prompt(f"Column with multi-values {cols}")
                delim = prompt("Delimiter inside cells (default: comma)") or ","
                out = get_output_path("expanded_rows")
                result = column_ops.pivot_column_to_rows(file, col, out, delimiter=delim)

            elif choice == "10":
                style = prompt("Style [snake_case / title_case / upper / lower] (default: snake_case)") or "snake_case"
                out = get_output_path("headers_normalized")
                result = column_ops.normalize_column_names(file, out, style=style)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 7 — GENERATE REPORTS
# =============================================================================

def menu_reports():
    while True:
        banner()
        section("7. GENERATE REPORTS")
        choice = menu_choice([
            ("1", "Summary statistics report  (one or more files)"),
            ("2", "Data Profile  (detailed column analysis)"),
            ("3", "KPI Report  (key metrics dashboard)"),
            ("4", "Top-N / Bottom-N report"),
            ("5", "Frequency / Value count report"),
            ("6", "Monthly summary report  (aggregate by month)"),
        ], "Select report type")

        if choice == "0":
            break

        try:
            if choice == "1":
                files = pick_files("Enter Excel file paths (or folder)")
                if not files: continue
                out = get_output_path("summary_report")
                result = reporter.generate_summary_report(files, out)

            elif choice == "2":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                out = get_output_path("data_profile")
                result = reporter.data_profile(file, out)

            elif choice == "3":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                kpi_cols = pick_columns(cols, "Select KPI (numeric) columns")
                label_col = prompt(f"Category / label column (or Enter to skip) {cols}") or None
                out = get_output_path("kpi_report")
                result = reporter.generate_kpi_report(file, kpi_cols, label_col, out)

            elif choice == "4":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                sort_col = prompt(f"Column to rank by {cols}")
                n = int(prompt("Top N value (e.g. 10)") or "10")
                asc = prompt("Ascending order? [y/n]").lower() == "y"
                out = get_output_path("top_n_report")
                result = reporter.top_n_report(file, sort_col, n, out, ascending=asc)

            elif choice == "5":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                selected = pick_columns(cols, "Select columns for frequency count")
                top_n = int(prompt("Show top N values per column (default: 20)") or "20")
                out = get_output_path("frequency_report")
                result = reporter.frequency_report(file, selected, out, top_n=top_n)

            elif choice == "6":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                date_col = prompt(f"Date column {cols}")
                val_cols = pick_columns(cols, "Numeric columns to aggregate")
                aggfunc = prompt("Aggregation [sum/mean/count] (default: sum)") or "sum"
                out = get_output_path("monthly_summary")
                result = reporter.monthly_summary_report(file, date_col, val_cols, out, aggfunc=aggfunc)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 8 — QUICK PREVIEW
# =============================================================================

def menu_preview():
    banner()
    section("QUICK FILE PREVIEW")
    file = pick_single_file("Enter Excel file path to preview")
    if not file:
        return
    n = int(prompt("Number of rows to preview (default: 10)") or "10")
    preview_file(file, n)
    pause()


# =============================================================================
# MENU 9 — FINANCE
# =============================================================================

def menu_finance():
    while True:
        banner()
        section("9. FINANCE")
        choice = menu_choice([
            ("1", "AR/AP Aging Analysis       — 0-30, 31-60, 61-90, 90+ buckets"),
            ("2", "Loan Amortization Schedule — EMI, principal, interest, balance"),
            ("3", "Depreciation Schedule      — Straight-line + declining balance"),
            ("4", "Financial Ratios           — Gross margin, ROI, current ratio"),
            ("5", "Payroll Calculator         — Gross→Net with HRA/PF/ESI/TDS"),
            ("6", "Budget vs Actual           — Variance + % variance report"),
            ("7", "Compound Interest Schedule — Future value growth table"),
        ], "Select operation")

        if choice == "0":
            break

        try:
            if choice == "1":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                date_col   = prompt(f"Date column {cols}")
                amount_col = prompt(f"Amount column {cols}")
                as_of      = prompt("As-of date [YYYY-MM-DD] (or Enter for today)") or None
                out = get_output_path("aging_analysis")
                result = finance.aging_analysis(file, date_col, amount_col, out, as_of)

            elif choice == "2":
                principal = float(prompt("Principal amount (e.g. 1000000)"))
                rate      = float(prompt("Annual interest rate % (e.g. 12)"))
                months    = int(prompt("Loan tenure in months (e.g. 60)"))
                out = get_output_path("loan_amortization")
                result = finance.loan_amortization(principal, rate, months, out)

            elif choice == "3":
                file = pick_single_file("Enter Excel file path (needs: Asset, Cost, Salvage_Value, Useful_Life_Years)")
                if not file: continue
                out = get_output_path("depreciation_schedule")
                result = finance.depreciation_schedule(file, out)

            elif choice == "4":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                out = get_output_path("financial_ratios")
                result = finance.financial_ratios(file, out)

            elif choice == "5":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                basic_col = prompt(f"Basic salary column {cols}")
                out = get_output_path("payroll")
                result = finance.payroll_calculator(file, basic_col, out)

            elif choice == "6":
                budget_file = pick_single_file("Enter BUDGET file path")
                if not budget_file: continue
                actual_file = pick_single_file("Enter ACTUAL file path")
                if not actual_file: continue
                cols = pd.read_excel(budget_file, nrows=0).columns.tolist()
                key_col = prompt(f"Key column (present in both files) {cols}")
                out = get_output_path("budget_vs_actual")
                result = finance.budget_vs_actual(budget_file, actual_file, key_col, out)

            elif choice == "7":
                principal = float(prompt("Principal amount (e.g. 100000)"))
                rate      = float(prompt("Annual interest rate % (e.g. 8)"))
                periods   = int(prompt("Number of years (e.g. 10)"))
                freq      = prompt("Compounding frequency [annual/semi-annual/quarterly/monthly]") or "annual"
                out = get_output_path("compound_interest")
                result = finance.compound_interest_schedule(principal, rate, periods, out, freq)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 10 — HR
# =============================================================================

def menu_hr():
    while True:
        banner()
        section("10. HR ANALYTICS")
        choice = menu_choice([
            ("1", "Attrition Analysis         — Turnover rate by department"),
            ("2", "Headcount Summary          — Count/% by group columns"),
            ("3", "Tenure Analysis            — Years-of-service bands"),
            ("4", "Age Band Analysis          — Workforce age demographics"),
            ("5", "Salary Analysis            — Min/max/median/percentiles"),
            ("6", "Performance Distribution   — Rating bell-curve + ranking"),
            ("7", "Salary Increment Calculator— Apply % increment to salary"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                status_col = prompt(f"Employee status column {cols}")
                dept_col   = prompt(f"Department column {cols}")
                active     = prompt("Value meaning 'Active' (default: Active)") or "Active"
                left       = prompt("Value meaning 'Left' (default: Left)") or "Left"
                out = get_output_path("attrition_analysis")
                result = hr.attrition_analysis(file, status_col, dept_col, out, active, left)

            elif choice == "2":
                group_cols = pick_columns(cols, "Group-by columns")
                out = get_output_path("headcount_summary")
                result = hr.headcount_summary(file, group_cols, out)

            elif choice == "3":
                join_col  = prompt(f"Join / Hire date column {cols}")
                exit_col  = prompt(f"Exit date column (or Enter to skip) {cols}") or None
                out = get_output_path("tenure_analysis")
                result = hr.tenure_analysis(file, join_col, out, exit_date_col=exit_col)

            elif choice == "4":
                dob_col = prompt(f"Date of birth column {cols}")
                out = get_output_path("age_band_analysis")
                result = hr.age_band_analysis(file, dob_col, out)

            elif choice == "5":
                salary_col = prompt(f"Salary column {cols}")
                dept_col   = prompt(f"Department/Group column {cols}")
                out = get_output_path("salary_analysis")
                result = hr.salary_analysis(file, salary_col, dept_col, out)

            elif choice == "6":
                rating_col = prompt(f"Performance rating column {cols}")
                out = get_output_path("performance_distribution")
                result = hr.performance_distribution(file, rating_col, out)

            elif choice == "7":
                salary_col  = prompt(f"Current salary column {cols}")
                pct_raw     = prompt("Flat increment % for all (e.g. 10) OR column name with individual %s")
                try:
                    pct_or_col = float(pct_raw)
                except ValueError:
                    pct_or_col = pct_raw
                out = get_output_path("salary_increment")
                result = hr.salary_increment_calculator(file, salary_col, pct_or_col, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 11 — SALES
# =============================================================================

def menu_sales():
    while True:
        banner()
        section("11. SALES ANALYTICS")
        choice = menu_choice([
            ("1", "Commission Calculator  — Flat % or tiered slab"),
            ("2", "RFM Segmentation       — Recency, Frequency, Monetary"),
            ("3", "Quota Attainment       — % attainment + status labels"),
            ("4", "Pipeline Analysis      — Funnel by stage + conversion"),
            ("5", "Sales by Territory     — Territory summary + rank"),
            ("6", "Customer ABC           — A/B/C by revenue contribution"),
            ("7", "Discount Analysis      — Discount % + revenue leakage"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                sales_col = prompt(f"Sales amount column {cols}")
                use_tiers = prompt("Use tiered commission? [y/n]").lower() == "y"
                tiers = None
                if use_tiers:
                    tiers = []
                    info("Enter tiers: (max_sales, commission_%). Press Enter with empty max_sales to finish.")
                    while True:
                        lim = prompt("Tier max sales (or Enter to add ∞ final tier)")
                        pct = prompt("Commission % for this tier")
                        if lim == "":
                            tiers.append((float("inf"), float(pct)))
                            break
                        tiers.append((float(lim), float(pct)))
                flat_pct = float(prompt("Flat commission % (only if not tiered, default: 5)") or "5") if not tiers else 5.0
                out = get_output_path("commission")
                result = sales.commission_calculator(file, sales_col, out, tiers=tiers, flat_pct=flat_pct)

            elif choice == "2":
                customer_col = prompt(f"Customer ID column {cols}")
                date_col     = prompt(f"Transaction date column {cols}")
                amount_col   = prompt(f"Transaction amount column {cols}")
                out = get_output_path("rfm_segmentation")
                result = sales.rfm_segmentation(file, customer_col, date_col, amount_col, out)

            elif choice == "3":
                actual_col = prompt(f"Actual sales column {cols}")
                quota_col  = prompt(f"Quota column {cols}")
                out = get_output_path("quota_attainment")
                result = sales.quota_attainment(file, actual_col, quota_col, out)

            elif choice == "4":
                stage_col = prompt(f"Pipeline stage column {cols}")
                value_col = prompt(f"Deal value column {cols}")
                out = get_output_path("pipeline_analysis")
                result = sales.pipeline_analysis(file, stage_col, value_col, out)

            elif choice == "5":
                territory_col = prompt(f"Territory column {cols}")
                sales_col     = prompt(f"Sales amount column {cols}")
                out = get_output_path("sales_by_territory")
                result = sales.sales_by_territory(file, territory_col, sales_col, out)

            elif choice == "6":
                customer_col = prompt(f"Customer column {cols}")
                revenue_col  = prompt(f"Revenue column {cols}")
                out = get_output_path("customer_abc")
                result = sales.customer_abc(file, customer_col, revenue_col, out)

            elif choice == "7":
                list_col = prompt(f"List price column {cols}")
                sell_col = prompt(f"Sell price column {cols}")
                out = get_output_path("discount_analysis")
                result = sales.discount_analysis(file, list_col, sell_col, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 12 — INVENTORY
# =============================================================================

def menu_inventory():
    while True:
        banner()
        section("12. INVENTORY MANAGEMENT")
        choice = menu_choice([
            ("1", "ABC Analysis        — Classify items by value contribution"),
            ("2", "Reorder Point       — ROP = usage × lead time + safety stock"),
            ("3", "Stock Aging         — Age-bucket inventory by receipt date"),
            ("4", "Inventory Turnover  — Turnover ratio + days on hand"),
            ("5", "OEE Calculator      — Availability × Performance × Quality"),
            ("6", "Dead Stock Analysis — Items with no movement beyond N days"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                item_col  = prompt(f"Item/SKU column {cols}")
                value_col = prompt(f"Value column {cols}")
                out = get_output_path("abc_analysis")
                result = inventory.abc_analysis(file, item_col, value_col, out)

            elif choice == "2":
                info("Expected columns: Avg_Daily_Usage, Lead_Time_Days, Safety_Stock, Current_Stock")
                out = get_output_path("reorder_point")
                result = inventory.reorder_point(file, out)

            elif choice == "3":
                receipt_col = prompt(f"Receipt date column {cols}")
                qty_col     = prompt(f"Quantity column {cols}")
                out = get_output_path("stock_aging")
                result = inventory.stock_aging(file, receipt_col, qty_col, out)

            elif choice == "4":
                cogs_col      = prompt(f"COGS / Cost column {cols}")
                inventory_col = prompt(f"Inventory value column {cols}")
                item_col      = prompt(f"Item group column for summary (or Enter to skip) {cols}") or None
                out = get_output_path("inventory_turnover")
                result = inventory.inventory_turnover(file, cogs_col, inventory_col, out, item_col=item_col)

            elif choice == "5":
                info("Expected columns: Planned_Time, Downtime, Ideal_Rate, Actual_Rate, Good_Units, Total_Units")
                out = get_output_path("oee")
                result = inventory.oee_calculator(file, out)

            elif choice == "6":
                last_move_col = prompt(f"Last movement date column {cols}")
                qty_col       = prompt(f"Quantity column {cols}")
                days          = int(prompt("Dead stock threshold in days (default: 180)") or "180")
                out = get_output_path("dead_stock")
                result = inventory.dead_stock_analysis(file, last_move_col, qty_col, out, days=days)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 13 — FORMAT / STYLE
# =============================================================================

def menu_formatter():
    while True:
        banner()
        section("13. FORMAT & STYLE")
        choice = menu_choice([
            ("1",  "Add Bar Chart          — Embedded column chart"),
            ("2",  "Add Line Chart         — Embedded line chart"),
            ("3",  "Add Pie Chart          — Embedded pie chart"),
            ("4",  "Traffic Light Colors   — Red/Yellow/Green cell fills"),
            ("5",  "Color Scale            — Gradient min→max fill"),
            ("6",  "Format as Table        — Apply Excel table style"),
            ("7",  "Freeze & Filter        — Freeze header + auto-filter"),
            ("8",  "Auto-fit Columns       — Width based on content"),
            ("9",  "Add Totals Row         — SUM row at bottom"),
            ("10", "Highlight Duplicates   — Yellow fill on duplicate cells"),
            ("11", "Apply Number Format    — Currency/percentage/comma"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                x_col = prompt(f"Category/X-axis column {cols}")
                y_col = prompt(f"Value/Y-axis column {cols}")
                title = prompt("Chart title (or Enter for default)") or "Bar Chart"
                out = get_output_path("bar_chart")
                result = formatter.add_bar_chart(file, x_col, y_col, out, title=title)

            elif choice == "2":
                x_col = prompt(f"X-axis column {cols}")
                y_col = prompt(f"Y-axis column {cols}")
                title = prompt("Chart title (or Enter for default)") or "Line Chart"
                out = get_output_path("line_chart")
                result = formatter.add_line_chart(file, x_col, y_col, out, title=title)

            elif choice == "3":
                cat_col = prompt(f"Category column {cols}")
                val_col = prompt(f"Value column {cols}")
                title   = prompt("Chart title (or Enter for default)") or "Pie Chart"
                out = get_output_path("pie_chart")
                result = formatter.add_pie_chart(file, cat_col, val_col, out, title=title)

            elif choice == "4":
                col = prompt(f"Numeric column to color {cols}")
                info("Leave thresholds blank to use auto 33rd/66th percentile")
                red_raw    = prompt("Red threshold (below this = red)")
                yellow_raw = prompt("Yellow threshold (below this = yellow)")
                red    = float(red_raw)    if red_raw    else None
                yellow = float(yellow_raw) if yellow_raw else None
                out = get_output_path("traffic_light")
                result = formatter.apply_traffic_light(file, col, out, red=red, yellow=yellow)

            elif choice == "5":
                col = prompt(f"Numeric column for color scale {cols}")
                out = get_output_path("color_scale")
                result = formatter.apply_color_scale(file, col, out)

            elif choice == "6":
                style = prompt("Table style (default: TableStyleMedium9)") or "TableStyleMedium9"
                out = get_output_path("table_styled")
                result = formatter.format_as_table(file, out, style=style)

            elif choice == "7":
                out = get_output_path("frozen_filtered")
                result = formatter.freeze_and_filter(file, out)

            elif choice == "8":
                out = get_output_path("auto_fit")
                result = formatter.auto_fit_columns(file, out)

            elif choice == "9":
                out = get_output_path("with_totals")
                result = formatter.add_totals_row(file, out)

            elif choice == "10":
                col = prompt(f"Column to check for duplicates {cols}")
                out = get_output_path("duplicates_highlighted")
                result = formatter.highlight_duplicates(file, col, out)

            elif choice == "11":
                selected = pick_columns(cols, "Columns to format")
                info("Format examples: '#,##0.00'  |  '\"$\"#,##0.00'  |  '0.00%'  |  '#,##0'")
                fmt = prompt("Number format string")
                out = get_output_path("number_formatted")
                result = formatter.apply_number_format(file, selected, fmt, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 14 — VALIDATE DATA
# =============================================================================

def menu_validator():
    while True:
        banner()
        section("14. VALIDATE DATA")
        choice = menu_choice([
            ("1", "Check Mandatory Fields  — Flag rows missing required values"),
            ("2", "Validate Email          — Regex email format check"),
            ("3", "Validate Phone          — Phone number format check"),
            ("4", "Numeric Range Check     — Flag out-of-range values"),
            ("5", "Date Range Check        — Flag dates outside boundaries"),
            ("6", "Referential Integrity   — Values must exist in lookup file"),
            ("7", "Data Quality Report     — Comprehensive quality score 0-100"),
            ("8", "Detect PII              — Flag columns with sensitive data"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                req_cols = pick_columns(cols, "Select mandatory / required columns")
                out = get_output_path("mandatory_check")
                result = validator.check_mandatory_fields(file, req_cols, out)

            elif choice == "2":
                email_col = prompt(f"Email column {cols}")
                out = get_output_path("email_validation")
                result = validator.validate_email(file, email_col, out)

            elif choice == "3":
                phone_col = prompt(f"Phone column {cols}")
                out = get_output_path("phone_validation")
                result = validator.validate_phone(file, phone_col, out)

            elif choice == "4":
                col     = prompt(f"Numeric column to check {cols}")
                min_val = float(prompt("Minimum allowed value"))
                max_val = float(prompt("Maximum allowed value"))
                out = get_output_path("range_check")
                result = validator.validate_numeric_range(file, col, min_val, max_val, out)

            elif choice == "5":
                date_col = prompt(f"Date column {cols}")
                min_date = prompt("Minimum allowed date [YYYY-MM-DD]")
                max_date = prompt("Maximum allowed date [YYYY-MM-DD]")
                out = get_output_path("date_range_check")
                result = validator.validate_date_range(file, date_col, min_date, max_date, out)

            elif choice == "6":
                col      = prompt(f"Column to validate {cols}")
                ref_file = pick_single_file("Enter reference/lookup file path")
                if not ref_file: continue
                ref_cols = pd.read_excel(ref_file, nrows=0).columns.tolist()
                ref_col  = prompt(f"Reference column (must contain valid values) {ref_cols}")
                out = get_output_path("referential_integrity")
                result = validator.referential_integrity(file, col, ref_file, ref_col, out)

            elif choice == "7":
                out = get_output_path("data_quality_report")
                result = validator.data_quality_report(file, out)

            elif choice == "8":
                out = get_output_path("pii_detection")
                result = validator.detect_pii(file, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 15 — ANALYTICS
# =============================================================================

def menu_analytics():
    while True:
        banner()
        section("15. STATISTICAL ANALYTICS")
        choice = menu_choice([
            ("1", "Correlation Matrix     — Pairwise Pearson correlation"),
            ("2", "Pareto Analysis        — 80/20 cumulative % chart data"),
            ("3", "Linear Regression      — OLS + R² + predictions"),
            ("4", "Trend Forecast         — Extrapolate trend N periods ahead"),
            ("5", "Frequency Distribution — Histogram bin counts"),
            ("6", "Z-Score Analysis       — Outlier detection via std deviation"),
            ("7", "Cohort Retention       — Monthly customer retention matrix"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                selected = pick_columns(cols, "Select numeric columns for correlation (or ALL)")
                out = get_output_path("correlation_matrix")
                result = analytics.correlation_matrix(file, selected, out)

            elif choice == "2":
                category_col = prompt(f"Category column {cols}")
                value_col    = prompt(f"Value column {cols}")
                out = get_output_path("pareto_analysis")
                result = analytics.pareto_analysis(file, category_col, value_col, out)

            elif choice == "3":
                x_col = prompt(f"Independent variable (X) column {cols}")
                y_col = prompt(f"Dependent variable (Y) column {cols}")
                out = get_output_path("linear_regression")
                result = analytics.linear_regression(file, x_col, y_col, out)

            elif choice == "4":
                date_col  = prompt(f"Date column {cols}")
                value_col = prompt(f"Value column {cols}")
                periods   = int(prompt("Number of future periods to forecast") or "6")
                out = get_output_path("trend_forecast")
                result = analytics.trend_forecast(file, date_col, value_col, periods, out)

            elif choice == "5":
                col  = prompt(f"Numeric column {cols}")
                bins = int(prompt("Number of bins (default: 10)") or "10")
                out = get_output_path("frequency_distribution")
                result = analytics.frequency_distribution(file, col, bins, out)

            elif choice == "6":
                col       = prompt(f"Numeric column to analyze {cols}")
                threshold = float(prompt("Outlier threshold in σ (default: 3.0)") or "3.0")
                out = get_output_path("z_score_analysis")
                result = analytics.z_score_analysis(file, col, out, threshold=threshold)

            elif choice == "7":
                customer_col = prompt(f"Customer ID column {cols}")
                date_col     = prompt(f"Transaction date column {cols}")
                out = get_output_path("cohort_retention")
                result = analytics.cohort_retention(file, customer_col, date_col, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 16 — CONVERT FILES
# =============================================================================

def menu_converter():
    while True:
        banner()
        section("16. CONVERT FILES")
        choice = menu_choice([
            ("1", "Excel → CSV        — Each sheet to separate CSV"),
            ("2", "CSV → Excel        — Multiple CSVs into one workbook"),
            ("3", "Excel → JSON       — All sheets to JSON"),
            ("4", "JSON → Excel       — JSON array/object to Excel"),
            ("5", "XLS → XLSX Batch   — Convert old .xls files"),
            ("6", "Excel → Text       — Tab/pipe/custom delimited export"),
            ("7", "Merge CSV Files    — Stack multiple CSVs into one Excel"),
        ], "Select operation")

        if choice == "0":
            break

        try:
            if choice == "1":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                out_dir = str(OUTPUT_DIR / f"csv_export_{datetime.datetime.now().strftime('%H%M%S')}")
                created = converter.excel_to_csv(file, out_dir)
                success(f"Created {len(created)} CSV file(s) in: {out_dir}")
                pause(); continue

            elif choice == "2":
                info("Select CSV files to merge into Excel sheets")
                raw = prompt("Enter CSV file paths (comma-separated or folder)")
                paths = []
                for part in raw.split(","):
                    part = part.strip().strip('"').strip("'")
                    p = Path(part)
                    if p.is_dir():
                        paths.extend([str(f) for f in p.glob("*.csv")])
                    elif p.is_file():
                        paths.append(str(p))
                if not paths:
                    error("No CSV files found"); continue
                out = get_output_path("csv_merged")
                result = converter.csv_to_excel(paths, out)

            elif choice == "3":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                out = get_output_path("excel_export", ext=".json")
                result = converter.excel_to_json(file, out)

            elif choice == "4":
                json_file = prompt("Enter JSON file path").strip().strip('"')
                if not Path(json_file).is_file():
                    error(f"File not found: {json_file}"); continue
                out = get_output_path("from_json")
                result = converter.json_to_excel(json_file, out)

            elif choice == "5":
                files = pick_files("Enter .xls file paths (or folder)")
                if not files: continue
                out_dir = str(OUTPUT_DIR / f"xlsx_batch_{datetime.datetime.now().strftime('%H%M%S')}")
                created = converter.xls_to_xlsx_batch(files, out_dir)
                success(f"Converted {len(created)} file(s) to: {out_dir}")
                pause(); continue

            elif choice == "6":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                delim = prompt("Delimiter: tab=\\t  pipe=|  comma=,  (default: tab)") or "\t"
                if delim.lower() in ("\\t", "tab"): delim = "\t"
                out_dir = str(OUTPUT_DIR / f"text_export_{datetime.datetime.now().strftime('%H%M%S')}")
                created = converter.excel_to_text(file, out_dir, delimiter=delim)
                success(f"Created {len(created)} file(s) in: {out_dir}")
                pause(); continue

            elif choice == "7":
                raw = prompt("Enter CSV file paths (comma-separated or folder)")
                paths = []
                for part in raw.split(","):
                    part = part.strip().strip('"').strip("'")
                    p = Path(part)
                    if p.is_dir():
                        paths.extend([str(f) for f in p.glob("*.csv")])
                    elif p.is_file():
                        paths.append(str(p))
                if not paths:
                    error("No CSV files found"); continue
                out = get_output_path("csv_stacked")
                result = converter.merge_csv_files(paths, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 17 — LOOKUP & MATCH
# =============================================================================

def menu_lookup():
    while True:
        banner()
        section("17. LOOKUP & MATCH")
        choice = menu_choice([
            ("1", "VLOOKUP           — Join columns from a reference file"),
            ("2", "Fuzzy Match       — Approximate string matching"),
            ("3", "Multi-Key Lookup  — Multi-column JOIN"),
            ("4", "Reverse Lookup    — Find key by value"),
            ("5", "Enrich from Ref   — Add columns from a master reference"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter main Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                lookup_col  = prompt(f"Lookup column in main file {cols}")
                ref_file    = pick_single_file("Enter reference/lookup file path")
                if not ref_file: continue
                ref_cols    = pd.read_excel(ref_file, nrows=0).columns.tolist()
                ref_col     = prompt(f"Key column in reference file {ref_cols}")
                return_raw  = prompt(f"Columns to return from reference (comma-sep or ALL) {ref_cols}")
                if return_raw.strip().upper() == "ALL":
                    return_cols = [c for c in ref_cols if c != ref_col]
                else:
                    return_cols = [c.strip() for c in return_raw.split(",") if c.strip()]
                how = prompt("Join type [left/inner/outer] (default: left)") or "left"
                out = get_output_path("vlookup_result")
                result = lookup.vlookup(file, lookup_col, ref_file, ref_col, return_cols, out, how=how)

            elif choice == "2":
                col      = prompt(f"Column to match {cols}")
                ref_file = pick_single_file("Enter reference file path")
                if not ref_file: continue
                ref_cols = pd.read_excel(ref_file, nrows=0).columns.tolist()
                ref_col  = prompt(f"Reference column to match against {ref_cols}")
                thresh   = float(prompt("Match threshold 0-1 (default: 0.75)") or "0.75")
                out = get_output_path("fuzzy_match")
                result = lookup.fuzzy_match(file, col, ref_file, ref_col, out, threshold=thresh)

            elif choice == "3":
                lookup_cols = pick_columns(cols, "Key columns for JOIN (must exist in reference file too)")
                ref_file    = pick_single_file("Enter reference file path")
                if not ref_file: continue
                how = prompt("Join type [left/inner/outer] (default: left)") or "left"
                out = get_output_path("multi_key_lookup")
                result = lookup.multi_key_lookup(file, lookup_cols, ref_file, out, how=how)

            elif choice == "4":
                value_col = prompt(f"Column with values to look up {cols}")
                ref_file  = pick_single_file("Enter reference file path")
                if not ref_file: continue
                ref_cols  = pd.read_excel(ref_file, nrows=0).columns.tolist()
                key_col   = prompt(f"Key column in reference {ref_cols}")
                val_col   = prompt(f"Value column in reference {ref_cols}")
                out = get_output_path("reverse_lookup")
                result = lookup.reverse_lookup(file, value_col, ref_file, key_col, val_col, out)

            elif choice == "5":
                join_col      = prompt(f"Join column in main file {cols}")
                ref_file      = pick_single_file("Enter reference file path")
                if not ref_file: continue
                ref_cols      = pd.read_excel(ref_file, nrows=0).columns.tolist()
                ref_join_col  = prompt(f"Join column in reference file {ref_cols}")
                enrich_raw    = prompt(f"Columns to bring from reference (comma-sep or ALL) {ref_cols}")
                if enrich_raw.strip().upper() == "ALL":
                    enrich_cols = [c for c in ref_cols if c != ref_join_col]
                else:
                    enrich_cols = [c.strip() for c in enrich_raw.split(",") if c.strip()]
                out = get_output_path("enriched_data")
                result = lookup.enrich_from_lookup(file, join_col, ref_file, ref_join_col, enrich_cols, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MAIN MENU
# =============================================================================

def main():
    while True:
        banner()
        print(Fore.WHITE + Style.BRIGHT + "  MAIN MENU\n")
        options = [
            ("1",  "Consolidate Files       — stack, join, merge multiple Excel files"),
            ("2",  "Calculate & Analyze     — efficiency, KPIs, variance, stats"),
            ("3",  "Clean Data              — duplicates, blanks, formats, types"),
            ("4",  "Transform Data          — pivot, split, transpose, reshape"),
            ("5",  "Compare Files           — diff, new/deleted rows, changes"),
            ("6",  "Column Operations       — rename, split, merge, calculate columns"),
            ("7",  "Generate Reports        — summary, profile, KPI, frequency"),
            ("8",  "Quick File Preview      — peek at any Excel file"),
            ("9",  "Finance                 — aging, payroll, ratios, amortization"),
            ("10", "HR Analytics            — attrition, headcount, tenure, salary"),
            ("11", "Sales Analytics         — commission, RFM, quota, pipeline"),
            ("12", "Inventory Management    — ABC, reorder point, OEE, dead stock"),
            ("13", "Format & Style          — charts, traffic lights, table styles"),
            ("14", "Validate Data           — email, phone, range, PII detection"),
            ("15", "Statistical Analytics   — correlation, Pareto, regression, Z-score"),
            ("16", "Convert Files           — Excel↔CSV, Excel↔JSON, XLS→XLSX"),
            ("17", "Lookup & Match          — VLOOKUP, fuzzy match, multi-key join"),
        ]
        for key, label in options:
            print(f"    {Fore.CYAN}{key:>2}{Style.RESET_ALL}  {label}")
        print(f"\n    {Fore.RED} 0{Style.RESET_ALL}  Exit\n")

        choice = prompt("Select menu")

        if choice == "0":
            print(Fore.GREEN + "\n  Goodbye! All outputs saved to: " + str(OUTPUT_DIR))
            sys.exit(0)
        elif choice == "1":  menu_consolidate()
        elif choice == "2":  menu_calculate()
        elif choice == "3":  menu_clean()
        elif choice == "4":  menu_transform()
        elif choice == "5":  menu_compare()
        elif choice == "6":  menu_columns()
        elif choice == "7":  menu_reports()
        elif choice == "8":  menu_preview()
        elif choice == "9":  menu_finance()
        elif choice == "10": menu_hr()
        elif choice == "11": menu_sales()
        elif choice == "12": menu_inventory()
        elif choice == "13": menu_formatter()
        elif choice == "14": menu_validator()
        elif choice == "15": menu_analytics()
        elif choice == "16": menu_converter()
        elif choice == "17": menu_lookup()
        else:
            error("Invalid choice — please enter a number from the menu")
            pause()


if __name__ == "__main__":
    MODULE_MAP = {
        # Original modules
        "consolidate": menu_consolidate,
        "calculate":   menu_calculate,
        "clean":       menu_clean,
        "transform":   menu_transform,
        "compare":     menu_compare,
        "columns":     menu_columns,
        "reports":     menu_reports,
        "preview":     menu_preview,
        # New modules
        "finance":     menu_finance,
        "hr":          menu_hr,
        "sales":       menu_sales,
        "inventory":   menu_inventory,
        "format":      menu_formatter,
        "validate":    menu_validator,
        "analytics":   menu_analytics,
        "convert":     menu_converter,
        "lookup":      menu_lookup,
    }

    if len(sys.argv) > 1:
        key = sys.argv[1].lower()
        target = MODULE_MAP.get(key)
        if target:
            banner()
            target()
        else:
            print(f"  [ERROR] Unknown module: '{key}'")
            print(f"  Available: {', '.join(MODULE_MAP.keys())}")
            sys.exit(1)
    else:
        main()
