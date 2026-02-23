"""
=============================================================================
  EXCEL AUTOMATION TOOLKIT  v1.0
  All-in-one Excel automation powered by Python + pandas
=============================================================================
  Run this file directly:   python main.py
  Or use the launcher:      run.bat  (Windows)
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
from modules import consolidator, calculator, cleaner, transformer, comparator, reporter, column_ops

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
║          EXCEL AUTOMATION TOOLKIT  v1.0                  ║
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
# MAIN MENU
# =============================================================================

def main():
    while True:
        banner()
        print(Fore.WHITE + Style.BRIGHT + "  MAIN MENU\n")
        options = [
            ("1", "Consolidate Files       — stack, join, merge multiple Excel files"),
            ("2", "Calculate & Analyze     — efficiency, KPIs, variance, stats"),
            ("3", "Clean Data              — duplicates, blanks, formats, types"),
            ("4", "Transform Data          — pivot, split, transpose, reshape"),
            ("5", "Compare Files           — diff, new/deleted rows, changes"),
            ("6", "Column Operations       — rename, split, merge, calculate columns"),
            ("7", "Generate Reports        — summary, profile, KPI, frequency"),
            ("8", "Quick File Preview      — peek at any Excel file"),
        ]
        for key, label in options:
            print(f"    {Fore.CYAN}{key}{Style.RESET_ALL}  {label}")
        print(f"\n    {Fore.RED}0{Style.RESET_ALL}  Exit\n")

        choice = prompt("Select menu")

        if choice == "0":
            print(Fore.GREEN + "\n  Goodbye! All outputs saved to: " + str(OUTPUT_DIR))
            sys.exit(0)
        elif choice == "1": menu_consolidate()
        elif choice == "2": menu_calculate()
        elif choice == "3": menu_clean()
        elif choice == "4": menu_transform()
        elif choice == "5": menu_compare()
        elif choice == "6": menu_columns()
        elif choice == "7": menu_reports()
        elif choice == "8": menu_preview()
        else:
            error("Invalid choice — please enter a number from the menu")
            pause()


if __name__ == "__main__":
    main()
