"""Standalone Transform — Excel Automation Toolkit"""
# ── Bootstrap ──────────────────────────────────────────────────────────────
def _check_deps():
    import subprocess as _sp, sys as _sys
    missing = []
    for pkg in ['pandas', 'openpyxl', 'xlrd', 'colorama', 'tabulate', 'numpy']:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"[INFO] Installing missing packages: {missing}")
        _sp.check_call([_sys.executable, "-m", "pip", "install"] + missing + ["-q"])

_check_deps()

# ── Imports ────────────────────────────────────────────────────────────────
import os, sys, glob, datetime
from pathlib import Path
import pandas as pd
from colorama import init, Fore, Style
init(autoreset=True)

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


# ── UI Helpers ─────────────────────────────────────────────────────────────
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


# ── Module ─────────────────────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
import transformer


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_transform()
