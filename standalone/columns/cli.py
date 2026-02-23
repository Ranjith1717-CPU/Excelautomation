"""Standalone Columns — Excel Automation Toolkit"""
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
import column_ops


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_columns()
