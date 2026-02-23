"""Standalone Consolidate — Excel Automation Toolkit"""
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
import consolidator


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_consolidate()
