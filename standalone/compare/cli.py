"""Standalone Compare — Excel Automation Toolkit"""
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
import comparator


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_compare()
