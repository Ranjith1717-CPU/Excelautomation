"""Standalone Clean — Excel Automation Toolkit"""
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
import cleaner


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_clean()
