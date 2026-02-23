"""Standalone Hr — Excel Automation Toolkit"""
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
import hr


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_hr()
