"""Standalone Calculate — Excel Automation Toolkit"""
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
import calculator


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_calculate()
