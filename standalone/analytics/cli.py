"""Standalone Analytics — Excel Automation Toolkit"""
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
import analytics


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_analytics()
