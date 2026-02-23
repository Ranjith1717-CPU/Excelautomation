"""Standalone Finance — Excel Automation Toolkit"""
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
import finance


# ── Menu ───────────────────────────────────────────────────────────────────
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


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_finance()
