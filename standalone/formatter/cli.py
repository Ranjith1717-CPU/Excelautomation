"""Standalone Formatter — Excel Automation Toolkit"""
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
import formatter


# ── Menu ───────────────────────────────────────────────────────────────────
def menu_formatter():
    while True:
        banner()
        section("13. FORMAT & STYLE")
        choice = menu_choice([
            ("1",  "Add Bar Chart          — Embedded column chart"),
            ("2",  "Add Line Chart         — Embedded line chart"),
            ("3",  "Add Pie Chart          — Embedded pie chart"),
            ("4",  "Traffic Light Colors   — Red/Yellow/Green cell fills"),
            ("5",  "Color Scale            — Gradient min→max fill"),
            ("6",  "Format as Table        — Apply Excel table style"),
            ("7",  "Freeze & Filter        — Freeze header + auto-filter"),
            ("8",  "Auto-fit Columns       — Width based on content"),
            ("9",  "Add Totals Row         — SUM row at bottom"),
            ("10", "Highlight Duplicates   — Yellow fill on duplicate cells"),
            ("11", "Apply Number Format    — Currency/percentage/comma"),
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
                x_col = prompt(f"Category/X-axis column {cols}")
                y_col = prompt(f"Value/Y-axis column {cols}")
                title = prompt("Chart title (or Enter for default)") or "Bar Chart"
                out = get_output_path("bar_chart")
                result = formatter.add_bar_chart(file, x_col, y_col, out, title=title)

            elif choice == "2":
                x_col = prompt(f"X-axis column {cols}")
                y_col = prompt(f"Y-axis column {cols}")
                title = prompt("Chart title (or Enter for default)") or "Line Chart"
                out = get_output_path("line_chart")
                result = formatter.add_line_chart(file, x_col, y_col, out, title=title)

            elif choice == "3":
                cat_col = prompt(f"Category column {cols}")
                val_col = prompt(f"Value column {cols}")
                title   = prompt("Chart title (or Enter for default)") or "Pie Chart"
                out = get_output_path("pie_chart")
                result = formatter.add_pie_chart(file, cat_col, val_col, out, title=title)

            elif choice == "4":
                col = prompt(f"Numeric column to color {cols}")
                info("Leave thresholds blank to use auto 33rd/66th percentile")
                red_raw    = prompt("Red threshold (below this = red)")
                yellow_raw = prompt("Yellow threshold (below this = yellow)")
                red    = float(red_raw)    if red_raw    else None
                yellow = float(yellow_raw) if yellow_raw else None
                out = get_output_path("traffic_light")
                result = formatter.apply_traffic_light(file, col, out, red=red, yellow=yellow)

            elif choice == "5":
                col = prompt(f"Numeric column for color scale {cols}")
                out = get_output_path("color_scale")
                result = formatter.apply_color_scale(file, col, out)

            elif choice == "6":
                style = prompt("Table style (default: TableStyleMedium9)") or "TableStyleMedium9"
                out = get_output_path("table_styled")
                result = formatter.format_as_table(file, out, style=style)

            elif choice == "7":
                out = get_output_path("frozen_filtered")
                result = formatter.freeze_and_filter(file, out)

            elif choice == "8":
                out = get_output_path("auto_fit")
                result = formatter.auto_fit_columns(file, out)

            elif choice == "9":
                out = get_output_path("with_totals")
                result = formatter.add_totals_row(file, out)

            elif choice == "10":
                col = prompt(f"Column to check for duplicates {cols}")
                out = get_output_path("duplicates_highlighted")
                result = formatter.highlight_duplicates(file, col, out)

            elif choice == "11":
                selected = pick_columns(cols, "Columns to format")
                info("Format examples: '#,##0.00'  |  '\"$\"#,##0.00'  |  '0.00%'  |  '#,##0'")
                fmt = prompt("Number format string")
                out = get_output_path("number_formatted")
                result = formatter.apply_number_format(file, selected, fmt, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 14 — VALIDATE DATA
# =============================================================================


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_formatter()
