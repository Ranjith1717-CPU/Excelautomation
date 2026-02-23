"""Standalone Converter — Excel Automation Toolkit"""
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
import converter


# ── Menu ───────────────────────────────────────────────────────────────────
def menu_converter():
    while True:
        banner()
        section("16. CONVERT FILES")
        choice = menu_choice([
            ("1", "Excel → CSV        — Each sheet to separate CSV"),
            ("2", "CSV → Excel        — Multiple CSVs into one workbook"),
            ("3", "Excel → JSON       — All sheets to JSON"),
            ("4", "JSON → Excel       — JSON array/object to Excel"),
            ("5", "XLS → XLSX Batch   — Convert old .xls files"),
            ("6", "Excel → Text       — Tab/pipe/custom delimited export"),
            ("7", "Merge CSV Files    — Stack multiple CSVs into one Excel"),
        ], "Select operation")

        if choice == "0":
            break

        try:
            if choice == "1":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                out_dir = str(OUTPUT_DIR / f"csv_export_{datetime.datetime.now().strftime('%H%M%S')}")
                created = converter.excel_to_csv(file, out_dir)
                success(f"Created {len(created)} CSV file(s) in: {out_dir}")
                pause(); continue

            elif choice == "2":
                info("Select CSV files to merge into Excel sheets")
                raw = prompt("Enter CSV file paths (comma-separated or folder)")
                paths = []
                for part in raw.split(","):
                    part = part.strip().strip('"').strip("'")
                    p = Path(part)
                    if p.is_dir():
                        paths.extend([str(f) for f in p.glob("*.csv")])
                    elif p.is_file():
                        paths.append(str(p))
                if not paths:
                    error("No CSV files found"); continue
                out = get_output_path("csv_merged")
                result = converter.csv_to_excel(paths, out)

            elif choice == "3":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                out = get_output_path("excel_export", ext=".json")
                result = converter.excel_to_json(file, out)

            elif choice == "4":
                json_file = prompt("Enter JSON file path").strip().strip('"')
                if not Path(json_file).is_file():
                    error(f"File not found: {json_file}"); continue
                out = get_output_path("from_json")
                result = converter.json_to_excel(json_file, out)

            elif choice == "5":
                files = pick_files("Enter .xls file paths (or folder)")
                if not files: continue
                out_dir = str(OUTPUT_DIR / f"xlsx_batch_{datetime.datetime.now().strftime('%H%M%S')}")
                created = converter.xls_to_xlsx_batch(files, out_dir)
                success(f"Converted {len(created)} file(s) to: {out_dir}")
                pause(); continue

            elif choice == "6":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                delim = prompt("Delimiter: tab=\\t  pipe=|  comma=,  (default: tab)") or "\t"
                if delim.lower() in ("\\t", "tab"): delim = "\t"
                out_dir = str(OUTPUT_DIR / f"text_export_{datetime.datetime.now().strftime('%H%M%S')}")
                created = converter.excel_to_text(file, out_dir, delimiter=delim)
                success(f"Created {len(created)} file(s) in: {out_dir}")
                pause(); continue

            elif choice == "7":
                raw = prompt("Enter CSV file paths (comma-separated or folder)")
                paths = []
                for part in raw.split(","):
                    part = part.strip().strip('"').strip("'")
                    p = Path(part)
                    if p.is_dir():
                        paths.extend([str(f) for f in p.glob("*.csv")])
                    elif p.is_file():
                        paths.append(str(p))
                if not paths:
                    error("No CSV files found"); continue
                out = get_output_path("csv_stacked")
                result = converter.merge_csv_files(paths, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MENU 17 — LOOKUP & MATCH
# =============================================================================


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_converter()
