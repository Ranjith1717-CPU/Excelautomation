"""Standalone Lookup — Excel Automation Toolkit"""
# ── Bootstrap ──────────────────────────────────────────────────────────────
def _check_deps():
    import subprocess as _sp, sys as _sys
    missing = []
    for pkg in ['pandas', 'openpyxl', 'xlrd', 'colorama', 'tabulate', 'numpy', 'rapidfuzz']:
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
import lookup


# ── Menu ───────────────────────────────────────────────────────────────────
def menu_lookup():
    while True:
        banner()
        section("17. LOOKUP & MATCH")
        choice = menu_choice([
            ("1", "VLOOKUP           — Join columns from a reference file"),
            ("2", "Fuzzy Match       — Approximate string matching"),
            ("3", "Multi-Key Lookup  — Multi-column JOIN"),
            ("4", "Reverse Lookup    — Find key by value"),
            ("5", "Enrich from Ref   — Add columns from a master reference"),
        ], "Select operation")

        if choice == "0":
            break

        file = pick_single_file("Enter main Excel file path")
        if not file: continue

        try:
            cols = pd.read_excel(file, nrows=0).columns.tolist()
        except Exception as e:
            error(str(e)); pause(); continue

        try:
            if choice == "1":
                lookup_col  = prompt(f"Lookup column in main file {cols}")
                ref_file    = pick_single_file("Enter reference/lookup file path")
                if not ref_file: continue
                ref_cols    = pd.read_excel(ref_file, nrows=0).columns.tolist()
                ref_col     = prompt(f"Key column in reference file {ref_cols}")
                return_raw  = prompt(f"Columns to return from reference (comma-sep or ALL) {ref_cols}")
                if return_raw.strip().upper() == "ALL":
                    return_cols = [c for c in ref_cols if c != ref_col]
                else:
                    return_cols = [c.strip() for c in return_raw.split(",") if c.strip()]
                how = prompt("Join type [left/inner/outer] (default: left)") or "left"
                out = get_output_path("vlookup_result")
                result = lookup.vlookup(file, lookup_col, ref_file, ref_col, return_cols, out, how=how)

            elif choice == "2":
                col      = prompt(f"Column to match {cols}")
                ref_file = pick_single_file("Enter reference file path")
                if not ref_file: continue
                ref_cols = pd.read_excel(ref_file, nrows=0).columns.tolist()
                ref_col  = prompt(f"Reference column to match against {ref_cols}")
                thresh   = float(prompt("Match threshold 0-1 (default: 0.75)") or "0.75")
                out = get_output_path("fuzzy_match")
                result = lookup.fuzzy_match(file, col, ref_file, ref_col, out, threshold=thresh)

            elif choice == "3":
                lookup_cols = pick_columns(cols, "Key columns for JOIN (must exist in reference file too)")
                ref_file    = pick_single_file("Enter reference file path")
                if not ref_file: continue
                how = prompt("Join type [left/inner/outer] (default: left)") or "left"
                out = get_output_path("multi_key_lookup")
                result = lookup.multi_key_lookup(file, lookup_cols, ref_file, out, how=how)

            elif choice == "4":
                value_col = prompt(f"Column with values to look up {cols}")
                ref_file  = pick_single_file("Enter reference file path")
                if not ref_file: continue
                ref_cols  = pd.read_excel(ref_file, nrows=0).columns.tolist()
                key_col   = prompt(f"Key column in reference {ref_cols}")
                val_col   = prompt(f"Value column in reference {ref_cols}")
                out = get_output_path("reverse_lookup")
                result = lookup.reverse_lookup(file, value_col, ref_file, key_col, val_col, out)

            elif choice == "5":
                join_col      = prompt(f"Join column in main file {cols}")
                ref_file      = pick_single_file("Enter reference file path")
                if not ref_file: continue
                ref_cols      = pd.read_excel(ref_file, nrows=0).columns.tolist()
                ref_join_col  = prompt(f"Join column in reference file {ref_cols}")
                enrich_raw    = prompt(f"Columns to bring from reference (comma-sep or ALL) {ref_cols}")
                if enrich_raw.strip().upper() == "ALL":
                    enrich_cols = [c for c in ref_cols if c != ref_join_col]
                else:
                    enrich_cols = [c.strip() for c in enrich_raw.split(",") if c.strip()]
                out = get_output_path("enriched_data")
                result = lookup.enrich_from_lookup(file, join_col, ref_file, ref_join_col, enrich_cols, out)
            else:
                continue

            success(f"Saved → {result}")
        except Exception as e:
            error(str(e))
        pause()


# =============================================================================
# MAIN MENU
# =============================================================================


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    menu_lookup()
