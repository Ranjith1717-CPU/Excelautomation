"""
generate_standalone.py
======================
Run once (or re-run after updates) from the toolkit root to produce
fully self-contained standalone module folders.

Usage:
    python3 generate_standalone.py

Output:
    standalone/
        consolidate/   (consolidator.py + cli.py + run.bat)
        calculate/
        clean/
        ... (16 folders total)
"""
import re
import shutil
from pathlib import Path

ROOT         = Path(__file__).parent
MODULES_DIR  = ROOT / "modules"
STANDALONE   = ROOT / "standalone"
MAIN_PY      = ROOT / "main.py"

# ─────────────────────────────────────────────────────────────────────────────
# Module table
# (folder, module_file, import_name, menu_fn, extra_pip)
# ─────────────────────────────────────────────────────────────────────────────
MODULES = [
    ("consolidate", "consolidator.py", "consolidator", "menu_consolidate", []),
    ("calculate",   "calculator.py",   "calculator",   "menu_calculate",   []),
    ("clean",       "cleaner.py",      "cleaner",      "menu_clean",       []),
    ("transform",   "transformer.py",  "transformer",  "menu_transform",   []),
    ("compare",     "comparator.py",   "comparator",   "menu_compare",     []),
    ("columns",     "column_ops.py",   "column_ops",   "menu_columns",     []),
    ("reports",     "reporter.py",     "reporter",     "menu_reports",     []),
    ("finance",     "finance.py",      "finance",      "menu_finance",     []),
    ("hr",          "hr.py",           "hr",           "menu_hr",          []),
    ("sales",       "sales.py",        "sales",        "menu_sales",       []),
    ("inventory",   "inventory.py",    "inventory",    "menu_inventory",   []),
    ("formatter",   "formatter.py",    "formatter",    "menu_formatter",   []),
    ("validator",   "validator.py",    "validator",    "menu_validator",   []),
    ("analytics",   "analytics.py",    "analytics",    "menu_analytics",   []),
    ("converter",   "converter.py",    "converter",    "menu_converter",   []),
    ("lookup",      "lookup.py",       "lookup",       "menu_lookup",      ["rapidfuzz"]),
]

# UI helper functions to extract verbatim from main.py
HELPERS = [
    "banner",
    "section",
    "success",
    "error",
    "info",
    "prompt",
    "menu_choice",
    "pause",
    "get_output_path",
    "pick_files",
    "pick_single_file",
    "pick_columns",
    "preview_file",
]


# ─────────────────────────────────────────────────────────────────────────────
# Extraction helpers
# ─────────────────────────────────────────────────────────────────────────────

def extract_function(src: str, fn_name: str) -> str:
    """Extract a complete top-level `def fn_name(...):` block from source."""
    start_re = re.compile(rf'^def {re.escape(fn_name)}\b', re.MULTILINE)
    m = start_re.search(src)
    if not m:
        raise ValueError(f"Function '{fn_name}' not found in source")

    start = m.start()

    # Find the next top-level `def ` (col 0) that comes AFTER this one
    next_def_re = re.compile(r'^def ', re.MULTILINE)
    end = len(src)
    for nm in next_def_re.finditer(src, start + 1):
        end = nm.start()
        break

    return src[start:end].rstrip()


# ─────────────────────────────────────────────────────────────────────────────
# Content builders
# ─────────────────────────────────────────────────────────────────────────────

def build_check_deps(extra_pip: list) -> str:
    """Return a _check_deps() function string that installs required packages."""
    base = ["pandas", "openpyxl", "xlrd", "colorama", "tabulate", "numpy"]
    all_pkgs = base + extra_pip
    pkgs_repr = repr(all_pkgs)
    return f"""\
def _check_deps():
    import subprocess as _sp, sys as _sys
    missing = []
    for pkg in {pkgs_repr}:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"[INFO] Installing missing packages: {{missing}}")
        _sp.check_call([_sys.executable, "-m", "pip", "install"] + missing + ["-q"])"""


def build_cli(folder: str, import_name: str, menu_fn: str,
              extra_pip: list, main_src: str) -> str:
    """Build the full cli.py content for one standalone module."""
    title = folder.capitalize()

    # Gather helpers
    helper_blocks = []
    for fn in HELPERS:
        helper_blocks.append(extract_function(main_src, fn))
    helpers_str = "\n\n\n".join(helper_blocks)

    # Extract the menu function
    menu_block = extract_function(main_src, menu_fn)

    check_deps = build_check_deps(extra_pip)

    return f'''\
"""Standalone {title} — Excel Automation Toolkit"""
# ── Bootstrap ──────────────────────────────────────────────────────────────
{check_deps}

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
{helpers_str}


# ── Module ─────────────────────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
import {import_name}


# ── Menu ───────────────────────────────────────────────────────────────────
{menu_block}


# ── Entry ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    {menu_fn}()
'''


def build_run_bat(folder: str, extra_pip: list) -> str:
    """Build run.bat content for one standalone module."""
    title = folder.capitalize()
    extra_line = ""
    if extra_pip:
        extra_line = "\npip install " + " ".join(extra_pip) + " --quiet\n"
    return (
        f"@echo off\n"
        f"title Excel Automation \u2014 {title}\n"
        f"color 0B\n"
        f"pip install pandas openpyxl xlrd colorama tabulate numpy --quiet --upgrade\n"
        f"{extra_line}"
        f'cd /d "%~dp0"\n'
        f"python cli.py\n"
        f"pause\n"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    if not MAIN_PY.exists():
        raise FileNotFoundError(f"main.py not found at {MAIN_PY}")

    main_src = MAIN_PY.read_text(encoding="utf-8")
    STANDALONE.mkdir(exist_ok=True)

    count = 0
    errors = []

    for folder, module_file, import_name, menu_fn, extra_pip in MODULES:
        dest = STANDALONE / folder
        dest.mkdir(exist_ok=True)

        try:
            # 1. Copy module file
            shutil.copy2(MODULES_DIR / module_file, dest / module_file)

            # 2. Build and write cli.py
            cli_content = build_cli(folder, import_name, menu_fn, extra_pip, main_src)
            (dest / "cli.py").write_text(cli_content, encoding="utf-8")

            # 3. Write run.bat
            bat_content = build_run_bat(folder, extra_pip)
            (dest / "run.bat").write_text(bat_content, encoding="utf-8")

            count += 1
            print(f"  \u2713  standalone/{folder}/")

        except Exception as exc:
            errors.append((folder, exc))
            print(f"  \u2717  standalone/{folder}/  ERROR: {exc}")

    print()
    if errors:
        print(f"{len(errors)} error(s) encountered (see above).")
    else:
        print(f"{count} standalone modules created")
    print(f"Location: {STANDALONE}")


if __name__ == "__main__":
    main()
