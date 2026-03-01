"""
ask.py — Natural Language CLI for Excel Automation Toolkit
===========================================================
Usage:
    python ask.py sales.xlsx "remove duplicates"
    python ask.py reports/ "consolidate all files"
    python ask.py                         # interactive mode
    ask.bat sales.xlsx "remove duplicates" # Windows

Flow:
  1. Parse CLI args → (files, query)
  2. nl_router.parse_intent → ranked matches
  3. Display matches with confidence bars
  4. Confirm / select operation
  5. Collect missing parameters interactively
  6. Confirm final call
  7. Execute → print output path
"""

import sys
import os
import glob
import importlib
from pathlib import Path

# ── Bootstrap ─────────────────────────────────────────────────────────────────
def _bootstrap():
    for pkg in ["pandas", "openpyxl", "colorama"]:
        try:
            __import__(pkg)
        except ImportError:
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_bootstrap()

from colorama import init, Fore, Style
init(autoreset=True)

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))
from nl_router import (
    parse_intent, inspect_file, get_columns_from_file,
    get_output_path, get_output_dir, extract_number_from_query,
    split_compound_query,
)


# =============================================================================
# UI HELPERS
# =============================================================================

def _c(color, text: str) -> str:
    return color + text + Style.RESET_ALL


def banner():
    print(_c(Fore.CYAN + Style.BRIGHT, """
╔══════════════════════════════════════════════════════════╗
║   🧠  Excel Natural Language Toolkit                      ║
║       Tell it what to do — no menus needed               ║
╚══════════════════════════════════════════════════════════╝"""))


def conf_bar(score: float, width: int = 20) -> str:
    filled = int(score * width)
    bar = "█" * filled + "░" * (width - filled)
    pct = f"{score*100:.0f}%"
    if score >= 0.6:
        color = Fore.GREEN
    elif score >= 0.3:
        color = Fore.YELLOW
    else:
        color = Fore.RED
    return color + bar + Style.RESET_ALL + f"  {pct}"


def ask_input(prompt_text: str, default: str = "") -> str:
    default_hint = f" [{default}]" if default else ""
    raw = input(_c(Fore.WHITE + Style.BRIGHT, f"\n  → {prompt_text}{default_hint}: ")).strip()
    return raw if raw else default


def ask_yn(prompt_text: str, default: bool = True) -> bool:
    hint = "[Y/n]" if default else "[y/N]"
    raw = input(_c(Fore.WHITE + Style.BRIGHT, f"\n  → {prompt_text} {hint}: ")).strip().lower()
    if not raw:
        return default
    return raw.startswith("y")


def show_columns(columns: list):
    if not columns:
        return
    print(_c(Fore.CYAN, "\n  Available columns:"))
    for i, c in enumerate(columns, 1):
        print(f"    {i:3}. {c}")


def ask_column(prompt_text: str, columns: list, required: bool = True) -> str:
    """Prompt for a single column name. Shows available columns."""
    show_columns(columns)
    while True:
        raw = ask_input(prompt_text + " (name or #)")
        if not raw:
            if not required:
                return ""
            print(_c(Fore.RED, "  This field is required."))
            continue
        # Number shortcut
        if raw.isdigit():
            idx = int(raw) - 1
            if 0 <= idx < len(columns):
                return columns[idx]
        # Name match (case-insensitive)
        match = next((c for c in columns if c.lower() == raw.lower()), None)
        if match:
            return match
        # Partial match
        partial = [c for c in columns if raw.lower() in c.lower()]
        if len(partial) == 1:
            print(f"  (using: {partial[0]})")
            return partial[0]
        print(_c(Fore.RED, f"  Not found: {raw}. Try again."))


def ask_columns(prompt_text: str, columns: list, required: bool = True) -> list:
    """Prompt for comma-separated column names."""
    show_columns(columns)
    hint = "(comma-sep, #'s, or Enter=all)" if not required else "(comma-sep or #'s)"
    raw = ask_input(f"{prompt_text} {hint}")
    if not raw:
        if not required:
            return []  # caller interprets as "all"
        print(_c(Fore.RED, "  At least one column required."))
        return ask_columns(prompt_text, columns, required)
    result = []
    for part in raw.split(","):
        part = part.strip()
        if part.isdigit():
            idx = int(part) - 1
            if 0 <= idx < len(columns):
                result.append(columns[idx])
        else:
            match = next((c for c in columns if c.lower() == part.lower()), None)
            if match:
                result.append(match)
    return result if result else []


def ask_mapping(prompt_text: str) -> dict:
    """Prompt for key:value pairs. e.g. 'old_name:new_name, A:B'"""
    raw = ask_input(f"{prompt_text} (old:new, comma-sep)")
    result = {}
    for pair in raw.split(","):
        pair = pair.strip()
        if ":" in pair:
            k, v = pair.split(":", 1)
            result[k.strip()] = v.strip()
    return result


# =============================================================================
# FILE RESOLUTION
# =============================================================================

def resolve_files(raw_paths: list) -> list:
    """Expand paths: single files, directories, globs."""
    files = []
    for raw in raw_paths:
        raw = raw.strip().strip('"').strip("'")
        p = Path(raw)
        if "*" in raw or "?" in raw:
            matched = glob.glob(raw)
            files.extend([m for m in matched
                          if m.lower().endswith((".xlsx", ".xls", ".xlsm", ".csv"))])
        elif p.is_dir():
            for ext in ["*.xlsx", "*.xls", "*.xlsm", "*.csv"]:
                files.extend([str(f) for f in p.glob(ext)])
        elif p.is_file():
            files.append(str(p))
        else:
            print(_c(Fore.RED, f"  Not found: {raw}"))
    return files


# =============================================================================
# PARAMETER COLLECTION
# =============================================================================

def collect_params(intent: dict, files: list, query: str) -> dict:
    """
    Collect all required parameters for the intent interactively.
    Returns a kwargs dict ready for function call.
    """
    kwargs = {}
    primary_file = files[0] if files else ""
    columns = get_columns_from_file(primary_file) if primary_file else []

    for p in intent["params"]:
        ptype = p["type"]
        name  = p["name"]
        prompt_text = p.get("prompt", name)
        default = p.get("default", "")
        options = p.get("options", [])

        if ptype == "file":
            if not primary_file:
                primary_file = ask_input("Input file path")
                columns = get_columns_from_file(primary_file)
            kwargs[name] = primary_file

        elif ptype == "files":
            kwargs[name] = files if files else [ask_input("Input file(s) path")]

        elif ptype == "file1":
            kwargs[name] = files[0] if len(files) >= 1 else ask_input("First file path")

        elif ptype == "file2":
            if len(files) >= 2:
                kwargs[name] = files[1]
            else:
                kwargs[name] = ask_input("Second file path")

        elif ptype == "ref_file":
            kwargs[name] = ask_input("Reference file path")

        elif ptype == "output":
            ext = p.get("ext", ".xlsx")
            kwargs[name] = get_output_path(intent["fn"], ext)

        elif ptype == "output_dir":
            kwargs[name] = get_output_dir(intent["fn"])

        elif ptype == "output_csv":
            kwargs[name] = get_output_path(intent["fn"], ".csv")

        elif ptype == "output_json":
            kwargs[name] = get_output_path(intent["fn"], ".json")

        elif ptype == "col_req":
            kwargs[name] = ask_column(prompt_text, columns, required=True)

        elif ptype == "col_opt":
            val = ask_column(prompt_text + " (Enter=skip)", columns, required=False)
            kwargs[name] = val if val else None

        elif ptype == "cols_req":
            result = ask_columns(prompt_text, columns, required=True)
            kwargs[name] = result if result else None

        elif ptype == "cols_opt":
            result = ask_columns(prompt_text, columns, required=False)
            kwargs[name] = result if result else None

        elif ptype == "number":
            extracted = extract_number_from_query(query, None)
            if extracted is not None and not kwargs.get("_number_used"):
                kwargs["_number_used"] = True
                val = extracted
                print(f"  → {prompt_text}: {val} (extracted from query)")
            else:
                raw = ask_input(prompt_text, str(default))
                val = int(raw) if raw.isdigit() else default
            kwargs[name] = val

        elif ptype == "float_val":
            raw = ask_input(prompt_text, str(default))
            try:
                kwargs[name] = float(raw)
            except (ValueError, TypeError):
                kwargs[name] = float(default) if default else 0.0

        elif ptype == "choice":
            choices_str = " / ".join(f"{i+1}={o}" for i, o in enumerate(options))
            raw = ask_input(f"{prompt_text} ({choices_str})", str(default))
            if raw.isdigit():
                idx = int(raw) - 1
                kwargs[name] = options[idx] if 0 <= idx < len(options) else default
            elif raw in options:
                kwargs[name] = raw
            else:
                kwargs[name] = default

        elif ptype == "string":
            kwargs[name] = ask_input(prompt_text, str(default) if default else "")

        elif ptype == "bool_val":
            kwargs[name] = ask_yn(prompt_text, bool(default))

        elif ptype == "mapping":
            kwargs[name] = ask_mapping(prompt_text)

    # Remove internal tracking key
    kwargs.pop("_number_used", None)
    return kwargs


# =============================================================================
# EXECUTION
# =============================================================================

def execute_intent(intent: dict, kwargs: dict) -> any:
    """Dynamically import the module and call the function."""
    module_name = intent["module"]
    fn_name     = intent["fn"]

    # Import the module from the modules/ package
    sys.path.insert(0, str(Path(__file__).parent))
    mod = importlib.import_module(f"modules.{module_name}")
    fn  = getattr(mod, fn_name)

    return fn(**kwargs)


def show_result(result, kwargs: dict):
    """Print result path(s) after execution."""
    if isinstance(result, str):
        print(_c(Fore.GREEN + Style.BRIGHT, f"\n  ✓  Done!  Output: {result}"))
    elif isinstance(result, list):
        print(_c(Fore.GREEN + Style.BRIGHT, f"\n  ✓  Done!  {len(result)} output(s):"))
        for r in result:
            print(f"     {r}")
    else:
        out = kwargs.get("output_path") or kwargs.get("output_dir") or ""
        print(_c(Fore.GREEN + Style.BRIGHT, f"\n  ✓  Done!" + (f"  Output: {out}" if out else "")))


# =============================================================================
# MAIN
# =============================================================================

def home_menu():
    """
    Unified home screen: shows all 18 module options first, then
    offers the NL prompt at the bottom so the user can either:
      • Pick a module by number  → launches that module's classic menu
      • Type a plain-English query → NL routing
    Loops until the user exits with 0.
    """
    import importlib
    _main = importlib.import_module("main")   # load main.py once

    MODULES = [
        ("1",  "Consolidate Files",      "stack, join, merge multiple Excel files",       _main.menu_consolidate),
        ("2",  "Calculate & Analyze",    "efficiency, KPIs, variance, stats",             _main.menu_calculate),
        ("3",  "Clean Data",             "duplicates, blanks, formats, types",            _main.menu_clean),
        ("4",  "Transform Data",         "pivot, split, transpose, reshape",              _main.menu_transform),
        ("5",  "Compare Files",          "diff, new/deleted rows, changes",               _main.menu_compare),
        ("6",  "Column Operations",      "rename, split, merge, calculate columns",       _main.menu_columns),
        ("7",  "Generate Reports",       "summary, profile, KPI, frequency",              _main.menu_reports),
        ("8",  "Quick File Preview",     "peek at any Excel file",                        _main.menu_preview),
        ("9",  "Finance",                "aging, payroll, ratios, amortization",          _main.menu_finance),
        ("10", "HR Analytics",           "attrition, headcount, tenure, salary",          _main.menu_hr),
        ("11", "Sales Analytics",        "commission, RFM, quota, pipeline",              _main.menu_sales),
        ("12", "Inventory Management",   "ABC, reorder point, OEE, dead stock",           _main.menu_inventory),
        ("13", "Format & Style",         "charts, traffic lights, table styles",          _main.menu_formatter),
        ("14", "Validate Data",          "email, phone, range, PII detection",            _main.menu_validator),
        ("15", "Statistical Analytics",  "correlation, Pareto, regression, Z-score",      _main.menu_analytics),
        ("16", "Convert Files",          "Excel↔CSV, Excel↔JSON, XLS→XLSX",              _main.menu_converter),
        ("17", "Lookup & Match",         "VLOOKUP, fuzzy match, multi-key join",          _main.menu_lookup),
        ("18", "Project Management",     "timesheets, milestones, RACI, risks, sprints",  _main.menu_project_mgmt),
    ]

    while True:
        banner()

        # ── Built-in module options ───────────────────────────────────────────
        print(_c(Fore.WHITE + Style.BRIGHT, "  MODULES\n"))
        for key, name, desc, _ in MODULES:
            print(f"    {_c(Fore.CYAN, f'{key:>2}')}  "
                  f"{name:<24}  {_c(Fore.WHITE, desc)}")

        print(f"\n    {_c(Fore.RED, ' 0')}  Exit\n")

        # ── NL prompt ─────────────────────────────────────────────────────────
        print(_c(Fore.YELLOW + Style.BRIGHT,
                 "  ─────────────────────────────────────────────────────────"))
        print(_c(Fore.YELLOW + Style.BRIGHT,
                 "  💬 Or just tell me what you want to do (plain English):"))
        print()

        raw = input(_c(Fore.WHITE + Style.BRIGHT,
                       "  → Pick [1-18 / 0] or describe what you want: ")).strip()

        if not raw or raw == "0":
            print(_c(Fore.GREEN, "\n  Goodbye!  All outputs saved to: output/"))
            break

        # ── Module picked by number ───────────────────────────────────────────
        if raw.isdigit():
            idx = int(raw)
            if 1 <= idx <= 18:
                _main.banner()
                MODULES[idx - 1][3]()   # call the module's menu function
                continue                # back to home menu when it returns
            else:
                print(_c(Fore.RED, f"\n  Invalid choice: {raw}  (enter 1–18 or 0)"))
                input("  Press Enter to continue...")
                continue

        # ── Natural language query ────────────────────────────────────────────
        print()
        raw_path = ask_input("File path or folder (Enter to skip)").strip()
        files = resolve_files([raw_path]) if raw_path else []
        if not files:
            print(_c(Fore.YELLOW, "  No files — continuing with query only."))
        run(files, raw)

        # After NL op, offer to run another without going back to the full menu
        print()
        if not ask_yn("Run another operation?", default=False):
            break


def run(files: list, query: str):
    """Core: parse intent → collect params → execute."""

    # ── Compound query detection ──────────────────────────────────────────────
    # If the query contains a clear sequential connector ("and then", "then",
    # "followed by", etc.), split it into independent sub-queries and run each.
    compound_parts = split_compound_query(query)
    if compound_parts:
        print(_c(Fore.CYAN + Style.BRIGHT,
                 f"\n  💡 Multi-step query detected ({len(compound_parts)} operations):"))
        for i, part in enumerate(compound_parts, 1):
            print(f"     {i}. \"{part}\"")

        choice = ask_yn(
            f"\n  Run all {len(compound_parts)} steps sequentially?",
            default=True,
        )
        if choice:
            for i, part in enumerate(compound_parts, 1):
                print(_c(Fore.CYAN + Style.BRIGHT,
                         f"\n  ═══ Step {i}/{len(compound_parts)}: \"{part}\" ═══"))
                run(files, part)
            return
        else:
            nums = "/".join(str(i) for i in range(1, len(compound_parts) + 1))
            raw = ask_input(f"Which step to run? [{nums}]", "1")
            idx = (int(raw) - 1) if raw.isdigit() else 0
            idx = max(0, min(len(compound_parts) - 1, idx))
            query = compound_parts[idx]
            print(_c(Fore.CYAN, f"\n  Running step {idx+1}: \"{query}\""))

    # ── File inspection ───────────────────────────────────────────────────────
    if files:
        print(_c(Fore.CYAN, f"\n  📂 Inspecting file: {Path(files[0]).name}"))
        try:
            fi = inspect_file(files[0])
            n_sheets = fi["sheet_count"]
            first_sheet = fi["sheets"][0] if fi["sheets"] else "Sheet1"
            cols = fi["columns"].get(first_sheet, [])
            n_rows = fi["row_counts"].get(first_sheet, "?")
            print(f"     Sheets  : {n_sheets}  {fi['sheets'][:5]}")
            print(f"     Rows    : {n_rows}  |  Cols: {len(cols)}")
            if cols:
                print(f"     Columns : {', '.join(cols[:8])}" +
                      (" ..." if len(cols) > 8 else ""))
            if fi["domain_hint"]:
                print(_c(Fore.YELLOW, f"     Domain  : {fi['domain_hint']}"))
        except Exception as e:
            print(_c(Fore.YELLOW, f"  (file inspection skipped: {e})"))

    # Parse intent
    print(_c(Fore.CYAN, f"\n  🔍 Analysing: \"{query}\"\n"))
    results = parse_intent(query, files, top_n=5)

    if not results:
        print(_c(Fore.RED, "  No matching operation found. Try rephrasing."))
        print("  Tip: be specific, e.g. 'remove duplicates', 'pivot by department'")
        return

    # Display matches
    print(_c(Fore.CYAN + Style.BRIGHT, "  Matched operations:"))
    for i, r in enumerate(results, 1):
        intent = r["intent"]
        score  = r["score"]
        conf   = r["confidence"]
        icon   = "✅" if conf == "high" else ("🔸" if conf == "medium" else "🔹")
        bar    = conf_bar(score)
        print(f"    [{i}] {icon}  {intent['module']:<14} → {intent['fn']:<35}  {bar}")
        print(f"         {_c(Fore.WHITE, intent['desc'])}")

    print()

    # High-confidence auto-select
    top = results[0]
    if top["confidence"] == "high":
        confirmed = ask_yn(
            f"Run [{top['intent']['module']} → {top['intent']['fn']}]?",
            default=True
        )
        if not confirmed:
            choice = ask_input("Pick another [1-5] or Enter to exit", "")
            if not choice:
                return
            idx = int(choice) - 1 if choice.isdigit() else 0
            top = results[idx] if 0 <= idx < len(results) else results[0]
    else:
        choice = ask_input(
            f"Which operation? [1-{len(results)}]",
            "1"
        )
        idx = (int(choice) - 1) if choice.isdigit() else 0
        top = results[idx] if 0 <= idx < len(results) else results[0]

    intent = top["intent"]
    print(_c(Fore.CYAN + Style.BRIGHT,
             f"\n  Collecting parameters for: {intent['module']} → {intent['fn']}"))

    # Collect params
    kwargs = collect_params(intent, files, query)

    # Confirm
    print(_c(Fore.YELLOW + Style.BRIGHT, "\n  Will execute:"))
    for k, v in kwargs.items():
        display = str(v)
        if len(display) > 60:
            display = display[:57] + "..."
        print(f"    {k} = {display}")

    if not ask_yn("Proceed?", default=True):
        print(_c(Fore.YELLOW, "  Cancelled."))
        return

    # Execute
    print(_c(Fore.CYAN, f"\n  ⚡ Running {intent['fn']}..."))
    try:
        result = execute_intent(intent, kwargs)
        show_result(result, kwargs)
    except TypeError as e:
        print(_c(Fore.RED, f"\n  ✗  Parameter error: {e}"))
        print("  Tip: check that all column names are correct.")
    except Exception as e:
        print(_c(Fore.RED, f"\n  ✗  Error: {e}"))


def main():
    args = sys.argv[1:]

    # No args → unified home menu (modules + NL)
    if not args:
        home_menu()
        return

    # Last arg is likely the query (doesn't look like a file path)
    # Try: ask.py file.xlsx "query"  OR  ask.py folder/ "query"
    # Also handle: ask.py "query"  (no file)
    files = []
    query = ""

    if len(args) == 1:
        # Could be a query only or a file only
        a = args[0]
        p = Path(a)
        if p.exists():
            # It's a file/folder — no query
            files = resolve_files([a])
            query = ask_input("What do you want to do?").strip()
        else:
            # Treat it as a query with no file
            query = a

    else:
        # Last arg = query, everything before = file paths
        query = args[-1]
        file_args = args[:-1]
        files = resolve_files(file_args)

    if not query:
        print(_c(Fore.RED, "  No query provided. Exiting."))
        sys.exit(1)

    banner()
    run(files, query)
    print()


if __name__ == "__main__":
    main()
