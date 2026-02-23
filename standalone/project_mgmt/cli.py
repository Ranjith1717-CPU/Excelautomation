"""Standalone Project_mgmt — Excel Automation Toolkit"""
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
import project_mgmt


# ── Menu ───────────────────────────────────────────────────────────────────
def menu_project_mgmt():
    while True:
        banner()
        section("18. PROJECT MANAGEMENT")
        choice = menu_choice([
            ("1",  "Team Consolidator     — Merge team data from multiple files/sheets"),
            ("2",  "Split by Team         — One file per department / team"),
            ("3",  "Timesheet Rollup      — Consolidate N timesheets → person×project pivot"),
            ("4",  "Resource Allocation   — Allocation % per resource per project"),
            ("5",  "Milestone Tracker     — RAG status, slippage days, overdue flags"),
            ("6",  "RACI Matrix           — Build and validate responsibility matrix"),
            ("7",  "Risk Register         — Score by Prob×Impact, heat map, owner summary"),
            ("8",  "Action Tracker        — Consolidate meeting actions, flag overdue"),
            ("9",  "Capacity Planner      — Available vs allocated, over-allocation alerts"),
            ("10", "Sprint Tracker        — Velocity, completion %, backlog health"),
        ], "Select operation")

        if choice == "0":
            break

        try:
            if choice == "1":
                files = pick_files("Enter Excel files (or folder)")
                if not files: continue
                add_src  = prompt("Add source file/sheet columns? [y/n]").lower() != "n"
                id_col   = prompt("Unique ID column for deduplication (or Enter to skip)") or None
                out = get_output_path("team_consolidated")
                result = project_mgmt.team_consolidator(files, out, add_source=add_src, id_col=id_col)

            elif choice == "2":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                split_col = prompt(f"Column to split by (e.g. Team, Department) {cols}")
                out_dir = str(OUTPUT_DIR / f"team_split_{datetime.datetime.now().strftime('%H%M%S')}")
                created = project_mgmt.split_by_team(file, split_col, out_dir)
                success(f"Created {len(created)} file(s) in {out_dir}")
                pause(); continue

            elif choice == "3":
                files = pick_files("Enter timesheet files (or folder)")
                if not files: continue
                cols = pd.read_excel(files[0], nrows=0).columns.tolist()
                person_col  = prompt(f"Person / Name column {cols}")
                project_col = prompt(f"Project column {cols}")
                hours_col   = prompt(f"Hours column {cols}")
                date_col    = prompt(f"Date column {cols}")
                out = get_output_path("timesheet_rollup")
                result = project_mgmt.timesheet_rollup(files, person_col, project_col,
                                                        hours_col, date_col, out)

            elif choice == "4":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                resource_col = prompt(f"Resource / Person column {cols}")
                project_col  = prompt(f"Project column {cols}")
                hours_col    = prompt(f"Allocated hours column {cols}")
                capacity_col = prompt(f"Capacity / Available hours column {cols}")
                out = get_output_path("resource_allocation")
                result = project_mgmt.resource_allocation(file, resource_col, project_col,
                                                           hours_col, capacity_col, out)

            elif choice == "5":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                task_col    = prompt(f"Task / Milestone column {cols}")
                owner_col   = prompt(f"Owner column {cols}")
                planned_col = prompt(f"Planned date column {cols}")
                actual_col  = prompt(f"Actual completion date column (blank = not done) {cols}")
                out = get_output_path("milestone_tracker")
                result = project_mgmt.milestone_tracker(file, task_col, owner_col,
                                                         planned_col, actual_col, out)

            elif choice == "6":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                task_col  = prompt(f"Task column {cols}")
                role_cols = pick_columns(cols, "Select role columns (values should be R / A / C / I)")
                out = get_output_path("raci_matrix")
                result = project_mgmt.raci_matrix(file, task_col, role_cols, out)

            elif choice == "7":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                desc_col   = prompt(f"Risk description column {cols}")
                prob_col   = prompt(f"Probability column (1-5) {cols}")
                impact_col = prompt(f"Impact column (1-5) {cols}")
                owner_col  = prompt(f"Risk owner column {cols}")
                out = get_output_path("risk_register")
                result = project_mgmt.risk_register(file, desc_col, prob_col,
                                                     impact_col, owner_col, out)

            elif choice == "8":
                files = pick_files("Enter meeting action files (or folder)")
                if not files: continue
                cols = pd.read_excel(files[0], nrows=0).columns.tolist()
                action_col = prompt(f"Action description column {cols}")
                owner_col  = prompt(f"Owner column {cols}")
                due_col    = prompt(f"Due date column {cols}")
                status_col = prompt(f"Status column {cols}")
                out = get_output_path("action_tracker")
                result = project_mgmt.action_tracker(files, action_col, owner_col,
                                                      due_col, status_col, out)

            elif choice == "9":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                resource_col  = prompt(f"Resource / Name column {cols}")
                role_col      = prompt(f"Role / Team column {cols}")
                available_col = prompt(f"Available hours column {cols}")
                allocated_col = prompt(f"Allocated / Demand hours column {cols}")
                out = get_output_path("capacity_planner")
                result = project_mgmt.capacity_planner(file, resource_col, role_col,
                                                        available_col, allocated_col, out)

            elif choice == "10":
                file = pick_single_file("Enter Excel file path")
                if not file: continue
                cols = pd.read_excel(file, nrows=0).columns.tolist()
                story_col  = prompt(f"Story / Task title column {cols}")
                points_col = prompt(f"Story points column {cols}")
                status_col = prompt(f"Status column {cols}")
                sprint_col = prompt(f"Sprint name / number column {cols}")
                out = get_output_path("sprint_tracker")
                result = project_mgmt.sprint_tracker(file, story_col, points_col,
                                                      status_col, sprint_col, out)
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
    menu_project_mgmt()
