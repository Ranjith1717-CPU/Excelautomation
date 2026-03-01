"""
project_mgmt.py — Project Management Tools
==========================================
10 operations covering team consolidation, timesheets, resource planning,
milestones, RACI, risk registers, action tracking, capacity planning and sprints.
"""
import pandas as pd
import numpy as np
from pathlib import Path
import datetime


# ─────────────────────────────────────────────────────────────────────────────
# 1. TEAM CONSOLIDATOR
# ─────────────────────────────────────────────────────────────────────────────

def team_consolidator(files: list, output_path: str,
                       add_source: bool = True,
                       id_col: str = None) -> str:
    """
    Merge team member data from multiple files (all sheets from each file).
    Optionally deduplicates on id_col.
    Adds _Source_File and _Source_Sheet columns when add_source=True.
    """
    frames = []
    for f in files:
        xl = pd.ExcelFile(f)
        for sheet in xl.sheet_names:
            df = pd.read_excel(f, sheet_name=sheet).copy()
            if df.empty:
                continue
            if add_source:
                df = df.assign(_Source_File=Path(f).stem, _Source_Sheet=sheet)
            frames.append(df)

    if not frames:
        raise ValueError("No data found in the provided files.")

    combined = pd.concat(frames, ignore_index=True)

    if id_col and id_col in combined.columns:
        before = len(combined)
        combined = combined.drop_duplicates(subset=[id_col], keep='first')
        dupes = before - len(combined)
        if dupes:
            print(f"  [INFO] Removed {dupes} duplicate row(s) on '{id_col}'")

    combined.to_excel(output_path, index=False)
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 2. SPLIT BY TEAM / DEPARTMENT
# ─────────────────────────────────────────────────────────────────────────────

def split_by_team(file: str, split_col: str, output_dir: str) -> list:
    """
    Split a master sheet into one file per unique value in split_col.
    Safe file names are generated from the column values.
    """
    df = pd.read_excel(file)
    if split_col not in df.columns:
        raise ValueError(f"Column '{split_col}' not found.")

    Path(output_dir).mkdir(parents=True, exist_ok=True)
    created = []

    for value, group in df.groupby(split_col, dropna=False):
        safe = str(value).replace('/', '_').replace('\\', '_').replace(' ', '_').replace(':', '_')
        out_path = str(Path(output_dir) / f"{safe}.xlsx")
        group.reset_index(drop=True).to_excel(out_path, index=False)
        created.append(out_path)

    return created


# ─────────────────────────────────────────────────────────────────────────────
# 3. TIMESHEET ROLLUP
# ─────────────────────────────────────────────────────────────────────────────

def timesheet_rollup(files: list, person_col: str, project_col: str,
                      hours_col: str, date_col: str,
                      output_path: str) -> str:
    """
    Consolidate N timesheet files into four sheets:
      - Detail      : all rows combined with source tag
      - By_Person   : total hours per person (sorted desc)
      - By_Project  : total hours per project (sorted desc)
      - Person_x_Project : pivot matrix (person rows, project columns)
    """
    frames = []
    for f in files:
        df = pd.read_excel(f)
        df['_Source'] = Path(f).stem
        frames.append(df)

    detail = pd.concat(frames, ignore_index=True)
    detail[date_col]  = pd.to_datetime(detail[date_col], errors='coerce')
    detail[hours_col] = pd.to_numeric(detail[hours_col], errors='coerce').fillna(0)

    by_person = (detail.groupby(person_col)[hours_col]
                       .sum()
                       .reset_index()
                       .rename(columns={hours_col: 'Total_Hours'})
                       .sort_values('Total_Hours', ascending=False))

    by_project = (detail.groupby(project_col)[hours_col]
                        .sum()
                        .reset_index()
                        .rename(columns={hours_col: 'Total_Hours'})
                        .sort_values('Total_Hours', ascending=False))

    pivot = detail.pivot_table(index=person_col, columns=project_col,
                                values=hours_col, aggfunc='sum', fill_value=0)
    pivot['TOTAL'] = pivot.sum(axis=1)
    pivot = pivot.reset_index()

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        detail.to_excel(writer,    sheet_name='Detail',           index=False)
        by_person.to_excel(writer, sheet_name='By_Person',        index=False)
        by_project.to_excel(writer,sheet_name='By_Project',       index=False)
        pivot.to_excel(writer,     sheet_name='Person_x_Project', index=False)

    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 4. RESOURCE ALLOCATION
# ─────────────────────────────────────────────────────────────────────────────

def resource_allocation(file: str, resource_col: str, project_col: str,
                         hours_col: str, capacity_col: str,
                         output_path: str) -> str:
    """
    Resource allocation analysis.
    Pivot: resource × project (hours), then appends:
      Total_Allocated, Capacity, Available_Hours, Allocation_%, Status.
    Status: Over-Allocated / Fully Allocated / Balanced / Under-Utilised.
    """
    df = pd.read_excel(file)
    df[hours_col]    = pd.to_numeric(df[hours_col],    errors='coerce').fillna(0)
    df[capacity_col] = pd.to_numeric(df[capacity_col], errors='coerce').fillna(0)

    pivot = df.pivot_table(index=resource_col, columns=project_col,
                            values=hours_col, aggfunc='sum', fill_value=0)
    pivot.columns = [str(c) for c in pivot.columns]
    pivot['Total_Allocated'] = pivot.sum(axis=1)

    capacity = df.groupby(resource_col)[capacity_col].max()
    pivot = pivot.join(capacity.rename('Capacity'))

    pivot['Available_Hours'] = (pivot['Capacity'] - pivot['Total_Allocated']).round(1)
    pivot['Allocation_%']    = (pivot['Total_Allocated'] /
                                pivot['Capacity'].replace(0, np.nan) * 100).round(1)
    pivot['Status'] = pivot['Allocation_%'].apply(
        lambda x: 'Over-Allocated'  if pd.notna(x) and x > 100
        else      ('Fully Allocated' if pd.notna(x) and x >= 90
        else      ('Balanced'        if pd.notna(x) and x >= 60
        else       'Under-Utilised'))
    )

    result = pivot.reset_index()
    result.to_excel(output_path, index=False)
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 5. MILESTONE TRACKER
# ─────────────────────────────────────────────────────────────────────────────

def milestone_tracker(file: str, task_col: str, owner_col: str,
                       planned_col: str, actual_col: str,
                       output_path: str) -> str:
    """
    Milestone / delivery tracker.
    Calculates Slippage_Days, Days_Overdue, RAG_Status, Complete flag.
    Sheets: Milestones (full detail) + Owner_Summary.
    RAG: Green = on time, Amber = 1-5 days late, Red = >5 days late.
    """
    df = pd.read_excel(file)
    today = pd.Timestamp.today().normalize()

    df[planned_col] = pd.to_datetime(df[planned_col], errors='coerce')
    df[actual_col]  = pd.to_datetime(df[actual_col],  errors='coerce')

    df['Slippage_Days'] = (df[actual_col] - df[planned_col]).dt.days

    df['Days_Overdue'] = df.apply(
        lambda r: max(0, (today - r[planned_col]).days)
        if pd.isna(r[actual_col]) and pd.notna(r[planned_col]) else 0,
        axis=1
    )

    def _rag(row):
        if pd.notna(row[actual_col]):
            slip = row['Slippage_Days'] if pd.notna(row['Slippage_Days']) else 0
            if slip <= 0:  return 'Green'
            if slip <= 5:  return 'Amber'
            return 'Red'
        else:
            if row['Days_Overdue'] == 0: return 'Green'
            if row['Days_Overdue'] <= 5: return 'Amber'
            return 'Red'

    df['RAG_Status'] = df.apply(_rag, axis=1)
    df['Complete']   = df[actual_col].notna().map({True: 'Yes', False: 'No'})

    summary = df.groupby(owner_col).agg(
        Total_Milestones=(task_col,       'count'),
        Completed=       (actual_col,     lambda x: x.notna().sum()),
        Overdue=         ('Days_Overdue', lambda x: (x > 0).sum()),
        Red_Items=       ('RAG_Status',   lambda x: (x == 'Red').sum()),
        Amber_Items=     ('RAG_Status',   lambda x: (x == 'Amber').sum()),
    ).reset_index()

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer,      sheet_name='Milestones',    index=False)
        summary.to_excel(writer, sheet_name='Owner_Summary', index=False)

    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 6. RACI MATRIX
# ─────────────────────────────────────────────────────────────────────────────

def raci_matrix(file: str, task_col: str, role_cols: list,
                output_path: str) -> str:
    """
    Build and validate a RACI matrix.
    Produces:
      - RACI_Matrix : original data + Has_Accountable / Has_Responsible / Issues columns
      - Role_Summary: R/A/C/I count per role
    Flags tasks with no Accountable or no Responsible.
    """
    df = pd.read_excel(file)
    raci_df = df[[task_col] + role_cols].copy()

    raci_df['Has_Accountable'] = raci_df[role_cols].apply(
        lambda row: any(str(v).strip().upper() == 'A' for v in row), axis=1
    )
    raci_df['Has_Responsible'] = raci_df[role_cols].apply(
        lambda row: any(str(v).strip().upper() == 'R' for v in row), axis=1
    )
    raci_df['Issues'] = raci_df.apply(
        lambda r: '; '.join(filter(None, [
            'No Accountable' if not r['Has_Accountable'] else '',
            'No Responsible' if not r['Has_Responsible'] else '',
        ])), axis=1
    )

    summary_rows = []
    for role in role_cols:
        vals = raci_df[role].astype(str).str.strip().str.upper()
        summary_rows.append({
            'Role':            role,
            'Responsible (R)': (vals == 'R').sum(),
            'Accountable (A)': (vals == 'A').sum(),
            'Consulted (C)':   (vals == 'C').sum(),
            'Informed (I)':    (vals == 'I').sum(),
            'Blank':           raci_df[role].isna().sum(),
        })
    summary = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        raci_df.to_excel(writer, sheet_name='RACI_Matrix',  index=False)
        summary.to_excel(writer, sheet_name='Role_Summary', index=False)

    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 7. RISK REGISTER
# ─────────────────────────────────────────────────────────────────────────────

def risk_register(file: str, desc_col: str, prob_col: str,
                   impact_col: str, owner_col: str,
                   output_path: str) -> str:
    """
    Score and rank a risk register.
    Risk Score = Probability × Impact (1-5 scale each).
    Levels: Critical (20-25), High (10-19), Medium (5-9), Low (1-4).
    Sheets: Risk_Register (sorted by score) + Owner_Summary + Heat_Map_Data.
    """
    df = pd.read_excel(file)
    df[prob_col]   = pd.to_numeric(df[prob_col],   errors='coerce').fillna(1).clip(1, 5)
    df[impact_col] = pd.to_numeric(df[impact_col], errors='coerce').fillna(1).clip(1, 5)

    df['Risk_Score'] = (df[prob_col] * df[impact_col]).astype(int)
    df['Risk_Level'] = df['Risk_Score'].apply(
        lambda s: 'Critical' if s >= 20
        else      ('High'    if s >= 10
        else      ('Medium'  if s >= 5
        else       'Low'))
    )
    df = df.sort_values('Risk_Score', ascending=False).reset_index(drop=True)

    summary = df.groupby(owner_col).agg(
        Total_Risks=('Risk_Score',  'count'),
        Critical=   ('Risk_Level',  lambda x: (x == 'Critical').sum()),
        High=       ('Risk_Level',  lambda x: (x == 'High').sum()),
        Medium=     ('Risk_Level',  lambda x: (x == 'Medium').sum()),
        Low=        ('Risk_Level',  lambda x: (x == 'Low').sum()),
        Avg_Score=  ('Risk_Score',  lambda x: round(x.mean(), 1)),
    ).reset_index()

    heat = df.groupby([prob_col, impact_col]).size().reset_index(name='Count')
    heat_pivot = (heat.pivot(index=prob_col, columns=impact_col, values='Count')
                      .fillna(0)
                      .astype(int))

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer,          sheet_name='Risk_Register',  index=False)
        summary.to_excel(writer,     sheet_name='Owner_Summary',  index=False)
        heat_pivot.to_excel(writer,  sheet_name='Heat_Map_Data')

    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 8. ACTION TRACKER
# ─────────────────────────────────────────────────────────────────────────────

def action_tracker(files: list, action_col: str, owner_col: str,
                    due_col: str, status_col: str,
                    output_path: str) -> str:
    """
    Consolidate action items from multiple meeting files.
    Flags overdue open items; adds Days_Overdue and Priority columns.
    Sheets: All_Actions + By_Owner + Overdue_Only.
    Priority: Critical (>14 days overdue), High (>7), Medium (>0), Normal.
    """
    frames = []
    for f in files:
        df = pd.read_excel(f)
        df['_Source'] = Path(f).stem
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    today = pd.Timestamp.today().normalize()
    combined[due_col] = pd.to_datetime(combined[due_col], errors='coerce')

    done_vals = {'done', 'complete', 'completed', 'closed', 'resolved'}
    combined['Is_Open']    = ~combined[status_col].astype(str).str.strip().str.lower().isin(done_vals)
    combined['Is_Overdue'] = combined['Is_Open'] & (combined[due_col] < today)
    combined['Days_Overdue'] = combined.apply(
        lambda r: max(0, (today - r[due_col]).days)
        if r['Is_Overdue'] and pd.notna(r[due_col]) else 0,
        axis=1
    )
    combined['Priority'] = combined['Days_Overdue'].apply(
        lambda d: 'Critical' if d > 14
        else      ('High'    if d > 7
        else      ('Medium'  if d > 0
        else       'Normal'))
    )
    combined = combined.sort_values(['Is_Overdue', 'Days_Overdue'],
                                     ascending=[False, False]).reset_index(drop=True)

    summary = combined.groupby(owner_col).agg(
        Total=           (action_col,    'count'),
        Open=            ('Is_Open',     'sum'),
        Overdue=         ('Is_Overdue',  'sum'),
        Max_Days_Overdue=('Days_Overdue','max'),
    ).reset_index()

    overdue = combined[combined['Is_Overdue']].copy()

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        combined.to_excel(writer, sheet_name='All_Actions',  index=False)
        summary.to_excel(writer,  sheet_name='By_Owner',     index=False)
        overdue.to_excel(writer,  sheet_name='Overdue_Only', index=False)

    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 9. CAPACITY PLANNER
# ─────────────────────────────────────────────────────────────────────────────

def capacity_planner(file: str, resource_col: str, role_col: str,
                      available_col: str, allocated_col: str,
                      output_path: str) -> str:
    """
    Capacity vs demand planning.
    Adds Net_Capacity, Utilisation_%, Status per resource.
    Team_Summary sheet: headcount, totals, over-allocation count per role/team.
    Over_Allocated sheet: filtered list for immediate action.
    """
    df = pd.read_excel(file)
    df[available_col] = pd.to_numeric(df[available_col], errors='coerce').fillna(0)
    df[allocated_col] = pd.to_numeric(df[allocated_col], errors='coerce').fillna(0)

    df['Net_Capacity']  = (df[available_col] - df[allocated_col]).round(1)
    df['Utilisation_%'] = (df[allocated_col] /
                           df[available_col].replace(0, np.nan) * 100).round(1)
    df['Status'] = df['Utilisation_%'].apply(
        lambda x: 'Over-Allocated'  if pd.notna(x) and x > 100
        else      ('Fully Allocated' if pd.notna(x) and x >= 90
        else      ('Balanced'        if pd.notna(x) and x >= 60
        else       'Under-Utilised'))
    )

    role_summary = df.groupby(role_col).agg(
        Headcount=      (resource_col,  'count'),
        Total_Available=(available_col, 'sum'),
        Total_Allocated=(allocated_col, 'sum'),
        Over_Allocated= ('Status',       lambda x: (x == 'Over-Allocated').sum()),
        Under_Utilised= ('Status',       lambda x: (x == 'Under-Utilised').sum()),
    ).reset_index()
    role_summary['Team_Utilisation_%'] = (
        role_summary['Total_Allocated'] /
        role_summary['Total_Available'].replace(0, np.nan) * 100
    ).round(1)

    overloaded = df[df['Status'] == 'Over-Allocated'].copy()

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer,          sheet_name='Resource_Detail', index=False)
        role_summary.to_excel(writer,sheet_name='Team_Summary',    index=False)
        overloaded.to_excel(writer,  sheet_name='Over_Allocated',  index=False)

    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 10. SPRINT TRACKER
# ─────────────────────────────────────────────────────────────────────────────

def sprint_tracker(file: str, story_col: str, points_col: str,
                    status_col: str, sprint_col: str,
                    output_path: str) -> str:
    """
    Sprint / iteration tracker.
    Sheets:
      - All_Stories    : raw data + Is_Done flag
      - Sprint_Summary : velocity, completion %, planned vs done per sprint
      - Backlog        : incomplete stories
      - Metrics        : avg velocity, backlog size, sprints to clear
    Done statuses: done / complete / completed / closed.
    """
    df = pd.read_excel(file)
    df[points_col] = pd.to_numeric(df[points_col], errors='coerce').fillna(0)

    done_vals = {'done', 'complete', 'completed', 'closed'}
    df['Is_Done'] = df[status_col].astype(str).str.strip().str.lower().isin(done_vals)

    total_pts   = df.groupby(sprint_col)[points_col].sum().rename('Planned_Points')
    total_count = df.groupby(sprint_col)[story_col].count().rename('Total_Stories')

    done_df       = df[df['Is_Done']]
    done_pts      = done_df.groupby(sprint_col)[points_col].sum().rename('Completed_Points')
    done_count    = done_df.groupby(sprint_col).size().rename('Stories_Done')

    sprint_summary = (pd.concat([total_count, total_pts, done_pts, done_count], axis=1)
                        .fillna(0)
                        .reset_index())
    sprint_summary['Velocity']      = sprint_summary['Completed_Points']
    sprint_summary['Completion_%']  = (
        sprint_summary['Completed_Points'] /
        sprint_summary['Planned_Points'].replace(0, np.nan) * 100
    ).round(1)

    avg_velocity   = sprint_summary['Velocity'].mean()
    backlog        = df[~df['Is_Done']].copy()
    backlog_points = backlog[points_col].sum()

    metrics = pd.DataFrame([
        {'Metric': 'Average Velocity (pts/sprint)',
         'Value':  round(float(avg_velocity), 1) if avg_velocity > 0 else 0},
        {'Metric': 'Total Backlog Items',
         'Value':  int(len(backlog))},
        {'Metric': 'Total Backlog Points',
         'Value':  float(backlog_points)},
        {'Metric': 'Sprints to Clear Backlog',
         'Value':  round(float(backlog_points) / float(avg_velocity), 1)
                   if avg_velocity > 0 else 'N/A'},
    ])

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer,             sheet_name='All_Stories',    index=False)
        sprint_summary.to_excel(writer, sheet_name='Sprint_Summary', index=False)
        backlog.to_excel(writer,        sheet_name='Backlog',        index=False)
        metrics.to_excel(writer,        sheet_name='Metrics',        index=False)

    return output_path
