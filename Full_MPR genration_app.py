import io
import sys
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

HAS_PLOTLY = True

# Styling constants for OpenPyXL excel generation
FONT = 'Arial'
HEADER_FILL = PatternFill('solid', fgColor='991B1B')
HEADER_FONT = Font(name=FONT, size=10, bold=True, color='FFFFFF')
TITLE_FONT = Font(name=FONT, size=14, bold=True, color='991B1B')
SECTION_FONT = Font(name=FONT, size=11, bold=True, color='991B1B')
NOTE_FONT = Font(name=FONT, size=9, italic=True, color='64748B')
CELL_FONT = Font(name=FONT, size=10)
BOLD_CELL = Font(name=FONT, size=10, bold=True)
BORDER = Border(*(Side(style='thin', color='E2E8F0'),) * 4)
DATE_FMT = 'dd-mmm-yy'
PCT_FMT = '0.0%'
ISSUE_RANGE_END = 5000
PM_RANGE_END = 5000

PM_QUARTER_BLOCKS = {
    'FY2627-Q1': {'start': 60, 'qcol': 72},
    'FY2627-Q2': {'start': 74, 'qcol': 86},
}

PM_STATION_COLS = list(range(0, 13))


def get_val(row_tuple, idx, default=None):
    """Safely fetch item from tuple/list without IndexError."""
    if row_tuple is not None and 0 <= idx < len(row_tuple):
        val = row_tuple[idx]
        return val if val is not None else default
    return default


def select_sheet_name(sheetnames, preferred_names, fallback_keyword):
    """
    Selects the target sheet name from a workbook's sheetnames list.
    Prioritizes exact matching (case-insensitive & whitespace trimmed) against preferred_names
    (e.g., 'PM Tracker B2C- B2B', 'Issue Tracker'), then partial matching, then keyword match.
    Only takes the specified tracker sheet from the workbook.
    """
    if not sheetnames:
        return None
    
    clean_map = {str(s).strip().lower(): s for s in sheetnames if s is not None}
    
    # 1. Exact match (case-insensitive, trimmed)
    for pref in preferred_names:
        pref_clean = pref.strip().lower()
        if pref_clean in clean_map:
            return clean_map[pref_clean]
            
    # 2. Substring match for preferred names
    for pref in preferred_names:
        pref_clean = pref.strip().lower()
        for s in sheetnames:
            if s and pref_clean in str(s).strip().lower():
                return s
                
    # 3. Fallback keyword search
    for s in sheetnames:
        if s and fallback_keyword.lower() in str(s).strip().lower():
            return s
            
    # 4. Fallback to first sheet
    return sheetnames[0]


def make_unique_headers(headers):
    """Ensure all header column names are unique string values to prevent PyArrow rendering issues."""
    seen = {}
    unique_headers = []
    for idx, h in enumerate(headers):
        name = str(h).strip() if h is not None else f"Unnamed_{idx+1}"
        if not name:
            name = f"Unnamed_{idx+1}"
        if name in seen:
            seen[name] += 1
            unique_headers.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 0
            unique_headers.append(name)
    return unique_headers


def ensure_unique_columns(df):
    """Ensure DataFrame columns are strictly unique string values."""
    if df is None or df.empty:
        return df
    seen = {}
    new_cols = []
    for idx, c in enumerate(df.columns):
        c_str = str(c).strip() if (c is not None and str(c).strip()) else f"Unnamed_{idx+1}"
        if c_str in seen:
            seen[c_str] += 1
            new_cols.append(f"{c_str}_{seen[c_str]}")
        else:
            seen[c_str] = 0
            new_cols.append(c_str)
    df.columns = new_cols
    return df


def find_col(df, possible_names):
    """Robust case-insensitive and whitespace-flexible column lookup."""
    if df is None or df.empty:
        return None
    clean_cols = {str(c).strip().lower(): c for c in df.columns}
    for p in possible_names:
        p_clean = p.strip().lower()
        if p_clean in clean_cols:
            return clean_cols[p_clean]
    for p in possible_names:
        p_clean = p.strip().lower()
        for c_lower, orig in clean_cols.items():
            if p_clean in c_lower:
                return orig
    return None


ISSUE_HEADER_KEYWORDS = ['sr no', 'issue id', 'ocpp id', 'issue date', 'issue type', 'severity']


def find_header_row(rows, keywords, max_scan=10, min_matches=2):
    """
    Locates the real header row when a sheet leads with summary/KPI rows
    (e.g. 'Total Issues', 654) before the actual column headers.
    Returns the 0-based index of the best-matching row within the first
    max_scan rows, falling back to row 0 if nothing scores well enough.
    """
    best_idx, best_score = 0, -1
    for idx, row in enumerate(rows[:max_scan]):
        cells = [str(c).strip().lower() for c in row if c is not None]
        score = sum(1 for kw in keywords if any(kw in c for c in cells))
        if score > best_score:
            best_idx, best_score = idx, score
    return best_idx if best_score >= min_matches else 0


def rows_to_df(rows, header_idx=0):
    """Builds a DataFrame from raw sheet rows given the header row's index."""
    if not rows or header_idx >= len(rows):
        return pd.DataFrame()
    headers = make_unique_headers([str(h).strip() if h is not None else f"Unnamed_{i+1}" for i, h in enumerate(rows[header_idx])])
    return pd.DataFrame([r[:len(headers)] for r in rows[header_idx + 1:]], columns=headers).dropna(how='all')


def load_issue_tracker(source):
    preferred_sheets = ['Issue Tracker', 'Issue Data']
    if isinstance(source, pd.DataFrame):
        df = source.copy()
    elif isinstance(source, bytes):
        if source.startswith(b'PK\x03\x04'):
            source_stream = io.BytesIO(source)
            wb = openpyxl.load_workbook(source_stream, data_only=True, read_only=True)
            sheet_name = select_sheet_name(wb.sheetnames, preferred_sheets, 'issue')
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
            if not rows:
                return pd.DataFrame()
            header_idx = find_header_row(rows, ISSUE_HEADER_KEYWORDS)
            df = rows_to_df(rows, header_idx)
        else:
            df = pd.read_csv(io.BytesIO(source))
    elif hasattr(source, 'name') and source.name.lower().endswith('.csv'):
        df = pd.read_csv(source)
    else:
        if not hasattr(source, 'sheetnames'):
            if isinstance(source, (str, io.BytesIO)):
                wb = openpyxl.load_workbook(source, data_only=True, read_only=True)
            else:
                wb = openpyxl.load_workbook(io.BytesIO(source.read()), data_only=True, read_only=True)
        else:
            wb = source
        sheet_name = select_sheet_name(wb.sheetnames, preferred_sheets, 'issue')
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if hasattr(wb, 'close'):
            wb.close()
        if not rows:
            return pd.DataFrame()
        header_idx = find_header_row(rows, ISSUE_HEADER_KEYWORDS)
        df = rows_to_df(rows, header_idx)

    df = ensure_unique_columns(df)
    if 'Issue Date' in df.columns:
        df['Issue Date Parsed'] = pd.to_datetime(df['Issue Date'], errors='coerce')

    if 'TAT Compliance' in df.columns:
        tat_upper = df['TAT Compliance'].astype(str).str.upper()
        df['Is_TAT_Compliant'] = (tat_upper == 'YES')
        df['Is_TAT_Breached'] = (tat_upper == 'NO')
    else:
        df['Is_TAT_Compliant'] = False
        df['Is_TAT_Breached'] = False

    return df


def load_pm_tracker(source):
    preferred_sheets = ['PM Tracker B2C- B2B', 'PM Tracker B2C-B2B', 'PM Tracker B2C - B2B', 'PM Tracker', 'PM Data']
    if isinstance(source, pd.DataFrame):
        df = source.copy()
    elif isinstance(source, bytes):
        if source.startswith(b'PK\x03\x04'):
            source_stream = io.BytesIO(source)
            wb = openpyxl.load_workbook(source_stream, data_only=True, read_only=True)
        else:
            df = pd.read_csv(io.BytesIO(source))
            df = ensure_unique_columns(df)
            if 'Due Date' in df.columns:
                df['Due Date Parsed'] = pd.to_datetime(df['Due Date'], errors='coerce')
            if 'Actual Completion Date' in df.columns:
                df['Actual Completion Date Parsed'] = pd.to_datetime(df['Actual Completion Date'], errors='coerce')
            if 'PM Status' in df.columns:
                pm_upper = df['PM Status'].astype(str).str.upper()
                df['Is_PM_Done'] = (pm_upper == 'YES')
                df['Is_PM_Pending'] = (pm_upper == 'NO')
            else:
                df['Is_PM_Done'] = False
                df['Is_PM_Pending'] = False
            return df
        sheet_name = select_sheet_name(wb.sheetnames, preferred_sheets, 'pm')
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            return pd.DataFrame()

        first_row_str = [str(h).strip().lower() for h in rows[0] if h is not None]
        is_flat_table = any(k in first_row_str for k in ['zme', 'station id', 'charger id', 'pm status', 'due date']) or len(rows) <= 5

        if is_flat_table:
            headers = make_unique_headers([str(h).strip() if h is not None else f"Unnamed_{i+1}" for i, h in enumerate(rows[0])])
            df = pd.DataFrame([r[:len(headers)] for r in rows[1:]], columns=headers).dropna(how='all')
        else:
            row_date = rows[3] if len(rows) > 3 else ()
            row_headers = rows[4] if len(rows) > 4 else ()
            station_fields = [get_val(row_headers, i, f"Col_{i}") for i in PM_STATION_COLS]

            records = []
            for r in rows[5:]:
                if get_val(r, 0) is None and get_val(r, 10) is None:
                    continue
                station = {name: get_val(r, i) for name, i in zip(station_fields, PM_STATION_COLS)}
                for quarter, blk in PM_QUARTER_BLOCKS.items():
                    compliance = get_val(r, blk['qcol'])
                    col = blk['start']
                    for _ in range(3):
                        records.append({
                            **station,
                            'Quarter': quarter,
                            'Due Date': get_val(row_date, col),
                            'PM Status': get_val(r, col),
                            'F.E. Inspection': get_val(r, col + 1),
                            'HSE Inspection': get_val(r, col + 2),
                            'Actual Completion Date': get_val(r, col + 3),
                            'Quarterly Compliance': compliance,
                        })
                        col += 4
            df = pd.DataFrame(records).rename(columns={'Route ': 'Route'})
    elif hasattr(source, 'name') and source.name.lower().endswith('.csv'):
        df = pd.read_csv(source)
    else:
        if not hasattr(source, 'sheetnames'):
            if isinstance(source, (str, io.BytesIO)):
                wb = openpyxl.load_workbook(source, data_only=True, read_only=True)
            else:
                wb = openpyxl.load_workbook(io.BytesIO(source.read()), data_only=True, read_only=True)
        else:
            wb = source
        sheet_name = select_sheet_name(wb.sheetnames, preferred_sheets, 'pm')
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if hasattr(wb, 'close'):
            wb.close()
        if not rows:
            return pd.DataFrame()

        first_row_str = [str(h).strip().lower() for h in rows[0] if h is not None]
        is_flat_table = any(k in first_row_str for k in ['zme', 'station id', 'charger id', 'pm status', 'due date']) or len(rows) <= 5

        if is_flat_table:
            headers = make_unique_headers([str(h).strip() if h is not None else f"Unnamed_{i+1}" for i, h in enumerate(rows[0])])
            df = pd.DataFrame([r[:len(headers)] for r in rows[1:]], columns=headers).dropna(how='all')
        else:
            row_date = rows[3] if len(rows) > 3 else ()
            row_headers = rows[4] if len(rows) > 4 else ()
            station_fields = [get_val(row_headers, i, f"Col_{i}") for i in PM_STATION_COLS]

            records = []
            for r in rows[5:]:
                if get_val(r, 0) is None and get_val(r, 10) is None:
                    continue
                station = {name: get_val(r, i) for name, i in zip(station_fields, PM_STATION_COLS)}
                for quarter, blk in PM_QUARTER_BLOCKS.items():
                    compliance = get_val(r, blk['qcol'])
                    col = blk['start']
                    for _ in range(3):
                        records.append({
                            **station,
                            'Quarter': quarter,
                            'Due Date': get_val(row_date, col),
                            'PM Status': get_val(r, col),
                            'F.E. Inspection': get_val(r, col + 1),
                            'HSE Inspection': get_val(r, col + 2),
                            'Actual Completion Date': get_val(r, col + 3),
                            'Quarterly Compliance': compliance,
                        })
                        col += 4
            df = pd.DataFrame(records).rename(columns={'Route ': 'Route'})

    df = ensure_unique_columns(df)

    # 1. Parse Due Date
    due_col_found = find_col(df, ['Due Date', 'PM Due Date', 'Scheduled Date', 'Schedule Date', 'PM Date', 'Date', 'Target Date', 'Completion Date', 'Due_Date'])
    if due_col_found and due_col_found in df.columns:
        df['Due Date Parsed'] = pd.to_datetime(df[due_col_found], errors='coerce')
    elif 'Due Date Parsed' not in df.columns:
        df['Due Date Parsed'] = pd.NaT

    # 2. Parse Actual Completion Date
    act_col_found = find_col(df, ['Actual Completion Date', 'Completion Date', 'PM Completion Date', 'Actual Date'])
    if act_col_found and act_col_found in df.columns:
        df['Actual Completion Date Parsed'] = pd.to_datetime(df[act_col_found], errors='coerce')
    elif 'Actual Completion Date Parsed' not in df.columns:
        df['Actual Completion Date Parsed'] = pd.NaT

    # 3. Parse Go Live Date
    go_live_col = find_col(df, ['Go Live Date', 'Go-Live Date', 'Live Date', 'Commissioning Date', 'Go Live'])
    if go_live_col and go_live_col in df.columns:
        df['Go Live Date Parsed'] = pd.to_datetime(df[go_live_col], errors='coerce')
    elif 'Go Live Date Parsed' not in df.columns:
        df['Go Live Date Parsed'] = pd.NaT

    # 4. Derived Period Columns for Date, Month, Quarter, Year grouping
    # --- Month ---
    raw_m_col = find_col(df, ['Scheduled Month', 'PM Month', 'Month', 'PM_Month', 'Schedule Month'])
    if raw_m_col and raw_m_col in df.columns and not df[raw_m_col].dropna().empty:
        df['Scheduled Month'] = df[raw_m_col].astype(str).str.strip()
    elif 'Due Date Parsed' in df.columns and df['Due Date Parsed'].notna().any():
        df['Scheduled Month'] = df['Due Date Parsed'].dt.strftime('%b-%Y').fillna('Unscheduled / General')
    elif 'Go Live Date Parsed' in df.columns and df['Go Live Date Parsed'].notna().any():
        df['Scheduled Month'] = df['Go Live Date Parsed'].dt.strftime('%b-%Y').fillna('Unscheduled / General')
    else:
        df['Scheduled Month'] = 'Unscheduled / General'

    # --- Quarter ---
    raw_q_col = find_col(df, ['Scheduled Quarter', 'PM Quarter', 'Quarter', 'PM_Quarter', 'Qtr'])
    if raw_q_col and raw_q_col in df.columns and not df[raw_q_col].dropna().empty:
        q_vals = df[raw_q_col].astype(str).str.strip()
        df['Scheduled Quarter'] = q_vals.apply(lambda x: f"Q{x}" if x.isdigit() and len(x) == 1 else x)
    elif 'Due Date Parsed' in df.columns and df['Due Date Parsed'].notna().any():
        df['Scheduled Quarter'] = ('Q' + df['Due Date Parsed'].dt.quarter.astype(str)).fillna('Unscheduled / General')
    elif 'Go Live Date Parsed' in df.columns and df['Go Live Date Parsed'].notna().any():
        df['Scheduled Quarter'] = ('Q' + df['Go Live Date Parsed'].dt.quarter.astype(str)).fillna('Unscheduled / General')
    else:
        df['Scheduled Quarter'] = 'Unscheduled / General'

    # --- Year ---
    raw_y_col = find_col(df, ['Scheduled Year', 'PM Year', 'Year', 'PM_Year', 'FY', 'Financial Year'])
    if raw_y_col and raw_y_col in df.columns and not df[raw_y_col].dropna().empty:
        df['Scheduled Year'] = df[raw_y_col].astype(str).str.strip()
    elif 'Due Date Parsed' in df.columns and df['Due Date Parsed'].notna().any():
        df['Scheduled Year'] = df['Due Date Parsed'].dt.year.astype(str).replace('<NA>', 'Unscheduled / General').fillna('Unscheduled / General')
    elif 'Go Live Date Parsed' in df.columns and df['Go Live Date Parsed'].notna().any():
        df['Scheduled Year'] = df['Go Live Date Parsed'].dt.year.astype(str).replace('<NA>', 'Unscheduled / General').fillna('Unscheduled / General')
    else:
        df['Scheduled Year'] = 'Unscheduled / General'

    # --- Date ---
    raw_d_col = find_col(df, ['Scheduled Date', 'PM Date', 'Date', 'Schedule Date'])
    if raw_d_col and raw_d_col in df.columns and not df[raw_d_col].dropna().empty:
        df['Scheduled Date'] = df[raw_d_col].astype(str).str.strip()
    elif 'Due Date Parsed' in df.columns and df['Due Date Parsed'].notna().any():
        df['Scheduled Date'] = df['Due Date Parsed'].dt.strftime('%Y-%m-%d').fillna('Unscheduled / General')
    elif 'Go Live Date Parsed' in df.columns and df['Go Live Date Parsed'].notna().any():
        df['Scheduled Date'] = df['Go Live Date Parsed'].dt.strftime('%Y-%m-%d').fillna('Unscheduled / General')
    else:
        df['Scheduled Date'] = 'Unscheduled / General'

    # 5. Live Station Filtering (Strictly catch LIVE stations only)
    st_status_col = find_col(df, ['Station Status', 'Status', 'Site Status', 'Live Status', 'Station_Status'])
    if st_status_col and st_status_col in df.columns:
        live_mask = df[st_status_col].astype(str).str.strip().str.upper().str.contains('LIVE|ACTIVE|COMMISSIONED|OPERATIONAL')
        if live_mask.any():
            df = df[live_mask].copy()

    if 'PM Status' in df.columns:
        pm_upper = df['PM Status'].astype(str).str.upper()
        df['Is_PM_Done'] = (pm_upper == 'YES')
        df['Is_PM_Pending'] = (pm_upper == 'NO')
    else:
        df['Is_PM_Done'] = False
        df['Is_PM_Pending'] = False

    # 6. Advance PM Done Formula (Quarterly Week Number vs Actual Completion Date)
    week_col = find_col(df, ['Scheduled Week', 'Week Number', 'Week', 'PM Week', 'Quarterly Week'])

    def check_advance_pm(row):
        due_dt = row.get('Due Date Parsed') if 'Due Date Parsed' in row else None
        act_dt = row.get('Actual Completion Date Parsed') if 'Actual Completion Date Parsed' in row else None
        wk_val = row.get(week_col) if week_col and week_col in row else None

        if pd.notna(act_dt):
            if wk_val is not None and str(wk_val).isdigit():
                if act_dt.isocalendar().week < int(wk_val):
                    return "Yes"
            if pd.notna(due_dt) and act_dt < due_dt:
                return "Yes"
        return "No"

    df['Advance PM Done'] = df.apply(check_advance_pm, axis=1)

    # Compute OCPP / Charger Compliance Status
    def get_pm_compliance(row):
        due_dt = row.get('Due Date Parsed') if 'Due Date Parsed' in row else None
        act_dt = row.get('Actual Completion Date Parsed') if 'Actual Completion Date Parsed' in row else None
        pm_st = str(row.get('PM Status', '')).strip().upper()
        is_done = (pm_st == 'YES') or (pd.notna(act_dt))
        
        if is_done:
            if pd.notna(act_dt) and pd.notna(due_dt):
                if act_dt < due_dt:
                    return '🟡 Advance PM Done (Before Schedule)'
                elif act_dt.month == due_dt.month and act_dt.year == due_dt.year:
                    return '🟢 On-Time (As Scheduled)'
                elif act_dt > due_dt:
                    return '🟠 Completed Delayed'
                else:
                    return '🟢 On-Time (As Scheduled)'
            else:
                return '🟢 Completed'
        else:
            if pd.notna(due_dt) and datetime.now() > due_dt:
                return '🔴 Overdue / Breached Schedule'
            else:
                return '⚪ Pending (In Schedule)'

    df['PM Compliance Status'] = df.apply(get_pm_compliance, axis=1)
    return df


def write_data_sheet(wb, name, df, table_name, date_cols):
    ws = wb.create_sheet(name)
    exclude_cols = {'Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed', 'Is_TAT_Compliant', 'Is_TAT_Breached', 'Is_PM_Done', 'Is_PM_Pending'}
    clean_df = df.drop(columns=[c for c in exclude_cols if c in df.columns])

    # Header
    ws.append(list(clean_df.columns))
    for c in range(1, len(clean_df.columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    date_col_indices = {i + 1 for i, col in enumerate(clean_df.columns) if col in date_cols}

    # Fast row appending without per-cell object allocations
    for rec in clean_df.itertuples(index=False):
        row_vals = []
        for val in rec:
            if pd.isna(val):
                row_vals.append(None)
            elif isinstance(val, pd.Timestamp):
                row_vals.append(val.to_pydatetime())
            else:
                row_vals.append(val)
        ws.append(row_vals)

    # Format date cells only
    if date_col_indices:
        for r in range(2, len(clean_df) + 2):
            for c in date_col_indices:
                cell = ws.cell(row=r, column=c)
                if cell.value is not None:
                    cell.number_format = DATE_FMT

    for c, col in enumerate(clean_df.columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(12, min(28, len(str(col)) + 2))

    last_row = len(clean_df) + 1
    table = Table(displayName=table_name, ref=f"A1:{get_column_letter(len(clean_df.columns))}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium3', showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = 'A2'
    return ws, last_row


def add_pm_helper_columns(ws, pm_df, pcol, last_row):
    exclude_cols = {'Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed', 'Is_TAT_Compliant', 'Is_TAT_Breached', 'Is_PM_Done', 'Is_PM_Pending'}
    clean_cols = [c for c in pm_df.columns if c not in exclude_cols]
    adv_idx = len(clean_cols) + 1
    adv_letter = get_column_letter(adv_idx)

    h1 = ws.cell(row=1, column=adv_idx, value='Advance PM Done')
    h1.font, h1.fill = HEADER_FONT, HEADER_FILL

    due_col_name = find_col(pm_df, ['Due Date', 'Due Date Parsed', 'Scheduled Date'])
    done_col_name = find_col(pm_df, ['Actual Completion Date', 'Actual Completion Date Parsed', 'Completion Date'])
    zme_col_name = find_col(pm_df, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'])
    station_col_name = find_col(pm_df, ['Station ID', 'Station_ID', 'Station'])

    for idx, r_data in enumerate(pm_df.itertuples(), start=2):
        adv_val = "No"
        if due_col_name and done_col_name and hasattr(r_data, due_col_name) and hasattr(r_data, done_col_name):
            due_val = getattr(r_data, due_col_name)
            done_val = getattr(r_data, done_col_name)
            if pd.notna(done_val) and pd.notna(due_val):
                try:
                    if pd.to_datetime(done_val) < pd.to_datetime(due_val):
                        adv_val = "Yes"
                except Exception:
                    pass
        c_adv = ws.cell(row=idx, column=adv_idx, value=adv_val)
        c_adv.font, c_adv.border = CELL_FONT, BORDER

    ws.column_dimensions[adv_letter].width = 16

    occ_idx = adv_idx + 1
    occ_letter = get_column_letter(occ_idx)
    h2 = ws.cell(row=1, column=occ_idx, value='First Station Occurrence')
    h2.font, h2.fill = HEADER_FONT, HEADER_FILL

    seen_pairs = set()
    for idx, r_data in enumerate(pm_df.itertuples(), start=2):
        occ_val = 1
        z_v = getattr(r_data, zme_col_name) if zme_col_name and hasattr(r_data, zme_col_name) else 'N/A'
        s_v = getattr(r_data, station_col_name) if station_col_name and hasattr(r_data, station_col_name) else 'N/A'
        pair = (str(z_v).strip(), str(s_v).strip())
        if pair in seen_pairs:
            occ_val = 0
        else:
            seen_pairs.add(pair)
        c_occ = ws.cell(row=idx, column=occ_idx, value=occ_val)
        c_occ.font, c_occ.border = CELL_FONT, BORDER

    ws.column_dimensions[occ_letter].width = 20
    ws.tables['PMTable'].ref = f"A1:{occ_letter}{last_row}"
    return adv_letter, occ_letter


def section_title(ws, row, text, span):
    ws.cell(row=row, column=1, value=text).font = SECTION_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    return row + 1


def header_row(ws, row, headers):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font, c.fill = HEADER_FONT, HEADER_FILL
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    return row + 1


def data_row(ws, row, values, pct_cols=()):
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font, c.border = CELL_FONT, BORDER
        if i - 1 in pct_cols:
            c.number_format = PCT_FMT
    return row + 1


def build_issue_dashboard(wb, issue_df, irange):
    ws = wb.create_sheet('Dashboard - Issues', 0)
    ws.sheet_view.showGridLines = False
    ws['A1'] = 'CHARGEZONE ISSUE TRACKER MPR DASHBOARD'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:E1')
    ws['A2'] = 'Source: Issue Tracker Data. Filter-wise summary dashboard.'
    ws['A2'].font = NOTE_FONT
    ws.merge_cells('A2:H2')

    status_col = find_col(issue_df, ['Status', 'Ticket Status', 'Issue Status', 'State'])
    tat_col = find_col(issue_df, ['TAT Compliance', 'SLA Compliance', 'Compliance', 'TAT Status'])

    if status_col and status_col in issue_df.columns:
        s_status = issue_df[status_col].astype(str).str.strip().str.upper()
        is_closed = s_status.isin(['CLOSED', 'RESOLVED'])
        issue_df['_Is_Open_'] = ~is_closed
        issue_df['_Is_Closed_'] = is_closed
    else:
        issue_df['_Is_Closed_'] = True
        issue_df['_Is_Open_'] = False

    if tat_col and tat_col in issue_df.columns:
        s_tat = issue_df[tat_col].astype(str).str.strip().str.upper()
        issue_df['_Is_Within_'] = (s_tat == 'YES')
        issue_df['_Is_Without_'] = (s_tat == 'NO')
        issue_df['_Is_Closed_Within_'] = issue_df['_Is_Closed_'] & (s_tat == 'YES')
        issue_df['_Is_Closed_Without_'] = issue_df['_Is_Closed_'] & (s_tat == 'NO')
    else:
        issue_df['_Is_Within_'] = issue_df.get('Is_TAT_Compliant', False)
        issue_df['_Is_Without_'] = issue_df.get('Is_TAT_Breached', False)
        issue_df['_Is_Closed_Within_'] = issue_df['_Is_Closed_'] & issue_df['_Is_Within_']
        issue_df['_Is_Closed_Without_'] = issue_df['_Is_Closed_'] & issue_df['_Is_Without_']

    row = 4
    zme_header_row = row + 1
    row = section_title(ws, row, '1. Issue Summary, Overall CM Efficiency & CM-TAT Efficiency by ZME', 8)
    row = header_row(ws, row, ['ZME Name', 'Faults Registered', 'Open Faults', 'Closed Faults', 'Closed Within TAT', 'Closed Without TAT', 'Overall CM Efficiency %', 'CM-TAT Efficiency %'])

    zme_col = find_col(issue_df, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'])
    zme_start = row
    if zme_col and zme_col in issue_df.columns:
        for zme in sorted(issue_df[zme_col].dropna().unique(), key=str):
            zme_sub = issue_df[issue_df[zme_col].astype(str).str.strip() == str(zme)]
            total_val = len(zme_sub)
            open_val = int(zme_sub['_Is_Open_'].sum())
            closed_val = int(zme_sub['_Is_Closed_'].sum())
            within_val = int(zme_sub['_Is_Closed_Within_'].sum())
            without_val = int(zme_sub['_Is_Closed_Without_'].sum())
            overall_cm_eff_val = (closed_val / total_val) if total_val > 0 else 0.0
            cm_tat_eff_val = (within_val / closed_val) if closed_val > 0 else 0.0

            row = data_row(ws, row, [str(zme), total_val, open_val, closed_val, within_val, without_val, overall_cm_eff_val, cm_tat_eff_val], pct_cols={6, 7})

    zme_end = row - 1
    if zme_end >= zme_start:
        chart1 = BarChart()
        chart1.type = "col"
        chart1.style = 10
        chart1.title = "Faults Registered vs Open vs Within TAT vs Without TAT by ZME"
        chart1.y_axis.title = "Fault Count"
        chart1.x_axis.title = "ZME"

        data1 = Reference(ws, min_col=2, min_row=zme_header_row, max_col=6, max_row=zme_end)
        cats1 = Reference(ws, min_col=1, min_row=zme_start, max_row=zme_end)
        chart1.add_data(data1, titles_from_data=True)
        chart1.set_categories(cats1)
        chart1.width = 16
        chart1.height = 9
        ws.add_chart(chart1, "J4")

    row += 1
    zone_header_row = row + 1
    row = section_title(ws, row, '2. Issue Summary, Overall CM Efficiency & CM-TAT Efficiency by Zone', 8)
    row = header_row(ws, row, ['Zone', 'Faults Registered', 'Open Faults', 'Closed Faults', 'Closed Within TAT', 'Closed Without TAT', 'Overall CM Efficiency %', 'CM-TAT Efficiency %'])

    zone_start = row
    zone_col = find_col(issue_df, ['Zone', 'Zone Name', 'Region'])
    if zone_col and zone_col in issue_df.columns:
        for zone in sorted(issue_df[zone_col].dropna().unique(), key=str):
            zone_sub = issue_df[issue_df[zone_col].astype(str).str.strip() == str(zone)]
            total_val = len(zone_sub)
            open_val = int(zone_sub['_Is_Open_'].sum())
            closed_val = int(zone_sub['_Is_Closed_'].sum())
            within_val = int(zone_sub['_Is_Closed_Within_'].sum())
            without_val = int(zone_sub['_Is_Closed_Without_'].sum())
            overall_cm_eff_val = (closed_val / total_val) if total_val > 0 else 0.0
            cm_tat_eff_val = (within_val / closed_val) if closed_val > 0 else 0.0

            row = data_row(ws, row, [str(zone), total_val, open_val, closed_val, within_val, without_val, overall_cm_eff_val, cm_tat_eff_val], pct_cols={6, 7})

    zone_end = row - 1
    if zone_end >= zone_start:
        chart2 = BarChart()
        chart2.type = "col"
        chart2.style = 11
        chart2.title = "Fault Breakdown by Zone"
        chart2.y_axis.title = "Fault Count"
        chart2.x_axis.title = "Zone"

        data2 = Reference(ws, min_col=2, min_row=zone_header_row, max_col=6, max_row=zone_end)
        cats2 = Reference(ws, min_col=1, min_row=zone_start, max_row=zone_end)
        chart2.add_data(data2, titles_from_data=True)
        chart2.set_categories(cats2)
        chart2.width = 16
        chart2.height = 9
        ws.add_chart(chart2, f"J{zone_header_row - 1}")

    row += 1
    row = section_title(ws, row, '3. Repetitive Faults (same Station ID + Issue Sub-Type, 2+ occurrences)', 5)
    stn_col = find_col(issue_df, ['Station ID', 'Station_ID', 'Station', 'Site ID'])
    stn_name_col = find_col(issue_df, ['Station Name', 'Station_Name', 'Site Name'])
    zme_name_col = find_col(issue_df, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'])
    sub_col = find_col(issue_df, ['Issue Sub-Type', 'Issue Sub Type', 'Sub Type', 'Fault Subtype'])

    headers = ['Station ID']
    if stn_name_col:
        headers.append('Station Name')
    if zme_name_col:
        headers.append('ZME Name')
    headers.extend(['Issue Sub-Type', 'Occurrences'])

    row = header_row(ws, row, headers)
    if stn_col and sub_col and stn_col in issue_df.columns and sub_col in issue_df.columns:
        group_keys = [stn_col]
        if stn_name_col and stn_name_col != stn_col and stn_name_col in issue_df.columns:
            group_keys.append(stn_name_col)
        if zme_name_col and zme_name_col in issue_df.columns:
            group_keys.append(zme_name_col)
        group_keys.append(sub_col)

        pair_counts = issue_df.groupby(group_keys).size().reset_index(name='Occurrences')
        repeats = pair_counts[pair_counts['Occurrences'] >= 2].sort_values(by='Occurrences', ascending=False)
        if not repeats.empty:
            for idx, r in repeats.iterrows():
                vals = [r[stn_col]]
                if stn_name_col and stn_name_col in r:
                    vals.append(r[stn_name_col])
                if zme_name_col and zme_name_col in r:
                    vals.append(r[zme_name_col])
                vals.extend([r[sub_col], int(r['Occurrences'])])
                row = data_row(ws, row, vals)
        else:
            row = data_row(ws, row, ['None found in current data'] + [''] * (len(headers) - 1))

    row += 1
    status_header_row = row + 1
    row = section_title(ws, row, '4. Status Breakdown', 2)
    row = header_row(ws, row, ['Status', 'Count'])
    status_start = row
    if status_col and status_col in issue_df.columns:
        status_vc = issue_df[status_col].dropna().value_counts()
        for status, cnt in status_vc.items():
            row = data_row(ws, row, [str(status), int(cnt)])
        status_end = row - 1
        # Add Total Row
        row = data_row(ws, row, ['Total', int(status_vc.sum())])
    else:
        status_end = row - 1

    if status_end >= status_start:
        pie_status = PieChart()
        pie_status.title = "Status Breakdown"
        data_s = Reference(ws, min_col=2, min_row=status_header_row, max_row=status_end)
        cats_s = Reference(ws, min_col=1, min_row=status_start, max_row=status_end)
        pie_status.add_data(data_s, titles_from_data=True)
        pie_status.set_categories(cats_s)
        pie_status.width = 14
        pie_status.height = 8
        ws.add_chart(pie_status, f"D{status_header_row - 1}")

    row += 1
    row = section_title(ws, row, '5. Customer Filter (B2B / B2C)', 2)
    row = header_row(ws, row, ['Segment', 'Count'])
    seg_col = find_col(issue_df, ['B2B/ B2C', 'B2B/B2C', 'Segment', 'Customer Segment'])
    if seg_col and seg_col in issue_df.columns:
        for seg, cnt in issue_df[seg_col].dropna().value_counts().items():
            row = data_row(ws, row, [str(seg), int(cnt)])

    for col, width in zip('ABCDEFGH', [26, 20, 16, 22, 26, 14]):
        ws.column_dimensions[col].width = width


def build_pm_dashboard(wb, pm_df, prange):
    ws = wb.create_sheet('Dashboard - PM F-01', 1)
    ws.sheet_view.showGridLines = False
    ws['A1'] = 'PM F-01 — PREVENTIVE MAINTENANCE DASHBOARD'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:J1')
    ws['A2'] = 'Scope: Filter-wise PM F-01 summary report.'
    ws['A2'].font = NOTE_FONT
    ws.merge_cells('A2:J2')

    zme_col = find_col(pm_df, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'])
    zone_col = find_col(pm_df, ['Zone', 'Zone Name', 'Region'])
    chg_col = find_col(pm_df, ['Charger ID', 'Charger_ID', 'Charger', 'OCPP ID', 'OCPP_ID'])
    stn_col = find_col(pm_df, ['Station ID', 'Station_ID', 'Station'])
    st_col = find_col(pm_df, ['PM Status', 'PM_Status', 'Status'])
    adv_col = find_col(pm_df, ['Advance PM Done', 'Advance PM'])

    row = 4
    pm_header_row = row + 1
    if zme_col and zme_col in pm_df.columns:
        group_cols = [zme_col]
        if zone_col and zone_col in pm_df.columns:
            group_cols.append(zone_col)

        zme_summary = []
        for name_tuple, group in pm_df.groupby(group_cols):
            z_name = name_tuple[0] if isinstance(name_tuple, tuple) else name_tuple
            z_zone = name_tuple[1] if isinstance(name_tuple, tuple) and len(name_tuple) > 1 else (group[zone_col].iloc[0] if zone_col and zone_col in group.columns else 'N/A')

            t_chg = len(group[chg_col].dropna().unique()) if chg_col and chg_col in group.columns else len(group)
            t_stn = len(group[stn_col].dropna().unique()) if stn_col and stn_col in group.columns else len(group)
            planning = len(group)

            if st_col and st_col in group.columns:
                st_u = group[st_col].astype(str).str.strip().str.upper()
                done = int((st_u == 'YES').sum())
                pending = int((st_u == 'NO').sum())
            else:
                done = int(group['Is_PM_Done'].sum()) if 'Is_PM_Done' in group.columns else 0
                pending = int(group['Is_PM_Pending'].sum()) if 'Is_PM_Pending' in group.columns else 0

            if adv_col and adv_col in group.columns:
                adv = int((group[adv_col].astype(str).str.strip().str.upper() == 'YES').sum())
            else:
                adv = 0

            eff = (done / planning) if planning > 0 else 0.0
            zme_summary.append((str(z_name), str(z_zone), t_chg, t_stn, planning, done, pending, adv, eff))

        row = section_title(ws, row, 'PM Summary by ZME', 9)
        row = header_row(ws, row, ['ZME Name', 'Zone', 'Total Chargers', 'Total Stations', 'PM Planning',
                                    'PM Done', 'PM Pending', 'Advance PM Done', 'PM Efficiency'])
        pm_start = row
        for item in sorted(zme_summary, key=lambda x: x[4], reverse=True):
            row = data_row(ws, row, list(item), pct_cols={8})

        pm_end = row - 1
        if pm_end >= pm_start:
            chart_pm = BarChart()
            chart_pm.type = "col"
            chart_pm.style = 10
            chart_pm.title = "PM Planning vs PM Done by ZME"
            chart_pm.y_axis.title = "PM Work Orders"
            chart_pm.x_axis.title = "ZME"

            data_pm = Reference(ws, min_col=5, min_row=pm_header_row, max_col=7, max_row=pm_end)
            cats_pm = Reference(ws, min_col=1, min_row=pm_start, max_row=pm_end)
            chart_pm.add_data(data_pm, titles_from_data=True)
            chart_pm.set_categories(cats_pm)
            chart_pm.width = 18
            chart_pm.height = 10
            ws.add_chart(chart_pm, "K4")

    for col, width in zip('ABCDEFGHIJ', [18, 10, 14, 14, 13, 11, 12, 16, 14, 10]):
        ws.column_dimensions[col].width = width


@st.cache_data(show_spinner=False)
def get_dataframes_cached(source_bytes_or_path, pm_bytes_or_path=None):
    """Cached DataFrame parser for instant dashboard rendering (<0.2s)."""
    if pm_bytes_or_path is None:
        issue_df = load_issue_tracker(source_bytes_or_path)
        pm_df = load_pm_tracker(source_bytes_or_path)
    else:
        issue_df = load_issue_tracker(source_bytes_or_path)
        pm_df = load_pm_tracker(pm_bytes_or_path)
    return issue_df, pm_df


def generate_mpr_workbook_from_dfs(issue_df, pm_df):
    """Generates the Excel report package matching the active sidebar filters."""
    wb = Workbook()
    wb.remove(wb.active)

    issue_ws, issue_last = write_data_sheet(
        wb, 'Issue Data', issue_df, 'IssueTable',
        date_cols=['Issue Date', 'Resolution Date', 'Restoration Date'])
    exclude_cols = {'Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed', 'Is_TAT_Compliant', 'Is_TAT_Breached', 'Is_PM_Done', 'Is_PM_Pending', '_Is_Open_', '_Is_Closed_', '_Is_Within_', '_Is_Without_', '_Is_Closed_Within_', '_Is_Closed_Without_'}
    clean_issue_cols = [c for c in issue_df.columns if c not in exclude_cols]
    icol = {name: get_column_letter(i + 1) for i, name in enumerate(clean_issue_cols)}

    pm_ws, pm_last = write_data_sheet(
        wb, 'PM Data', pm_df, 'PMTable',
        date_cols=['Go Live Date', 'Due Date', 'Actual Completion Date'])
    clean_pm_cols = [c for c in pm_df.columns if c not in exclude_cols]
    pcol = {name: get_column_letter(i + 1) for i, name in enumerate(clean_pm_cols)}
    add_pm_helper_columns(pm_ws, pm_df, pcol, pm_last)
    pcol_full = {name: get_column_letter(i + 1)
                 for i, name in enumerate(list(clean_pm_cols) + ['Advance PM Done', 'First Station Occurrence'])}

    def irange(col):
        col_aliases = {
            'ZME': ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'],
            'Zone': ['Zone', 'Zone Name', 'Region'],
            'Status': ['Status', 'Ticket Status', 'Issue Status', 'State'],
            'TAT Compliance': ['TAT Compliance', 'SLA Compliance', 'Compliance', 'TAT Status'],
            'Station ID': ['Station ID', 'Station_ID', 'Station', 'Site ID'],
            'Issue Sub-Type': ['Issue Sub-Type', 'Issue Sub Type', 'Sub Type', 'Fault Subtype'],
            'Severity': ['Severity', 'Ticket Severity', 'Priority'],
            'B2B/ B2C': ['B2B/ B2C', 'B2B/B2C', 'Segment', 'Customer Segment']
        }
        possible = col_aliases.get(col, [col])
        target_name = find_col(issue_df, possible)
        letter = icol.get(target_name, 'A') if target_name else 'A'
        return f"'Issue Data'!${letter}$2:${letter}${max(100, len(issue_df) + 1)}"

    def prange(col):
        col_aliases = {
            'ZME': ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'],
            'Zone': ['Zone', 'Zone Name', 'Region'],
            'Station ID': ['Station ID', 'Station_ID', 'Station'],
            'Due Date': ['Due Date', 'PM Due Date', 'Scheduled Date', 'Schedule Date', 'Due_Date'],
            'PM Status': ['PM Status', 'PM_Status', 'Status'],
            'Advance PM Done': ['Advance PM Done', 'Advance PM'],
            'First Station Occurrence': ['First Station Occurrence']
        }
        possible = col_aliases.get(col, [col])
        target_name = find_col(pm_df, possible)
        letter = pcol_full.get(target_name, 'A') if target_name else 'A'
        return f"'PM Data'!${letter}$2:${letter}${max(100, len(pm_df) + 1)}"

    build_issue_dashboard(wb, issue_df, irange)
    build_pm_dashboard(wb, pm_df, prange)
    wb.active = 0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def generate_workbook_cached(source_bytes_or_path, pm_bytes_or_path=None):
    """Cached workbook builder for Excel export generation."""
    issue_df, pm_df = get_dataframes_cached(source_bytes_or_path, pm_bytes_or_path)
    return generate_mpr_workbook_from_dfs(issue_df, pm_df)


def plot_pie_chart(labels, values, title, colors=None, hole=0.45):
    """Renders a responsive, interactive Donut/Pie Chart in Red & White Theme."""
    if HAS_PLOTLY:
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=hole,
            marker_colors=colors if colors else ['#991B1B', '#DC2626', '#EF4444', '#F87171', '#FCA5A5'],
            textinfo='percent+value',
            hovertemplate='<b>%{label}</b><br>Count: <b>%{value:,}</b> (%{percent})<extra></extra>',
            insidetextfont=dict(color='#FFFFFF', size=13, family='Inter')
        )])
        fig.update_layout(
            title=dict(text=title, font=dict(size=14, color='#991B1B', family='Inter', weight='bold')),
            margin=dict(l=15, r=15, t=45, b=15),
            height=320,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': True, 'responsive': True, 'displaylogo': False})
    else:
        df_pie = pd.DataFrame({'Category': labels, 'Value': values})
        st.write(f"**{title}**")
        st.dataframe(df_pie, use_container_width=True)


def plot_vertical_bar(df, x_col, y_col, title, color_hex="#DC2626"):
    """Renders a clean, interactive vertical bar chart in Red & White Theme."""
    if HAS_PLOTLY:
        fig = px.bar(
            df,
            x=x_col,
            y=y_col,
            title=title,
            text=y_col,
            color_discrete_sequence=[color_hex]
        )
        fig.update_traces(
            texttemplate='%{text}',
            textposition='outside',
            textfont=dict(size=12, family='Inter'),
            hovertemplate='<b>%{x}</b><br>Total Count: <b>%{y:,}</b><extra></extra>'
        )
        fig.update_layout(
            margin=dict(l=15, r=15, t=45, b=15),
            height=340,
            font=dict(family='Inter', color='#0F172A'),
            xaxis_title=x_col,
            yaxis_title=y_col,
            plot_bgcolor='rgba(0,0,0,0)',
            yaxis=dict(showgrid=True, gridcolor='#FEE2E2')
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': True, 'responsive': True, 'displaylogo': False})
    else:
        st.write(f"**{title}**")
        st.bar_chart(df.set_index(x_col)[y_col], height=280)


def plot_grouped_bar(df, x_col, y_cols, title, colors=None):
    """Renders a responsive, interactive multi-series grouped column chart in Red & White Theme."""
    if HAS_PLOTLY:
        fig = go.Figure()
        palette = colors if colors else ['#991B1B', '#EF4444', '#DC2626', '#7F1D1D']
        for idx, col in enumerate(y_cols):
            fig.add_trace(go.Bar(
                name=col,
                x=df[x_col],
                y=df[col],
                marker_color=palette[idx % len(palette)],
                text=df[col],
                textposition='auto',
                textfont=dict(size=11, family='Inter'),
                hovertemplate='<b>%{x}</b><br>' + col + ': <b>%{y:,}</b><extra></extra>'
            ))
        fig.update_layout(
            barmode='group',
            hovermode='x unified',
            title=dict(text=title, font=dict(size=14, color='#991B1B', family='Inter', weight='bold')),
            margin=dict(l=15, r=15, t=45, b=15),
            height=340,
            legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="center", x=0.5),
            plot_bgcolor='rgba(0,0,0,0)',
            yaxis=dict(showgrid=True, gridcolor='#FEE2E2')
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': True, 'responsive': True, 'displaylogo': False})
    else:
        st.write(f"**{title}**")
        st.dataframe(df.set_index(x_col)[y_cols], use_container_width=True)


def run_streamlit_app():
    st.set_page_config(
        page_title="ChargeZone | Executive MPR & PM Dashboard",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Professional Executive Styling - Red, White & Slate Grey Theme
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
        
        /* Force Crisp Clean Background on Main Container & Header */
        .stApp, [data-testid="stMain"], [data-testid="stHeader"] {
            background-color: #F8FAFC !important;
        }

        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #E2E8F0 !important;
        }

        /* Executive Typography */
        html, body, [class*="css"], p, span, label, div {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            color: #1E293B !important;
        }

        h1, h2, h3, h4, h5, h6 {
            color: #0F172A !important;
            font-weight: 800 !important;
            letter-spacing: -0.02em !important;
        }

        /* Executive Crimson Buttons */
        .stButton > button, div[data-testid="stDownloadButton"] > button {
            background: linear-gradient(135deg, #801B1B 0%, #B91C1C 50%, #DC2626 100%) !important;
            color: #FFFFFF !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            font-size: 0.92rem !important;
            border: none !important;
            padding: 0.65rem 1.4rem !important;
            box-shadow: 0 4px 14px rgba(185, 28, 28, 0.25) !important;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
            letter-spacing: 0.01em !important;
        }
        .stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {
            background: linear-gradient(135deg, #7F1D1D 0%, #991B1B 50%, #B91C1C 100%) !important;
            box-shadow: 0 6px 20px rgba(153, 27, 27, 0.38) !important;
            transform: translateY(-2px) !important;
        }

        /* Streamlit Tabs Red Accent */
        button[data-baseweb="tab"] {
            font-weight: 700 !important;
            color: #64748B !important;
            font-size: 0.95rem !important;
            padding: 0.75rem 1.25rem !important;
            border-radius: 8px 8px 0 0 !important;
            transition: color 0.15s ease !important;
        }
        button[data-baseweb="tab"]:hover {
            color: #991B1B !important;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            color: #991B1B !important;
            border-bottom: 3px solid #DC2626 !important;
            background-color: rgba(220, 38, 38, 0.04) !important;
        }
        div[data-baseweb="tab-highlight"] {
            background-color: #DC2626 !important;
            height: 3px !important;
        }

        /* Executive Banner Header Box */
        .exec-header-box {
            background: linear-gradient(135deg, #6B1111 0%, #991B1B 45%, #B91C1C 100%);
            padding: 1.6rem 2.2rem;
            border-radius: 16px;
            color: #FFFFFF;
            box-shadow: 0 10px 28px rgba(153, 27, 27, 0.20);
            margin-bottom: 1.5rem;
            border-left: 8px solid #EF4444;
        }

        .exec-badge {
            background: rgba(255, 255, 255, 0.18);
            color: #FFFFFF;
            font-size: 0.72rem;
            font-weight: 800;
            padding: 5px 14px;
            border-radius: 20px;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            display: inline-block;
            margin-bottom: 0.6rem;
            border: 1px solid rgba(255, 255, 255, 0.35);
            backdrop-filter: blur(6px);
        }

        .exec-title {
            font-size: 2.1rem;
            font-weight: 900;
            color: #FFFFFF !important;
            margin: 0;
            line-height: 1.2;
            letter-spacing: -0.025em;
        }

        .exec-subtitle {
            font-size: 0.95rem;
            color: #FEE2E2 !important;
            margin-top: 0.35rem;
            margin-bottom: 0;
            font-weight: 500;
        }

        /* Metric Cards - Pure Red & White Corporate Theme */
        .metric-card {
            background: #FFFFFF;
            border: 1px solid #FEE2E2;
            border-radius: 14px;
            padding: 1.15rem 1.35rem;
            box-shadow: 0 4px 14px rgba(153, 27, 27, 0.05);
            transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .metric-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 25px rgba(153, 27, 27, 0.12);
            border-color: #FCA5A5;
        }
        .metric-card.red { border-top: 5px solid #DC2626; }
        .metric-card.darkred { border-top: 5px solid #991B1B; }
        .metric-card.grey { border-top: 5px solid #7F1D1D; }
        .metric-card.green { border-top: 5px solid #B91C1C; }
        .metric-card.amber { border-top: 5px solid #EF4444; }

        .metric-label {
            font-size: 0.73rem;
            font-weight: 800;
            color: #991B1B !important;
            text-transform: uppercase;
            letter-spacing: 0.07em;
        }
        .metric-val {
            font-size: 1.85rem;
            font-weight: 900;
            color: #7F1D1D !important;
            margin-top: 0.25rem;
            letter-spacing: -0.02em;
        }
        .metric-sub {
            font-size: 0.74rem;
            color: #991B1B !important;
            margin-top: 0.2rem;
            font-weight: 600;
        }

        .section-header {
            font-size: 1.15rem;
            font-weight: 800;
            color: #991B1B !important;
            margin-top: 1.2rem;
            margin-bottom: 0.9rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            border-bottom: 2px solid #FEE2E2;
            padding-bottom: 0.45rem;
            letter-spacing: -0.01em;
        }

        /* Streamlit Container & Table Styling */
        [data-testid="stForm"], .stDataFrame, div[data-testid="stExpander"] {
            background-color: #FFFFFF !important;
            border-radius: 12px !important;
            border: 1px solid #E2E8F0 !important;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.02) !important;
        }

        /* Radio & Selectbox Styling */
        div[class*="stRadio"] label, div[class*="stSelectbox"] label {
            font-weight: 700 !important;
            color: #334155 !important;
            font-size: 0.88rem !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # Executive Meeting Presentation Header
    st.markdown("""
        <div class="exec-header-box">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;">
                <div>
                    <span class="exec-badge">⚡ CHARGEZONE EXECUTIVE BOARD • LIVE OPERATIONS REVIEW</span>
                    <h1 class="exec-title">Monthly Progress Report (MPR) & PM Governance</h1>
                    <p class="exec-subtitle">C-Suite Operations Deck: SLA Performance, Preventive Maintenance (PM F-01) & Infrastructure Analytics.</p>
                </div>
                <div style="text-align: right; background: rgba(255, 255, 255, 0.15); padding: 8px 18px; border-radius: 10px; border: 1px solid rgba(255, 255, 255, 0.3);">
                    <div style="font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.09em; color: #FEE2E2; font-weight: 700;">Executive Deck</div>
                    <div style="font-size: 1.05rem; font-weight: 900; color: #FFFFFF;">FY 2026-27 Review</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Sidebar Control Panel
    st.sidebar.markdown("### ⚙️ Control Panel")
    st.sidebar.markdown("---")

    MAX_FILE_SIZE_MB = 20
    MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

    st.sidebar.markdown("#### 1. Data Upload (Max 20 MB per file)")
    issue_file = st.sidebar.file_uploader("1️⃣ Issue Tracker File (.xlsx / .csv)", type=["xlsx", "csv"], key="issue_file", help="Max allowed file size: 20 MB")
    pm_file = st.sidebar.file_uploader("2️⃣ PM Tracker File (.xlsx / .csv)", type=["xlsx", "csv"], key="pm_file", help="Max allowed file size: 20 MB")

    # Enforce 20 MB File Size Restriction
    if issue_file is not None and issue_file.size > MAX_FILE_SIZE_BYTES:
        st.sidebar.error(f"**Issue Tracker file exceeds 20 MB limit** ({issue_file.size / (1024*1024):.1f} MB). Please upload a file smaller than 20 MB.")
        issue_file = None

    if pm_file is not None and pm_file.size > MAX_FILE_SIZE_BYTES:
        st.sidebar.error(f"**PM Tracker file exceeds 20 MB limit** ({pm_file.size / (1024*1024):.1f} MB). Please upload a file smaller than 20 MB.")
        pm_file = None

    issue_input = None
    pm_input = None
    file_status_text = ""

    use_default = False
    default_issue = "Demo for mpr cm.xlsx"
    default_pm = "Demo for mpr pm.xlsx"

    if issue_file is None and pm_file is None:
        try:
            import os
            if os.path.exists(default_issue) and os.path.exists(default_pm):
                use_default = st.sidebar.checkbox("Load Sample Datasets (`Demo for mpr cm.xlsx` & `Demo for mpr pm.xlsx`)", value=True)
        except Exception:
            pass

    if issue_file is not None and pm_file is not None:
        issue_input = issue_file.getvalue()
        pm_input = pm_file.getvalue()
        file_status_text = f"✓ Connected Issue: `{issue_file.name}` & PM: `{pm_file.name}`"
    elif use_default:
        issue_input = default_issue
        pm_input = default_pm
        file_status_text = f"✓ Connected Sample Files: `{default_issue}` & `{default_pm}`"
    elif issue_file is not None or pm_file is not None:
        st.sidebar.warning("⚠️ Please upload BOTH the Issue Tracker file and the PM Tracker file.")

    if issue_input is None or pm_input is None:
        st.info("📌 **Upload Required**: Upload BOTH the Issue Tracker file and the PM Tracker file in the sidebar to load the dashboard (Max 20 MB each).")
        
        st.markdown("### 📥 Upload Tracker Files (Max 20 MB)")
        st.markdown("""
        #### 📂 Upload 2 Separate Files (up to 20 MB each):
        - **File 1**: Issue Tracker Data (`.xlsx` or `.csv`) containing `Issue Tracker` sheet.
        - **File 2**: PM Tracker Data (`.xlsx` or `.csv`) containing `PM Tracker B2C- B2B` sheet.
        """)
        return

    st.sidebar.success(file_status_text)
    st.sidebar.markdown("---")

    # Processing Workbook & Dataframes with Caching
    with st.spinner("Processing Operational Data Engine..."):
        try:
            raw_issue_df, raw_pm_df = get_dataframes_cached(issue_input, pm_input)
        except Exception as e:
            st.error(f"⚠️ **Processing Error**: {e}")
            st.exception(e)
            return

    st.sidebar.markdown("#### 📅 Dashboard Filter: Month(s) Selection")

    # Gather all unique available months across Issue Tracker and PM Tracker
    available_months_set = set()
    if 'Issue Date Parsed' in raw_issue_df.columns:
        valid_issue_d = raw_issue_df['Issue Date Parsed'].dropna()
        for dt in valid_issue_d:
            available_months_set.add((dt.year, dt.month))

    for dt_col in ['Due Date Parsed', 'Actual Completion Date Parsed', 'Go Live Date Parsed']:
        if dt_col in raw_pm_df.columns:
            valid_pm_d = raw_pm_df[dt_col].dropna()
            for dt in valid_pm_d:
                available_months_set.add((dt.year, dt.month))

    sorted_ym = sorted(list(available_months_set))
    month_labels = [datetime(y, m, 1).strftime('%b-%Y') for y, m in sorted_ym]

    if month_labels:
        selected_months = st.sidebar.multiselect(
            "Select Month(s) to View Dashboard Data:",
            options=month_labels,
            default=month_labels,
            help="Select one or multiple months to dynamically update all charts and tables across the dashboard."
        )
    else:
        selected_months = []

    # Optional Fine-tune Custom Date Range
    start_d, end_d = None, None
    with st.sidebar.expander("📆 Optional: Fine-tune by Custom Date Range", expanded=False):
        all_dates = []
        if 'Issue Date Parsed' in raw_issue_df.columns:
            all_dates.extend(raw_issue_df['Issue Date Parsed'].dropna().dt.date.tolist())
        for dt_col in ['Due Date Parsed', 'Actual Completion Date Parsed', 'Go Live Date Parsed']:
            if dt_col in raw_pm_df.columns:
                all_dates.extend(raw_pm_df[dt_col].dropna().dt.date.tolist())

        if all_dates:
            min_date = min(all_dates)
            max_date = max(all_dates)
            selected_dates = st.date_input("Exact Date Range:", value=(min_date, max_date), min_value=min_date, max_value=max_date)
            if selected_dates and isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
                start_d, end_d = selected_dates[0], selected_dates[1]

    # Apply Strict Multi-Month Filter to Issue Dataframe
    filtered_issue_df = raw_issue_df.copy()
    if 'Issue Date Parsed' in filtered_issue_df.columns and filtered_issue_df['Issue Date Parsed'].notna().any():
        if selected_months:
            filtered_issue_df = filtered_issue_df[
                filtered_issue_df['Issue Date Parsed'].dt.strftime('%b-%Y').isin(selected_months)
            ]
        if start_d and end_d:
            filtered_issue_df = filtered_issue_df[
                (filtered_issue_df['Issue Date Parsed'].dt.date >= start_d) &
                (filtered_issue_df['Issue Date Parsed'].dt.date <= end_d)
            ]

    # Apply Strict Multi-Month Filter to PM Dataframe
    filtered_pm_df = raw_pm_df.copy()
    if selected_months:
        pm_month_match = pd.Series(False, index=filtered_pm_df.index)
        for dt_col in ['Due Date Parsed', 'Actual Completion Date Parsed', 'Go Live Date Parsed']:
            if dt_col in filtered_pm_df.columns and filtered_pm_df[dt_col].notna().any():
                pm_month_match |= filtered_pm_df[dt_col].dt.strftime('%b-%Y').isin(selected_months)
        # If any date column matched, filter strictly
        if pm_month_match.any():
            filtered_pm_df = filtered_pm_df[pm_month_match]

    if start_d and end_d:
        pm_range_match = pd.Series(False, index=filtered_pm_df.index)
        for dt_col in ['Due Date Parsed', 'Actual Completion Date Parsed', 'Go Live Date Parsed']:
            if dt_col in filtered_pm_df.columns and filtered_pm_df[dt_col].notna().any():
                pm_range_match |= (filtered_pm_df[dt_col].dt.date >= start_d) & (filtered_pm_df[dt_col].dt.date <= end_d)
        if pm_range_match.any():
            filtered_pm_df = filtered_pm_df[pm_range_match]

    st.sidebar.markdown("---")

    # Download Button
    st.sidebar.markdown("#### 3. Export MPR Excel Package")
    timestamp = datetime.now().strftime("%Y-%m")
    out_filename = f"ChargeZone_MPR_Report_{timestamp}.xlsx"

    if st.sidebar.button("⚡ Prepare MPR Excel Report", use_container_width=True, type="primary"):
        with st.spinner("Building Excel report..."):
            st.session_state['mpr_excel_bytes'] = generate_mpr_workbook_from_dfs(filtered_issue_df, filtered_pm_df)
            st.session_state['mpr_excel_filename'] = out_filename

    if 'mpr_excel_bytes' in st.session_state:
        st.sidebar.download_button(
            label="⬇️ Download MPR Excel Report",
            data=st.session_state['mpr_excel_bytes'],
            file_name=st.session_state.get('mpr_excel_filename', out_filename),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Main Tabs
    tab_issues, tab_pm, tab_raw = st.tabs([
        "📉 Issue MPR Charts & Tables",
        "🛠️ PM F-01 Charts & Tables",
        "📋 Data Explorer"
    ])

    # ---------------------------------------------------------
    # TAB 1: ISSUE MPR DASHBOARD (CHARTS + TABLES SIDE-BY-SIDE)
    # ---------------------------------------------------------
    with tab_issues:
        st.markdown('<div class="section-header">📊 Operational Issue & SLA Analytics Performance</div>', unsafe_allow_html=True)

        status_col = find_col(filtered_issue_df, ['Status', 'Ticket Status', 'Issue Status', 'State'])
        tat_col = find_col(filtered_issue_df, ['TAT Compliance', 'SLA Compliance', 'Compliance', 'TAT Status'])
        
        # Prepare status helper columns
        if status_col and status_col in filtered_issue_df.columns:
            s_status = filtered_issue_df[status_col].astype(str).str.strip().str.upper()
            is_closed = s_status.isin(['CLOSED', 'RESOLVED'])
            filtered_issue_df['_Is_Open_'] = ~is_closed
            filtered_issue_df['_Is_Closed_'] = is_closed
        else:
            filtered_issue_df['_Is_Closed_'] = True
            filtered_issue_df['_Is_Open_'] = False

        # Prepare TAT compliance helper columns
        if tat_col and tat_col in filtered_issue_df.columns:
            s_tat = filtered_issue_df[tat_col].astype(str).str.strip().str.upper()
            filtered_issue_df['_Is_Within_'] = (s_tat == 'YES')
            filtered_issue_df['_Is_Without_'] = (s_tat == 'NO')
            filtered_issue_df['_Is_Closed_Within_'] = filtered_issue_df['_Is_Closed_'] & (s_tat == 'YES')
            filtered_issue_df['_Is_Closed_Without_'] = filtered_issue_df['_Is_Closed_'] & (s_tat == 'NO')
        else:
            filtered_issue_df['_Is_Within_'] = filtered_issue_df.get('Is_TAT_Compliant', False)
            filtered_issue_df['_Is_Without_'] = filtered_issue_df.get('Is_TAT_Breached', False)
            filtered_issue_df['_Is_Closed_Within_'] = filtered_issue_df['_Is_Closed_'] & filtered_issue_df['_Is_Within_']
            filtered_issue_df['_Is_Closed_Without_'] = filtered_issue_df['_Is_Closed_'] & filtered_issue_df['_Is_Without_']

        # Interactive Instant Search & Ticket State Filter Bar
        c_srch, c_fltr = st.columns([7, 5])
        with c_srch:
            search_txt = st.text_input("🔍 Instant Search (Station ID, Station Name, ZME, or OCPP ID):", placeholder="e.g. STN-101, Ahmedabad, Rahul...", key="quick_tab1_search")
        with c_fltr:
            quick_status = st.selectbox("📌 Filter Dashboard View by Ticket State:", ["All Active Tickets", "Open Tickets Only", "Closed Within TAT Only", "Closed Without TAT Only"], key="quick_tab1_status_select")

        # Apply Instant Search & Quick Ticket Filter dynamically
        if search_txt and search_txt.strip() != "":
            q = search_txt.strip().lower()
            match_mask = pd.Series(False, index=filtered_issue_df.index)
            for col in filtered_issue_df.columns:
                if not col.startswith('_'):
                    match_mask |= filtered_issue_df[col].astype(str).str.lower().str.contains(q, na=False)
            filtered_issue_df = filtered_issue_df[match_mask]

        if quick_status == "Open Tickets Only":
            filtered_issue_df = filtered_issue_df[filtered_issue_df['_Is_Open_']]
        elif quick_status == "Closed Within TAT Only":
            filtered_issue_df = filtered_issue_df[filtered_issue_df['_Is_Closed_Within_']]
        elif quick_status == "Closed Without TAT Only":
            filtered_issue_df = filtered_issue_df[filtered_issue_df['_Is_Closed_Without_']]

        total_issues = len(filtered_issue_df)

        total_open = int(filtered_issue_df['_Is_Open_'].sum())
        total_closed = int(filtered_issue_df['_Is_Closed_'].sum())
        closed_within = int(filtered_issue_df['_Is_Closed_Within_'].sum())
        closed_without = int(filtered_issue_df['_Is_Closed_Without_'].sum())

        # Formula 1: Overall CM Efficiency = Closed Faults / Faults Registered * 100
        overall_cm_eff = (total_closed / total_issues * 100) if total_issues > 0 else 0.0
        # Formula 2: CM-TAT Efficiency = Closed Within TAT / Total Closed Faults * 100
        cm_tat_eff = (closed_within / total_closed * 100) if total_closed > 0 else 0.0

        # High-Impact KPI Row (7 Columns)
        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
        with c1:
            st.markdown(f"""
                <div class="metric-card darkred">
                    <div class="metric-label">Faults Registered</div>
                    <div class="metric-val">{total_issues:,}</div>
                    <div class="metric-sub">Total Tickets Logged</div>
                </div>
            """, unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
                <div class="metric-card amber">
                    <div class="metric-label">Open Faults</div>
                    <div class="metric-val">{total_open:,}</div>
                    <div class="metric-sub">Pending / In-Progress</div>
                </div>
            """, unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
                <div class="metric-card grey">
                    <div class="metric-label">Closed Faults</div>
                    <div class="metric-val">{total_closed:,}</div>
                    <div class="metric-sub">Resolved Tickets</div>
                </div>
            """, unsafe_allow_html=True)
        with c4:
            st.markdown(f"""
                <div class="metric-card green">
                    <div class="metric-label">Closed Within TAT</div>
                    <div class="metric-val">{closed_within:,}</div>
                    <div class="metric-sub">SLA Compliant</div>
                </div>
            """, unsafe_allow_html=True)
        with c5:
            st.markdown(f"""
                <div class="metric-card red">
                    <div class="metric-label">Closed Without TAT</div>
                    <div class="metric-val">{closed_without:,}</div>
                    <div class="metric-sub">SLA Breached</div>
                </div>
            """, unsafe_allow_html=True)
        with c6:
            st.markdown(f"""
                <div class="metric-card {'green' if overall_cm_eff >= 85 else 'amber'}">
                    <div class="metric-label">Overall CM Efficiency</div>
                    <div class="metric-val">{overall_cm_eff:.1f}%</div>
                    <div class="metric-sub">Closed / Registered</div>
                </div>
            """, unsafe_allow_html=True)
        with c7:
            st.markdown(f"""
                <div class="metric-card {'green' if cm_tat_eff >= 80 else 'amber'}">
                    <div class="metric-label">CM-TAT Efficiency</div>
                    <div class="metric-val">{cm_tat_eff:.1f}%</div>
                    <div class="metric-sub">Within TAT / Closed</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 1. ZME Performance, Overall CM & CM-TAT Efficiency Breakdown
        st.markdown('<div class="section-header">1. ZME Performance, Overall CM & CM-TAT Efficiency Breakdown</div>', unsafe_allow_html=True)
        zme_col = find_col(filtered_issue_df, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'])

        if zme_col:
            zme_df = filtered_issue_df.groupby(zme_col).agg(
                Faults_Received=(zme_col, 'count'),
                Open_Faults=('_Is_Open_', 'sum'),
                Closed_Faults=('_Is_Closed_', 'sum'),
                Closed_Within_TAT=('_Is_Closed_Within_', 'sum'),
                Closed_Without_TAT=('_Is_Closed_Without_', 'sum')
            ).reset_index()

            # Overall CM Efficiency = Closed Faults / Faults Registered * 100
            zme_df['Overall CM Efficiency %'] = (zme_df['Closed_Faults'] / zme_df['Faults_Received'].replace(0, pd.NA) * 100).fillna(0.0).round(1)
            # CM-TAT Efficiency = Closed Within TAT / Total Closed Faults * 100
            zme_df['CM-TAT Efficiency %'] = (zme_df['Closed_Within_TAT'] / zme_df['Closed_Faults'].replace(0, pd.NA) * 100).fillna(0.0).round(1)
            zme_df = zme_df.rename(columns={zme_col: 'ZME Name'}).sort_values(by='Faults_Received', ascending=False)

            st.write("##### 📊 ZME Fault & Efficiency Table")
            st.dataframe(
                zme_df[['ZME Name', 'Faults_Received', 'Open_Faults', 'Closed_Faults', 'Closed_Within_TAT', 'Closed_Without_TAT', 'Overall CM Efficiency %', 'CM-TAT Efficiency %']].style.format({
                    'Faults_Received': '{:,}',
                    'Open_Faults': '{:,}',
                    'Closed_Faults': '{:,}',
                    'Closed_Within_TAT': '{:,}',
                    'Closed_Without_TAT': '{:,}',
                    'Overall CM Efficiency %': '{:.1f}%',
                    'CM-TAT Efficiency %': '{:.1f}%'
                }).background_gradient(subset=['Overall CM Efficiency %', 'CM-TAT Efficiency %'], cmap='Reds'),
                use_container_width=True,
                height=320
            )
        else:
            st.info("ℹ️ ZME column not found in dataset.")

        st.markdown("---")

        # 2. Zone Performance, Overall CM & CM-TAT Efficiency Breakdown
        st.markdown('<div class="section-header">🏢 2. Zone Performance, Overall CM & CM-TAT Efficiency Breakdown</div>', unsafe_allow_html=True)
        col_zone_c, col_zone_t = st.columns([6, 6])
        zone_col = find_col(filtered_issue_df, ['Zone', 'Zone Name', 'Region'])

        if zone_col:
            zone_df = filtered_issue_df.groupby(zone_col).agg(
                Faults_Received=(zone_col, 'count'),
                Open_Faults=('_Is_Open_', 'sum'),
                Closed_Faults=('_Is_Closed_', 'sum'),
                Closed_Within_TAT=('_Is_Closed_Within_', 'sum'),
                Closed_Without_TAT=('_Is_Closed_Without_', 'sum')
            ).reset_index()

            zone_df['Overall CM Efficiency %'] = (zone_df['Closed_Faults'] / zone_df['Faults_Received'].replace(0, pd.NA) * 100).fillna(0.0).round(1)
            zone_df['CM-TAT Efficiency %'] = (zone_df['Closed_Within_TAT'] / zone_df['Closed_Faults'].replace(0, pd.NA) * 100).fillna(0.0).round(1)
            zone_df = zone_df.rename(columns={zone_col: 'Zone'}).sort_values(by='Faults_Received', ascending=False)

            with col_zone_c:
                plot_grouped_bar(
                    df=zone_df,
                    x_col='Zone',
                    y_cols=['Faults_Received', 'Open_Faults', 'Closed_Within_TAT', 'Closed_Without_TAT'],
                    title="Faults Registered vs Open vs Within TAT vs Without TAT by Zone",
                    colors=['#991B1B', '#EF4444', '#DC2626', '#7F1D1D']
                )

            with col_zone_t:
                st.write("##### 🏢 Zone Fault & Efficiency Table")
                st.dataframe(
                    zone_df[['Zone', 'Faults_Received', 'Open_Faults', 'Closed_Faults', 'Closed_Within_TAT', 'Closed_Without_TAT', 'Overall CM Efficiency %', 'CM-TAT Efficiency %']].style.format({
                        'Faults_Received': '{:,}',
                        'Open_Faults': '{:,}',
                        'Closed_Faults': '{:,}',
                        'Closed_Within_TAT': '{:,}',
                        'Closed_Without_TAT': '{:,}',
                        'Overall CM Efficiency %': '{:.1f}%',
                        'CM-TAT Efficiency %': '{:.1f}%'
                    }).background_gradient(subset=['Overall CM Efficiency %', 'CM-TAT Efficiency %'], cmap='Reds'),
                    use_container_width=True,
                    height=300
                )

            # -------------------------------------------------------------
            # INTERACTIVE TICKET & STATION DETAILS INSPECTOR (DRILL-DOWN)
            # -------------------------------------------------------------
            st.markdown("##### 🔍 Interactive Station & OCPP Ticket Inspector")
            zone_list = ["All Zones in Selected Data"] + sorted(list(zone_df['Zone'].dropna().astype(str).unique()))
            sel_zone = st.selectbox(
                "Select a Zone to drill down into underlying Station ID, OCPP ID & ZME ticket details:",
                zone_list,
                key="interactive_zone_inspector"
            )

            if sel_zone != "All Zones in Selected Data":
                inspect_df = filtered_issue_df[filtered_issue_df[zone_col].astype(str).str.strip() == str(sel_zone)]
            else:
                inspect_df = filtered_issue_df

            ocpp_col = find_col(inspect_df, ['OCPP ID', 'OCPP_ID', 'Charger ID', 'Connector ID', 'EVSE ID', 'OCPP', 'Station ID'])
            stn_id_col = find_col(inspect_df, ['Station ID', 'Station_ID', 'Station', 'Site ID'])
            stn_name_col = find_col(inspect_df, ['Station Name', 'Station_Name', 'Site Name'])
            zme_n_col = find_col(inspect_df, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'])
            sub_t_col = find_col(inspect_df, ['Issue Sub-Type', 'Issue Sub Type', 'Sub Type', 'Fault Subtype', 'Issue Subtype'])
            stat_c = find_col(inspect_df, ['Status', 'Ticket Status', 'Issue Status', 'State'])
            tat_c = find_col(inspect_df, ['TAT Compliance', 'SLA Compliance', 'Compliance', 'TAT Status'])

            show_cols = []
            for c in [ocpp_col, stn_id_col, stn_name_col, zme_n_col, sub_t_col, stat_c, tat_c]:
                if c and c in inspect_df.columns and c not in show_cols:
                    show_cols.append(c)

            if show_cols:
                st.dataframe(
                    inspect_df[show_cols],
                    use_container_width=True,
                    height=260
                )
                st.caption(f"Showing {len(inspect_df):,} ticket records for Zone: **{sel_zone}**")
        else:
            st.info("ℹ️ Zone column not found in dataset.")

        st.markdown("---")

        # 3. Status Breakdown & Customer Segment (Side-by-Side)
        c_stat, c_cust = st.columns(2)

        with c_stat:
            st.markdown('<div class="section-header">📌 3. Status Breakdown</div>', unsafe_allow_html=True)
            status_col = find_col(filtered_issue_df, ['Status', 'Ticket Status', 'Issue Status', 'State'])
            if status_col:
                status_counts = filtered_issue_df[status_col].dropna().value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                plot_vertical_bar(status_counts, x_col='Status', y_col='Count', title="Status Pipeline", color_hex="#8B5CF6")
                
                tot_cnt = int(status_counts['Count'].sum())
                status_tbl_df = pd.concat([
                    status_counts,
                    pd.DataFrame([{'Status': 'Total', 'Count': tot_cnt}])
                ], ignore_index=True)
                
                st.dataframe(status_tbl_df.style.format({'Count': '{:,}'}), use_container_width=True)
            else:
                st.info("ℹ️ Status column not found.")

        with c_cust:
            st.markdown('<div class="section-header">👥 4. Customer Segment (B2B/B2C)</div>', unsafe_allow_html=True)
            seg_col = find_col(filtered_issue_df, ['B2B/ B2C', 'B2B/B2C', 'Segment', 'Customer Segment'])
            if seg_col:
                cust_counts = filtered_issue_df[seg_col].dropna().value_counts().reset_index()
                cust_counts.columns = ['Segment', 'Count']
                plot_pie_chart(
                    labels=cust_counts['Segment'].tolist(),
                    values=cust_counts['Count'].tolist(),
                    title="Segment Share",
                    colors=['#2563EB', '#10B981'],
                    hole=0.4
                )
                st.dataframe(cust_counts, use_container_width=True)
            else:
                st.info("ℹ️ B2B/B2C Segment column not found.")

        st.markdown("---")

        # 5. Repetitive Faults & Region-wise Summary
        st.markdown('<div class="section-header">⚠️ 5. Repetitive Faults & Region-wise Summary (Station ID & Sub-Type ≥ 2)</div>', unsafe_allow_html=True)
        stn_col = find_col(filtered_issue_df, ['Station ID', 'Station_ID', 'Station', 'Site ID'])
        stn_name_col = find_col(filtered_issue_df, ['Station Name', 'Station_Name', 'Site Name'])
        zme_col = find_col(filtered_issue_df, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'])
        zone_col = find_col(filtered_issue_df, ['Zone', 'Zone Name', 'Region'])
        sub_col = find_col(filtered_issue_df, ['Issue Sub-Type', 'Issue Sub Type', 'Sub Type', 'Fault Subtype', 'Issue Subtype'])
        type_col = find_col(filtered_issue_df, ['Issue Type', 'Issue_Type', 'Category'])

        if stn_col and sub_col:
            group_cols = [stn_col]
            if stn_name_col and stn_name_col != stn_col:
                group_cols.append(stn_name_col)
            if zone_col:
                group_cols.append(zone_col)
            if zme_col:
                group_cols.append(zme_col)
            if type_col:
                group_cols.append(type_col)
            group_cols.append(sub_col)

            pair_counts = filtered_issue_df.groupby(group_cols).size().reset_index(name='Occurrences')
            repeats = pair_counts[pair_counts['Occurrences'] >= 2].sort_values(by='Occurrences', ascending=False)

            if not repeats.empty:
                # ---------------------------------------------------------
                # A. Region-wise (Zone-wise) Repetitive Faults Summary Table & Chart
                # ---------------------------------------------------------
                if zone_col and zone_col in repeats.columns:
                    st.markdown("##### 🌍 Region-wise Repetitive Fault Breakdown")
                    col_reg_c, col_reg_t = st.columns([6, 6])
                    
                    region_rep_df = repeats.groupby(zone_col).agg(
                        Total_Repetitive_Faults=('Occurrences', 'sum'),
                        Affected_Stations=(stn_col, 'nunique')
                    ).reset_index().rename(columns={zone_col: 'Region'}).sort_values(by='Total_Repetitive_Faults', ascending=False)
                    
                    with col_reg_c:
                        plot_vertical_bar(
                            region_rep_df,
                            x_col='Region',
                            y_col='Total_Repetitive_Faults',
                            title="Total Repetitive Faults by Region / Zone",
                            color_hex="#E11D48"
                        )
                    
                    with col_reg_t:
                        st.write("**Region-wise Summary Table:**")
                        st.dataframe(
                            region_rep_df.style.format({
                                'Total_Repetitive_Faults': '{:,}',
                                'Affected_Stations': '{:,}'
                            }).background_gradient(subset=['Total_Repetitive_Faults'], cmap='Reds'),
                            use_container_width=True,
                            height=220
                        )
                    st.markdown("---")

                # ---------------------------------------------------------
                # B. Top 20 Stations with Repetitive Faults (Descending Order Explorer)
                # ---------------------------------------------------------
                # Compute total occurrences per station
                stn_totals = repeats.groupby(stn_col)['Occurrences'].sum().reset_index(name='Total_Station_Occurrences')
                stn_totals = stn_totals.sort_values(by='Total_Station_Occurrences', ascending=False)
                
                # Get Top 20 stations
                top_20_stn_ids = stn_totals.head(20)[stn_col].tolist()
                top_20_repeats = repeats[repeats[stn_col].isin(top_20_stn_ids)].copy()
                
                st.write("##### 📁 Top 20 Repetitive Stations Explorer (Descending Order by Occurrences)")
                st.caption("Click on any station below (ranked by total repetitive fault count) to inspect sub-types and ticket records:")
                
                for _, row_stn in stn_totals.head(20).iterrows():
                        stn_id = row_stn[stn_col]
                        stn_group = top_20_repeats[top_20_repeats[stn_col] == stn_id]
                        
                        s_name = ""
                        if stn_name_col and stn_name_col in stn_group.columns:
                            first_val = stn_group[stn_name_col].iloc[0]
                            if pd.notna(first_val) and str(first_val).strip() != "":
                                s_name = str(first_val).strip()
                        
                        total_stn_reps = int(row_stn['Total_Station_Occurrences'])
                        sub_count = len(stn_group)
                        
                        header_title = f"🚉 Station {stn_id}"
                        if s_name:
                            header_title += f" ({s_name})"
                        header_title += f" — {total_stn_reps:,} Repetitive Faults across {sub_count} Sub-Type(s)"
                        
                        with st.expander(header_title, expanded=False):
                            st.write("📌 **Sub-Type Breakdown:**")
                            disp_sub_cols = [c for c in [sub_col, type_col, zme_col, zone_col, 'Occurrences'] if c and c in stn_group.columns]
                            st.dataframe(stn_group[disp_sub_cols].style.format({'Occurrences': '{:,}'}), use_container_width=True)
                            
                            stn_tickets = filtered_issue_df[filtered_issue_df[stn_col].astype(str).str.strip() == str(stn_id).strip()]
                            ocpp_c = find_col(stn_tickets, ['OCPP ID', 'OCPP_ID', 'Charger ID', 'Connector ID', 'EVSE ID'])
                            stat_c = find_col(stn_tickets, ['Status', 'Ticket Status', 'Issue Status'])
                            tat_c = find_col(stn_tickets, ['TAT Compliance', 'SLA Compliance', 'Compliance'])
                            
                            t_cols = [c for c in [ocpp_c, stn_col, sub_col, stat_c, tat_c] if c and c in stn_tickets.columns]
                            if t_cols:
                                st.write("📝 **Underlying Station Ticket Records:**")
                                st.dataframe(stn_tickets[t_cols], use_container_width=True, height=180)
            else:
                st.success("✅ Zero repetitive station faults detected in current period dataset.")
        else:
            st.info("ℹ️ Station ID / Issue Sub-Type columns not found for repetitive fault tracking.")

    # ---------------------------------------------------------
    # TAB 2: PM F-01 DASHBOARD (CHARTS + TABLES SIDE-BY-SIDE)
    # ---------------------------------------------------------
    with tab_pm:
        st.markdown('<div class="section-header">🛠️ Preventive Maintenance (PM F-01) Operational Analytics</div>', unsafe_allow_html=True)

        pm_df = filtered_pm_df.copy()

        pm_st_col = find_col(pm_df, ['PM Status', 'PM_Status', 'Status', 'PM Status (Yes/No)'])
        pm_chg_col = find_col(pm_df, ['Charger ID', 'Charger_ID', 'Charger', 'OCPP ID', 'OCPP_ID'])
        pm_stn_col = find_col(pm_df, ['Station ID', 'Station_ID', 'Station'])
        stn_name_col_pm = find_col(pm_df, ['Station Name', 'Station_Name', 'Site Name'])
        pm_zme_col = find_col(pm_df, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager'])
        pm_zone_col = find_col(pm_df, ['Zone', 'Zone Name', 'Region'])
        adv_col = find_col(pm_df, ['Advance PM Done', 'Advance PM', 'Advance_PM_Done'])

        total_pm_planning = len(pm_df)
        if pm_st_col:
            pm_st_upper = pm_df[pm_st_col].astype(str).str.strip().str.upper()
            pm_done = int((pm_st_upper == 'YES').sum())
            pm_pending = int((pm_st_upper == 'NO').sum())
        else:
            pm_done = int(pm_df['Is_PM_Done'].sum()) if 'Is_PM_Done' in pm_df.columns else 0
            pm_pending = int(pm_df['Is_PM_Pending'].sum()) if 'Is_PM_Pending' in pm_df.columns else 0
        
        if adv_col:
            advance_done = int((pm_df[adv_col].astype(str).str.strip().str.upper() == 'YES').sum())
        elif 'Actual Completion Date Parsed' in pm_df.columns and 'Due Date Parsed' in pm_df.columns:
            advance_done = int(((pm_df['Actual Completion Date Parsed'].notna()) &
                                (pm_df['Due Date Parsed'].notna()) &
                                (pm_df['Actual Completion Date Parsed'] < pm_df['Due Date Parsed'])).sum())
        else:
            advance_done = 0

        pm_eff = (pm_done / total_pm_planning * 100) if total_pm_planning > 0 else 0.0

        # Calculate unique counts of Live Chargers and Live Stations
        total_chargers = int(pm_df[pm_chg_col].dropna().nunique()) if pm_chg_col and pm_chg_col in pm_df.columns else len(pm_df)
        total_stations = int(pm_df[pm_stn_col].dropna().nunique()) if pm_stn_col and pm_stn_col in pm_df.columns else len(pm_df)
        live_stations = total_stations  # PM df is pre-filtered for LIVE stations

        # PM High-Impact KPI Row (6 Columns)
        p1, p2, p3, p4, p5, p6 = st.columns(6)
        with p1:
            st.markdown(f"""
                <div class="metric-card darkred">
                    <div class="metric-label">Live Chargers</div>
                    <div class="metric-val">{total_chargers:,}</div>
                    <div class="metric-sub">Unique Live OCPP IDs</div>
                </div>
            """, unsafe_allow_html=True)
        with p2:
            st.markdown(f"""
                <div class="metric-card grey">
                    <div class="metric-label">Live Stations</div>
                    <div class="metric-val">{live_stations:,}</div>
                    <div class="metric-sub">Unique Live Sites</div>
                </div>
            """, unsafe_allow_html=True)
        with p3:
            st.markdown(f"""
                <div class="metric-card green">
                    <div class="metric-label">PM Planning</div>
                    <div class="metric-val">{total_pm_planning:,}</div>
                    <div class="metric-sub">Scheduled Work Orders</div>
                </div>
            """, unsafe_allow_html=True)
        with p4:
            st.markdown(f"""
                <div class="metric-card green">
                    <div class="metric-label">PM Done</div>
                    <div class="metric-val">{pm_done:,}</div>
                    <div class="metric-sub">Verified & Completed</div>
                </div>
            """, unsafe_allow_html=True)
        with p5:
            st.markdown(f"""
                <div class="metric-card red">
                    <div class="metric-label">PM Pending</div>
                    <div class="metric-val">{pm_pending:,}</div>
                    <div class="metric-sub">Scheduled Pending</div>
                </div>
            """, unsafe_allow_html=True)
        with p6:
            st.markdown(f"""
                <div class="metric-card {'green' if pm_eff >= 90 else 'amber'}">
                    <div class="metric-label">PM Efficiency</div>
                    <div class="metric-val">{pm_eff:.1f}%</div>
                    <div class="metric-sub">Target Benchmark: ≥ 90.0%</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 1. PM Execution Status Overview (Vertical Bar Chart replacing Pie Chart)
        st.markdown('<div class="section-header">📊 1. PM Execution Status Overview</div>', unsafe_allow_html=True)
        pm_exec_df = pd.DataFrame({
            'Metric': ['PM Planning', 'PM Done', 'PM Pending', 'Advance PM Done'],
            'Count': [total_pm_planning, pm_done, pm_pending, advance_done]
        })

        c_exec_c, c_exec_t = st.columns([7, 5])
        with c_exec_c:
            plot_vertical_bar(
                pm_exec_df,
                x_col='Metric',
                y_col='Count',
                title="PM Work Orders Execution Distribution",
                color_hex="#991B1B"
            )

        with c_exec_t:
            st.write("##### PM Execution Summary Table")
            pm_exec_df['Share %'] = (pm_exec_df['Count'] / total_pm_planning * 100).round(1) if total_pm_planning > 0 else 0.0
            st.dataframe(
                pm_exec_df.style.format({
                    'Count': '{:,}',
                    'Share %': '{:.1f}%'
                }).background_gradient(subset=['Count'], cmap='Reds'),
                use_container_width=True,
                height=250
            )

        st.markdown("---")

        # 2. Zone-wise PM Planning vs Actual PM Done & Advance PM Done (Grouped Bar Chart)
        st.markdown('<div class="section-header">🏢 2. Zone-wise PM Planning vs Actual PM Done & Advance PM Done</div>', unsafe_allow_html=True)
        pm_zone_col = find_col(pm_df, ['Zone', 'Zone Name', 'Region'])

        if pm_zone_col and pm_zone_col in pm_df.columns:
            zone_summary_list = []
            for z_val, group in pm_df.groupby(pm_zone_col):
                z_str = str(z_val).strip()
                z_plan = len(group)
                if pm_st_col and pm_st_col in group.columns:
                    st_u = group[pm_st_col].astype(str).str.strip().str.upper()
                    z_done = int((st_u == 'YES').sum())
                    z_pend = int((st_u == 'NO').sum())
                else:
                    z_done = int(group['Is_PM_Done'].sum()) if 'Is_PM_Done' in group.columns else 0
                    z_pend = int(group['Is_PM_Pending'].sum()) if 'Is_PM_Pending' in group.columns else 0

                if 'Advance PM Done' in group.columns:
                    z_adv = int((group['Advance PM Done'].astype(str).str.strip().str.upper() == 'YES').sum())
                else:
                    z_adv = 0

                z_eff = (z_done / z_plan * 100) if z_plan > 0 else 0.0

                zone_summary_list.append({
                    'Zone': z_str,
                    'PM Planning': z_plan,
                    'Actual PM Done': z_done,
                    'Advance PM Done': z_adv,
                    'PM Pending': z_pend,
                    'PM Efficiency (%)': round(z_eff, 1)
                })

            zone_pm_df = pd.DataFrame(zone_summary_list).sort_values(by='PM Planning', ascending=False)

            c_zone_c, c_zone_t = st.columns([6, 6])
            with c_zone_c:
                plot_grouped_bar(
                    df=zone_pm_df,
                    x_col='Zone',
                    y_cols=['PM Planning', 'Actual PM Done', 'Advance PM Done'],
                    title="PM Planning vs Actual PM Done vs Advance PM Done by Zone",
                    colors=['#991B1B', '#16A34A', '#EF4444']
                )

            with c_zone_t:
                st.write("##### Zone-wise PM Summary Table")
                st.dataframe(
                    zone_pm_df.style.format({
                        'PM Planning': '{:,}',
                        'Actual PM Done': '{:,}',
                        'Advance PM Done': '{:,}',
                        'PM Pending': '{:,}',
                        'PM Efficiency (%)': '{:.1f}%'
                    }).background_gradient(subset=['PM Efficiency (%)'], cmap='Reds'),
                    use_container_width=True,
                    height=280
                )

        # 3. B2B / B2C Customer Segment Breakdown
        pm_seg_col_tab = find_col(pm_df, ['B2B/ B2C', 'B2B/B2C', 'Segment', 'Customer Segment'])
        if pm_seg_col_tab and pm_seg_col_tab in pm_df.columns:
            st.markdown("---")
            st.markdown('<div class="section-header">💼 3. Customer Segment (B2B vs B2C) PM Performance Breakdown</div>', unsafe_allow_html=True)
            col_seg_c, col_seg_t = st.columns([5, 7])

            seg_summary_list = []
            for seg_val, group in pm_df.groupby(pm_seg_col_tab):
                seg_str = str(seg_val).strip()
                s_chg = int(group[pm_chg_col].dropna().nunique()) if pm_chg_col and pm_chg_col in group.columns else len(group)
                s_stn = int(group[pm_stn_col].dropna().nunique()) if pm_stn_col and pm_stn_col in group.columns else len(group)
                s_plan = len(group)
                if pm_st_col and pm_st_col in group.columns:
                    st_u = group[pm_st_col].astype(str).str.strip().str.upper()
                    s_done = int((st_u == 'YES').sum())
                    s_pend = int((st_u == 'NO').sum())
                else:
                    s_done = int(group['Is_PM_Done'].sum()) if 'Is_PM_Done' in group.columns else 0
                    s_pend = int(group['Is_PM_Pending'].sum()) if 'Is_PM_Pending' in group.columns else 0

                s_eff = (s_done / s_plan * 100) if s_plan > 0 else 0.0
                seg_summary_list.append({
                    'Customer Segment': seg_str,
                    'Live Chargers': s_chg,
                    'Live Stations': s_stn,
                    'PM Planning': s_plan,
                    'PM Done': s_done,
                    'PM Pending': s_pend,
                    'PM Efficiency (%)': round(s_eff, 1)
                })

            seg_df = pd.DataFrame(seg_summary_list).sort_values(by='PM Planning', ascending=False)

            with col_seg_c:
                plot_pie_chart(
                    labels=seg_df['Customer Segment'].tolist(),
                    values=seg_df['PM Planning'].tolist(),
                    title="PM Work Orders by Customer Segment (B2B vs B2C)",
                    colors=['#991B1B', '#DC2626', '#EF4444', '#7F1D1D']
                )

            with col_seg_t:
                st.write("##### B2B / B2C Customer Segment Summary Data Table")
                st.dataframe(
                    seg_df.style.format({
                        'Live Chargers': '{:,}',
                        'Live Stations': '{:,}',
                        'PM Planning': '{:,}',
                        'PM Done': '{:,}',
                        'PM Pending': '{:,}',
                        'PM Efficiency (%)': '{:.1f}%'
                    }).background_gradient(subset=['PM Efficiency (%)'], cmap='Reds'),
                    use_container_width=True,
                    height=240
                )

        st.markdown("---")

        # 4. Detailed PM F-01 Breakdown Data Table (Full-Width)
        st.markdown('<div class="section-header">⚙️ 4. Detailed PM F-01 Summary Data Table</div>', unsafe_allow_html=True)
        if pm_zme_col:
            group_keys = [pm_zme_col]
            if pm_zone_col:
                group_keys.append(pm_zone_col)

            pm_summary_list = []
            for name_tuple, group in pm_df.groupby(group_keys):
                zme_name = name_tuple[0] if isinstance(name_tuple, tuple) else name_tuple
                zone_val = name_tuple[1] if isinstance(name_tuple, tuple) and len(name_tuple) > 1 else (group[pm_zone_col].iloc[0] if pm_zone_col and pm_zone_col in group.columns else 'N/A')

                chargers_cnt = int(group[pm_chg_col].dropna().nunique()) if pm_chg_col and pm_chg_col in group.columns else len(group)
                stations_cnt = int(group[pm_stn_col].dropna().nunique()) if pm_stn_col and pm_stn_col in group.columns else len(group)
                planning_cnt = len(group)

                if pm_st_col and pm_st_col in group.columns:
                    st_u = group[pm_st_col].astype(str).str.strip().str.upper()
                    done_cnt = int((st_u == 'YES').sum())
                    pending_cnt = int((st_u == 'NO').sum())
                else:
                    done_cnt = int(group['Is_PM_Done'].sum()) if 'Is_PM_Done' in group.columns else 0
                    pending_cnt = int(group['Is_PM_Pending'].sum()) if 'Is_PM_Pending' in group.columns else 0

                if adv_col and adv_col in group.columns:
                    adv_cnt = int((group[adv_col].astype(str).str.strip().str.upper() == 'YES').sum())
                else:
                    adv_cnt = 0

                eff_val = (done_cnt / planning_cnt * 100) if planning_cnt > 0 else 0.0

                pm_summary_list.append({
                    'ZME Name': zme_name,
                    'Zone': zone_val,
                    'Total Chargers': chargers_cnt,
                    'Total Stations': stations_cnt,
                    'PM Planning': planning_cnt,
                    'PM Done': done_cnt,
                    'PM Pending': pending_cnt,
                    'Advance PM Done': adv_cnt,
                    'PM Efficiency (%)': round(eff_val, 1)
                })

            pm_summary_df = pd.DataFrame(pm_summary_list).sort_values(by='PM Planning', ascending=False)
            st.dataframe(
                pm_summary_df.style.format({
                    'Total Chargers': '{:,}',
                    'Total Stations': '{:,}',
                    'PM Planning': '{:,}',
                    'PM Done': '{:,}',
                    'PM Pending': '{:,}',
                    'Advance PM Done': '{:,}',
                    'PM Efficiency (%)': '{:.1f}%'
                }).background_gradient(subset=['PM Efficiency (%)'], cmap='Reds'),
                use_container_width=True,
                height=320
            )

        st.markdown("---")

        # 5. Count number of stations PM scheduled on basis of selection of Date, Month, Quarter, Year
        st.markdown('<div class="section-header">🗓️ 5. PM Scheduled Station Breakdown by Selection (Date, Month, Quarter, Year)</div>', unsafe_allow_html=True)
        period_choice = st.radio(
            "Group PM Scheduled Stations by Period:",
            ["Month", "Quarter", "Year", "Date"],
            horizontal=True,
            key="pm_period_choice"
        )
        period_col_map = {
            "Month": "Scheduled Month",
            "Quarter": "Scheduled Quarter",
            "Year": "Scheduled Year",
            "Date": "Scheduled Date"
        }
        target_period_col = period_col_map[period_choice]

        # Resolve period series with fallback
        if target_period_col in pm_df.columns and not pm_df[target_period_col].dropna().empty:
            period_series = pm_df[target_period_col]
        else:
            if period_choice == "Month":
                m_col = find_col(pm_df, ['Month', 'PM Month', 'Scheduled Month', 'Due Date Parsed'])
                period_series = pm_df[m_col].astype(str) if m_col and m_col in pm_df.columns else pd.Series(['General / All-Period'] * len(pm_df))
            elif period_choice == "Quarter":
                q_col = find_col(pm_df, ['Quarter', 'PM Quarter', 'Scheduled Quarter'])
                period_series = pm_df[q_col].astype(str) if q_col and q_col in pm_df.columns else pd.Series(['General / All-Period'] * len(pm_df))
            elif period_choice == "Year":
                y_col = find_col(pm_df, ['Year', 'PM Year', 'FY', 'Financial Year', 'Scheduled Year'])
                period_series = pm_df[y_col].astype(str) if y_col and y_col in pm_df.columns else pd.Series(['General / All-Period'] * len(pm_df))
            else:
                d_col = find_col(pm_df, ['Date', 'Due Date', 'PM Date', 'Scheduled Date'])
                period_series = pm_df[d_col].astype(str) if d_col and d_col in pm_df.columns else pd.Series(['General / All-Period'] * len(pm_df))

        pm_df_temp = pm_df.copy()
        pm_df_temp['_Group_Period_'] = period_series.astype(str).str.strip().replace({'nan': 'Unscheduled / General', 'None': 'Unscheduled / General', 'NaT': 'Unscheduled / General', '<NA>': 'Unscheduled / General', '': 'Unscheduled / General'})

        period_summary_list = []
        for p_val, group in pm_df_temp.groupby('_Group_Period_'):
            p_str = str(p_val)
            p_chargers = len(group[pm_chg_col].dropna().unique()) if pm_chg_col and pm_chg_col in group.columns else len(group)
            p_stations = len(group[pm_stn_col].dropna().unique()) if pm_stn_col and pm_stn_col in group.columns else len(group)
            p_planning = len(group)

            if pm_st_col and pm_st_col in group.columns:
                st_u = group[pm_st_col].astype(str).str.strip().str.upper()
                p_done = int((st_u == 'YES').sum())
                p_pending = int((st_u == 'NO').sum())
            else:
                p_done = int(group['Is_PM_Done'].sum()) if 'Is_PM_Done' in group.columns else 0
                p_pending = int(group['Is_PM_Pending'].sum()) if 'Is_PM_Pending' in group.columns else 0

            p_eff = (p_done / p_planning * 100) if p_planning > 0 else 0.0

            period_summary_list.append({
                f'Period ({period_choice})': p_str,
                'Stations Scheduled': p_stations,
                'Chargers Scheduled': p_chargers,
                'PM Planning': p_planning,
                'PM Done': p_done,
                'PM Pending': p_pending,
                'PM Efficiency (%)': round(p_eff, 1)
            })

        period_summary_df = pd.DataFrame(period_summary_list)

        c_per_chart, c_per_tbl = st.columns([6, 6])
        with c_per_chart:
            plot_grouped_bar(
                df=period_summary_df,
                x_col=f'Period ({period_choice})',
                y_cols=['Stations Scheduled', 'PM Done', 'PM Pending'],
                title=f"Scheduled Stations vs Completion by {period_choice}",
                colors=['#991B1B', '#DC2626', '#EF4444']
            )

        with c_per_tbl:
            st.write(f"##### Scheduled Stations Data Table ({period_choice} Level)")
            st.dataframe(
                period_summary_df.style.format({
                    'Stations Scheduled': '{:,}',
                    'Chargers Scheduled': '{:,}',
                    'PM Planning': '{:,}',
                    'PM Done': '{:,}',
                    'PM Pending': '{:,}',
                    'PM Efficiency (%)': '{:.1f}%'
                }).background_gradient(subset=['PM Efficiency (%)'], cmap='Reds'),
                use_container_width=True,
                height=280
            )

        st.markdown("---")

        # 6. Check if PM of specific OCPP ID has been done as per Scheduled Month, Quarter, Year, Date
        st.markdown('<div class="section-header">🔌 6. OCPP ID / Charger PM Schedule & Compliance Status Checker</div>', unsafe_allow_html=True)
        ocpp_col = find_col(pm_df, ['OCPP ID', 'OCPP_ID', 'Charger ID', 'Charger_ID', 'Charger', 'EVSE ID', 'Connector ID'])

        if ocpp_col and ocpp_col in pm_df.columns:
            all_ocpps = sorted([str(x) for x in pm_df[ocpp_col].dropna().unique().tolist()])
            selected_ocpp = st.selectbox("Search or Select OCPP ID / Charger ID to Check Compliance:", ["All OCPP IDs"] + all_ocpps)

            if selected_ocpp != "All OCPP IDs":
                ocpp_records = pm_df[pm_df[ocpp_col].astype(str) == str(selected_ocpp)]
            else:
                ocpp_records = pm_df.copy()

            # Build clean display table
            disp_ocpp_cols = [ocpp_col]
            if pm_stn_col and pm_stn_col in ocpp_records.columns:
                disp_ocpp_cols.append(pm_stn_col)
            if stn_name_col_pm and stn_name_col_pm in ocpp_records.columns and stn_name_col_pm != pm_stn_col:
                disp_ocpp_cols.append(stn_name_col_pm)
            if pm_zme_col and pm_zme_col in ocpp_records.columns:
                disp_ocpp_cols.append(pm_zme_col)

            for p_col in ['Scheduled Date', 'Scheduled Month', 'Scheduled Quarter', 'Scheduled Year', 'Due Date', 'Actual Completion Date', 'PM Compliance Status']:
                if p_col in ocpp_records.columns:
                    disp_ocpp_cols.append(p_col)

            disp_ocpp_df = ocpp_records[[c for c in disp_ocpp_cols if c in ocpp_records.columns]].drop_duplicates()

            st.write(f"##### Compliance Check Results ({len(disp_ocpp_df):,} Record(s))")
            st.dataframe(disp_ocpp_df, use_container_width=True, height=300)
        else:
            st.info("ℹ️ OCPP ID / Charger ID column not found in dataset for compliance verification.")

    # ---------------------------------------------------------
    # TAB 3: DATA EXPLORER
    # ---------------------------------------------------------
    with tab_raw:
        st.markdown('<div class="section-header">🔍 Data Explorer</div>', unsafe_allow_html=True)
        data_choice = st.radio("Select Dataset:", ["Issue Tracker Dataset", "PM Tracker Dataset"], horizontal=True)
        exclude_cols = {'Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed', 'Is_TAT_Compliant', 'Is_TAT_Breached', 'Is_PM_Done', 'Is_PM_Pending'}

        if data_choice == "Issue Tracker Dataset":
            df_disp = ensure_unique_columns(filtered_issue_df.drop(columns=[c for c in exclude_cols if c in filtered_issue_df.columns]))
            st.dataframe(df_disp, use_container_width=True)
        else:
            df_disp = ensure_unique_columns(filtered_pm_df.drop(columns=[c for c in exclude_cols if c in filtered_pm_df.columns]))
            st.dataframe(df_disp, use_container_width=True)


if __name__ == '__main__':
    # CLI execution support
    if len(sys.argv) >= 2 and not sys.argv[1].startswith('-'):
        if len(sys.argv) == 2 or (len(sys.argv) >= 3 and sys.argv[2].endswith('.xlsx') and not sys.argv[1].endswith('.xlsx')):
            src_path = sys.argv[1]
            out_path = sys.argv[2] if len(sys.argv) > 2 else f"MPR_Report_{datetime.now():%Y-%m}.xlsx"
            wb = generate_workbook_cached(src_path)
            with open(out_path, 'wb') as f:
                f.write(wb)
            print(f'Written: {out_path}')
        else:
            issue_path = sys.argv[1]
            pm_path = sys.argv[2]
            out_path = sys.argv[3] if len(sys.argv) > 3 else f"MPR_Report_{datetime.now():%Y-%m}.xlsx"
            wb = generate_workbook_cached(issue_path, pm_path)
            with open(out_path, 'wb') as f:
                f.write(wb)
            print(f'Written: {out_path}')
    else:
        run_streamlit_app()
