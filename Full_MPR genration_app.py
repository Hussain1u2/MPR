import io
import sys
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl import Workbook
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
            rows = [tuple(cell for cell in row) for row in ws.iter_rows(values_only=True)]
            wb.close()
            if not rows:
                return pd.DataFrame()
            headers = make_unique_headers([str(h).strip() if h is not None else f"Unnamed_{i+1}" for i, h in enumerate(rows[0])])
            df = pd.DataFrame([r[:len(headers)] for r in rows[1:]], columns=headers).dropna(how='all')
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
        rows = [tuple(cell for cell in row) for row in ws.iter_rows(values_only=True)]
        if hasattr(wb, 'close'):
            wb.close()
        if not rows:
            return pd.DataFrame()
        headers = make_unique_headers([str(h).strip() if h is not None else f"Unnamed_{i+1}" for i, h in enumerate(rows[0])])
        df = pd.DataFrame([r[:len(headers)] for r in rows[1:]], columns=headers).dropna(how='all')

    df = ensure_unique_columns(df)
    if 'Issue Date' in df.columns:
        df['Issue Date Parsed'] = pd.to_datetime(df['Issue Date'], errors='coerce')
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
            return df
        sheet_name = select_sheet_name(wb.sheetnames, preferred_sheets, 'pm')
        ws = wb[sheet_name]
        rows = [tuple(cell for cell in row) for row in ws.iter_rows(values_only=True)]
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
        rows = [tuple(cell for cell in row) for row in ws.iter_rows(values_only=True)]
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
    if 'Due Date' in df.columns:
        df['Due Date Parsed'] = pd.to_datetime(df['Due Date'], errors='coerce')
    if 'Actual Completion Date' in df.columns:
        df['Actual Completion Date Parsed'] = pd.to_datetime(df['Actual Completion Date'], errors='coerce')
    return df


def write_data_sheet(wb, name, df, table_name, date_cols):
    ws = wb.create_sheet(name)
    clean_df = df.drop(columns=[c for c in ['Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed'] if c in df.columns])
    for c, col in enumerate(clean_df.columns, start=1):
        cell = ws.cell(row=1, column=c, value=col)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for r, rec in enumerate(clean_df.itertuples(index=False), start=2):
        for c, val in enumerate(rec, start=1):
            col_name = clean_df.columns[c - 1]
            if pd.isna(val):
                val = None
            elif isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            cell = ws.cell(row=r, column=c, value=val)
            cell.font, cell.border = CELL_FONT, BORDER
            if col_name in date_cols and val is not None:
                cell.number_format = DATE_FMT

    for c, col in enumerate(clean_df.columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(12, min(28, len(str(col)) + 2))

    last_row = len(clean_df) + 1
    table = Table(displayName=table_name, ref=f"A1:{get_column_letter(len(clean_df.columns))}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9', showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = 'A2'
    return ws, last_row


def add_pm_helper_columns(ws, pm_df, pcol, last_row):
    due_col, done_col = pcol.get('Due Date', 'G'), pcol.get('Actual Completion Date', 'J')
    adv_idx = len(pm_df.columns) - (1 if 'Due Date Parsed' in pm_df.columns else 0) - (1 if 'Actual Completion Date Parsed' in pm_df.columns else 0) + 1
    adv_letter = get_column_letter(adv_idx)
    ws.cell(row=1, column=adv_idx, value='Advance PM Done').font = HEADER_FONT
    ws.cell(row=1, column=adv_idx).fill = HEADER_FILL
    for r in range(2, last_row + 1):
        formula = (f'=IF(AND(${done_col}{r}<>"",${due_col}{r}<>"",${done_col}{r}<${due_col}{r}),"Yes",'
                   f'IF(${done_col}{r}<>"","No",""))')
        cell = ws.cell(row=r, column=adv_idx, value=formula)
        cell.font, cell.border = CELL_FONT, BORDER
    ws.column_dimensions[adv_letter].width = 16

    zme_col, station_col = pcol.get('ZME', 'A'), pcol.get('Station ID', 'B')
    occ_idx = adv_idx + 1
    occ_letter = get_column_letter(occ_idx)
    ws.cell(row=1, column=occ_idx, value='First Station Occurrence').font = HEADER_FONT
    ws.cell(row=1, column=occ_idx).fill = HEADER_FILL
    for r in range(2, last_row + 1):
        formula = (f'=IF(COUNTIFS(${zme_col}$2:${zme_col}{r},${zme_col}{r},'
                   f'${station_col}$2:${station_col}{r},${station_col}{r})=1,1,0)')
        cell = ws.cell(row=r, column=occ_idx, value=formula)
        cell.font, cell.border = CELL_FONT, BORDER
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
    ws['A2'] = ('Source: Issue Data tab (live formulas). "Within TAT" = TAT Compliance = Yes, '
                '"Without TAT" = TAT Compliance = No.')
    ws['A2'].font = NOTE_FONT
    ws.merge_cells('A2:H2')

    row = 4
    row = section_title(ws, row, '1. Issue Summary by ZME', 5)
    row = header_row(ws, row, ['ZME Name', 'Total Issues', 'Within TAT', 'Without TAT', 'TAT Efficiency'])
    if 'ZME' in issue_df.columns:
        for zme in sorted(issue_df['ZME'].dropna().unique()):
            total = f'=COUNTIFS({irange("ZME")},A{row})'
            within = f'=COUNTIFS({irange("ZME")},A{row},{irange("TAT Compliance")},"Yes")'
            without = f'=COUNTIFS({irange("ZME")},A{row},{irange("TAT Compliance")},"No")'
            eff = f'=IFERROR(C{row}/B{row},0)'
            row = data_row(ws, row, [zme, total, within, without, eff], pct_cols={4})

    row += 1
    row = section_title(ws, row, '2. Issue Summary by Zone (CM Efficiency)', 4)
    row = header_row(ws, row, ['Zone', 'Total Issues', 'CM Efficiency (Within TAT)', 'CM Efficiency (Without TAT)'])
    if 'Zone' in issue_df.columns:
        for zone in sorted(issue_df['Zone'].dropna().unique()):
            total = f'=COUNTIFS({irange("Zone")},A{row})'
            within = f'=IFERROR(COUNTIFS({irange("Zone")},A{row},{irange("TAT Compliance")},"Yes")/B{row},0)'
            without = f'=IFERROR(COUNTIFS({irange("Zone")},A{row},{irange("TAT Compliance")},"No")/B{row},0)'
            row = data_row(ws, row, [zone, total, within, without], pct_cols={2, 3})

    row += 1
    row = section_title(ws, row, '3. Repetitive Faults (same Station ID + Issue Sub-Type, 2+ occurrences)', 3)
    row = header_row(ws, row, ['Station ID', 'Issue Sub-Type', 'Occurrences'])
    if 'Station ID' in issue_df.columns and 'Issue Sub-Type' in issue_df.columns:
        pair_counts = issue_df.groupby(['Station ID', 'Issue Sub-Type']).size()
        repeats = pair_counts[pair_counts >= 2].index.tolist()
        if repeats:
            for station_id, subtype in repeats:
                cnt = f'=COUNTIFS({irange("Station ID")},A{row},{irange("Issue Sub-Type")},B{row})'
                row = data_row(ws, row, [station_id, subtype, cnt])
        else:
            row = data_row(ws, row, ['None found in current data', '', ''])

    row += 1
    row = section_title(ws, row, '4. Status Breakdown', 2)
    row = header_row(ws, row, ['Status', 'Count'])
    if 'Status' in issue_df.columns:
        for status in sorted(issue_df['Status'].dropna().unique()):
            row = data_row(ws, row, [status, f'=COUNTIFS({irange("Status")},A{row})'])

    row += 1
    row = section_title(ws, row, '5. Severity Breakdown', 2)
    row = header_row(ws, row, ['Severity', 'Count'])
    if 'Severity' in issue_df.columns:
        for sev in sorted(issue_df['Severity'].dropna().unique()):
            row = data_row(ws, row, [sev, f'=COUNTIFS({irange("Severity")},A{row})'])

    row += 1
    row = section_title(ws, row, '6. Customer Filter (B2B / B2C)', 2)
    row = header_row(ws, row, ['Segment', 'Count'])
    if 'B2B/ B2C' in issue_df.columns:
        for seg in sorted(issue_df['B2B/ B2C'].dropna().unique()):
            row = data_row(ws, row, [seg, f'=COUNTIFS({irange("B2B/ B2C")},A{row})'])

    row += 1
    ws.cell(row=row, column=1,
            value=('Tip: use the dropdown filter arrows on the "Issue Data" tab to slice by '
                   'customer, status, severity, zone, or ZME.')).font = NOTE_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)

    for col, width in zip('ABCDEFGH', [26, 20, 16, 22, 26, 14]):
        ws.column_dimensions[col].width = width


def build_pm_dashboard(wb, pm_df, prange):
    ws = wb.create_sheet('Dashboard - PM F-01', 1)
    ws.sheet_view.showGridLines = False
    ws['A1'] = 'PM F-01 — PREVENTIVE MAINTENANCE DASHBOARD (YTD, FY2627 Q1 + Q2)'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:J1')
    ws['A2'] = ('Scope: current fiscal year to date (Apr-26 to Sep-26). "PM Planning" = monthly PM instances '
                'scheduled (non-blank PM Status), "PM Done" = PM Status = Yes, "PM Pending" = scheduled but not '
                'done (PM Status = No), "Advance PM Done" = completed before the scheduled month started.')
    ws['A2'].font = NOTE_FONT
    ws.merge_cells('A2:J2')

    row = 4
    if 'ZME' in pm_df.columns and 'Zone' in pm_df.columns:
        zme_zone_pairs = pm_df[['ZME', 'Zone']].drop_duplicates().sort_values(['Zone', 'ZME']).values.tolist()
        row = section_title(ws, row, 'PM Summary by ZME', 9)
        row = header_row(ws, row, ['ZME Name', 'Zone', 'Total Chargers', 'Total Stations', 'PM Planning',
                                    'PM Done', 'PM Pending', 'Advance PM Done', 'PM Efficiency'])
        for zme, zone in zme_zone_pairs:
            total_chargers = f'=COUNTIFS({prange("ZME")},A{row},{prange("Due Date")},DATE(2026,4,1))'
            total_stations = (f'=SUMIFS({prange("First Station Occurrence")},{prange("ZME")},A{row},'
                               f'{prange("Due Date")},DATE(2026,4,1))')
            pm_planning = f'=COUNTIFS({prange("ZME")},A{row},{prange("PM Status")},"<>")'
            pm_done = f'=COUNTIFS({prange("ZME")},A{row},{prange("PM Status")},"Yes")'
            pm_pending = f'=E{row}-F{row}'
            advance_done = f'=COUNTIFS({prange("ZME")},A{row},{prange("Advance PM Done")},"Yes")'
            pm_eff = f'=IFERROR(F{row}/E{row},0)'
            row = data_row(ws, row, [zme, zone, total_chargers, total_stations, pm_planning, pm_done,
                                      pm_pending, advance_done, pm_eff], pct_cols={8})

    for col, width in zip('ABCDEFGHIJ', [18, 10, 14, 14, 13, 11, 12, 16, 14, 10]):
        ws.column_dimensions[col].width = width


@st.cache_data(show_spinner=False)
def generate_workbook_cached(source_bytes_or_path, pm_bytes_or_path=None):
    """Cached workbook & DataFrame generator for instant responsive updates."""
    if pm_bytes_or_path is None:
        issue_df = load_issue_tracker(source_bytes_or_path)
        pm_df = load_pm_tracker(source_bytes_or_path)
    else:
        issue_df = load_issue_tracker(source_bytes_or_path)
        pm_df = load_pm_tracker(pm_bytes_or_path)

    wb = Workbook()
    wb.remove(wb.active)

    issue_ws, issue_last = write_data_sheet(
        wb, 'Issue Data', issue_df, 'IssueTable',
        date_cols=['Issue Date', 'Resolution Date', 'Restoration Date'])
    clean_issue_cols = [c for c in issue_df.columns if c not in ['Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed']]
    icol = {name: get_column_letter(i + 1) for i, name in enumerate(clean_issue_cols)}

    pm_ws, pm_last = write_data_sheet(
        wb, 'PM Data', pm_df, 'PMTable',
        date_cols=['Go Live Date', 'Due Date', 'Actual Completion Date'])
    clean_pm_cols = [c for c in pm_df.columns if c not in ['Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed']]
    pcol = {name: get_column_letter(i + 1) for i, name in enumerate(clean_pm_cols)}
    add_pm_helper_columns(pm_ws, pm_df, pcol, pm_last)
    pcol_full = {name: get_column_letter(i + 1)
                 for i, name in enumerate(list(clean_pm_cols) + ['Advance PM Done', 'First Station Occurrence'])}

    def irange(col):
        col_letter = icol.get(col, 'A')
        return f"'Issue Data'!${col_letter}$2:${col_letter}${ISSUE_RANGE_END}"

    def prange(col):
        col_letter = pcol_full.get(col, 'A')
        return f"'PM Data'!${col_letter}$2:${col_letter}${PM_RANGE_END}"

    build_issue_dashboard(wb, issue_df, irange)
    build_pm_dashboard(wb, pm_df, prange)
    wb.active = 0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), issue_df, pm_df


def plot_pie_chart(labels, values, title, colors=None, hole=0.45):
    """Renders a responsive, high-contrast Donut/Pie Chart."""
    dark_mode = st.session_state.get('dark_mode', False)
    text_color = '#ffffff' if dark_mode else '#0F172A'
    legend_color = '#cbd5e1' if dark_mode else '#475569'

    if HAS_PLOTLY:
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=hole,
            marker_colors=colors if colors else ['#10b981', '#DC2626', '#475569', '#f59e0b', '#991B1B'],
            textinfo='percent+value',
            hoverinfo='label+percent+value',
            insidetextfont=dict(color='#ffffff', size=13, family='Outfit')
        )])
        fig.update_layout(
            title=dict(text=title, font=dict(size=15, color=text_color, family='Outfit', weight='600')),
            margin=dict(l=15, r=15, t=45, b=15),
            height=300,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5, font=dict(color=legend_color)),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False, 'responsive': True})
    else:
        df_pie = pd.DataFrame({'Category': labels, 'Value': values})
        st.write(f"**{title}**")
        st.dataframe(df_pie, use_container_width=True)


def plot_vertical_bar(df, x_col, y_col, title, color_hex="#DC2626"):
    """Renders a clean vertical bar chart with exact value labels."""
    dark_mode = st.session_state.get('dark_mode', False)
    text_color = '#f8fafc' if dark_mode else '#0F172A'
    grid_color = 'rgba(255,255,255,0.05)' if dark_mode else '#F1F5F9'
    label_color = '#cbd5e1' if dark_mode else '#475569'

    if HAS_PLOTLY:
        fig = px.bar(
            df,
            x=x_col,
            y=y_col,
            title=title,
            text=y_col,
            color_discrete_sequence=[color_hex]
        )
        fig.update_traces(texttemplate='%{text}', textposition='outside', textfont=dict(size=12, family='Outfit', color=label_color))
        fig.update_layout(
            margin=dict(l=15, r=15, t=45, b=15),
            height=320,
            font=dict(family='Outfit', color=text_color),
            xaxis_title=x_col,
            yaxis_title=y_col,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            yaxis=dict(showgrid=True, gridcolor=grid_color)
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False, 'responsive': True})
    else:
        st.write(f"**{title}**")
        st.bar_chart(df.set_index(x_col)[y_col], height=280)


def plot_grouped_bar(df, x_col, y_cols, title, colors=None):
    """Renders a responsive multi-series grouped column chart."""
    dark_mode = st.session_state.get('dark_mode', False)
    text_color = '#f8fafc' if dark_mode else '#0F172A'
    grid_color = 'rgba(255,255,255,0.05)' if dark_mode else '#F1F5F9'
    legend_color = '#cbd5e1' if dark_mode else '#475569'

    if HAS_PLOTLY:
        fig = go.Figure()
        palette = colors if colors else ['#DC2626', '#991B1B', '#475569', '#10b981']
        for idx, col in enumerate(y_cols):
            fig.add_trace(go.Bar(
                name=col,
                x=df[x_col],
                y=df[col],
                marker_color=palette[idx % len(palette)],
                text=df[col],
                textposition='auto',
                textfont=dict(size=11, family='Outfit', color='#ffffff')
            ))
        fig.update_layout(
            barmode='group',
            title=dict(text=title, font=dict(size=15, color=text_color, family='Outfit', weight='600')),
            margin=dict(l=15, r=15, t=45, b=15),
            height=320,
            legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5, font=dict(color=legend_color)),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            yaxis=dict(showgrid=True, gridcolor=grid_color)
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False, 'responsive': True})
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

    st.sidebar.markdown("### ⚙️ Theme Settings")
    dark_mode = st.sidebar.toggle("🌙 Dark Mode", value=st.session_state.get('dark_mode', False))
    st.session_state['dark_mode'] = dark_mode
    st.sidebar.markdown("---")

    if dark_mode:
        bg_main = "#0f172a"
        bg_sidebar = "#020617"
        text_color = "#f8fafc"
        text_sub = "#94a3b8"
        card_bg = "rgba(30, 41, 59, 0.5)"
        card_hover = "rgba(30, 41, 59, 0.8)"
        border_color = "rgba(255, 255, 255, 0.05)"
    else:
        bg_main = "#FFFFFF"
        bg_sidebar = "#F8FAFC"
        text_color = "#0F172A"
        text_sub = "#64748B"
        card_bg = "#FFFFFF"
        card_hover = "#FFFFFF"
        border_color = "#E2E8F0"

    # Red & White theme CSS (with dynamic dark/light)
    css = f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800;900&display=swap');
        
        .stApp, [data-testid="stMain"], [data-testid="stHeader"] {{
            background-color: {bg_main} !important;
        }}

        [data-testid="stSidebar"] {{
            background-color: {bg_sidebar} !important;
            border-right: 1px solid {border_color} !important;
        }}

        html, body, [class*="css"], p, span, label, div {{
            font-family: 'Outfit', -apple-system, BlinkMacSystemFont, sans-serif;
            color: {text_color};
        }}

        /* Red Gradient Buttons */
        .stButton > button, div[data-testid="stDownloadButton"] > button {{
            background: linear-gradient(135deg, #991B1B 0%, #DC2626 100%) !important;
            color: #ffffff !important;
            border-radius: 10px !important;
            font-weight: 600 !important;
            font-size: 1rem !important;
            border: none !important;
            padding: 0.65rem 1.4rem !important;
            box-shadow: 0 4px 15px rgba(220, 38, 38, 0.3) !important;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
            letter-spacing: 0.02em !important;
        }}
        .stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {{
            box-shadow: 0 6px 20px rgba(220, 38, 38, 0.5) !important;
            transform: translateY(-2px) !important;
        }}

        /* Streamlit Tabs */
        button[data-baseweb="tab"] {{
            font-weight: 600 !important;
            color: {text_sub} !important;
            font-size: 1.05rem !important;
            padding: 0.75rem 1.25rem !important;
            transition: color 0.3s ease !important;
        }}
        button[data-baseweb="tab"][aria-selected="true"] {{
            color: #DC2626 !important;
            border-bottom-color: #DC2626 !important;
        }}
        div[data-baseweb="tab-highlight"] {{
            background-color: #DC2626 !important;
            height: 3px !important;
            border-radius: 3px 3px 0 0 !important;
        }}

        /* Header Box - Always Red Gradient for brand */
        .exec-header-box {{
            background: linear-gradient(135deg, #7F1D1D 0%, #B91C1C 45%, #991B1B 100%);
            padding: 1.8rem 2.5rem;
            border-radius: 16px;
            color: #ffffff;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
            margin-bottom: 2rem;
            border-left: 6px solid #DC2626;
            position: relative;
            overflow: hidden;
        }}

        .exec-badge {{
            background-color: rgba(255, 255, 255, 0.18);
            color: #ffffff;
            font-size: 0.75rem;
            font-weight: 700;
            padding: 6px 16px;
            border-radius: 20px;
            letter-spacing: 0.1em;
            text-transform: uppercase;
            display: inline-block;
            margin-bottom: 0.8rem;
            border: 1px solid rgba(255, 255, 255, 0.3);
        }}

        .exec-title {{
            font-size: 2.2rem;
            font-weight: 800;
            color: #ffffff;
            margin: 0;
            line-height: 1.2;
            letter-spacing: -0.02em;
        }}

        .exec-subtitle {{
            font-size: 1rem;
            color: #FEE2E2;
            margin-top: 0.5rem;
            margin-bottom: 0;
            font-weight: 400;
        }}

        /* Metric Cards */
        .metric-card {{
            background: {card_bg};
            border: 1px solid {border_color};
            border-radius: 16px;
            padding: 1.5rem;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
        }}
        .metric-card::after {{
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 4px;
        }}
        .metric-card:hover {{
            transform: translateY(-4px);
            box-shadow: 0 12px 30px rgba(220, 38, 38, 0.15);
            background: {card_hover};
            border-color: rgba(220, 38, 38, 0.3);
        }}
        .metric-card.red::after {{ background: linear-gradient(90deg, #f43f5e, #e11d48); }}
        .metric-card.darkred::after {{ background: linear-gradient(90deg, #991B1B, #7F1D1D); }} 
        .metric-card.grey::after {{ background: linear-gradient(90deg, #64748b, #475569); }}
        .metric-card.green::after {{ background: linear-gradient(90deg, #10b981, #059669); }}
        .metric-card.amber::after {{ background: linear-gradient(90deg, #f59e0b, #d97706); }}

        .metric-label {{
            font-size: 0.8rem;
            font-weight: 600;
            color: {text_sub};
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}
        .metric-val {{
            font-size: 2.2rem;
            font-weight: 700;
            color: {text_color};
            margin-top: 0.5rem;
            letter-spacing: -0.02em;
        }}
        .metric-sub {{
            font-size: 0.8rem;
            color: {text_sub};
            margin-top: 0.2rem;
            font-weight: 400;
        }}

        .section-header {{
            font-size: 1.3rem;
            font-weight: 600;
            color: #DC2626;
            margin-top: 1.5rem;
            margin-bottom: 1rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            border-bottom: 1px solid {border_color};
            padding-bottom: 0.6rem;
            letter-spacing: -0.01em;
        }}

        .stDataFrame {{
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid {border_color};
        }}
        </style>
    """
    st.markdown(css, unsafe_allow_html=True)

    # Executive Meeting Presentation Header
    st.markdown("""
        <div class="exec-header-box">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;">
                <div>
                    <span class="exec-badge">⚡ CHARGEZONE EXECUTIVE BOARD • LIVE OPERATIONS REVIEW</span>
                    <h1 class="exec-title">Monthly Progress Report (MPR) & PM Governance</h1>
                    <p class="exec-subtitle">C-Suite Operations Deck: SLA Performance, Preventive Maintenance (PM F-01) & Infrastructure Analytics.</p>
                </div>
                <div style="text-align: right; background: rgba(255, 255, 255, 0.1); padding: 10px 20px; border-radius: 12px; border: 1px solid rgba(255, 255, 255, 0.15); backdrop-filter: blur(8px);">
                    <div style="font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.1em; color: #94a3b8; font-weight: 600;">Executive Deck</div>
                    <div style="font-size: 1.15rem; font-weight: 800; color: #ffffff;">FY 2026-27 Review</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # ---------------------------------------------------------
    # OFFICE EMAIL AUTHENTICATION GATE (MAX 10 AUTHORIZED USERS)
    # ---------------------------------------------------------
    PUBLIC_DOMAINS = {
        'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'icloud.com',
        'yandex.com', 'mail.com', 'zoho.com', 'aol.com', 'gmx.com', 'protonmail.com', 'live.com'
    }

    if 'auth_email' not in st.session_state:
        st.session_state['auth_email'] = None

    if 'authorized_office_emails' not in st.session_state:
        st.session_state['authorized_office_emails'] = []

    # Authentication Lock Screen
    if st.session_state['auth_email'] is None:
        st.markdown("---")
        st.markdown("### 🔒 Office Domain Access Authentication")
        st.markdown("Please verify your official corporate office email address to access the Executive MPR Governance Dashboard.")

        c_auth1, c_auth2 = st.columns([7, 5])
        with c_auth1:
            input_email = st.text_input(
                "🏢 Enter your Official Office Email Address (.co.in domain):",
                placeholder="e.g. alex.smith@chargezone.co.in",
                key="office_email_input"
            )
            login_btn = st.button("🔐 Verify & Access Dashboard", type="primary")

            if login_btn and input_email:
                clean_email = input_email.strip().lower()
                if '@' not in clean_email or '.' not in clean_email.split('@')[-1]:
                    st.error("❌ **Invalid Email Format**: Please enter a complete email address (e.g. user@domain.co.in).")
                elif not clean_email.endswith('.co.in'):
                    st.error("❌ **Access Restricted**: Only official corporate email addresses ending with **`.co.in`** (e.g. `user@chargezone.co.in`, `user@domain.co.in`) are allowed.")
                elif clean_email.split('@')[-1] in PUBLIC_DOMAINS:
                    st.error(f"❌ **Public Email Restricted**: Public email providers are restricted. Please use your official corporate `.co.in` email address.")
                else:
                    # Check 10 User Limit
                    if clean_email in st.session_state['authorized_office_emails']:
                        st.session_state['auth_email'] = clean_email
                        st.success(f"✓ Access Granted! Welcome back, `{clean_email}`.")
                        st.rerun()
                    elif len(st.session_state['authorized_office_emails']) < 10:
                        st.session_state['authorized_office_emails'].append(clean_email)
                        st.session_state['auth_email'] = clean_email
                        st.success(f"✓ Account Registered ({len(st.session_state['authorized_office_emails'])}/10 slots used). Welcome, `{clean_email}`!")
                        st.rerun()
                    else:
                        st.error("⛔ **Access Denied**: Maximum limit of **10 authorized `.co.in` office email users** has been reached. Please contact your administrator.")

        with c_auth2:
            st.info(f"""
            #### 🛡️ Access Policy & Governance
            - **Domain Restriction**: Strictly restricted to Official **`.co.in`** Corporate Email Domains (e.g. `@chargezone.co.in`, `@domain.co.in`).
            - **User Access Limit**: Strictly capped at **10 Authorized `.co.in` Person Accounts**.
            - **Current Registered Users**: `{len(st.session_state['authorized_office_emails'])} / 10 Slots Used`.
            """)
        return

    # Sidebar Profile & Control Panel
    st.sidebar.markdown("### ⚙️ Control Panel")
    st.sidebar.markdown("---")

    st.sidebar.markdown("#### 🔒 Authenticated Session")
    st.sidebar.info(f"👤 **User**: `{st.session_state['auth_email']}`\n\n🏢 **Status**: Corporate Office User ({len(st.session_state['authorized_office_emails'])}/10 Slots Used)")
    if st.sidebar.button("🚪 Logout"):
        st.session_state['auth_email'] = None
        st.rerun()
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

    # Processing Workbook with Caching
    with st.spinner("Processing Operational Data Engine..."):
        try:
            excel_bytes, raw_issue_df, raw_pm_df = generate_workbook_cached(issue_input, pm_input)
        except Exception as e:
            st.error(f"⚠️ **Processing Error**: {e}")
            st.exception(e)
            return

    # Sidebar Interactive Filters
    st.sidebar.markdown("#### 2. Dashboard Filters")

    segment_options = ["All Segments"]
    if 'B2B/ B2C' in raw_issue_df.columns:
        segment_options += sorted(raw_issue_df['B2B/ B2C'].dropna().unique().tolist())
    selected_segment = st.sidebar.selectbox("Filter by Customer Segment (B2B / B2C):", segment_options)

    min_date, max_date = None, None
    if 'Issue Date Parsed' in raw_issue_df.columns and not raw_issue_df['Issue Date Parsed'].dropna().empty:
        min_date = raw_issue_df['Issue Date Parsed'].min().date()
        max_date = raw_issue_df['Issue Date Parsed'].max().date()

    if min_date and max_date:
        selected_dates = st.sidebar.date_input("Filter by Issue Date Range:", value=(min_date, max_date), min_value=min_date, max_value=max_date)
    else:
        selected_dates = None

    # Apply Filters to Issue Dataframe
    filtered_issue_df = raw_issue_df.copy()
    if selected_segment != "All Segments" and 'B2B/ B2C' in filtered_issue_df.columns:
        filtered_issue_df = filtered_issue_df[filtered_issue_df['B2B/ B2C'] == selected_segment]

    if selected_dates and isinstance(selected_dates, tuple) and len(selected_dates) == 2:
        start_d, end_d = selected_dates
        if 'Issue Date Parsed' in filtered_issue_df.columns:
            filtered_issue_df = filtered_issue_df[
                (filtered_issue_df['Issue Date Parsed'].dt.date >= start_d) &
                (filtered_issue_df['Issue Date Parsed'].dt.date <= end_d)
            ]

    st.sidebar.markdown("---")

    # Download Button
    st.sidebar.markdown("#### 3. Export MPR Excel Package")
    timestamp = datetime.now().strftime("%Y-%m")
    out_filename = f"ChargeZone_MPR_Report_{timestamp}.xlsx"

    st.sidebar.download_button(
        label="⚡ Download MPR Excel Report",
        data=excel_bytes,
        file_name=out_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

    # Main Tabs
    tab_issues, tab_pm, tab_raw = st.tabs([
        "📉 Issue MPR Charts & Tables",
        "🛠️ PM F-01 Charts & Tables",
        "📋 Master Data Explorer"
    ])

    # ---------------------------------------------------------
    # TAB 1: ISSUE MPR DASHBOARD (CHARTS + TABLES SIDE-BY-SIDE)
    # ---------------------------------------------------------
    with tab_issues:
        st.markdown('<div class="section-header">📊 Operational Issue & SLA Analytics Performance</div>', unsafe_allow_html=True)

        total_issues = len(filtered_issue_df)
        within_tat = len(filtered_issue_df[filtered_issue_df['TAT Compliance'].astype(str).str.upper() == 'YES']) if 'TAT Compliance' in filtered_issue_df.columns else 0
        without_tat = len(filtered_issue_df[filtered_issue_df['TAT Compliance'].astype(str).str.upper() == 'NO']) if 'TAT Compliance' in filtered_issue_df.columns else 0
        tat_eff = (within_tat / total_issues * 100) if total_issues > 0 else 0.0

        # High-Impact KPI Row
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
                <div class="metric-card darkred">
                    <div class="metric-label">Total Issues Logged</div>
                    <div class="metric-val">{total_issues:,}</div>
                    <div class="metric-sub">Active & Closed Tickets</div>
                </div>
            """, unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
                <div class="metric-card green">
                    <div class="metric-label">Within TAT</div>
                    <div class="metric-val">{within_tat:,}</div>
                    <div class="metric-sub">SLA Compliant</div>
                </div>
            """, unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
                <div class="metric-card red">
                    <div class="metric-label">Without TAT</div>
                    <div class="metric-val">{without_tat:,}</div>
                    <div class="metric-sub">SLA Breached</div>
                </div>
            """, unsafe_allow_html=True)
        with c4:
            st.markdown(f"""
                <div class="metric-card {'green' if tat_eff >= 85 else 'amber'}">
                    <div class="metric-label">TAT Efficiency %</div>
                    <div class="metric-val">{tat_eff:.1f}%</div>
                    <div class="metric-sub">Target Benchmark: ≥ 85.0%</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 1. ZME Summary: Chart & Table Side-by-Side
        st.markdown('<div class="section-header">1. ZME Performance & TAT Efficiency Breakdown</div>', unsafe_allow_html=True)
        col_zme_c, col_zme_t = st.columns([6, 6])

        if 'ZME' in filtered_issue_df.columns and 'TAT Compliance' in filtered_issue_df.columns:
            zme_df = filtered_issue_df.groupby('ZME').agg(
                Total_Issues=('Status', 'count'),
                Within_TAT=('TAT Compliance', lambda s: (s.astype(str).str.upper() == 'YES').sum()),
                Without_TAT=('TAT Compliance', lambda s: (s.astype(str).str.upper() == 'NO').sum())
            ).reset_index()
            zme_df['TAT Efficiency %'] = (zme_df['Within_TAT'] / zme_df['Total_Issues'] * 100).round(1)
            zme_df = zme_df.rename(columns={'ZME': 'ZME Name'}).sort_values(by='Total_Issues', ascending=False)

            with col_zme_c:
                plot_vertical_bar(zme_df, x_col='ZME Name', y_col='Total_Issues', title="Total Issues Logged by ZME", color_hex="#DC2626")

            with col_zme_t:
                st.write("##### 📊 ZME SLA Data Table")
                st.dataframe(
                    zme_df.style.format({
                        'Total_Issues': '{:,}',
                        'Within_TAT': '{:,}',
                        'Without_TAT': '{:,}',
                        'TAT Efficiency %': '{:.1f}%'
                    }).background_gradient(subset=['TAT Efficiency %'], cmap='Blues'),
                    use_container_width=True,
                    height=290
                )

        st.markdown("---")

        # 2. SLA Compliance Ratio & Zone Breakdown: Side-by-Side
        col_sla_chart, col_zone_chart = st.columns([5, 7])

        with col_sla_chart:
            st.markdown('<div class="section-header">🍩 2. SLA Compliance Ratio (Pie Chart)</div>', unsafe_allow_html=True)
            plot_pie_chart(
                labels=['Within TAT (Compliant)', 'Without TAT (Breached)'],
                values=[within_tat, without_tat],
                title="Overall SLA Compliance Share",
                colors=['#10b981', '#DC2626'],
                hole=0.45
            )

        with col_zone_chart:
            st.markdown('<div class="section-header">🏢 3. Zone CM Efficiency (Grouped Chart & Table)</div>', unsafe_allow_html=True)
            if 'Zone' in filtered_issue_df.columns and 'TAT Compliance' in filtered_issue_df.columns:
                zone_df = filtered_issue_df.groupby('Zone').agg(
                    Total_Issues=('Status', 'count'),
                    Within_TAT=('TAT Compliance', lambda s: (s.astype(str).str.upper() == 'YES').sum()),
                    Without_TAT=('TAT Compliance', lambda s: (s.astype(str).str.upper() == 'NO').sum())
                ).reset_index()
                zone_df['CM Efficiency (Within TAT) %'] = (zone_df['Within_TAT'] / zone_df['Total_Issues'] * 100).round(1)
                zone_df['CM Efficiency (Without TAT) %'] = (zone_df['Without_TAT'] / zone_df['Total_Issues'] * 100).round(1)
                zone_df = zone_df.sort_values(by='Total_Issues', ascending=False)

                plot_grouped_bar(
                    df=zone_df,
                    x_col='Zone',
                    y_cols=['Within_TAT', 'Without_TAT'],
                    title="Work Orders Within vs Without TAT by Zone",
                    colors=['#10b981', '#DC2626']
                )

                st.dataframe(
                    zone_df[['Zone', 'Total_Issues', 'CM Efficiency (Within TAT) %', 'CM Efficiency (Without TAT) %']].style.format({
                        'Total_Issues': '{:,}',
                        'CM Efficiency (Within TAT) %': '{:.1f}%',
                        'CM Efficiency (Without TAT) %': '{:.1f}%'
                    }).background_gradient(subset=['CM Efficiency (Within TAT) %'], cmap='Greens'),
                    use_container_width=True
                )

        st.markdown("---")

        # 3. Status & Severity & Customer Segment (Chart + Table Pairs)
        c_stat, c_sev, c_cust = st.columns(3)

        with c_stat:
            st.markdown('<div class="section-header">📌 4. Status Breakdown</div>', unsafe_allow_html=True)
            if 'Status' in filtered_issue_df.columns:
                status_counts = filtered_issue_df['Status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                plot_vertical_bar(status_counts, x_col='Status', y_col='Count', title="Status Pipeline", color_hex="#475569")
                st.dataframe(status_counts, use_container_width=True)

        with c_sev:
            st.markdown('<div class="section-header">🚨 5. Severity Risk Profile</div>', unsafe_allow_html=True)
            if 'Severity' in filtered_issue_df.columns:
                sev_counts = filtered_issue_df['Severity'].value_counts().reset_index()
                sev_counts.columns = ['Severity', 'Count']
                plot_pie_chart(
                    labels=sev_counts['Severity'].tolist(),
                    values=sev_counts['Count'].tolist(),
                    title="Severity Share",
                    colors=['#DC2626', '#f59e0b', '#475569', '#10b981'],
                    hole=0.4
                )
                st.dataframe(sev_counts, use_container_width=True)

        with c_cust:
            st.markdown('<div class="section-header">👥 6. Customer Segment (B2B/B2C)</div>', unsafe_allow_html=True)
            if 'B2B/ B2C' in filtered_issue_df.columns:
                cust_counts = filtered_issue_df['B2B/ B2C'].value_counts().reset_index()
                cust_counts.columns = ['Segment', 'Count']
                plot_pie_chart(
                    labels=cust_counts['Segment'].tolist(),
                    values=cust_counts['Count'].tolist(),
                    title="Segment Share",
                    colors=['#475569', '#10b981'],
                    hole=0.4
                )
                st.dataframe(cust_counts, use_container_width=True)

        st.markdown("---")

        # 4. Repetitive Faults: Side-by-Side Chart + Table
        st.markdown('<div class="section-header">⚠️ 7. Repetitive Faults (Station ID & Sub-Type ≥ 2)</div>', unsafe_allow_html=True)

        if 'Station ID' in filtered_issue_df.columns and 'Issue Sub-Type' in filtered_issue_df.columns:
            group_cols = ['Station ID']
            if 'Issue Type' in filtered_issue_df.columns:
                group_cols.append('Issue Type')
            group_cols.append('Issue Sub-Type')

            pair_counts = filtered_issue_df.groupby(group_cols).size().reset_index(name='Occurrences')
            repeats = pair_counts[pair_counts['Occurrences'] >= 2].sort_values(by='Occurrences', ascending=False)

            if not repeats.empty:
                col_rep_chart, col_rep_tbl = st.columns([6, 6])
                with col_rep_chart:
                    repeats['Station_Fault'] = repeats['Station ID'].astype(str) + " - " + repeats['Issue Sub-Type'].astype(str)
                    plot_vertical_bar(repeats.head(10), x_col='Station_Fault', y_col='Occurrences', title="Top Repetitive Fault Patterns", color_hex="#991B1B")
                with col_rep_tbl:
                    st.write("##### Repetitive Station Faults Table")
                    st.dataframe(
                        repeats.style.format({'Occurrences': '{:,}'}),
                        use_container_width=True,
                        height=280
                    )
            else:
                st.success("✅ Zero repetitive station faults detected in current period dataset.")

    # ---------------------------------------------------------
    # TAB 2: PM F-01 DASHBOARD (CHARTS + TABLES SIDE-BY-SIDE)
    # ---------------------------------------------------------
    with tab_pm:
        st.markdown('<div class="section-header">🛠️ Preventive Maintenance (PM F-01) Operational Analytics</div>', unsafe_allow_html=True)

        pm_df = raw_pm_df.copy()

        total_pm_planning = len(pm_df)
        pm_done = len(pm_df[pm_df['PM Status'].astype(str).str.upper() == 'YES']) if 'PM Status' in pm_df.columns else 0
        pm_pending = len(pm_df[pm_df['PM Status'].astype(str).str.upper() == 'NO']) if 'PM Status' in pm_df.columns else 0
        
        if 'Advance PM Done' in pm_df.columns:
            advance_done = len(pm_df[pm_df['Advance PM Done'].astype(str).str.upper() == 'YES'])
        elif 'Actual Completion Date Parsed' in pm_df.columns and 'Due Date Parsed' in pm_df.columns:
            advance_done = len(pm_df[
                (pm_df['Actual Completion Date Parsed'].notna()) &
                (pm_df['Due Date Parsed'].notna()) &
                (pm_df['Actual Completion Date Parsed'] < pm_df['Due Date Parsed'])
            ])
        else:
            advance_done = 0

        pm_eff = (pm_done / total_pm_planning * 100) if total_pm_planning > 0 else 0.0

        total_chargers = len(pm_df['Charger ID'].dropna().unique()) if 'Charger ID' in pm_df.columns else len(pm_df)
        total_stations = len(pm_df['Station ID'].dropna().unique()) if 'Station ID' in pm_df.columns else len(pm_df)

        # PM High-Impact KPI Row
        p1, p2, p3, p4, p5 = st.columns(5)
        with p1:
            st.markdown(f"""
                <div class="metric-card darkred">
                    <div class="metric-label">Total Chargers</div>
                    <div class="metric-val">{total_chargers:,}</div>
                    <div class="metric-sub">Active Infrastructure</div>
                </div>
            """, unsafe_allow_html=True)
        with p2:
            st.markdown(f"""
                <div class="metric-card grey">
                    <div class="metric-label">Total Stations</div>
                    <div class="metric-val">{total_stations:,}</div>
                    <div class="metric-sub">Station Sites</div>
                </div>
            """, unsafe_allow_html=True)
        with p3:
            st.markdown(f"""
                <div class="metric-card green">
                    <div class="metric-label">PM Done</div>
                    <div class="metric-val">{pm_done:,}</div>
                    <div class="metric-sub">Verified & Completed</div>
                </div>
            """, unsafe_allow_html=True)
        with p4:
            st.markdown(f"""
                <div class="metric-card red">
                    <div class="metric-label">PM Pending</div>
                    <div class="metric-val">{pm_pending:,}</div>
                    <div class="metric-sub">Scheduled Pending</div>
                </div>
            """, unsafe_allow_html=True)
        with p5:
            st.markdown(f"""
                <div class="metric-card {'green' if pm_eff >= 90 else 'amber'}">
                    <div class="metric-label">PM Efficiency</div>
                    <div class="metric-val">{pm_eff:.1f}%</div>
                    <div class="metric-sub">Target Benchmark: ≥ 90.0%</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # PM Status Chart & ZME Grouped Bar Chart: Side-by-Side
        col_pm_pie, col_pm_bar = st.columns([5, 7])

        with col_pm_pie:
            st.markdown('<div class="section-header">🍩 1. PM Execution Status Ratio (Pie Chart)</div>', unsafe_allow_html=True)
            plot_pie_chart(
                labels=['PM Done (Completed)', 'PM Pending (Scheduled)', 'Advance PM Done'],
                values=[pm_done, pm_pending, advance_done],
                title="PM Execution Distribution Share",
                colors=['#10b981', '#DC2626', '#991B1B'],
                hole=0.45
            )

        with col_pm_bar:
            st.markdown('<div class="section-header">📊 2. PM Planning vs Completion by ZME (Grouped Chart)</div>', unsafe_allow_html=True)

            if 'ZME' in pm_df.columns and 'PM Status' in pm_df.columns:
                group_keys = ['ZME']
                if 'Zone' in pm_df.columns:
                    group_keys.append('Zone')

                pm_summary_list = []
                for name_tuple, group in pm_df.groupby(group_keys):
                    zme_name = name_tuple[0] if isinstance(name_tuple, tuple) else name_tuple
                    zone_val = name_tuple[1] if isinstance(name_tuple, tuple) and len(name_tuple) > 1 else (group['Zone'].iloc[0] if 'Zone' in group.columns else 'N/A')

                    chargers_cnt = len(group['Charger ID'].dropna().unique()) if 'Charger ID' in group.columns else len(group)
                    stations_cnt = len(group['Station ID'].dropna().unique()) if 'Station ID' in group.columns else len(group)
                    planning_cnt = len(group)
                    done_cnt = (group['PM Status'].astype(str).str.upper() == 'YES').sum()
                    pending_cnt = (group['PM Status'].astype(str).str.upper() == 'NO').sum()
                    
                    if 'Advance PM Done' in group.columns:
                        adv_cnt = (group['Advance PM Done'].astype(str).str.upper() == 'YES').sum()
                    elif 'Actual Completion Date Parsed' in group.columns and 'Due Date Parsed' in group.columns:
                        adv_cnt = ((group['Actual Completion Date Parsed'].notna()) & (group['Due Date Parsed'].notna()) & (group['Actual Completion Date Parsed'] < group['Due Date Parsed'])).sum()
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

                plot_grouped_bar(
                    df=pm_summary_df,
                    x_col='ZME Name',
                    y_cols=['PM Planning', 'PM Done', 'PM Pending'],
                    title="Scheduled vs Completed PM Work Orders by ZME",
                    colors=['#991B1B', '#10b981', '#DC2626']
                )

        st.markdown("---")

        # PM Asset Density Chart & Detailed Table: Side-by-Side
        st.markdown('<div class="section-header">⚙️ 3. Asset Infrastructure & Detailed PM Summary Table</div>', unsafe_allow_html=True)
        col_density_c, col_density_t = st.columns([6, 6])

        if 'ZME' in pm_df.columns and 'PM Status' in pm_df.columns:
            with col_density_c:
                plot_grouped_bar(
                    df=pm_summary_df,
                    x_col='ZME Name',
                    y_cols=['Total Chargers', 'Total Stations'],
                    title="Total Chargers vs Total Stations by ZME",
                    colors=['#991B1B', '#DC2626']
                )

            with col_density_t:
                st.write("##### Detailed PM F-01 Breakdown Data Table")
                st.dataframe(
                    pm_summary_df.style.format({
                        'Total Chargers': '{:,}',
                        'Total Stations': '{:,}',
                        'PM Planning': '{:,}',
                        'PM Done': '{:,}',
                        'PM Pending': '{:,}',
                        'Advance PM Done': '{:,}',
                        'PM Efficiency (%)': '{:.1f}%'
                    }).background_gradient(subset=['PM Efficiency (%)'], cmap='Greens'),
                    use_container_width=True,
                    height=300
                )

    # ---------------------------------------------------------
    # TAB 3: MASTER DATA EXPLORER
    # ---------------------------------------------------------
    with tab_raw:
        st.markdown('<div class="section-header">🔍 Master Data Governance</div>', unsafe_allow_html=True)
        data_choice = st.radio("Select Sheet:", ["Issue Tracker Master Data", "PM Tracker Master Data"], horizontal=True)

        if data_choice == "Issue Tracker Master Data":
            df_disp = ensure_unique_columns(raw_issue_df.drop(columns=[c for c in ['Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed'] if c in raw_issue_df.columns]))
            st.dataframe(df_disp, use_container_width=True)
        else:
            df_disp = ensure_unique_columns(raw_pm_df.drop(columns=[c for c in ['Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed'] if c in raw_pm_df.columns]))
            st.dataframe(df_disp, use_container_width=True)


if __name__ == '__main__':
    # CLI execution support
    if len(sys.argv) >= 2 and not sys.argv[1].startswith('-'):
        if len(sys.argv) == 2 or (len(sys.argv) >= 3 and sys.argv[2].endswith('.xlsx') and not sys.argv[1].endswith('.xlsx')):
            src_path = sys.argv[1]
            out_path = sys.argv[2] if len(sys.argv) > 2 else f"MPR_Report_{datetime.now():%Y-%m}.xlsx"
            wb, _, _ = generate_workbook_cached(src_path)
            with open(out_path, 'wb') as f:
                f.write(wb)
            print(f'Written: {out_path}')
        else:
            issue_path = sys.argv[1]
            pm_path = sys.argv[2]
            out_path = sys.argv[3] if len(sys.argv) > 3 else f"MPR_Report_{datetime.now():%Y-%m}.xlsx"
            wb, _, _ = generate_workbook_cached(issue_path, pm_path)
            with open(out_path, 'wb') as f:
                f.write(wb)
            print(f'Written: {out_path}')
    else:
        run_streamlit_app()
