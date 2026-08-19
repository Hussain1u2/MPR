import sys
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = 'Arial'
HEADER_FILL = PatternFill('solid', fgColor='1F4E78')
HEADER_FONT = Font(name=FONT, size=10, bold=True, color='FFFFFF')
TITLE_FONT = Font(name=FONT, size=14, bold=True, color='1F4E78')
SECTION_FONT = Font(name=FONT, size=11, bold=True, color='1F4E78')
NOTE_FONT = Font(name=FONT, size=9, italic=True, color='7F7F7F')
CELL_FONT = Font(name=FONT, size=10)
BOLD_CELL = Font(name=FONT, size=10, bold=True)
BORDER = Border(*(Side(style='thin', color='D9D9D9'),) * 4)
DATE_FMT = 'dd-mmm-yy'
PCT_FMT = '0.0%'
ISSUE_RANGE_END = 5000
PM_RANGE_END = 5000

PM_QUARTER_BLOCKS = {
    'FY2627-Q1': {'start': 60, 'qcol': 72},
    'FY2627-Q2': {'start': 74, 'qcol': 86},
}

PM_STATION_COLS = list(range(0, 13))

def load_issue_tracker(wb):
    rows = list(wb['Issue Tracker'].iter_rows(values_only=True))
    headers = [h.strip() if isinstance(h, str) else h for h in rows[0]]
    df = pd.DataFrame(rows[1:], columns=headers).dropna(how='all')
    return df

def load_pm_tracker(wb):
    rows = list(wb['PM Tracker'].iter_rows(values_only=True))
    row_date, row_headers = rows[3], rows[4]
    station_fields = [row_headers[i] for i in PM_STATION_COLS]

    records = []
    for r in rows[5:]:
        if r[0] is None and r[10] is None:
            continue
        station = {name: r[i] for name, i in zip(station_fields, PM_STATION_COLS)}
        for quarter, blk in PM_QUARTER_BLOCKS.items():
            compliance = r[blk['qcol']]
            col = blk['start']
            for _ in range(3):
                records.append({
                    **station,
                    'Quarter': quarter,
                    'Due Date': row_date[col],
                    'PM Status': r[col],
                    'F.E. Inspection': r[col + 1],
                    'HSE Inspection': r[col + 2],
                    'Actual Completion Date': r[col + 3],
                    'Quarterly Compliance': compliance,
                })
                col += 4
    return pd.DataFrame(records).rename(columns={'Route ': 'Route'})

def write_data_sheet(wb, name, df, table_name, date_cols):
    ws = wb.create_sheet(name)
    for c, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c, value=col)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for r, rec in enumerate(df.itertuples(index=False), start=2):
        for c, val in enumerate(rec, start=1):
            col_name = df.columns[c - 1]
            if pd.isna(val):
                val = None
            elif isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            cell = ws.cell(row=r, column=c, value=val)
            cell.font, cell.border = CELL_FONT, BORDER
            if col_name in date_cols and val is not None:
                cell.number_format = DATE_FMT

    for c, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(12, min(28, len(str(col)) + 2))

    last_row = len(df) + 1
    table = Table(displayName=table_name, ref=f"A1:{get_column_letter(len(df.columns))}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9', showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = 'A2'
    return ws, last_row


def add_pm_helper_columns(ws, pm_df, pcol, last_row):
    """Adds two formula columns PM Data needs but the raw tracker doesn't have:
    Advance PM Done (completed ahead of the due month) and First Station
    Occurrence (flags one row per ZME+Station, so distinct-station counts on
    the dashboard don't need array formulas)."""
    due_col, done_col = pcol['Due Date'], pcol['Actual Completion Date']
    adv_idx = len(pm_df.columns) + 1
    adv_letter = get_column_letter(adv_idx)
    ws.cell(row=1, column=adv_idx, value='Advance PM Done').font = HEADER_FONT
    ws.cell(row=1, column=adv_idx).fill = HEADER_FILL
    for r in range(2, last_row + 1):
        formula = (f'=IF(AND(${done_col}{r}<>"",${due_col}{r}<>"",${done_col}{r}<${due_col}{r}),"Yes",'
                   f'IF(${done_col}{r}<>"","No",""))')
        cell = ws.cell(row=r, column=adv_idx, value=formula)
        cell.font, cell.border = CELL_FONT, BORDER
    ws.column_dimensions[adv_letter].width = 16

    zme_col, station_col = pcol['ZME'], pcol['Station ID']
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
    for zme in sorted(issue_df['ZME'].dropna().unique()):
        total = f'=COUNTIFS({irange("ZME")},A{row})'
        within = f'=COUNTIFS({irange("ZME")},A{row},{irange("TAT Compliance")},"Yes")'
        without = f'=COUNTIFS({irange("ZME")},A{row},{irange("TAT Compliance")},"No")'
        eff = f'=IFERROR(C{row}/B{row},0)'
        row = data_row(ws, row, [zme, total, within, without, eff], pct_cols={4})

    row += 1
    row = section_title(ws, row, '2. Issue Summary by Zone (CM Efficiency)', 4)
    row = header_row(ws, row, ['Zone', 'Total Issues', 'CM Efficiency (Within TAT)', 'CM Efficiency (Without TAT)'])
    for zone in sorted(issue_df['Zone'].dropna().unique()):
        total = f'=COUNTIFS({irange("Zone")},A{row})'
        within = f'=IFERROR(COUNTIFS({irange("Zone")},A{row},{irange("TAT Compliance")},"Yes")/B{row},0)'
        without = f'=IFERROR(COUNTIFS({irange("Zone")},A{row},{irange("TAT Compliance")},"No")/B{row},0)'
        row = data_row(ws, row, [zone, total, within, without], pct_cols={2, 3})

    row += 1
    row = section_title(ws, row, '3. Repetitive Faults (same Station ID + Issue Sub-Type, 2+ occurrences)', 3)
    row = header_row(ws, row, ['Station ID', 'Issue Sub-Type', 'Occurrences'])
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
    for status in sorted(issue_df['Status'].dropna().unique()):
        row = data_row(ws, row, [status, f'=COUNTIFS({irange("Status")},A{row})'])

    row += 1
    row = section_title(ws, row, '5. Severity Breakdown', 2)
    row = header_row(ws, row, ['Severity', 'Count'])
    for sev in sorted(issue_df['Severity'].dropna().unique()):
        row = data_row(ws, row, [sev, f'=COUNTIFS({irange("Severity")},A{row})'])

    row += 1
    row = section_title(ws, row, '6. Customer Filter (B2B / B2C)', 2)
    row = header_row(ws, row, ['Segment', 'Count'])
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


def generate(source_path, output_path):
    import openpyxl
    src = openpyxl.load_workbook(source_path, data_only=True)
    issue_df = load_issue_tracker(src)
    pm_df = load_pm_tracker(src)

    wb = Workbook()
    wb.remove(wb.active)

    issue_ws, issue_last = write_data_sheet(
        wb, 'Issue Data', issue_df, 'IssueTable',
        date_cols=['Issue Date', 'Resolution Date', 'Restoration Date'])
    icol = {name: get_column_letter(i + 1) for i, name in enumerate(issue_df.columns)}

    pm_ws, pm_last = write_data_sheet(
        wb, 'PM Data', pm_df, 'PMTable',
        date_cols=['Go Live Date', 'Due Date', 'Actual Completion Date'])
    pcol = {name: get_column_letter(i + 1) for i, name in enumerate(pm_df.columns)}
    add_pm_helper_columns(pm_ws, pm_df, pcol, pm_last)
    pcol_full = {name: get_column_letter(i + 1)
                 for i, name in enumerate(list(pm_df.columns) + ['Advance PM Done', 'First Station Occurrence'])}

    def irange(col):
        return f"'Issue Data'!${icol[col]}$2:${icol[col]}${ISSUE_RANGE_END}"

    def prange(col):
        return f"'PM Data'!${pcol_full[col]}$2:${pcol_full[col]}${PM_RANGE_END}"

    build_issue_dashboard(wb, issue_df, irange)
    build_pm_dashboard(wb, pm_df, prange)
    wb.active = 0
    wb.save(output_path)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 generate_mpr_report.py <source.xlsx> [output.xlsx]')
        sys.exit(1)
    src_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else f"MPR_Report_{datetime.now():%Y-%m}.xlsx"
    generate(src_path, out_path)
    print(f'Written: {out_path}')
    print('Run recalc.py from the xlsx skill (or open in Excel and save) so formula values populate.')