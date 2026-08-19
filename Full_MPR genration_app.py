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

# Styling constants for OpenPyXL excel generation
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
    if 'Issue Date' in df.columns:
        df['Issue Date Parsed'] = pd.to_datetime(df['Issue Date'], errors='coerce')
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
    df = pd.DataFrame(records).rename(columns={'Route ': 'Route'})
    if 'Due Date' in df.columns:
        df['Due Date Parsed'] = pd.to_datetime(df['Due Date'], errors='coerce')
    if 'Actual Completion Date' in df.columns:
        df['Actual Completion Date Parsed'] = pd.to_datetime(df['Actual Completion Date'], errors='coerce')
    return df


def write_data_sheet(wb, name, df, table_name, date_cols):
    ws = wb.create_sheet(name)
    # Drop helper parsed columns before writing to openpyxl
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
    due_col, done_col = pcol['Due Date'], pcol['Actual Completion Date']
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


def generate_workbook(source_input):
    src = openpyxl.load_workbook(source_input, data_only=True)
    issue_df = load_issue_tracker(src)
    pm_df = load_pm_tracker(src)

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
        return f"'Issue Data'!${icol[col]}$2:${icol[col]}${ISSUE_RANGE_END}"

    def prange(col):
        return f"'PM Data'!${pcol_full[col]}$2:${pcol_full[col]}${PM_RANGE_END}"

    build_issue_dashboard(wb, issue_df, irange)
    build_pm_dashboard(wb, pm_df, prange)
    wb.active = 0
    return wb, issue_df, pm_df


def run_streamlit_app():
    st.set_page_config(
        page_title="ChargeZone | Monthly Progress Report (MPR) Executive Dashboard",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Professional Executive Styling
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        .exec-header-box {
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            padding: 1.5rem 2rem;
            border-radius: 12px;
            color: #FFFFFF;
            box-shadow: 0 4px 20px rgba(15, 23, 42, 0.15);
            margin-bottom: 1.5rem;
            border-left: 6px solid #2563EB;
        }

        .exec-badge {
            background-color: #2563EB;
            color: #FFFFFF;
            font-size: 0.75rem;
            font-weight: 700;
            padding: 4px 10px;
            border-radius: 20px;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            display: inline-block;
            margin-bottom: 0.5rem;
        }

        .exec-title {
            font-size: 2rem;
            font-weight: 700;
            color: #F8FAFC;
            margin: 0;
            line-height: 1.2;
        }

        .exec-subtitle {
            font-size: 0.95rem;
            color: #94A3B8;
            margin-top: 0.4rem;
            margin-bottom: 0;
        }

        .metric-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            padding: 1rem 1.2rem;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .metric-card:hover {
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
        }
        .metric-card.blue { border-top: 4px solid #2563EB; }
        .metric-card.green { border-top: 4px solid #10B981; }
        .metric-card.red { border-top: 4px solid #EF4444; }
        .metric-card.amber { border-top: 4px solid #F59E0B; }

        .metric-label {
            font-size: 0.75rem;
            font-weight: 600;
            color: #64748B;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        .metric-val {
            font-size: 1.7rem;
            font-weight: 700;
            color: #0F172A;
            margin-top: 0.2rem;
        }
        .metric-sub {
            font-size: 0.75rem;
            color: #64748B;
            margin-top: 0.1rem;
        }

        .section-header {
            font-size: 1.1rem;
            font-weight: 700;
            color: #0F172A;
            margin-top: 1rem;
            margin-bottom: 0.8rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .stDataFrame {
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid #E2E8F0;
        }
        </style>
    """, unsafe_allow_html=True)

    # Executive Header
    st.markdown("""
        <div class="exec-header-box">
            <span class="exec-badge">ChargeZone Executive Governance</span>
            <h1 class="exec-title">Monthly Progress Report (MPR) & PM F-01 Dashboard</h1>
            <p class="exec-subtitle">Executive Performance Analytics for Operations, Preventive Maintenance (PM) & SLA Tracking.</p>
        </div>
    """, unsafe_allow_html=True)

    # Sidebar Panel
    st.sidebar.markdown("### ⚙️ Control Panel")
    st.sidebar.markdown("---")

    st.sidebar.markdown("#### 1. Data Source")
    uploaded_file = st.sidebar.file_uploader("Upload Tracker Workbook (.xlsx)", type=["xlsx"])

    use_default = False
    default_path = "issue,pm tracker merged.xlsx"
    if uploaded_file is None:
        try:
            import os
            if os.path.exists(default_path):
                use_default = st.sidebar.checkbox("Load Sample Dataset (`issue,pm tracker merged.xlsx`)", value=True)
        except Exception:
            pass

    source_input = None
    file_name = None

    if uploaded_file is not None:
        source_input = uploaded_file
        file_name = uploaded_file.name
    elif use_default:
        source_input = default_path
        file_name = default_path

    if source_input is None:
        st.info("📌 **Action Required**: Please upload an Operational Tracker Workbook (`.xlsx`) containing `Issue Tracker` and `PM Tracker` worksheets.")
        return

    st.sidebar.success(f"✓ Connected: `{file_name}`")
    st.sidebar.markdown("---")

    # Processing Workbook
    with st.spinner("Processing Operational Engine..."):
        try:
            wb, raw_issue_df, raw_pm_df = generate_workbook(source_input)
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
        except Exception as e:
            st.error(f"⚠️ **Processing Error**: {e}")
            st.exception(e)
            return

    # Sidebar Interactive Filters
    st.sidebar.markdown("#### 2. Dashboard Filters")

    # Customer Segment Filter
    segment_options = ["All Segments"]
    if 'B2B/ B2C' in raw_issue_df.columns:
        segment_options += sorted(raw_issue_df['B2B/ B2C'].dropna().unique().tolist())
    selected_segment = st.sidebar.selectbox("Filter by Customer Segment (B2B / B2C):", segment_options)

    # Date Filter
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
        data=output_buffer,
        file_name=out_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

    # Main Tabs
    tab_issues, tab_pm, tab_raw = st.tabs([
        "📉 Issue MPR Dashboard",
        "🛠️ PM F-01 Dashboard",
        "📋 Master Data Explorer"
    ])

    # ---------------------------------------------------------
    # TAB 1: ISSUE MPR DASHBOARD
    # ---------------------------------------------------------
    with tab_issues:
        st.markdown('<div class="section-header">📊 Operational Issue & SLA Performance Analytics</div>', unsafe_allow_html=True)

        total_issues = len(filtered_issue_df)
        within_tat = len(filtered_issue_df[filtered_issue_df['TAT Compliance'].astype(str).str.upper() == 'YES']) if 'TAT Compliance' in filtered_issue_df.columns else 0
        without_tat = len(filtered_issue_df[filtered_issue_df['TAT Compliance'].astype(str).str.upper() == 'NO']) if 'TAT Compliance' in filtered_issue_df.columns else 0
        tat_eff = (within_tat / total_issues * 100) if total_issues > 0 else 0.0

        # KPI Summary Cards (Items 2, 3, 4, 5)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
                <div class="metric-card blue">
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
                    <div class="metric-sub">Target: ≥ 85.0%</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        col_zme, col_zone = st.columns([6, 6])

        # 1. ZME Name Summary (Items 1, 2, 3, 4, 5)
        with col_zme:
            st.markdown('<div class="section-header">1. Issue Summary by ZME Name</div>', unsafe_allow_html=True)
            if 'ZME' in filtered_issue_df.columns and 'TAT Compliance' in filtered_issue_df.columns:
                zme_df = filtered_issue_df.groupby('ZME').agg(
                    Total_Issues=('Status', 'count'),
                    Within_TAT=('TAT Compliance', lambda s: (s.astype(str).str.upper() == 'YES').sum()),
                    Without_TAT=('TAT Compliance', lambda s: (s.astype(str).str.upper() == 'NO').sum())
                ).reset_index()
                zme_df['TAT Efficiency %'] = (zme_df['Within_TAT'] / zme_df['Total_Issues'] * 100).round(1)
                zme_df = zme_df.rename(columns={'ZME': 'ZME Name'}).sort_values(by='Total_Issues', ascending=False)
                
                st.dataframe(
                    zme_df.style.format({
                        'Total_Issues': '{:,}',
                        'Within_TAT': '{:,}',
                        'Without_TAT': '{:,}',
                        'TAT Efficiency %': '{:.1f}%'
                    }).background_gradient(subset=['TAT Efficiency %'], cmap='Blues'),
                    use_container_width=True,
                    height=300
                )

        # 2. Zone Summary & CM Efficiency (Items 6, 7, 8, 9)
        with col_zone:
            st.markdown('<div class="section-header">2. Issue Summary by Zone (CM Efficiency)</div>', unsafe_allow_html=True)
            if 'Zone' in filtered_issue_df.columns and 'TAT Compliance' in filtered_issue_df.columns:
                zone_df = filtered_issue_df.groupby('Zone').agg(
                    Total_Issues=('Status', 'count'),
                    Within_TAT=('TAT Compliance', lambda s: (s.astype(str).str.upper() == 'YES').sum()),
                    Without_TAT=('TAT Compliance', lambda s: (s.astype(str).str.upper() == 'NO').sum())
                ).reset_index()
                zone_df['CM Efficiency (Within TAT) %'] = (zone_df['Within_TAT'] / zone_df['Total_Issues'] * 100).round(1)
                zone_df['CM Efficiency (Without TAT) %'] = (zone_df['Without_TAT'] / zone_df['Total_Issues'] * 100).round(1)
                zone_df = zone_df[['Zone', 'Total_Issues', 'CM Efficiency (Within TAT) %', 'CM Efficiency (Without TAT) %']].sort_values(by='Total_Issues', ascending=False)

                st.dataframe(
                    zone_df.style.format({
                        'Total_Issues': '{:,}',
                        'CM Efficiency (Within TAT) %': '{:.1f}%',
                        'CM Efficiency (Without TAT) %': '{:.1f}%'
                    }).background_gradient(subset=['CM Efficiency (Within TAT) %'], cmap='Greens'),
                    use_container_width=True,
                    height=300
                )

        st.markdown("---")

        # 3. Repetitive Faults (Item 10)
        st.markdown('<div class="section-header">3. Repetitive Faults (Top Issue Type & Sub-Type)</div>', unsafe_allow_html=True)
        st.caption("Stations with multiple occurrences (≥ 2) for identical Issue Sub-Types.")

        rep_cols = [c for c in ['Station ID', 'Issue Type', 'Issue Sub-Type'] if c in filtered_issue_df.columns]
        if 'Station ID' in filtered_issue_df.columns and 'Issue Sub-Type' in filtered_issue_df.columns:
            group_cols = ['Station ID']
            if 'Issue Type' in filtered_issue_df.columns:
                group_cols.append('Issue Type')
            group_cols.append('Issue Sub-Type')

            pair_counts = filtered_issue_df.groupby(group_cols).size().reset_index(name='Occurrences')
            repeats = pair_counts[pair_counts['Occurrences'] >= 2].sort_values(by='Occurrences', ascending=False)

            if not repeats.empty:
                st.dataframe(
                    repeats.style.format({'Occurrences': '{:,}'}),
                    use_container_width=True
                )
            else:
                st.success("✅ No repetitive faults found in selected dataset.")

        st.markdown("---")

        col_stat, col_sev, col_cust = st.columns(3)

        # 4. Status Breakdown (Item 11)
        with col_stat:
            st.markdown('<div class="section-header">4. Status Breakdown</div>', unsafe_allow_html=True)
            if 'Status' in filtered_issue_df.columns:
                status_counts = filtered_issue_df['Status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                st.dataframe(status_counts, use_container_width=True)
                st.bar_chart(status_counts.set_index('Status'), color="#2563EB", height=200)

        # 5. Severity Breakdown (Item 12)
        with col_sev:
            st.markdown('<div class="section-header">5. Severity Breakdown</div>', unsafe_allow_html=True)
            if 'Severity' in filtered_issue_df.columns:
                sev_counts = filtered_issue_df['Severity'].value_counts().reset_index()
                sev_counts.columns = ['Severity', 'Count']
                st.dataframe(sev_counts, use_container_width=True)
                st.bar_chart(sev_counts.set_index('Severity'), color="#EF4444", height=200)

        # 6. Customer Filter Segment (Item 13)
        with col_cust:
            st.markdown('<div class="section-header">6. Customer Segment (B2B / B2C)</div>', unsafe_allow_html=True)
            if 'B2B/ B2C' in filtered_issue_df.columns:
                cust_counts = filtered_issue_df['B2B/ B2C'].value_counts().reset_index()
                cust_counts.columns = ['Segment', 'Count']
                st.dataframe(cust_counts, use_container_width=True)
                st.bar_chart(cust_counts.set_index('Segment'), color="#10B981", height=200)

    # ---------------------------------------------------------
    # TAB 2: PM F-01 DASHBOARD
    # ---------------------------------------------------------
    with tab_pm:
        st.markdown('<div class="section-header">🛠️ Preventive Maintenance (PM F-01) Dashboard</div>', unsafe_allow_html=True)

        pm_df = raw_pm_df.copy()

        # Compute Metrics for PM F-01
        total_pm_planning = len(pm_df)
        pm_done = len(pm_df[pm_df['PM Status'].astype(str).str.upper() == 'YES']) if 'PM Status' in pm_df.columns else 0
        pm_pending = len(pm_df[pm_df['PM Status'].astype(str).str.upper() == 'NO']) if 'PM Status' in pm_df.columns else 0
        
        # Advance PM Done calculation
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

        # Unique Chargers and Stations Count
        total_chargers = len(pm_df['Charger ID'].dropna().unique()) if 'Charger ID' in pm_df.columns else len(pm_df)
        total_stations = len(pm_df['Station ID'].dropna().unique()) if 'Station ID' in pm_df.columns else len(pm_df)

        # PM KPI Cards Row (Items 2, 3, 5, 6, 7, 8, 9)
        p1, p2, p3, p4, p5 = st.columns(5)
        with p1:
            st.markdown(f"""
                <div class="metric-card blue">
                    <div class="metric-label">Total Chargers</div>
                    <div class="metric-val">{total_chargers:,}</div>
                    <div class="metric-sub">Active Assets</div>
                </div>
            """, unsafe_allow_html=True)
        with p2:
            st.markdown(f"""
                <div class="metric-card blue">
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
                    <div class="metric-sub">Completed</div>
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
                    <div class="metric-sub">Target: ≥ 90.0%</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # PM Summary by ZME Table (Items 1 to 9)
        st.markdown('<div class="section-header">⚙️ PM F-01 Detailed Summary by ZME Name</div>', unsafe_allow_html=True)

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
                use_container_width=True
            )

    # ---------------------------------------------------------
    # TAB 3: MASTER DATA EXPLORER
    # ---------------------------------------------------------
    with tab_raw:
        st.markdown('<div class="section-header">🔍 Master Data Governance</div>', unsafe_allow_html=True)
        data_choice = st.radio("Select Sheet:", ["Issue Tracker Master Data", "PM Tracker Master Data"], horizontal=True)

        if data_choice == "Issue Tracker Master Data":
            st.dataframe(raw_issue_df.drop(columns=[c for c in ['Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed'] if c in raw_issue_df.columns]), use_container_width=True)
        else:
            st.dataframe(raw_pm_df.drop(columns=[c for c in ['Issue Date Parsed', 'Due Date Parsed', 'Actual Completion Date Parsed'] if c in raw_pm_df.columns]), use_container_width=True)


if __name__ == '__main__':
    # CLI execution support
    if len(sys.argv) >= 2 and not sys.argv[1].startswith('-'):
        src_path = sys.argv[1]
        out_path = sys.argv[2] if len(sys.argv) > 2 else f"MPR_Report_{datetime.now():%Y-%m}.xlsx"
        wb, _, _ = generate_workbook(src_path)
        wb.save(out_path)
        print(f'Written: {out_path}')
    else:
        run_streamlit_app()
