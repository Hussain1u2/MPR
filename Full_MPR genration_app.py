from datetime import datetime
from io import BytesIO

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="ChargeZone | Executive MPR & PM Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Cleaning & Error Handling

BAD_VALUES = {
    '', '#REF!', '#N/A', '#VALUE!', '#NAME?', 'NONE', 'NULL', 'NAN', 'NAT',
    'N/A', 'NA', 'N.A.', 'N.A', '<NA>', '#N/A N/A', 'UNKNOWN', 'UNDEFINED',
    'REF!', '#NULL!', '#NUM!', '#DIV/0!', '-', '--'
}

def clean_val(v):
    if v is None or pd.isna(v):
        return None
    if isinstance(v, str):
        s = v.strip()
        if s.upper() in BAD_VALUES:
            return None
        return s
    return v

def norm_header(h):
    if h is None:
        return ''
    return ''.join(c for c in str(h).lower() if c.isalnum())

def find_col(df, possible_names):
    if df is None or df.empty:
        return None
    clean_cols = {norm_header(c): c for c in df.columns}
    for p in possible_names:
        p_norm = norm_header(p)
        if p_norm in clean_cols:
            return clean_cols[p_norm]
    for p in possible_names:
        p_norm = norm_header(p)
        for c_norm, orig in clean_cols.items():
            if p_norm in c_norm or c_norm in p_norm:
                return orig
    return None

def fiscal_year_quarter(dt):
    if not isinstance(dt, (datetime, pd.Timestamp)) or pd.isna(dt):
        return 'Unscheduled FY', 0, 'Unscheduled Q', 'Unscheduled'
    fy_start = dt.year if dt.month >= 4 else dt.year - 1
    fy_label = f"FY{str(fy_start)[-2:]}{str(fy_start + 1)[-2:]}"
    quarter_num = (dt.month - 4) % 12 // 3 + 1
    return fy_label, quarter_num, f"{fy_label} Q{quarter_num}", dt.strftime('%b-%y')

# Excel Parsers

def select_sheet_name(sheetnames, preferred_names, fallback_keyword):
    if not sheetnames:
        return None
    clean_map = {str(s).strip().lower(): s for s in sheetnames if s is not None}
    for pref in preferred_names:
        pref_clean = pref.strip().lower()
        if pref_clean in clean_map:
            return clean_map[pref_clean]
    for pref in preferred_names:
        pref_clean = pref.strip().lower()
        for s in sheetnames:
            if s and pref_clean in str(s).strip().lower():
                return s
    for s in sheetnames:
        if s and fallback_keyword.lower() in str(s).strip().lower():
            return s
    return sheetnames[0]


def parse_issue_tracker(wb):
    sheet_name = select_sheet_name(wb.sheetnames, ['Issue Tracker', 'Issue Data'], 'issue')
    if not sheet_name or sheet_name not in wb.sheetnames:
        return pd.DataFrame()
    
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame()

    header_idx = 0
    for idx, r in enumerate(rows[:15]):
        row_norms = [norm_header(c) for c in r if c is not None]
        if 'issueid' in row_norms or 'srno' in row_norms or ('zone' in row_norms and 'status' in row_norms):
            header_idx = idx
            break

    header_row = rows[header_idx]
    idx_map = {norm_header(h): i for i, h in enumerate(header_row) if h is not None}
    
    col_issueid = idx_map.get('issueid', idx_map.get('srno', idx_map.get('ticketid')))
    col_issuedate = idx_map.get('issuedate', idx_map.get('createddate', idx_map.get('date')))
    col_zone = idx_map.get('zone', idx_map.get('region'))
    col_zme = idx_map.get('zme', idx_map.get('leadzme', idx_map.get('managername', idx_map.get('zonemanager'))))
    col_status = idx_map.get('status', idx_map.get('ticketstatus', idx_map.get('issuestatus')))
    col_severity = idx_map.get('severity', idx_map.get('priority'))
    col_tat = idx_map.get('tatcompliance', idx_map.get('slacompliance', idx_map.get('compliance')))
    col_type = idx_map.get('issuetype', idx_map.get('category'))
    col_subtype = idx_map.get('issuesubtype', idx_map.get('subtype', idx_map.get('faultsubtype')))
    col_stationid = idx_map.get('stationid', idx_map.get('siteid', idx_map.get('ocppid')))
    col_stationname = idx_map.get('stationname', idx_map.get('sitename'))
    col_segment = idx_map.get('b2bb2c', idx_map.get('segment', idx_map.get('customersegment')))
    col_ocpp = idx_map.get('ocppid', idx_map.get('chargerid'))

    records = []
    for r in rows[header_idx + 1:]:
        if not r or len(r) <= max(filter(lambda x: x is not None, [col_issueid, col_zone, col_status]), default=0):
            continue
        
        issue_id = clean_val(r[col_issueid]) if col_issueid is not None else None
        if issue_id is None:
            continue
            
        issue_date = r[col_issuedate] if col_issuedate is not None and col_issuedate < len(r) else None
        if isinstance(issue_date, str):
            issue_date = pd.to_datetime(issue_date, errors='coerce', dayfirst=True)
            if pd.isna(issue_date):
                issue_date = None
        
        if isinstance(issue_date, (datetime, pd.Timestamp)) and not pd.isna(issue_date):
            fy, qnum, qlabel, mlabel = fiscal_year_quarter(issue_date)
        else:
            fy, qnum, qlabel, mlabel = 'Unscheduled FY', 0, 'Unscheduled Q', 'Unscheduled'

        status_raw = str(clean_val(r[col_status]) or 'Open').strip()
        # A blank TAT Compliance cell means "not yet determined" (e.g. the
        # issue is still open) -- it must NOT default to 'No', which would
        # falsely assert an SLA breach that was never actually recorded.
        tat_clean = clean_val(r[col_tat])
        tat_raw = str(tat_clean).strip().upper() if tat_clean is not None else ''
        
        records.append({
            'issueId': issue_id,
            'zme': clean_val(r[col_zme]) if col_zme is not None and col_zme < len(r) else 'Unassigned',
            'zone': clean_val(r[col_zone]) if col_zone is not None and col_zone < len(r) else 'Unknown',
            'status': status_raw,
            'severity': clean_val(r[col_severity]) if col_severity is not None and col_severity < len(r) else 'Normal',
            'tatCompliance': tat_raw,
            'issueType': clean_val(r[col_type]) if col_type is not None and col_type < len(r) else 'General',
            'issueSubType': clean_val(r[col_subtype]) if col_subtype is not None and col_subtype < len(r) else 'General',
            'stationId': clean_val(r[col_stationid]) if col_stationid is not None and col_stationid < len(r) else 'N/A',
            'stationName': clean_val(r[col_stationname]) if col_stationname is not None and col_stationname < len(r) else 'N/A',
            'ocppId': clean_val(r[col_ocpp]) if col_ocpp is not None and col_ocpp < len(r) else 'N/A',
            'segment': clean_val(r[col_segment]) if col_segment is not None and col_segment < len(r) else 'B2C',
            'fy': fy, 'quarter': qnum, 'quarterLabel': qlabel, 'month': mlabel,
            'issueDate': issue_date,
        })

    df = pd.DataFrame(records)
    if df.empty:
        return df

    df['zme'] = df['zme'].fillna('Unassigned')
    df['zone'] = df['zone'].fillna('Unknown')

    # Executive SLA & CM Formulas:
    # 1. Closed Faults = Tickets with status CLOSED or RESOLVED
    # 2. Open Faults = Tickets with status NOT CLOSED/RESOLVED
    # 3. Closed Within TAT = Closed Tickets with TAT Compliance == YES
    # 4. Closed Without TAT = Closed Tickets with TAT Compliance == NO
    # 5. Overall CM Efficiency % = Closed Faults / Faults Registered * 100
    # 6. CM-TAT Efficiency % = Closed Within TAT / Total Closed Faults * 100
    status_upper = df['status'].astype(str).str.strip().str.upper()
    df['_Is_Closed_'] = status_upper.isin(['CLOSED', 'RESOLVED'])
    df['_Is_Open_'] = ~df['_Is_Closed_']

    tat_upper = df['tatCompliance'].astype(str).str.strip().str.upper()
    df['_Is_Within_'] = (tat_upper == 'YES')
    df['_Is_Without_'] = (tat_upper == 'NO')
    df['_Is_Closed_Within_'] = df['_Is_Closed_'] & df['_Is_Within_']
    df['_Is_Closed_Without_'] = df['_Is_Closed_'] & df['_Is_Without_']

    return df


def classify_pm_blocks(rows, header_idx):
    due_date_row = rows[header_idx - 1] if header_idx > 0 else [None] * len(rows[header_idx])
    header_row = rows[header_idx]

    blocks = []
    current_block = None

    for i, h in enumerate(header_row):
        text = str(h or '').strip().lower()

        if any(x in text for x in ['quarterly schedule', 'repeatitive', 'compliance', 'inspection', 'first aid', 'hse', 'f.e.']):
            continue
        if 'actual completion date' in text or 'completion date' in text:
            if current_block is not None:
                current_block['completionCol'] = i
            continue
        if 'status' in text and 'station' not in text:
            due_dt = due_date_row[i] if i < len(due_date_row) else None
            if not isinstance(due_dt, (datetime, pd.Timestamp)):
                due_dt = None
            current_block = {
                'statusCol': i,
                'completionCol': None,
                'headerText': str(h),
                'dueDate': due_dt
            }
            blocks.append(current_block)

    return blocks


def parse_pm_tracker(wb):
    sheet_name = select_sheet_name(wb.sheetnames, ['PM Tracker B2C- B2B', 'PM Tracker B2C-B2B', 'PM Tracker', 'PM Data'], 'pm')
    if not sheet_name or sheet_name not in wb.sheetnames:
        return pd.DataFrame()

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame()

    header_idx = 1
    for idx, r in enumerate(rows[:10]):
        row_norms = [norm_header(c) for c in r if c is not None]
        if 'ocppid' in row_norms or 'leadzme' in row_norms:
            header_idx = idx
            break

    header_row = rows[header_idx]
    due_date_row = rows[header_idx - 1] if header_idx > 0 else [None] * len(header_row)

    station_fields = [norm_header(h) for h in header_row[:15]]
    ocpp_pos = station_fields.index('ocppid') if 'ocppid' in station_fields else 1
    zme_pos = station_fields.index('zme') if 'zme' in station_fields else 4
    lead_zme_pos = station_fields.index('leadzme') if 'leadzme' in station_fields else 3
    zone_pos = station_fields.index('zone') if 'zone' in station_fields else 7
    station_id_pos = station_fields.index('stationid') if 'stationid' in station_fields else 10
    station_name_pos = station_fields.index('stationname') if 'stationname' in station_fields else 5
    segment_pos = station_fields.index('b2bb2c') if 'b2bb2c' in station_fields else 8
    status_pos = station_fields.index('stationstatus') if 'stationstatus' in station_fields else 11

    blocks = classify_pm_blocks(rows, header_idx)

    records = []
    for r in rows[header_idx + 1:]:
        if len(r) <= ocpp_pos or r[ocpp_pos] is None:
            continue

        ocpp_val = clean_val(r[ocpp_pos])
        if ocpp_val is None:
            continue

        zme_val = clean_val(r[lead_zme_pos]) if lead_zme_pos < len(r) else None
        if zme_val is None and zme_pos < len(r):
            zme_val = clean_val(r[zme_pos])
        if zme_val is None:
            zme_val = 'Unassigned'

        zone_val = clean_val(r[zone_pos]) if zone_pos < len(r) else 'Unknown'
        station_id_val = clean_val(r[station_id_pos]) if station_id_pos < len(r) else 'N/A'
        station_name_val = clean_val(r[station_name_pos]) if station_name_pos < len(r) else 'N/A'
        segment_val = clean_val(r[segment_pos]) if segment_pos < len(r) else 'B2C'
        stn_status_val = clean_val(r[status_pos]) if status_pos < len(r) else 'Live'
        stn_status_str = str(stn_status_val or 'Live').strip().upper()

        # Strict Filter: Only process Live stations with valid OCPP ID
        if stn_status_str != 'LIVE':
            continue

        for b in blocks:
            sc = b['statusCol']
            if sc >= len(r):
                continue
            
            raw_st = r[sc]
            st_clean = clean_val(raw_st)
            
            due_date = due_date_row[sc] if sc < len(due_date_row) else None
            if not isinstance(due_date, (datetime, pd.Timestamp)):
                due_date = None

            # No real due date for this column -> not a genuine scheduled
            # month, skip rather than mis-attributing it to a fabricated
            # date (that previously forced stray columns into Apr-2025).
            if due_date is None:
                continue

            fy, qnum, qlabel, mlabel = fiscal_year_quarter(due_date)

            cc = b['completionCol']
            completion = r[cc] if cc is not None and cc < len(r) else None
            if isinstance(completion, str):
                completion = pd.to_datetime(completion, errors='coerce', dayfirst=True)
                if pd.isna(completion):
                    completion = None
            if not isinstance(completion, (datetime, pd.Timestamp)):
                completion = None

            # st_clean is already None for any BAD_VALUES token (including
            # bare 'NA', which showed up as a literal status value in the
            # source and was previously being counted as a real "pending"
            # PM instance).
            st_str = str(st_clean or '').strip().upper()
            has_status = st_clean is not None

            # Executive PM Governance Rules:
            # 1. PM Done: Marked as Yes/Done/Completed OR actual completion date is present
            is_done = (st_str in ['YES', 'DONE', 'COMPLETED', 'TRUE', '1']) or (completion is not None)

            # 2. PM Planned: Work order scheduled for selected month (has recorded status or is completed)
            is_planned = is_done or (st_clean is not None) or (st_str in ['NO', '0', 'FALSE', 'PENDING', 'OPEN', 'SCHED', 'SCHEDULED'])

            # 3. PM Pending: Work order scheduled for selected month that is not completed
            is_pending = is_planned and not is_done

            # 4. Advance PM Done: Completion date is earlier than schedule due date
            adv_done = False
            if is_done and completion is not None and due_date is not None:
                if completion < due_date or (completion.year == due_date.year and completion.month < due_date.month):
                    adv_done = True

            # PM Compliance Status Formula
            if is_planned:
                if is_done:
                    if completion is not None and due_date is not None:
                        if completion < due_date or (completion.year == due_date.year and completion.month < due_date.month):
                            pm_compliance = '🟡 Advance PM Done (Before Schedule)'
                        elif completion.month == due_date.month and completion.year == due_date.year:
                            pm_compliance = '🟢 On-Time (As Scheduled)'
                        elif completion > due_date:
                            pm_compliance = '🟠 Completed Delayed'
                        else:
                            pm_compliance = '🟢 On-Time (As Scheduled)'
                    else:
                        pm_compliance = '🟢 Completed'
                else:
                    if due_date is not None and datetime.now() > due_date:
                        pm_compliance = '🔴 Overdue / Breached Schedule'
                    else:
                        pm_compliance = '⚪ Pending (In Schedule)'
            else:
                pm_compliance = '⚪ Unscheduled'

            records.append({
                'ocppId': ocpp_val,
                'zme': zme_val,
                'zone': zone_val,
                'stationId': station_id_val,
                'stationName': station_name_val,
                'segment': segment_val,
                'stationStatus': stn_status_val,
                'status': st_clean,
                'dueDate': due_date,
                'completionDate': completion,
                'fy': fy, 'quarter': qnum, 'quarterLabel': qlabel, 'month': mlabel,
                'Is_PM_Planned': is_planned,
                'Is_PM_Done': is_done,
                'Is_PM_Pending': is_pending,
                'Advance PM Done': adv_done,
                'PM Compliance Status': pm_compliance,
            })

    df = pd.DataFrame(records)
    if not df.empty:
        df['zme'] = df['zme'].fillna('Unassigned')
        df['zone'] = df['zone'].fillna('Unknown')

        # Add Period Date helper columns for date/month/quarter/year breakdown
        if 'dueDate' in df.columns and df['dueDate'].notna().any():
            df['Scheduled Month'] = df['dueDate'].dt.strftime('%b-%Y').fillna('Unscheduled')
            df['Scheduled Quarter'] = ('Q' + df['dueDate'].dt.quarter.astype(str)).fillna('Unscheduled')
            df['Scheduled Year'] = df['dueDate'].dt.year.astype(str).fillna('Unscheduled')
            df['Scheduled Date'] = df['dueDate'].dt.strftime('%Y-%m-%d').fillna('Unscheduled')
        else:
            df['Scheduled Month'] = 'Unscheduled'
            df['Scheduled Quarter'] = 'Unscheduled'
            df['Scheduled Year'] = 'Unscheduled'
            df['Scheduled Date'] = 'Unscheduled'

    return df


@st.cache_data(show_spinner=False)
def load_workbook_data(file_bytes_or_dict):
    issue_df = pd.DataFrame()
    pm_df = pd.DataFrame()

    if isinstance(file_bytes_or_dict, dict):
        issue_bytes = file_bytes_or_dict.get('issue')
        pm_bytes = file_bytes_or_dict.get('pm')

        if issue_bytes:
            wb_i = openpyxl.load_workbook(BytesIO(issue_bytes), read_only=True, data_only=True)
            try:
                issue_df = parse_issue_tracker(wb_i)
                pm_df = parse_pm_tracker(wb_i)
            finally:
                wb_i.close()

        if pm_bytes and (pm_df.empty or issue_df.empty):
            wb_p = openpyxl.load_workbook(BytesIO(pm_bytes), read_only=True, data_only=True)
            try:
                if pm_df.empty:
                    pm_df = parse_pm_tracker(wb_p)
                if issue_df.empty:
                    issue_df = parse_issue_tracker(wb_p)
            finally:
                wb_p.close()

        return issue_df, pm_df
    else:
        wb = openpyxl.load_workbook(BytesIO(file_bytes_or_dict), read_only=True, data_only=True)
        try:
            issue_df = parse_issue_tracker(wb)
            pm_df = parse_pm_tracker(wb)
        finally:
            wb.close()
        return issue_df, pm_df



def quarter_sort_key(fy, q):
    try:
        return (2000 + int(fy[2:4])) * 10 + int(q)
    except Exception:
        return 20250

def build_quarters(issue_df, pm_df):
    dfs = []
    if not issue_df.empty and 'fy' in issue_df.columns:
        dfs.append(issue_df[['fy', 'quarter', 'quarterLabel']])
    if not pm_df.empty and 'fy' in pm_df.columns:
        dfs.append(pm_df[['fy', 'quarter', 'quarterLabel']])

    if not dfs:
        return {'All Quarters': []}

    combo = pd.concat(dfs).drop_duplicates()
    combo['sort_key'] = combo.apply(lambda r: quarter_sort_key(r['fy'], r['quarter']), axis=1)
    combo = combo.sort_values('sort_key')
    
    quarters = {}
    all_months_set = []

    for label in combo['quarterLabel']:
        m_list = []
        if not pm_df.empty and 'quarterLabel' in pm_df.columns and 'month' in pm_df.columns:
            pm_sub = pm_df[pm_df['quarterLabel'] == label]
            if 'dueDate' in pm_sub.columns:
                m_sub = pm_sub[['month', 'dueDate']].drop_duplicates().sort_values('dueDate')
                m_list = m_sub['month'].tolist()
            else:
                m_list = pm_sub['month'].unique().tolist()
                
        if not issue_df.empty and 'quarterLabel' in issue_df.columns:
            i_m_list = issue_df[issue_df['quarterLabel'] == label]['month'].unique().tolist()
            for m in i_m_list:
                if m not in m_list:
                    m_list.append(m)

        quarters[label] = m_list
        for m in m_list:
            if m not in all_months_set:
                all_months_set.append(m)

    res_quarters = {'All Quarters': all_months_set}
    res_quarters.update(quarters)
    return res_quarters

# Aggregation & Math Engine

def compute_zme_issue_table(issue_df, selected_months):
    if issue_df.empty:
        return pd.DataFrame()
    issues = issue_df[issue_df['month'].isin(selected_months)] if selected_months else issue_df
    if issues.empty:
        return pd.DataFrame()

    agg = issues.groupby(['zone', 'zme']).agg(
        total=('zme', 'size'),
        open=('_Is_Open_', 'sum'),
        closed=('_Is_Closed_', 'sum'),
        within=('_Is_Closed_Within_', 'sum'),
        outside=('_Is_Closed_Without_', 'sum'),
    ).reset_index()

    # Formulas from Executive Model:
    # 1. Overall CM Efficiency % = Closed Faults / Faults Registered * 100
    # 2. CM-TAT Efficiency % = Closed Within TAT / Total Closed Faults * 100
    agg['cm_efficiency'] = (agg['closed'] / agg['total']).fillna(0.0) * 100
    agg['tat_efficiency'] = (agg['within'] / agg['closed'].replace(0, pd.NA)).fillna(0.0) * 100
    return agg.sort_values(by='total', ascending=False)


def compute_zme_pm_table(pm_df, selected_months):
    if pm_df.empty:
        return pd.DataFrame()
    pm = pm_df[pm_df['month'].isin(selected_months)] if selected_months else pm_df
    if pm.empty:
        return pd.DataFrame()

    agg = pm.groupby(['zone', 'zme']).agg(
        total_chargers=('ocppId', 'count'),
        total_stations=('stationId', 'nunique'),
        planning=('Is_PM_Planned', 'sum'),
        done=('Is_PM_Done', 'sum'),
        pending=('Is_PM_Pending', 'sum'),
        advance=('Advance PM Done', 'sum'),
    ).reset_index()

    # Formula: PM Efficiency % = PM Done / PM Planning * 100
    agg['pm_efficiency'] = (agg['done'] / agg['planning'].replace(0, pd.NA)).fillna(0.0) * 100
    return agg.sort_values(by='planning', ascending=False)


@st.cache_data(show_spinner=False)
def generate_mpr_excel_report(issue_df, pm_df, selected_months):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        kpi_summary = []
        if not issue_df.empty and selected_months:
            sel_i = issue_df[issue_df['month'].isin(selected_months)]
            tot_i = len(sel_i)
            open_i = int(sel_i['_Is_Open_'].sum()) if not sel_i.empty else 0
            closed_i = int(sel_i['_Is_Closed_'].sum()) if not sel_i.empty else 0
            within_tat = int(sel_i['_Is_Closed_Within_'].sum()) if not sel_i.empty else 0
            cm_eff = (closed_i / tot_i * 100) if tot_i > 0 else 0.0
            tat_eff = (within_tat / closed_i * 100) if closed_i > 0 else 0.0
            kpi_summary.extend([
                {'Metric Category': 'Issue SLA Governance', 'Metric Name': 'Faults Registered', 'Value': tot_i},
                {'Metric Category': 'Issue SLA Governance', 'Metric Name': 'Open Faults', 'Value': open_i},
                {'Metric Category': 'Issue SLA Governance', 'Metric Name': 'Closed Faults', 'Value': closed_i},
                {'Metric Category': 'Issue SLA Governance', 'Metric Name': 'Closed Within TAT', 'Value': within_tat},
                {'Metric Category': 'Issue SLA Governance', 'Metric Name': 'Overall CM Efficiency %', 'Value': f"{cm_eff:.1f}%"},
                {'Metric Category': 'Issue SLA Governance', 'Metric Name': 'CM-TAT Efficiency %', 'Value': f"{tat_eff:.1f}%"},
            ])

        if not pm_df.empty and selected_months:
            sel_p = pm_df[pm_df['month'].isin(selected_months)]
            planned_p = int(sel_p['Is_PM_Planned'].sum()) if not sel_p.empty else 0
            done_p = int(sel_p['Is_PM_Done'].sum()) if not sel_p.empty else 0
            pending_p = int(sel_p['Is_PM_Pending'].sum()) if not sel_p.empty else 0
            adv_p = int(sel_p['Advance PM Done'].sum()) if not sel_p.empty else 0
            pm_eff = (done_p / planned_p * 100) if planned_p > 0 else 0.0
            kpi_summary.extend([
                {'Metric Category': 'PM Governance', 'Metric Name': 'PM Planning (Scheduled)', 'Value': planned_p},
                {'Metric Category': 'PM Governance', 'Metric Name': 'PM Done (Completed)', 'Value': done_p},
                {'Metric Category': 'PM Governance', 'Metric Name': 'PM Pending', 'Value': pending_p},
                {'Metric Category': 'PM Governance', 'Metric Name': 'Advance PM Done', 'Value': adv_p},
                {'Metric Category': 'PM Governance', 'Metric Name': 'PM Efficiency %', 'Value': f"{pm_eff:.1f}%"},
            ])

        if kpi_summary:
            pd.DataFrame(kpi_summary).to_excel(writer, sheet_name='Executive Summary', index=False)

        if not issue_df.empty and selected_months:
            sel_issues = issue_df[issue_df['month'].isin(selected_months)]
            if not sel_issues.empty:
                cols_clean = [c for c in sel_issues.columns if not c.startswith('_')]
                sel_issues[cols_clean].to_excel(writer, sheet_name='Issue Tracker', index=False)

        if not pm_df.empty and selected_months:
            sel_pm = pm_df[pm_df['month'].isin(selected_months)]
            if not sel_pm.empty:
                pm_planned = sel_pm[sel_pm['Is_PM_Planned']]
                if not pm_planned.empty:
                    cols_p = ['ocppId', 'zme', 'zone', 'stationId', 'stationName', 'dueDate', 'month', 'PM Compliance Status']
                    pm_planned[[c for c in cols_p if c in pm_planned.columns]].to_excel(writer, sheet_name='PM Planned', index=False)

                pm_done = sel_pm[sel_pm['Is_PM_Done']]
                if not pm_done.empty:
                    cols_d = ['ocppId', 'zme', 'zone', 'stationId', 'stationName', 'status', 'completionDate', 'month', 'PM Compliance Status']
                    pm_done[[c for c in cols_d if c in pm_done.columns]].to_excel(writer, sheet_name='PM Done', index=False)

                pm_pending = sel_pm[sel_pm['Is_PM_Pending']]
                if not pm_pending.empty:
                    cols_pend = ['ocppId', 'zme', 'zone', 'stationId', 'stationName', 'status', 'dueDate', 'month', 'PM Compliance Status']
                    pm_pending[[c for c in cols_pend if c in pm_pending.columns]].to_excel(writer, sheet_name='PM Pending', index=False)

                zme_summary = compute_zme_pm_table(pm_df, selected_months)
                if not zme_summary.empty:
                    zme_summary.to_excel(writer, sheet_name='ZME PM Summary', index=False)

    output.seek(0)
    return output.getvalue()


# Plotly Charts

RED_PALETTE = ['#991B1B', '#DC2626', '#EF4444', '#F87171', '#FCA5A5', '#7F1D1D', '#B91C1C']

def plot_interactive_bar(df, x_col, y_col, title, color_hex="#DC2626"):
    fig = px.bar(
        df, x=x_col, y=y_col, title=title, text=y_col,
        color_discrete_sequence=[color_hex]
    )
    fig.update_traces(
        texttemplate='%{text:,}', textposition='outside',
        textfont=dict(size=12, family='Inter'),
        hovertemplate='<b>%{x}</b><br>Count: <b>%{y:,}</b><extra></extra>'
    )
    fig.update_layout(
        margin=dict(l=15, r=15, t=45, b=15), height=340,
        font=dict(family='Inter', color='#0F172A'),
        title=dict(font=dict(size=15, color='#991B1B', family='Inter', weight='bold')),
        xaxis_title=x_col, yaxis_title=y_col,
        plot_bgcolor='rgba(0,0,0,0)',
        yaxis=dict(showgrid=True, gridcolor='#FEE2E2')
    )
    return fig


def plot_grouped_bar(df, x_col, y_cols, title, colors=None):
    fig = go.Figure()
    palette = colors if colors else ['#991B1B', '#DC2626', '#EF4444', '#7F1D1D']
    for idx, col in enumerate(y_cols):
        fig.add_trace(go.Bar(
            name=col, x=df[x_col], y=df[col],
            marker_color=palette[idx % len(palette)],
            text=df[col], textposition='auto',
            textfont=dict(size=11, family='Inter'),
            hovertemplate='<b>%{x}</b><br>' + col + ': <b>%{y:,}</b><extra></extra>'
        ))
    fig.update_layout(
        barmode='group', hovermode='x unified',
        title=dict(text=title, font=dict(size=15, color='#991B1B', family='Inter', weight='bold')),
        margin=dict(l=15, r=15, t=45, b=15), height=340,
        legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="center", x=0.5),
        plot_bgcolor='rgba(0,0,0,0)',
        yaxis=dict(showgrid=True, gridcolor='#FEE2E2')
    )
    return fig


def plot_donut_chart(labels, values, title, colors=None):
    fig = go.Figure(data=[go.Pie(
        labels=labels, values=values, hole=0.45,
        marker_colors=colors if colors else RED_PALETTE,
        textinfo='percent+value',
        hovertemplate='<b>%{label}</b><br>Count: <b>%{value:,}</b> (%{percent})<extra></extra>',
        insidetextfont=dict(color='#FFFFFF', size=12, family='Inter')
    )])
    fig.update_layout(
        title=dict(text=title, font=dict(size=15, color='#991B1B', family='Inter', weight='bold')),
        margin=dict(l=15, r=15, t=45, b=15), height=320,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5)
    )
    return fig

# Main Application UI

def main():
    # Executive Red & White Theme CSS
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
        
        .stApp, [data-testid="stMain"], [data-testid="stHeader"] {
            background-color: #F8FAFC !important;
        }
        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #E2E8F0 !important;
        }
        html, body, p, span, label, div {
            font-family: 'Inter', -apple-system, sans-serif;
            color: #1E293B !important;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #0F172A !important;
            font-weight: 800 !important;
            letter-spacing: -0.02em !important;
        }
        .stButton > button, div[data-testid="stDownloadButton"] > button {
            background: linear-gradient(135deg, #801B1B 0%, #B91C1C 50%, #DC2626 100%) !important;
            color: #FFFFFF !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            font-size: 0.92rem !important;
            border: none !important;
            padding: 0.65rem 1.4rem !important;
            box-shadow: 0 4px 14px rgba(185, 28, 28, 0.25) !important;
            transition: all 0.2s ease !important;
        }
        .stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {
            background: linear-gradient(135deg, #7F1D1D 0%, #991B1B 50%, #B91C1C 100%) !important;
            box-shadow: 0 6px 20px rgba(153, 27, 27, 0.38) !important;
            transform: translateY(-2px) !important;
        }
        button[data-baseweb="tab"] {
            font-weight: 700 !important;
            color: #64748B !important;
            font-size: 0.95rem !important;
            padding: 0.75rem 1.25rem !important;
            border-radius: 8px 8px 0 0 !important;
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
        }
        .exec-title {
            font-size: 2.1rem;
            font-weight: 900;
            color: #FFFFFF !important;
            margin: 0;
            line-height: 1.2;
        }
        .exec-subtitle {
            font-size: 0.95rem;
            color: #FEE2E2 !important;
            margin-top: 0.35rem;
            margin-bottom: 0;
            font-weight: 500;
        }
        .metric-card {
            background: #FFFFFF;
            border: 1px solid #FEE2E2;
            border-radius: 14px;
            padding: 1.15rem 1.35rem;
            box-shadow: 0 4px 14px rgba(153, 27, 27, 0.05);
            transition: all 0.25s ease;
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
            border-bottom: 2px solid #FEE2E2;
            padding-bottom: 0.45rem;
        }
        </style>
    """, unsafe_allow_html=True)

    # Executive Banner
    st.markdown("""
        <div class="exec-header-box">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;">
                <div>
                    <span class="exec-badge">⚡ CHARGEZONE EXECUTIVE BOARD • OPERATIONS DASHBOARD</span>
                    <h1 class="exec-title">Monthly Progress Report (MPR) & PM Governance</h1>
                    <p class="exec-subtitle">Interactive SLA analytics, Preventive Maintenance (PM F-01) tracking & PM Pending governance.</p>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Sidebar Controls
    st.sidebar.markdown("### ⚙️ Control Panel")
    st.sidebar.markdown("---")

    st.sidebar.markdown("**Upload Operational Files:**")
    issue_file = st.sidebar.file_uploader("1️⃣ Issue Tracker File", type=["xlsx", "csv"], key="sep_issue")
    pm_file = st.sidebar.file_uploader("2️⃣ PM Tracker File", type=["xlsx", "csv"], key="sep_pm")

    file_bytes_input = None
    if issue_file is not None or pm_file is not None:
        file_bytes_input = {
            'issue': issue_file.getvalue() if issue_file else None,
            'pm': pm_file.getvalue() if pm_file else None
        }

    if file_bytes_input is None:
        st.markdown("""
            <div style="background-color: #FFFFFF; border: 2px dashed #FCA5A5; border-radius: 16px; padding: 3rem 2rem; text-align: center; margin-top: 1.5rem; box-shadow: 0 4px 14px rgba(153, 27, 27, 0.04);">
                <div style="font-size: 3.5rem; margin-bottom: 0.8rem;">📥</div>
                <h2 style="color: #991B1B !important; font-weight: 800; margin-bottom: 0.5rem;">Clean Dashboard — Data Upload Required</h2>
                <p style="color: #475569 !important; font-size: 1.05rem; max-width: 620px; margin: 0 auto 1.8rem auto;">
                    Please upload your operational files (Issue Tracker & PM Tracker) using the <b>Control Panel</b> in the sidebar to populate interactive charts, SLA metrics, and PM Pending governance reports.
                </p>
                <div style="display: flex; justify-content: center; gap: 1.5rem; flex-wrap: wrap; text-align: left; max-width: 700px; margin: 0 auto; background: #FFF5F5; padding: 1.2rem 1.6rem; border-radius: 12px; border: 1px solid #FEE2E2;">
                    <div>
                        <b style="color: #991B1B;">1️⃣ Issue Tracker File</b>
                        <p style="margin: 0.2rem 0 0 0; font-size: 0.88rem; color: #64748B;">Upload your Issue Tracker file via the sidebar.</p>
                    </div>
                    <div style="border-left: 1px solid #FCA5A5; padding-left: 1.5rem;">
                        <b style="color: #991B1B;">2️⃣ PM Tracker File</b>
                        <p style="margin: 0.2rem 0 0 0; font-size: 0.88rem; color: #64748B;">Upload your PM Tracker file via the sidebar.</p>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        return

    with st.spinner("Parsing operational workbook data..."):
        try:
            issue_df, pm_df = load_workbook_data(file_bytes_input)
        except Exception as e:
            st.error(f"⚠️ Error parsing workbook: {e}")
            return

    quarters = build_quarters(issue_df, pm_df)
    quarter_labels = list(quarters.keys())

    st.sidebar.markdown("---")
    st.sidebar.markdown("#### 📅 Select Quarter & Months")

    # Default to the most recent quarter that actually has reported PM
    # status data, not "All Quarters" -- opening on a multi-year aggregate
    # by default made every KPI look inflated/confusing at first glance.
    reported_quarters = [
        q for q in quarter_labels
        if q != 'All Quarters' and not pm_df.empty
        and pm_df.loc[pm_df['quarterLabel'] == q, 'status'].notna().any()
    ]
    default_quarter = reported_quarters[-1] if reported_quarters else quarter_labels[0]
    default_idx = quarter_labels.index(default_quarter)

    selected_quarter = st.sidebar.selectbox("Quarter:", quarter_labels, index=default_idx)
    months_for_quarter = quarters[selected_quarter]

    selected_months = st.sidebar.multiselect("Months:", months_for_quarter, default=months_for_quarter)

    if not selected_months:
        selected_months = months_for_quarter

    st.sidebar.markdown("---")
    st.sidebar.markdown("#### 📥 Export Executive Report")
    month_slug = "_".join(selected_months[:3]) if len(selected_months) <= 3 else f"{selected_months[0]}_to_{selected_months[-1]}"
    report_filename = f"ChargeZone_MPR_Report_{selected_quarter}_{month_slug}.xlsx"

    report_bytes = generate_mpr_excel_report(issue_df, pm_df, tuple(selected_months))
    st.sidebar.download_button(
        label="📥 Download Executive Report (.xlsx)",
        data=report_bytes,
        file_name=report_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # Tabs
    tab_issues, tab_pm, tab_raw = st.tabs([
        "📉 Issue Dashboard (SLA Analytics)",
        "🛠️ PM Dashboard (PM F-01 & Pending)",
        "📋 Data Explorer"
    ])

    # TAB 1: ISSUE DASHBOARD
    with tab_issues:
        st.markdown('<div class="section-header">📊 Operational Issue & SLA Performance</div>', unsafe_allow_html=True)
        
        selected_issues = issue_df[issue_df['month'].isin(selected_months)] if not issue_df.empty and selected_months else issue_df

        # Executive Formulas:
        # Total Registered = len(selected_issues)
        # Open Faults = count(_Is_Open_)
        # Closed Faults = count(_Is_Closed_)
        # Closed Within TAT = count(_Is_Closed_Within_)
        # Closed Without TAT = count(_Is_Closed_Without_)
        # Overall CM Efficiency % = (Closed Faults / Faults Registered) * 100
        # CM-TAT Efficiency % = (Closed Within TAT / Total Closed Faults) * 100
        total_issues = len(selected_issues)
        total_open = int(selected_issues['_Is_Open_'].sum()) if not selected_issues.empty else 0
        total_closed = int(selected_issues['_Is_Closed_'].sum()) if not selected_issues.empty else 0
        closed_within = int(selected_issues['_Is_Closed_Within_'].sum()) if not selected_issues.empty else 0
        closed_without = int(selected_issues['_Is_Closed_Without_'].sum()) if not selected_issues.empty else 0

        overall_cm_eff = (total_closed / total_issues * 100) if total_issues > 0 else 0.0
        cm_tat_eff = (closed_within / total_closed * 100) if total_closed > 0 else 0.0

        # KPI Row (7 Cards)
        k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
        with k1:
            st.markdown(f'<div class="metric-card darkred"><div class="metric-label">Faults Registered</div><div class="metric-val">{total_issues:,}</div><div class="metric-sub">Total Logged</div></div>', unsafe_allow_html=True)
        with k2:
            st.markdown(f'<div class="metric-card amber"><div class="metric-label">Open Faults</div><div class="metric-val">{total_open:,}</div><div class="metric-sub">Pending / Active</div></div>', unsafe_allow_html=True)
        with k3:
            st.markdown(f'<div class="metric-card grey"><div class="metric-label">Closed Faults</div><div class="metric-val">{total_closed:,}</div><div class="metric-sub">Resolved</div></div>', unsafe_allow_html=True)
        with k4:
            st.markdown(f'<div class="metric-card green"><div class="metric-label">Closed Within TAT</div><div class="metric-val">{closed_within:,}</div><div class="metric-sub">SLA Compliant</div></div>', unsafe_allow_html=True)
        with k5:
            st.markdown(f'<div class="metric-card red"><div class="metric-label">Closed Without TAT</div><div class="metric-val">{closed_without:,}</div><div class="metric-sub">SLA Breached</div></div>', unsafe_allow_html=True)
        with k6:
            st.markdown(f'<div class="metric-card {"green" if overall_cm_eff >= 85 else "amber"}"><div class="metric-label">Overall CM Eff %</div><div class="metric-val">{overall_cm_eff:.1f}%</div><div class="metric-sub">Closed / Registered</div></div>', unsafe_allow_html=True)
        with k7:
            st.markdown(f'<div class="metric-card {"green" if cm_tat_eff >= 80 else "amber"}"><div class="metric-label">CM-TAT Eff %</div><div class="metric-val">{cm_tat_eff:.1f}%</div><div class="metric-sub">Within TAT / Closed</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 1. ZME Performance Table & Chart
        st.markdown('<div class="section-header">1. ZME Performance & SLA Compliance Breakdown</div>', unsafe_allow_html=True)
        zme_issue_table = compute_zme_issue_table(issue_df, selected_months)
        
        if not zme_issue_table.empty:
            col_zme_chart, col_zme_tbl = st.columns([6, 6])
            with col_zme_chart:
                fig_zme = plot_grouped_bar(
                    df=zme_issue_table.head(10),
                    x_col='zme', y_cols=['total', 'open', 'within', 'outside'],
                    title="Fault Breakdown for Top 10 ZMEs",
                    colors=['#991B1B', '#EF4444', '#DC2626', '#7F1D1D']
                )
                st.plotly_chart(fig_zme, use_container_width=True)
            with col_zme_tbl:
                st.write("##### ZME Summary Data Table")
                st.dataframe(
                    zme_issue_table.rename(columns={
                        'zme': 'ZME Name', 'zone': 'Zone', 'total': 'Faults Registered',
                        'open': 'Open Faults', 'closed': 'Closed Faults',
                        'within': 'Closed Within TAT', 'outside': 'Closed Without TAT',
                        'cm_efficiency': 'Overall CM Efficiency %', 'tat_efficiency': 'CM-TAT Efficiency %'
                    }).style.format({
                        'Faults Registered': '{:,}', 'Open Faults': '{:,}', 'Closed Faults': '{:,}',
                        'Closed Within TAT': '{:,}', 'Closed Without TAT': '{:,}',
                        'Overall CM Efficiency %': '{:.1f}%', 'CM-TAT Efficiency %': '{:.1f}%'
                    }).background_gradient(subset=['Overall CM Efficiency %', 'CM-TAT Efficiency %'], cmap='Reds'),
                    use_container_width=True, height=320
                )
        else:
            st.info("No issue records found for selected month(s). Try selecting 'All Quarters'.")

        st.markdown("---")

        # 2. Zone Performance Breakdown & Drill-Down Inspector
        st.markdown('<div class="section-header">🏢 2. Zone Performance & Efficiency</div>', unsafe_allow_html=True)
        if not selected_issues.empty:
            zone_agg = selected_issues.groupby('zone').agg(
                Registered=('zme', 'size'),
                Open=('_Is_Open_', 'sum'),
                Closed=('_Is_Closed_', 'sum'),
                Within_TAT=('_Is_Closed_Within_', 'sum'),
                Without_TAT=('_Is_Closed_Without_', 'sum'),
            ).reset_index()
            # New Formulas:
            # Overall CM Efficiency % = Closed / Registered * 100
            # CM-TAT Efficiency % = Within_TAT / Closed * 100
            zone_agg['Overall CM Efficiency %'] = (zone_agg['Closed'] / zone_agg['Registered'] * 100).round(1)
            zone_agg['CM-TAT Efficiency %'] = (zone_agg['Within_TAT'] / zone_agg['Closed'].replace(0, pd.NA) * 100).fillna(0.0).round(1)

            col_z_c, col_z_t = st.columns([6, 6])
            with col_z_c:
                st.plotly_chart(
                    plot_grouped_bar(
                        zone_agg, x_col='zone',
                        y_cols=['Registered', 'Open', 'Within_TAT', 'Without_TAT'],
                        title="Zone Performance Comparison",
                        colors=['#991B1B', '#EF4444', '#DC2626', '#7F1D1D']
                    ), use_container_width=True
                )
            with col_z_t:
                st.write("##### Zone Summary Data Table")
                st.dataframe(
                    zone_agg.rename(columns={'zone': 'Zone'}).style.format({
                        'Registered': '{:,}', 'Open': '{:,}', 'Closed': '{:,}',
                        'Within_TAT': '{:,}', 'Without_TAT': '{:,}',
                        'Overall CM Efficiency %': '{:.1f}%', 'CM-TAT Efficiency %': '{:.1f}%'
                    }).background_gradient(subset=['Overall CM Efficiency %', 'CM-TAT Efficiency %'], cmap='Reds'),
                    use_container_width=True, height=300
                )

        st.markdown("---")

        # 3. Status Pipeline Breakdown
        st.markdown('<div class="section-header">📌 3. Status Breakdown</div>', unsafe_allow_html=True)
        if not selected_issues.empty:
            st_counts = selected_issues['status'].value_counts().reset_index()
            st_counts.columns = ['Status', 'Count']
            col_st_c, col_st_t = st.columns([7, 5])
            with col_st_c:
                st.plotly_chart(plot_interactive_bar(st_counts, 'Status', 'Count', 'Status Pipeline', '#991B1B'), use_container_width=True)
            with col_st_t:
                st.write("##### Status Pipeline Table")
                st.dataframe(st_counts.style.format({'Count': '{:,}'}).background_gradient(subset=['Count'], cmap='Reds'), use_container_width=True, height=280)

        st.markdown("---")

        # 5. Repetitive Faults & Top Issue / Sub-Issue Explorer
        st.markdown('<div class="section-header">⚠️ 5. Repetitive Faults, Top 10 Issue Types & Top 5 Sub-Issues Analytics</div>', unsafe_allow_html=True)
        
        if not selected_issues.empty:
            stn_c = find_col(selected_issues, ['Station ID', 'Station_ID', 'Station', 'Site ID']) or 'stationId'
            stn_name_c = find_col(selected_issues, ['Station Name', 'Station_Name', 'Site Name']) or 'stationName'
            zme_c = find_col(selected_issues, ['ZME', 'ZME Name', 'ZME_Name', 'Zone Manager']) or 'zme'
            zone_c = find_col(selected_issues, ['Zone', 'Zone Name', 'Region']) or 'zone'
            ocpp_c = find_col(selected_issues, ['OCPP ID', 'OCPP_ID', 'Charger ID', 'Connector ID', 'EVSE ID', 'OCPP']) or 'ocppId'
            type_c = find_col(selected_issues, ['Issue Type', 'Issue_Type', 'Category']) or 'issueType'
            sub_c = find_col(selected_issues, ['Issue Sub-Type', 'Issue Sub Type', 'Sub Type', 'Fault Subtype', 'Issue Subtype']) or 'issueSubType'

            # A. Top 10 Overall Repetitive Issue Types & Top 5 Sub-Types Explorer
            if type_c in selected_issues.columns:
                st.markdown("##### ⚡ Top 10 Overall Repetitive Issue Types & Top 5 Sub-Types Explorer")
                st.caption("Click on any Issue Type below to view its Top 5 Sub-Types and affected station/charger records:")

                type_totals = selected_issues.groupby(type_c).size().reset_index(name='Total_Faults')
                type_totals = type_totals.sort_values(by='Total_Faults', ascending=False).head(10)

                c_type_chart, c_type_tbl = st.columns([6, 6])
                with c_type_chart:
                    st.plotly_chart(
                        plot_interactive_bar(
                            df=type_totals,
                            x_col=type_c,
                            y_col='Total_Faults',
                            title="Top 10 Overall Repetitive Issue Types",
                            color_hex="#991B1B"
                        ),
                        use_container_width=True
                    )

                with c_type_tbl:
                    st.write("**Top 10 Issue Types Summary Table:**")
                    type_totals['Share %'] = (type_totals['Total_Faults'] / len(selected_issues) * 100).round(1)
                    st.dataframe(
                        type_totals.rename(columns={type_c: 'Issue Type'}).style.format({'Total_Faults': '{:,}', 'Share %': '{:.1f}%'}).background_gradient(subset=['Total_Faults'], cmap='Reds'),
                        use_container_width=True,
                        height=260
                    )

                # Expanders for each Top 10 Issue Type showing Top 5 Sub-Types
                for _, row_t in type_totals.iterrows():
                    t_name = str(row_t[type_c]).strip()
                    t_count = int(row_t['Total_Faults'])

                    df_type_sub = selected_issues[selected_issues[type_c].astype(str).str.strip() == t_name]

                    if sub_c in df_type_sub.columns:
                        sub_counts = df_type_sub.groupby(sub_c).size().reset_index(name='Subtype_Count')
                        sub_counts = sub_counts.sort_values(by='Subtype_Count', ascending=False).head(5)
                    else:
                        sub_counts = pd.DataFrame()

                    expander_title = f"🔧 Issue Type: {t_name} — Total {t_count:,} Occurrences ({len(sub_counts)} Top Sub-Types)"
                    with st.expander(expander_title, expanded=False):
                        if not sub_counts.empty:
                            c_sub_chart, c_sub_tbl = st.columns([6, 6])
                            with c_sub_chart:
                                st.plotly_chart(
                                    plot_interactive_bar(
                                        df=sub_counts,
                                        x_col=sub_c,
                                        y_col='Subtype_Count',
                                        title=f"Top 5 Sub-Types for '{t_name}'",
                                        color_hex="#DC2626"
                                    ),
                                    use_container_width=True
                                )
                            with c_sub_tbl:
                                st.write(f"📌 **Top 5 Sub-Types Table for {t_name}:**")
                                sub_counts['Sub-Type Share %'] = (sub_counts['Subtype_Count'] / t_count * 100).round(1)
                                st.dataframe(
                                    sub_counts.rename(columns={sub_c: 'Issue Sub-Type'}).style.format({'Subtype_Count': '{:,}', 'Sub-Type Share %': '{:.1f}%'}).background_gradient(subset=['Subtype_Count'], cmap='Reds'),
                                    use_container_width=True,
                                    height=220
                                )

                            st.write(f"📝 **Underlying Ticket Records for Issue Type '{t_name}':**")
                            stat_c = find_col(df_type_sub, ['Status', 'Ticket Status', 'Issue Status']) or 'status'
                            tat_c = find_col(df_type_sub, ['TAT Compliance', 'SLA Compliance']) or 'tatCompliance'
                            type_disp_cols = [c for c in [ocpp_c, stn_c, stn_name_c, sub_c, stat_c, tat_c, zme_c, zone_c] if c and c in df_type_sub.columns]
                            if type_disp_cols:
                                st.dataframe(df_type_sub[type_disp_cols], use_container_width=True, height=200)

                st.markdown("---")

            # B. Overall Repetitive Issue Details Table (Station ID, OCPP ID, Station Name, ZME, Sub-Type, Occurrences >= 2)
            st.markdown("##### 📋 Overall Repetitive Issue Details Table (Occurrences ≥ 2)")
            st.caption("Comprehensive breakdown of repetitive faults per Station ID, OCPP ID, Station Name, ZME & Issue Sub-Type:")

            rep_group_cols = []
            for c in [type_c, sub_c, stn_c, ocpp_c, stn_name_c, zme_c, zone_c]:
                if c and c in selected_issues.columns and c not in rep_group_cols:
                    rep_group_cols.append(c)

            if rep_group_cols:
                rep_details = selected_issues.groupby(rep_group_cols).size().reset_index(name='Occurrences')
                rep_details = rep_details[rep_details['Occurrences'] >= 2].sort_values(by='Occurrences', ascending=False)

                if not rep_details.empty:
                    rename_dict = {
                        type_c: 'Issue Type',
                        sub_c: 'Issue Sub-Type',
                        stn_c: 'Station ID',
                        ocpp_c: 'OCPP ID',
                        stn_name_c: 'Station Name',
                        zme_c: 'ZME Name',
                        zone_c: 'Zone'
                    }
                    disp_rep_df = rep_details.rename(columns={k: v for k, v in rename_dict.items() if k in rep_details.columns})
                    st.dataframe(
                        disp_rep_df.style.format({'Occurrences': '{:,}'}).background_gradient(subset=['Occurrences'], cmap='Reds'),
                        use_container_width=True,
                        height=350
                    )
                    st.caption(f"Showing **{len(disp_rep_df):,}** repetitive fault combinations (Occurrences ≥ 2)")
                else:
                    st.success("✅ Zero repetitive station faults found in current selection.")

    # TAB 2: PM DASHBOARD (PM F-01 & PENDING GOVERNANCE)
    with tab_pm:
        st.markdown('<div class="section-header">🛠️ Preventive Maintenance (PM F-01) Analytics & PM Pending Governance</div>', unsafe_allow_html=True)
        
        # Sidebar Status Filter for PM Governance
        st.sidebar.markdown("---")
        st.sidebar.markdown("#### ⚡ PM Charger Status Scope")
        pm_scope = st.sidebar.radio(
            "PM Scope Filter:",
            ["Live Chargers Only (Standard)", "All Chargers (Include Decom/Offline)"],
            index=0,
            key="pm_scope_filter"
        )

        base_pm = pm_df.copy() if not pm_df.empty else pd.DataFrame()
        if not base_pm.empty and "Live Chargers Only" in pm_scope and 'stationStatus' in base_pm.columns:
            base_pm = base_pm[base_pm['stationStatus'].astype(str).str.strip().str.upper() == 'LIVE']

        # Dynamic filtering based on selected months
        selected_pm = base_pm[base_pm['month'].isin(selected_months)] if not base_pm.empty and selected_months else base_pm

        # Dynamic KPI calculations for selected month(s)
        total_chargers = len(selected_pm['ocppId'].dropna()) if not selected_pm.empty else 0
        total_stations = selected_pm['stationId'].nunique() if not selected_pm.empty else 0
        total_pm_planning = int(selected_pm['Is_PM_Planned'].sum()) if not selected_pm.empty else 0
        pm_done = int(selected_pm['Is_PM_Done'].sum()) if not selected_pm.empty else 0
        pm_pending = int(selected_pm['Is_PM_Pending'].sum()) if not selected_pm.empty else 0
        advance_done = int(selected_pm['Advance PM Done'].sum()) if not selected_pm.empty else 0
        pm_eff = (pm_done / total_pm_planning * 100) if total_pm_planning > 0 else 0.0

        # PM KPI Cards Row (7 Cards dynamically updating with month selection)
        p1, p2, p3, p4, p5, p6, p7 = st.columns(7)
        with p1:
            st.markdown(f'<div class="metric-card darkred"><div class="metric-label">Live Chargers</div><div class="metric-val">{total_chargers:,}</div><div class="metric-sub">OCPP ID Count</div></div>', unsafe_allow_html=True)
        with p2:
            st.markdown(f'<div class="metric-card grey"><div class="metric-label">Live Stations</div><div class="metric-val">{total_stations:,}</div><div class="metric-sub">Unique Sites</div></div>', unsafe_allow_html=True)
        with p3:
            st.markdown(f'<div class="metric-card green"><div class="metric-label">PM Planning</div><div class="metric-val">{total_pm_planning:,}</div><div class="metric-sub">Scheduled Orders</div></div>', unsafe_allow_html=True)
        with p4:
            st.markdown(f'<div class="metric-card green"><div class="metric-label">PM Done</div><div class="metric-val">{pm_done:,}</div><div class="metric-sub">Completed</div></div>', unsafe_allow_html=True)
        with p5:
            st.markdown(f'<div class="metric-card red"><div class="metric-label">PM Pending</div><div class="metric-val">{pm_pending:,}</div><div class="metric-sub">Scheduled Pending</div></div>', unsafe_allow_html=True)
        with p6:
            st.markdown(f'<div class="metric-card amber"><div class="metric-label">Advance PM Done</div><div class="metric-val">{advance_done:,}</div><div class="metric-sub">Early Completed</div></div>', unsafe_allow_html=True)
        with p7:
            st.markdown(f'<div class="metric-card {"green" if pm_eff >= 90 else "amber"}"><div class="metric-label">PM Efficiency %</div><div class="metric-val">{pm_eff:.1f}%</div><div class="metric-sub">Target: ≥ 90%</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 1. PM Execution Overview Bar & Table
        st.markdown('<div class="section-header">📊 1. PM Execution Status & Pending Breakdown (Selected Months)</div>', unsafe_allow_html=True)
        exec_df = pd.DataFrame({
            'Metric': ['PM Planning', 'PM Done', 'PM Pending', 'Advance PM Done'],
            'Count': [total_pm_planning, pm_done, pm_pending, advance_done]
        })
        
        col_ex_c, col_ex_t = st.columns([7, 5])
        with col_ex_c:
            st.plotly_chart(
                plot_interactive_bar(exec_df, 'Metric', 'Count', 'PM Work Orders Execution Distribution', '#991B1B'),
                use_container_width=True
            )
        with col_ex_t:
            st.write("##### PM Execution Summary Table")
            exec_df['Share %'] = (exec_df['Count'] / total_pm_planning * 100).fillna(0.0).round(1) if total_pm_planning > 0 else 0.0
            st.dataframe(
                exec_df.style.format({'Count': '{:,}', 'Share %': '{:.1f}%'}).background_gradient(subset=['Count'], cmap='Reds'),
                use_container_width=True, height=250
            )

        st.markdown("---")

        # 2. Zone PM Summary Table & Chart
        st.markdown('<div class="section-header">🏢 2. Zone-wise PM Planning, Done & PM Pending (Selected Months)</div>', unsafe_allow_html=True)
        if not selected_pm.empty:
            zone_pm = selected_pm.groupby('zone').agg(
                chargers=('ocppId', 'count'),
                stations=('stationId', 'nunique'),
                planning=('Is_PM_Planned', 'sum'),
                done=('Is_PM_Done', 'sum'),
                pending=('Is_PM_Pending', 'sum'),
                advance=('Advance PM Done', 'sum')
            ).reset_index()
            zone_pm['PM Efficiency %'] = (zone_pm['done'] / zone_pm['planning'] * 100).round(1)

            col_zp_c, col_zp_t = st.columns([6, 6])
            with col_zp_c:
                st.plotly_chart(
                    plot_grouped_bar(
                        zone_pm, x_col='zone',
                        y_cols=['planning', 'done', 'pending', 'advance'],
                        title="PM Planning vs PM Done vs PM Pending by Zone",
                        colors=['#991B1B', '#16A34A', '#DC2626', '#EF4444']
                    ), use_container_width=True
                )
            with col_zp_t:
                st.write("##### Zone Summary Data Table")
                st.dataframe(
                    zone_pm.rename(columns={
                        'zone': 'Zone', 'chargers': 'Chargers', 'stations': 'Stations',
                        'planning': 'PM Planning', 'done': 'PM Done', 'pending': 'PM Pending',
                        'advance': 'Advance PM'
                    }).style.format({
                        'Chargers': '{:,}', 'Stations': '{:,}', 'PM Planning': '{:,}',
                        'PM Done': '{:,}', 'PM Pending': '{:,}', 'Advance PM': '{:,}',
                        'PM Efficiency %': '{:.1f}%'
                    }).background_gradient(subset=['PM Pending', 'PM Efficiency %'], cmap='Reds'),
                    use_container_width=True, height=280
                )

        st.markdown("---")

        # 3. Customer Segment Breakdown (B2B / B2C)
        st.markdown('<div class="section-header">👥 3. Customer Segment Breakdown (B2B / B2C)</div>', unsafe_allow_html=True)
        if not selected_pm.empty and 'segment' in selected_pm.columns:
            seg_pm = selected_pm.groupby('segment').agg(
                Chargers=('ocppId', 'count'),
                Stations=('stationId', 'nunique'),
                Planned=('Is_PM_Planned', 'sum'),
                Done=('Is_PM_Done', 'sum'),
                Pending=('Is_PM_Pending', 'sum'),
            ).reset_index()
            seg_pm['PM Efficiency %'] = (seg_pm['Done'] / seg_pm['Planned'] * 100).fillna(0.0).round(1)

            col_seg_c, col_seg_t = st.columns([6, 6])
            with col_seg_c:
                st.plotly_chart(
                    plot_donut_chart(seg_pm['segment'].tolist(), seg_pm['Planned'].tolist(), 'PM Planned Distribution by Customer Segment'),
                    use_container_width=True
                )
            with col_seg_t:
                st.write("##### Customer Segment PM Summary Data Table")
                st.dataframe(
                    seg_pm.rename(columns={'segment': 'Customer Segment'}).style.format({
                        'Chargers': '{:,}', 'Stations': '{:,}', 'Planned': '{:,}',
                        'Done': '{:,}', 'Pending': '{:,}', 'PM Efficiency %': '{:.1f}%'
                    }).background_gradient(subset=['Pending', 'PM Efficiency %'], cmap='Reds'),
                    use_container_width=True, height=250
                )

        st.markdown("---")

        # 4. Detailed ZME PM Table
        st.markdown('<div class="section-header">⚙️ 4. Detailed PM Summary by ZME (Selected Months)</div>', unsafe_allow_html=True)
        zme_pm_table = compute_zme_pm_table(base_pm, selected_months)
        if not zme_pm_table.empty:
            st.dataframe(
                zme_pm_table.rename(columns={
                    'zme': 'ZME Name', 'zone': 'Zone', 'total_chargers': 'Total Chargers',
                    'total_stations': 'Total Stations', 'planning': 'PM Planning',
                    'done': 'PM Done', 'pending': 'PM Pending', 'advance': 'Advance PM',
                    'pm_efficiency': 'PM Efficiency (%)'
                }).style.format({
                    'Total Chargers': '{:,}', 'Total Stations': '{:,}', 'PM Planning': '{:,}',
                    'PM Done': '{:,}', 'PM Pending': '{:,}', 'Advance PM': '{:,}',
                    'PM Efficiency (%)': '{:.1f}%'
                }).background_gradient(subset=['PM Pending', 'PM Efficiency (%)'], cmap='Reds'),
                use_container_width=True, height=320
            )

        st.markdown("---")

        # 5. Period Breakdown (Month, Quarter, Year, Date)
        st.markdown('<div class="section-header">🗓️ 5. PM Scheduled Station Breakdown by Selection (Month, Quarter, Year, Date)</div>', unsafe_allow_html=True)
        if not selected_pm.empty:
            period_choice = st.radio("Group PM Scheduled Stations by:", ["Month", "Quarter", "Year", "Date"], horizontal=True, key="pm_period_choice")
            period_col_map = {
                "Month": "Scheduled Month",
                "Quarter": "Scheduled Quarter",
                "Year": "Scheduled Year",
                "Date": "Scheduled Date"
            }
            target_col = period_col_map[period_choice]

            if target_col in selected_pm.columns:
                period_agg = selected_pm.groupby(target_col).agg(
                    stations=('stationId', 'nunique'),
                    chargers=('ocppId', 'count'),
                    planning=('Is_PM_Planned', 'sum'),
                    done=('Is_PM_Done', 'sum'),
                    pending=('Is_PM_Pending', 'sum'),
                ).reset_index()
                period_agg['PM Efficiency %'] = (period_agg['done'] / period_agg['planning'] * 100).round(1)

                col_p_c, col_p_t = st.columns([6, 6])
                with col_p_c:
                    st.plotly_chart(
                        plot_grouped_bar(
                            period_agg, x_col=target_col,
                            y_cols=['stations', 'done', 'pending'],
                            title=f"PM Work Orders by {period_choice}",
                            colors=['#991B1B', '#16A34A', '#DC2626']
                        ), use_container_width=True
                    )
                with col_p_t:
                    st.write(f"##### Scheduled Stations Data Table ({period_choice} Level)")
                    st.dataframe(
                        period_agg.rename(columns={
                            target_col: f'Period ({period_choice})',
                            'stations': 'Stations Scheduled', 'chargers': 'Chargers Scheduled',
                            'planning': 'PM Planning', 'done': 'PM Done', 'pending': 'PM Pending'
                        }).style.format({
                            'Stations Scheduled': '{:,}', 'Chargers Scheduled': '{:,}',
                            'PM Planning': '{:,}', 'PM Done': '{:,}', 'PM Pending': '{:,}',
                            'PM Efficiency %': '{:.1f}%'
                        }).background_gradient(subset=['PM Efficiency %'], cmap='Reds'),
                        use_container_width=True, height=280
                    )

        st.markdown("---")

        # 5. Export PM Planned & PM Done Data Tables
        st.markdown('<div class="section-header">📥 5. Data Tables (PM Planned & PM Done)</div>', unsafe_allow_html=True)
        if not selected_pm.empty:
            sub_tab_planned, sub_tab_done, sub_tab_pending = st.tabs([
                "📋 PM Planned Table",
                "✅ PM Done Table",
                "⏳ PM Pending Table"
            ])

            with sub_tab_planned:
                planned_records = selected_pm[selected_pm['Is_PM_Planned']].copy()
                cols_p = [c for c in ['ocppId', 'zme', 'zone', 'stationId', 'stationName', 'dueDate', 'month', 'PM Compliance Status'] if c in planned_records.columns]
                disp_planned = planned_records[cols_p]
                
                c_p_info, c_p_dl = st.columns([8, 4])
                with c_p_info:
                    st.write(f"Showing **{len(planned_records):,}** PM Planned chargers for selected month(s):")
                with c_p_dl:
                    st.download_button(
                        "📥 Download PM Planned CSV",
                        data=disp_planned.to_csv(index=False).encode('utf-8'),
                        file_name=f"PM_Planned_{month_slug}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                st.dataframe(disp_planned, use_container_width=True, height=280)

            with sub_tab_done:
                done_records = selected_pm[selected_pm['Is_PM_Done']].copy()
                cols_d = [c for c in ['ocppId', 'zme', 'zone', 'stationId', 'stationName', 'status', 'completionDate', 'month', 'PM Compliance Status'] if c in done_records.columns]
                disp_done = done_records[cols_d]

                c_d_info, c_d_dl = st.columns([8, 4])
                with c_d_info:
                    st.write(f"Showing **{len(done_records):,}** PM Done chargers for selected month(s):")
                with c_d_dl:
                    st.download_button(
                        "📥 Download PM Done CSV",
                        data=disp_done.to_csv(index=False).encode('utf-8'),
                        file_name=f"PM_Done_{month_slug}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                st.dataframe(disp_done, use_container_width=True, height=280)

            with sub_tab_pending:
                pending_records = selected_pm[selected_pm['Is_PM_Pending']].copy()
                cols_pend = [c for c in ['ocppId', 'zme', 'zone', 'stationId', 'stationName', 'status', 'dueDate', 'month', 'PM Compliance Status'] if c in pending_records.columns]
                disp_pending = pending_records[cols_pend]

                c_pend_info, c_pend_dl = st.columns([8, 4])
                with c_pend_info:
                    st.write(f"Showing **{len(pending_records):,}** PM Pending chargers for selected month(s):")
                with c_pend_dl:
                    st.download_button(
                        "📥 Download PM Pending CSV",
                        data=disp_pending.to_csv(index=False).encode('utf-8'),
                        file_name=f"PM_Pending_{month_slug}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                st.dataframe(disp_pending, use_container_width=True, height=280)

    # TAB 3: DATA EXPLORER
    with tab_raw:
        st.markdown('<div class="section-header">🔍 Raw Operational Data Explorer</div>', unsafe_allow_html=True)
        dataset_choice = st.radio("Select Dataset:", ["Issue Tracker Dataset", "PM Tracker Dataset"], horizontal=True)

        if dataset_choice == "Issue Tracker Dataset":
            if not issue_df.empty:
                st.write(f"Showing **{len(issue_df):,}** Issue records:")
                st.dataframe(issue_df, use_container_width=True)
            else:
                st.info("Issue Tracker dataset is empty.")
        else:
            if not pm_df.empty:
                st.write(f"Showing **{len(pm_df):,}** PM records:")
                st.dataframe(pm_df, use_container_width=True)
            else:
                st.info("PM Tracker dataset is empty.")


if __name__ == '__main__':
    main()
