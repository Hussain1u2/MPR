#!/usr/bin/env python3
"""
ChargeZone MPR Dashboard - Streamlit App

Run:
    pip install -r requirements.txt
    streamlit run mpr_streamlit_app.py

Upload the merged Issue Tracker / PM Tracker workbook, tap a quarter,
multi-select months, and the ZME/Zone summary tables update live.
"""
from datetime import datetime
from io import BytesIO

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

st.set_page_config(page_title='ChargeZone MPR Dashboard', layout='wide')


# ---------------------------------------------------------------- parsing --

def fiscal_year_quarter(dt):
    fy_start = dt.year if dt.month >= 4 else dt.year - 1
    fy_label = f"FY{str(fy_start)[-2:]}{str(fy_start + 1)[-2:]}"
    quarter_num = (dt.month - 4) % 12 // 3 + 1
    return fy_label, quarter_num, f"{fy_label} Q{quarter_num}", dt.strftime('%b-%y')


def parse_issue_tracker(wb):
    rows = list(wb['Issue Tracker'].iter_rows(values_only=True))
    headers = [h.strip() if isinstance(h, str) else h for h in rows[0]]
    idx = {h: i for i, h in enumerate(headers)}
    records = []
    for r in rows[1:]:
        if r[idx['Sr No.']] is None and r[idx['Issue Id']] is None:
            continue
        issue_date = r[idx['Issue Date']]
        if not isinstance(issue_date, datetime):
            continue
        fy, qnum, qlabel, mlabel = fiscal_year_quarter(issue_date)
        records.append({
            'zme': r[idx['Manager Name']], 'zone': r[idx['Zone']], 'status': r[idx['Status']],
            'severity': r[idx['Severity']], 'tatCompliance': r[idx['TAT Compliance']],
            'issueType': r[idx['Issue Type']], 'issueSubType': r[idx['Issue Sub-Type']],
            'stationId': r[idx['Station ID']], 'segment': r[idx['B2B/ B2C']],
            'fy': fy, 'quarter': qnum, 'quarterLabel': qlabel, 'month': mlabel,
        })
    return pd.DataFrame(records)


def classify_pm_columns(row5):
    blocks, pending, current = [], [], None
    for i, h in enumerate(row5):
        text = str(h) if h else ''
        if 'Quarterly Schedule' in text or 'Repeatitive PM' in text or text in ('', 'Blank'):
            continue
        if 'Quarterly once compliance' in text:
            blocks.extend(pending)
            pending, current = [], None
            continue
        if 'F.E. Inspection' in text or 'HSE Inspection' in text or 'First Aid' in text:
            continue
        if 'Actual Completion Date' in text:
            if current is not None:
                current['completionCol'] = i
            continue
        if 'Status' in text:
            current = {'statusCol': i, 'completionCol': None}
            pending.append(current)
    return blocks


def parse_pm_tracker(wb):
    rows = list(wb['PM Tracker'].iter_rows(values_only=True))
    row4, row5 = rows[3], rows[4]
    station_cols = list(range(0, 13))
    station_fields = [row5[i] for i in station_cols]
    blocks = classify_pm_columns(row5)
    records = []
    for r in rows[5:]:
        if r[0] is None and r[10] is None:
            continue
        station = {name: r[i] for name, i in zip(station_fields, station_cols)}
        for b in blocks:
            due_date = row4[b['statusCol']]
            if not isinstance(due_date, datetime):
                continue
            fy, qnum, qlabel, mlabel = fiscal_year_quarter(due_date)
            completion = r[b['completionCol']] if b['completionCol'] is not None else None
            records.append({
                'zme': station.get('ZME'), 'zone': station.get('Zone'),
                'ocppId': station.get('OCPP ID'), 'stationId': station.get('Station ID'),
                'status': r[b['statusCol']], 'dueDate': due_date,
                'completionDate': completion if isinstance(completion, datetime) else None,
                'fy': fy, 'quarter': qnum, 'quarterLabel': qlabel, 'month': mlabel,
            })
    return pd.DataFrame(records)


@st.cache_data(show_spinner=False)
def load_workbook_data(file_bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    return parse_issue_tracker(wb), parse_pm_tracker(wb)


# ------------------------------------------------------------- aggregation --

def quarter_sort_key(fy, q):
    return (2000 + int(fy[2:4])) * 10 + q


def build_quarters(issue_df, pm_df):
    combo = pd.concat([issue_df[['fy', 'quarter', 'quarterLabel']],
                        pm_df[['fy', 'quarter', 'quarterLabel']]]).drop_duplicates()
    combo = combo.sort_values(by=['fy', 'quarter'], key=lambda s: s if s.name != 'quarterLabel' else s)
    combo['sort_key'] = combo.apply(lambda r: quarter_sort_key(r['fy'], r['quarter']), axis=1)
    combo = combo.sort_values('sort_key')
    quarters = {}
    for label in combo['quarterLabel']:
        months = pm_df.loc[pm_df['quarterLabel'] == label, ['month', 'dueDate']].drop_duplicates()
        months = months.sort_values('dueDate')['month'].tolist()
        quarters[label] = months
    return quarters


def compute_zme_table(issue_df, pm_df, months):
    all_zme = pd.concat([issue_df[['zme', 'zone']], pm_df[['zme', 'zone']]]).dropna().drop_duplicates('zme')
    issues = issue_df[issue_df['month'].isin(months)]
    pm = pm_df[pm_df['month'].isin(months)]

    i_agg = issues.groupby('zme').agg(
        total=('zme', 'size'),
        open=('status', lambda s: (s == 'Open').sum()),
        within=('tatCompliance', lambda s: (s == 'Yes').sum()),
        outside=('tatCompliance', lambda s: (s == 'No').sum()),
    )

    pm_reported = pm[pm['status'].notna() & (pm['status'] != '')]
    p_agg = pm_reported.groupby('zme').agg(
        planned=('zme', 'size'),
        done=('status', lambda s: (s == 'Yes').sum()),
        advance=('zme', lambda s: (
            (pm_reported.loc[s.index, 'completionDate'].notna()) &
            (pm_reported.loc[s.index, 'completionDate'] < pm_reported.loc[s.index, 'dueDate'])
        ).sum()),
    )

    table = all_zme.set_index('zme').join(i_agg).join(p_agg).fillna(0)
    for col in ['total', 'open', 'within', 'outside', 'planned', 'done', 'advance']:
        table[col] = table[col].astype(int)
    table['cm_efficiency'] = (table['within'] / table['total']).fillna(0).where(table['total'] > 0, 0)
    denom = table['within'] + table['outside']
    table['tat_efficiency'] = (table['within'] / denom).where(denom > 0, 0)
    table['pm_efficiency'] = (table['done'] / table['planned']).where(table['planned'] > 0, 0)
    return table.reset_index()


def zone_rollup(zme_table):
    zone = zme_table.groupby('zone').agg(
        total=('total', 'sum'), within=('within', 'sum'), outside=('outside', 'sum'),
        planned=('planned', 'sum'), done=('done', 'sum'),
    )
    zone['cm_efficiency'] = (zone['within'] / zone['total']).where(zone['total'] > 0, 0)
    denom = zone['within'] + zone['outside']
    zone['tat_efficiency'] = (zone['within'] / denom).where(denom > 0, 0)
    zone['pm_efficiency'] = (zone['done'] / zone['planned']).where(zone['planned'] > 0, 0)
    return zone.reset_index()


def monthly_trend(issue_df, pm_df):
    all_months = pm_df.groupby('month')['dueDate'].min().sort_values().index.tolist()
    rows = []
    for m in all_months:
        i = issue_df[issue_df['month'] == m]
        p = pm_df[(pm_df['month'] == m) & pm_df['status'].notna() & (pm_df['status'] != '')]
        within, outside = (i['tatCompliance'] == 'Yes').sum(), (i['tatCompliance'] == 'No').sum()
        planned, done = len(p), (p['status'] == 'Yes').sum()
        rows.append({
            'month': m, 'total_issues': len(i),
            'tat_efficiency': within / (within + outside) if (within + outside) > 0 else 0,
            'pm_efficiency': done / planned if planned > 0 else 0,
        })
    return pd.DataFrame(rows)


# ------------------------------------------------------------------- UI ----

PCT_COLS = {
    'cm_efficiency': 'CM Efficiency', 'tat_efficiency': 'TAT Efficiency', 'pm_efficiency': 'PM Efficiency',
}
COL_LABELS = {
    'zone': 'Zone', 'zme': 'ZME Name', 'total': 'Total', 'open': 'Open', 'within': 'Within TAT',
    'outside': 'Outside TAT', 'planned': 'PM Planned', 'done': 'PM Done', 'advance': 'Advance PM',
    **PCT_COLS,
}


def render_table(df, columns):
    view = df[columns].rename(columns=COL_LABELS)
    fmt = {COL_LABELS[c]: '{:.1%}' for c in columns if c in PCT_COLS}
    st.dataframe(view.style.format(fmt), use_container_width=True, hide_index=True)


def trend_chart(trend_df):
    fig = make_subplots(specs=[[{'secondary_y': True}]])
    fig.add_trace(go.Bar(x=trend_df['month'], y=trend_df['total_issues'],
                          name='Total Issues', marker_color='#1FAE6B', opacity=0.75), secondary_y=False)
    fig.add_trace(go.Scatter(x=trend_df['month'], y=trend_df['tat_efficiency'] * 100,
                              name='TAT Efficiency', mode='lines+markers', line=dict(color='#2166D6')),
                  secondary_y=True)
    fig.add_trace(go.Scatter(x=trend_df['month'], y=trend_df['pm_efficiency'] * 100,
                              name='PM Efficiency', mode='lines+markers', line=dict(color='#D98C2B')),
                  secondary_y=True)
    fig.update_yaxes(title_text='Total Issues', secondary_y=False)
    fig.update_yaxes(title_text='Efficiency (%)', range=[0, 100], secondary_y=True)
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=30, b=10),
                       legend=dict(orientation='h', y=1.12), hovermode='x unified')
    return fig


def zone_efficiency_chart(zone_table):
    long_df = zone_table.melt(id_vars='zone', value_vars=['cm_efficiency', 'tat_efficiency', 'pm_efficiency'],
                               var_name='metric', value_name='value')
    long_df['metric'] = long_df['metric'].map(PCT_COLS)
    long_df['value'] *= 100
    fig = px.bar(long_df, x='zone', y='value', color='metric', barmode='group',
                 color_discrete_sequence=['#1FAE6B', '#2166D6', '#D98C2B'])
    fig.update_layout(height=340, margin=dict(l=10, r=10, t=30, b=10), yaxis_title='%',
                       xaxis_title='', legend_title='', legend=dict(orientation='h', y=1.15))
    return fig


def top_zme_chart(zme_table):
    top = zme_table[zme_table['total'] > 0].sort_values('total', ascending=True).tail(10)
    if top.empty:
        return None
    fig = px.bar(top, x='total', y='zme', orientation='h', color='zone',
                 color_discrete_sequence=px.colors.qualitative.Set2)
    fig.update_layout(height=340, margin=dict(l=10, r=10, t=30, b=10), xaxis_title='Total Issues',
                       yaxis_title='', legend_title='Zone')
    return fig


def breakdown_donut(df, column, title):
    counts = df[column].dropna()
    counts = counts[counts != '']
    if counts.empty:
        return None
    vc = counts.value_counts().rename_axis(column).reset_index(name='count')
    fig = px.pie(vc, names=column, values='count', hole=0.5, title=title)
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=40, b=10), showlegend=True)
    return fig


st.title('ChargeZone \u00b7 MPR Dashboard')

uploaded = st.file_uploader('Upload the merged Issue Tracker + PM Tracker workbook', type='xlsx')
if not uploaded:
    st.info('Upload a .xlsx file with "Issue Tracker" and "PM Tracker" sheets to get started.')
    st.stop()

issue_df, pm_df = load_workbook_data(uploaded.getvalue())
quarters = build_quarters(issue_df, pm_df)
quarter_labels = list(quarters.keys())

st.subheader('Monthly Trend (all quarters)')
st.plotly_chart(trend_chart(monthly_trend(issue_df, pm_df)), use_container_width=True)

reported = [q for q in quarter_labels if pm_df.loc[pm_df['quarterLabel'] == q, 'status'].notna().any()]
default_quarter = reported[-1] if reported else quarter_labels[-1]

if 'active_quarter' not in st.session_state:
    st.session_state.active_quarter = default_quarter

st.subheader('Quarter')
cols = st.columns(len(quarter_labels))
for col, label in zip(cols, quarter_labels):
    is_active = label == st.session_state.active_quarter
    if col.button(label.replace(' Q', ' \u00b7 Q'), key=f'q_{label}',
                  type='primary' if is_active else 'secondary'):
        st.session_state.active_quarter = label
        st.rerun()

months_for_quarter = quarters[st.session_state.active_quarter]
selected_months = st.multiselect('Months', months_for_quarter,
                                  default=months_for_quarter,
                                  key=f'month_picker_{st.session_state.active_quarter}')

zme_table = compute_zme_table(issue_df, pm_df, selected_months)
zone_table = zone_rollup(zme_table)

k1, k2, k3, k4 = st.columns(4)
k1.metric('Total Issues', int(zme_table['total'].sum()))
k2.metric('Open Issues', int(zme_table['open'].sum()))
overall_tat = zme_table['within'].sum() / max(zme_table['within'].sum() + zme_table['outside'].sum(), 1)
overall_pm = zme_table['done'].sum() / max(zme_table['planned'].sum(), 1)
k3.metric('TAT Efficiency', f'{overall_tat:.1%}')
k4.metric('PM Efficiency', f'{overall_pm:.1%}')

st.subheader('Issue & PM Summary by ZME')
zme_cols = ['zone', 'zme', 'total', 'open', 'within', 'outside', 'cm_efficiency',
            'tat_efficiency', 'planned', 'done', 'advance', 'pm_efficiency']
render_table(zme_table.sort_values(['zone', 'zme']), zme_cols)

st.subheader('Zone Rollup')
zone_cols = ['zone', 'total', 'cm_efficiency', 'tat_efficiency', 'pm_efficiency']
render_table(zone_table.sort_values('zone'), zone_cols)

st.subheader('Zone Efficiency Comparison')
st.plotly_chart(zone_efficiency_chart(zone_table), use_container_width=True)

top_chart = top_zme_chart(zme_table)
if top_chart is not None:
    st.subheader('Top ZMEs by Issue Volume')
    st.plotly_chart(top_chart, use_container_width=True)

selected_issues = issue_df[issue_df['month'].isin(selected_months)]
donut_col1, donut_col2 = st.columns(2)
status_fig = breakdown_donut(selected_issues, 'status', 'Status Breakdown')
severity_fig = breakdown_donut(selected_issues, 'severity', 'Severity Breakdown')
if status_fig is not None:
    donut_col1.plotly_chart(status_fig, use_container_width=True)
if severity_fig is not None:
    donut_col2.plotly_chart(severity_fig, use_container_width=True)
