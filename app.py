"""
Streamlit UI for the Checkwise Check TAT Analysis MIS Report Generator.

Run with: streamlit run app.py

Two views:
- Executive Dashboard: TAT performance KPIs, top/bottom performing clients,
  and an auto-generated summary, for an explicitly chosen set of clients
  (an unfiltered all-clients scan was tested against this database and takes
  several minutes at best - see report_generator.get_dashboard_summary).
- Generate & Send Reports: pick one or more clients + a date range, build
  each client's workbook from live data, and email it to the client's
  recipients from an uploaded mapping file (client_external_id, company_name,
  To_address, CC_address). A client absent from the mapping is generated but
  not emailed - no address is ever guessed.
"""
import io
import os
import subprocess
import traceback
import zipfile
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import dashboard_cache
import report_generator
from db_connect import ConfigError
from email_sender import EmailConfigError, send_report_email
from recipient_mapping import MappingError, parse_mapping_file
from report_generator import NoDataError

st.set_page_config(page_title='Checkwise MIS Report Generator', page_icon='📊', layout='wide')

st.markdown("""
<style>
.app-header {
    background: linear-gradient(135deg, #1f3c88 0%, #4a69bd 100%);
    padding: 1.6rem 2rem;
    border-radius: 14px;
    color: white;
    margin-bottom: 1.6rem;
}
.app-header h1 { margin: 0; font-size: 1.7rem; font-weight: 700; }
.app-header p { margin: 0.35rem 0 0; opacity: 0.9; font-size: 0.95rem; }
div[data-testid="stMetric"] {
    background: white;
    border: 1px solid #e3e6ea;
    border-radius: 12px;
    padding: 0.9rem 1rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.exec-summary {
    background: #eef2fb;
    border-left: 4px solid #4a69bd;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    margin: 0.8rem 0 1.2rem;
    font-size: 0.98rem;
    line-height: 1.5;
}
.status-card {
    border: 1px solid #e3e6ea;
    border-radius: 10px;
    padding: 0.8rem 1rem;
    margin-bottom: 0.6rem;
    background: white;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="app-header">
  <h1>📊 Checkwise Check TAT Analysis</h1>
  <p>Live MIS reporting straight from checkpoint_live - TAT/ageing analytics, workbook generation, and automated client email delivery.</p>
</div>
""", unsafe_allow_html=True)


def is_excel_running():
    result = subprocess.run(
        ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE'],
        capture_output=True, text=True,
    )
    return 'EXCEL.EXE' in result.stdout


@st.cache_data(ttl=600)
def _cached_client_list():
    return report_generator.list_available_clients()


@st.cache_data(ttl=1800, show_spinner=False)
def _cached_dashboard_summary(client_ids, date_from, date_to):
    return report_generator.get_dashboard_summary(list(client_ids), date_from, date_to)


MIN_CHECKS_FOR_RANKING = 10
TAT_TARGET_PCT = 95


def _load_all_clients_view(date_from, date_to):
    """Reads the background-computed All-Clients snapshot (see
    dashboard_cache.py for why this can't be a live query) and offers a
    refresh trigger. Returns (summary_df, prev_summary_df_or_None,
    prev_meta_or_None) or (None, None, None) if there's nothing to show yet."""
    status = dashboard_cache.get_status()
    latest = dashboard_cache.get_latest_snapshot()

    col_msg, col_btn = st.columns([4, 1.3])
    with col_btn:
        refreshing = status.get('state') == 'running'
        if st.button('🔄 Refresh All-Clients Cache', disabled=refreshing, use_container_width=True):
            if dashboard_cache.start_refresh(date_from, date_to):
                st.info('Refresh started in the background - there are 5,129 clients in total, so this can take a long time. This page will show live progress; you can keep working elsewhere and check back.')
                st.rerun()

    if status.get('state') == 'running':
        progress = status.get('progress', {})
        done, total = progress.get('done', 0), progress.get('total', 0)
        pct = (done / total) if total else 0.0
        with col_msg:
            st.warning(f"Refresh in progress: batch {done}/{total} ({pct * 100:.0f}%) - started {status.get('started_at', '?')}.")
        st.progress(pct)
    elif status.get('state') == 'error':
        err = (status.get('error') or 'unknown error').strip().splitlines()
        with col_msg:
            st.error(f"Last refresh failed: {err[-1] if err else 'unknown error'}")

    if latest is None or not latest.get('rows'):
        st.info(
            'No All-Clients snapshot has been computed yet. Click "Refresh All-Clients Cache" above to '
            'compute one now (this queries every client in the system and can take a long time), or select '
            'specific clients above for an instant live view instead.'
        )
        return None, None, None

    summary = pd.DataFrame(latest['rows'])
    st.caption(
        f"Showing the All-Clients snapshot computed {latest['computed_at']} "
        f"(data range {latest['date_from']} to {latest['date_to']}; {len(summary)} of "
        f"{latest.get('client_count_total', '?')} total clients had activity in that range). "
        f"Not live - click Refresh above for current data."
    )

    prev = dashboard_cache.get_previous_snapshot()
    prev_summary, prev_meta = None, None
    if prev and prev.get('rows'):
        prev_summary = pd.DataFrame(prev['rows'])
        prev_meta = {'date_from': prev['date_from'], 'date_to': prev['date_to'], 'computed_at': prev['computed_at']}

    return summary, prev_summary, prev_meta


def render_dashboard(client_options, name_by_id):
    selected_labels = st.multiselect(
        'Clients (leave empty for All Clients)', options=list(client_options.keys()), default=[], key='dash_clients',
    )
    date_range = st.date_input('Date range', value=(date(2025, 1, 1), date.today()), key='dash_dates')
    if not (isinstance(date_range, tuple) and len(date_range) == 2):
        st.info('Select a full date range (start and end date).')
        return
    date_from, date_to = date_range

    all_clients_mode = len(selected_labels) == 0
    scope_label = 'All Clients' if all_clients_mode else f"{len(selected_labels)} Selected Client{'s' if len(selected_labels) != 1 else ''}"
    st.markdown(f'#### Dashboard Scope: {scope_label}')

    prev_summary, prev_meta = None, None

    if all_clients_mode:
        summary, prev_summary, prev_meta = _load_all_clients_view(date_from, date_to)
        if summary is None:
            return
    else:
        client_ids = tuple(sorted(client_options[label] for label in selected_labels))
        try:
            with st.spinner(f'Computing TAT performance for {len(client_ids)} client(s)...'):
                summary = _cached_dashboard_summary(client_ids, date_from, date_to)
        except NoDataError as e:
            st.warning(str(e))
            return
        except ConfigError as e:
            st.error(f'Configuration error: {e}')
            return
        except Exception:
            st.error('Could not compute the dashboard. See details below.')
            st.code(traceback.format_exc())
            return

        if summary.empty:
            st.info('No data found for the selected clients/date range.')
            return

        no_data_clients = [name_by_id.get(cid, str(cid)) for cid in client_ids if cid not in set(summary['client_external_id'])]
        if no_data_clients:
            st.caption(f"No data in this date range for: {', '.join(no_data_clients)} - excluded from the KPIs below.")

        prev_from, prev_to = report_generator.previous_period(date_from, date_to)
        try:
            with st.spinner('Computing prior-period comparison...'):
                candidate = _cached_dashboard_summary(client_ids, prev_from, prev_to)
            if not candidate.empty:
                prev_summary = candidate
                prev_meta = {'date_from': str(prev_from), 'date_to': str(prev_to)}
        except Exception:
            prev_summary = None

    totals = report_generator.overall_totals(summary)
    prev_totals = report_generator.overall_totals(prev_summary) if prev_summary is not None and not prev_summary.empty else None

    render_overall_kpis(totals)
    render_executive_summary(scope_label, date_from, date_to, summary, totals, prev_totals)
    render_flat_recal_split(totals)
    render_gauges(totals)
    render_trend(totals, prev_totals, prev_meta)
    render_performance_health(summary, len(selected_labels))
    render_full_breakdown(summary)


def render_overall_kpis(totals):
    st.markdown('#### Overall KPIs')
    st.caption('"Overall" is reported on the Flat basis (the primary TAT measure) - see Flat vs Recal below for the reopen/insufficiency-adjusted Recal figures.')
    flat = totals['flat']
    row1 = st.columns(3)
    row1[0].metric('Case Count', f"{flat['case_count']:,}")
    row1[1].metric('Check Count', f"{flat['check_count']:,}")
    row1[2].metric('Avg TAT %', f"{flat['it_pct']}%" if flat['it_pct'] is not None else 'N/A')
    row2 = st.columns(3)
    row2[0].metric('IT', f"{flat['it_count']:,}")
    row2[1].metric('OT', f"{flat['ot_count']:,}")
    row2[2].metric('Ageing (days)', f"{flat['avg_ageing']}" if flat['avg_ageing'] is not None else 'N/A')


def render_executive_summary(scope_label, date_from, date_to, summary, totals, prev_totals):
    st.markdown('#### Executive Summary')
    flat, recal = totals['flat'], totals['recal']
    lines = []

    if not flat['check_count']:
        st.markdown('<div class="exec-summary">No checks found for this scope/date range.</div>', unsafe_allow_html=True)
        return

    vs_target = 'meeting or above' if flat['it_pct'] >= TAT_TARGET_PCT else 'below'
    lines.append(
        f"<b>{scope_label}</b>, {date_from} to {date_to}: overall TAT compliance is <b>{flat['it_pct']}%</b> "
        f"against the {TAT_TARGET_PCT}% target ({vs_target} target), across <b>{flat['case_count']:,} cases</b> "
        f"(<b>{flat['check_count']:,} checks</b>) and <b>{len(summary)} client(s)</b>."
    )
    lines.append(
        f"IT (on-time): <b>{flat['it_count']:,}</b> checks. OT (overdue): <b>{flat['ot_count']:,}</b> checks. "
        f"Average ageing is <b>{flat['avg_ageing']} days</b>."
    )
    if recal['check_count']:
        diff = round(recal['it_pct'] - flat['it_pct'], 1)
        relation = 'in line with' if abs(diff) <= 2 else ('higher than' if diff > 0 else 'lower than')
        lines.append(
            f"Recal (reopen/insufficiency-adjusted) TAT compliance is <b>{recal['it_pct']}%</b>, {relation} the Flat figure "
            f"({'+' if diff >= 0 else ''}{diff} pts)."
        )

    if prev_totals and prev_totals['flat']['it_pct'] is not None:
        change = round(flat['it_pct'] - prev_totals['flat']['it_pct'], 1)
        if change > 0:
            lines.append(f"TAT performance has <b>improved</b> by {change} percentage point(s) versus the prior comparable period.")
        elif change < 0:
            lines.append(f"TAT performance has <b>declined</b> by {abs(change)} percentage point(s) versus the prior comparable period.")
        else:
            lines.append('TAT performance is unchanged versus the prior comparable period.')
    else:
        lines.append('Trend: historical comparison unavailable for this scope/date range.')

    ranked = summary[summary['check_count_flat'] >= MIN_CHECKS_FOR_RANKING]
    below_target = ranked[ranked['it_pct_flat'] < TAT_TARGET_PCT]
    if len(ranked):
        best = ranked.sort_values('it_pct_flat', ascending=False).iloc[0]
        worst = ranked.sort_values('it_pct_flat', ascending=True).iloc[0]
        lines.append(
            f"<b>{len(below_target)}</b> of {len(ranked)} ranked client(s) (≥{MIN_CHECKS_FOR_RANKING} checks) are below "
            f"the {TAT_TARGET_PCT}% target and may need attention. Top performer: <b>{best['company_name']}</b> "
            f"({best['it_pct_flat']}%). Needs attention: <b>{worst['company_name']}</b> ({worst['it_pct_flat']}%)."
        )
    else:
        lines.append(f'No client reached the {MIN_CHECKS_FOR_RANKING}-check minimum for individual ranking in this scope.')

    st.markdown(f'<div class="exec-summary">{" ".join(lines)}</div>', unsafe_allow_html=True)


def render_flat_recal_split(totals):
    st.markdown('#### Flat vs Recal Performance')
    col_flat, col_recal = st.columns(2)
    for col, key, title in ((col_flat, 'flat', 'FLAT'), (col_recal, 'recal', 'RECAL')):
        t = totals[key]
        with col:
            st.markdown(f'**{title}**')
            r1c1, r1c2 = st.columns(2)
            r1c1.metric('Case Count', f"{t['case_count']:,}")
            r1c2.metric('Check Count', f"{t['check_count']:,}")
            r2c1, r2c2 = st.columns(2)
            r2c1.metric('Overall IT', f"{t['it_count']:,}")
            r2c2.metric('Overall OT', f"{t['ot_count']:,}")
            r3c1, r3c2 = st.columns(2)
            r3c1.metric('Ageing (days)', f"{t['avg_ageing']}" if t['avg_ageing'] is not None else 'N/A')
            r3c2.metric('TAT Performance', f"{t['it_pct']}%" if t['it_pct'] is not None else 'N/A')


def _gauge_bar_color(value):
    if value is None:
        return '#999999'
    if value >= TAT_TARGET_PCT:
        return '#2e7d32'
    if value >= TAT_TARGET_PCT - 15:
        return '#f9a825'
    return '#c62828'


def _build_gauge_figure(value, title):
    display_value = value if value is not None else 0
    fig = go.Figure(go.Indicator(
        mode='gauge+number',
        value=display_value,
        number={'suffix': '%'},
        gauge={
            'axis': {'range': [0, 100]},
            'bar': {'color': _gauge_bar_color(value)},
            'steps': [
                {'range': [0, TAT_TARGET_PCT - 15], 'color': '#fde0e0'},
                {'range': [TAT_TARGET_PCT - 15, TAT_TARGET_PCT], 'color': '#fff4d6'},
                {'range': [TAT_TARGET_PCT, 100], 'color': '#e1f5e1'},
            ],
            'threshold': {'line': {'color': 'black', 'width': 3}, 'thickness': 0.85, 'value': TAT_TARGET_PCT},
        },
        title={'text': title},
    ))
    fig.update_layout(height=240, margin=dict(l=25, r=25, t=50, b=10))
    return fig


def render_gauges(totals):
    st.markdown(f'#### TAT Gauges (Target: {TAT_TARGET_PCT}%)')
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(_build_gauge_figure(totals['flat']['it_pct'], 'Flat TAT Performance'), use_container_width=True)
    with col2:
        st.plotly_chart(_build_gauge_figure(totals['recal']['it_pct'], 'Recal TAT Performance'), use_container_width=True)


def _trend_metric(container, label, current_pct, previous_pct):
    with container:
        st.markdown(f'**{label}**')
        if current_pct is None:
            st.caption('No data.')
            return
        if previous_pct is None:
            st.metric(label, f'{current_pct}%')
            st.caption('Trend: historical comparison unavailable')
            return
        change = round(current_pct - previous_pct, 1)
        st.metric(label, f'{current_pct}%', delta=f'{change:+.1f} pts')
        verdict = 'Performance Improved' if change > 0 else ('Performance Declined' if change < 0 else 'No Change')
        st.caption(f'Previous: {previous_pct}% -  {verdict}')


def render_trend(totals, prev_totals, prev_meta):
    st.markdown('#### TAT Performance Trend')
    if prev_meta:
        st.caption(f"Compared against {prev_meta['date_from']} to {prev_meta['date_to']}" +
                   (f" (snapshot computed {prev_meta['computed_at']})" if 'computed_at' in prev_meta else ''))
    else:
        st.caption('Trend: historical comparison unavailable - no prior-period data found for this scope.')

    c1, c2 = st.columns(2)
    prev_flat = prev_totals['flat']['it_pct'] if prev_totals else None
    prev_recal = prev_totals['recal']['it_pct'] if prev_totals else None
    _trend_metric(c1, 'Flat TAT', totals['flat']['it_pct'], prev_flat)
    _trend_metric(c2, 'Recal TAT', totals['recal']['it_pct'], prev_recal)


def _health_status(it_pct):
    """Shares the same 95%/80% banding used for the TAT gauges (see
    _gauge_bar_color), so 'Healthy' here always means the same thing a green
    gauge means - one status definition, reused everywhere."""
    if it_pct is None or pd.isna(it_pct):
        return 'Unknown'
    if it_pct >= TAT_TARGET_PCT:
        return 'Healthy'
    if it_pct >= TAT_TARGET_PCT - 15:
        return 'Watch'
    return 'Needs Attention'


def render_performance_health(summary, num_selected):
    """Dynamic per the current client-selection context: a single selected
    client shows its own status (no top/bottom comparison to make); All
    Clients or multiple selected clients shows the best/worst performer
    *within that same scope* - never a client outside the current selection."""
    st.markdown('#### Performance Health')
    ranked = summary[summary['check_count_flat'] >= MIN_CHECKS_FOR_RANKING]

    if num_selected == 1:
        row = summary.iloc[0]
        st.markdown(f"**Selected Client:** {row['company_name']}")
        c1, c2 = st.columns(2)
        c1.metric('TAT Performance', f"{row['it_pct_flat']}%" if pd.notna(row['it_pct_flat']) else 'N/A')
        c2.metric('Status', _health_status(row['it_pct_flat']))
        return

    st.caption(f'Best/worst Flat TAT performer within the current scope, among clients with ≥{MIN_CHECKS_FOR_RANKING} checks.')
    if not len(ranked):
        st.caption('No clients in the current scope meet the minimum check count for ranking.')
        return

    best = ranked.sort_values('it_pct_flat', ascending=False).iloc[0]
    worst = ranked.sort_values('it_pct_flat', ascending=True).iloc[0]
    col_top, col_bottom = st.columns(2)
    with col_top:
        st.markdown('**🏆 Top Performing Client**')
        st.markdown(f"**{best['company_name']}**")
        c1, c2 = st.columns(2)
        c1.metric('TAT Performance', f"{best['it_pct_flat']}%")
        c2.metric('Status', _health_status(best['it_pct_flat']))
    with col_bottom:
        st.markdown('**⚠️ Bottom Performing Client**')
        st.markdown(f"**{worst['company_name']}**")
        c1, c2 = st.columns(2)
        c1.metric('TAT Performance', f"{worst['it_pct_flat']}%")
        c2.metric('Status', _health_status(worst['it_pct_flat']))


def _variant_breakdown_table(summary, variant):
    it_pct_col = f'it_pct_{variant}'
    table = summary.sort_values(it_pct_col, ascending=False).copy()
    table['Performance Health'] = table[it_pct_col].apply(_health_status)
    table['TAT Performance'] = table[it_pct_col].apply(lambda v: f'{v}%' if pd.notna(v) else 'N/A')
    table = table.rename(columns={
        'company_name': 'Client', f'case_count_{variant}': 'Case Count', f'check_count_{variant}': 'Check Count',
        f'it_count_{variant}': 'Overall IT', f'ot_count_{variant}': 'Overall OT', f'avg_ageing_{variant}': 'Ageing',
    })
    return table[['Client', 'Case Count', 'Check Count', 'Overall IT', 'Overall OT', 'Ageing',
                  'TAT Performance', 'Performance Health']]


def render_full_breakdown(summary):
    """Always reflects the current client-selection scope (the same
    `summary` every other section uses) - All Clients, the selected
    multiple, or the single selected client. Shown as two separate tables,
    one per TAT variant, since Flat and Recal use different due-date/ageing
    bases and so can rank/flag clients differently."""
    st.markdown('#### Full Client Breakdown')
    st.markdown('**Flat**')
    st.dataframe(_variant_breakdown_table(summary, 'flat'), hide_index=True, use_container_width=True)
    st.markdown('**Recal**')
    st.dataframe(_variant_breakdown_table(summary, 'recal'), hide_index=True, use_container_width=True)


def render_generate_and_send(client_options, name_by_id):
    if is_excel_running():
        st.error(
            'Excel is currently running. This tool drives Excel via COM automation and closes it when '
            'done - if you have a workbook open, close Excel first or you may lose unsaved changes / '
            'have your window closed unexpectedly.'
        )
        return

    selected_labels = st.multiselect('Client(s)', options=list(client_options.keys()), key='gen_clients')
    date_range = st.date_input('Date range (filters on received_date)', value=(date(2024, 1, 1), date.today()), key='gen_dates')

    st.markdown('##### Recipient mapping (optional)')
    st.caption(
        'Upload a CSV or Excel file with columns client_external_id, company_name, To_address, CC_address. '
        'A report is only emailed if its client_external_id appears in this file - multiple addresses in '
        'To_address/CC_address can be separated by comma or semicolon. Without a mapping, reports are '
        'generated and available to download, but not emailed.'
    )
    mapping_file = st.file_uploader('Mapping file', type=['csv', 'xlsx', 'xls'], key='mapping_upload')

    mapping = {}
    if mapping_file is not None:
        try:
            mapping, skipped = parse_mapping_file(mapping_file)
            st.success(f'Mapping loaded: {len(mapping)} client(s) with a valid To address.')
            if skipped:
                st.warning(f'{len(skipped)} row(s) skipped (no To_address): {skipped}')
            with st.expander('Preview mapping'):
                preview = pd.DataFrame([
                    {'client_external_id': cid, 'company_name': v['company_name'],
                     'to': ', '.join(v['to']), 'cc': ', '.join(v['cc'])}
                    for cid, v in mapping.items()
                ])
                st.dataframe(preview, hide_index=True, use_container_width=True)
        except MappingError as e:
            st.error(str(e))
            mapping_file = None

    if selected_labels and mapping_file is not None:
        selected_ids = [client_options[label] for label in selected_labels]
        unmapped = [name_by_id.get(cid, str(cid)) for cid in selected_ids if cid not in mapping]
        if unmapped:
            st.warning(f"No mapping row for: {', '.join(unmapped)} - these will be generated but not emailed.")

    run_clicked = st.button('Generate & Send Reports', type='primary')

    if run_clicked:
        if not selected_labels:
            st.error('Select at least one client.')
        elif not (isinstance(date_range, tuple) and len(date_range) == 2):
            st.error('Please select a full date range (both a start and an end date).')
        else:
            date_from, date_to = date_range
            date_range_label = f'{date_from} to {date_to}'
            selected_ids = [client_options[label] for label in selected_labels]

            st.subheader('Progress')
            progress_bar = st.progress(0.0)
            progress_text = st.empty()

            def on_progress(done, total, company_name):
                progress_bar.progress(done / total)
                progress_text.text(f'{done}/{total} - just finished {company_name}')

            try:
                with st.spinner(f'Fetching data and building {len(selected_ids)} report(s)...'):
                    results = report_generator.generate_reports_for_clients(
                        selected_ids, date_from, date_to, client_names=name_by_id, progress_callback=on_progress,
                    )
                for r in results:
                    if r['output_path']:
                        r['email_status'] = None
                        recipients = mapping.get(r['client_external_id'])
                        if recipients:
                            try:
                                send_report_email(r['company_name'], date_range_label, r['output_path'],
                                                   recipients['to'], recipients['cc'])
                                r['email_status'] = 'sent'
                            except EmailConfigError as e:
                                r['email_status'] = f'config error: {e}'
                            except Exception as e:
                                r['email_status'] = f'failed: {e}'
                        else:
                            r['email_status'] = 'skipped (no mapping)'

                # Persisted in session_state (not just a local variable) because
                # clicking a download_button below triggers its own Streamlit
                # rerun - without this, that rerun would find run_clicked False
                # again and the whole Results/Download section (including every
                # OTHER download button) would vanish before it could be used.
                st.session_state['generate_results'] = results
                st.session_state['generate_date_range_label'] = date_range_label
            except NoDataError as e:
                st.error(str(e))
            except ConfigError as e:
                st.error(f'Configuration error: {e}')
            except Exception:
                st.error('Report generation failed. See details below.')
                st.code(traceback.format_exc())

    results = st.session_state.get('generate_results')
    if not results:
        return
    date_range_label = st.session_state.get('generate_date_range_label', '')

    st.subheader('Results')
    successful = [r for r in results if r['output_path']]
    for r in results:
        with st.container():
            if r['output_path']:
                email = r['email_status']
                if email == 'sent':
                    email_line = '📧 Email sent'
                elif email == 'skipped (no mapping)':
                    email_line = '⚠️ Email skipped - no mapping for this client'
                else:
                    email_line = f'❌ Email failed - {email}'
                st.markdown(
                    f'<div class="status-card">✅ <b>{r["company_name"]}</b> - report generated<br>{email_line}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div class="status-card">❌ <b>{r["company_name"]}</b> - {r["error"]}</div>',
                    unsafe_allow_html=True,
                )

    if successful:
        with st.expander(f'Download reports ({len(successful)})', expanded=True):
            for r in successful:
                if not os.path.exists(r['output_path']):
                    st.warning(f"{os.path.basename(r['output_path'])} is no longer on disk.")
                    continue
                with open(r['output_path'], 'rb') as f:
                    st.download_button(
                        f"Download {os.path.basename(r['output_path'])}",
                        data=f.read(),
                        file_name=os.path.basename(r['output_path']),
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=f"dl_{r['client_external_id']}",
                    )
            if len(successful) > 1:
                buffer = io.BytesIO()
                with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for r in successful:
                        if os.path.exists(r['output_path']):
                            zf.write(r['output_path'], arcname=os.path.basename(r['output_path']))
                st.download_button(
                    'Download all as ZIP', data=buffer.getvalue(),
                    file_name=f"checkwise_reports_{date_range_label.replace(' ', '_')}.zip", mime='application/zip',
                    key='dl_zip',
                )


try:
    clients = _cached_client_list()
except ConfigError as e:
    st.error(f'Configuration error: {e}')
    st.stop()
except Exception:
    st.error('Could not connect to the database. Check the connection settings in .env and try again.')
    st.stop()

if not clients:
    st.warning('No clients found in the database.')
    st.stop()

client_options = {f"{c['company_name']} ({c['client_external_id']})": c['client_external_id'] for c in clients}
name_by_id = {c['client_external_id']: c['company_name'] for c in clients}

PAGE_DASHBOARD = '📊 Executive Dashboard'
PAGE_GENERATE = '📧 Generate & Send Reports'

# st.tabs() resets to the first tab on every rerun (it has no persistent
# selection state), which would kick users back to the Dashboard tab after
# every single widget interaction on the Generate & Send page (client
# picker, file upload, etc). st.segmented_control stores its value in
# session_state like any other widget, so it survives reruns correctly.
page = st.segmented_control('Navigation', options=[PAGE_DASHBOARD, PAGE_GENERATE],
                             default=PAGE_DASHBOARD, label_visibility='collapsed')

if page == PAGE_GENERATE:
    render_generate_and_send(client_options, name_by_id)
else:
    render_dashboard(client_options, name_by_id)
