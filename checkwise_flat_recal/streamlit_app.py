"""
Streamlit UI for the Check TAT Analysis report generator.

Run with: streamlit run streamlit_app.py

Wraps build_report.py's build_reports_for_clients() - no pipeline logic lives
here, so this and build_report.py never drift apart.
"""
import contextlib
import io
import os
import subprocess
import traceback
from datetime import date

import streamlit as st

import build_report
import db_source

st.set_page_config(page_title='Check TAT Analysis Report Generator', page_icon='📊')


def is_excel_running():
    result = subprocess.run(
        ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE'],
        capture_output=True, text=True,
    )
    return 'EXCEL.EXE' in result.stdout


class StreamlitLogStream(io.TextIOBase):
    """Redirected stdout target - appends every print() to a live placeholder."""

    def __init__(self, placeholder):
        self.placeholder = placeholder
        self.buffer = ''

    def write(self, s):
        if s:
            self.buffer += s
            self.placeholder.code(self.buffer, language=None)
        return len(s)

    def flush(self):
        pass


st.title('Check TAT Analysis - Report Generator')
st.caption(
    "Pulls the latest checkwise Flat / Recal data live from the checkpoint_live database for the "
    "client and date range you choose, and builds that client's workbook from scratch "
    "(data tables + pivots) - no dependency on any pre-existing template file."
)

if is_excel_running():
    st.error(
        'Excel is currently running. This pipeline drives Excel via COM automation and '
        'closes it when done - if you have a workbook open, close Excel first or you may '
        'lose unsaved changes / have your window closed unexpectedly.'
    )
    st.stop()

@st.cache_data(ttl=3600, show_spinner='Loading client list from the database...')
def load_clients():
    return db_source.available_clients()


try:
    clients = load_clients()
except Exception:
    st.error('Could not load the client list from the database.')
    st.code(traceback.format_exc())
    st.stop()

label_to_id = {f'{cid} - {name}': cid for cid, name in clients}
selected_labels = st.multiselect(
    'Clients (type a client_external_id or a client name to search; select one or more)',
    options=list(label_to_id),
    placeholder='Search by ID or name...',
)

date_range = st.date_input('Date range (filters on received_date)', value=(date(2020, 1, 1), date.today()))

run_clicked = st.button('Generate report(s)', type='primary')

if run_clicked:
    if not selected_labels:
        st.error('Please select at least one client.')
        st.stop()
    if not (isinstance(date_range, tuple) and len(date_range) == 2):
        st.error('Please select a full date range (both a start and an end date).')
        st.stop()
    date_from, date_to = date_range
    selected_ids = [label_to_id[l] for l in selected_labels]
    names = {cid: name for cid, name in clients}

    log_placeholder = st.empty()
    stream = StreamlitLogStream(log_placeholder)
    progress = st.progress(0.0)
    try:
        with st.spinner(f'Building report(s) for {len(selected_ids)} client(s)...'):
            with contextlib.redirect_stdout(stream):
                results = build_report.build_reports_for_clients(
                    selected_ids, date_from=date_from, date_to=date_to,
                    progress_callback=lambda done, total, cid: progress.progress(done / total),
                )
        # Persist: a download button click reruns the script, which would otherwise wipe the results.
        st.session_state['generate_results'] = [
            dict(r, company_name=names.get(r['client_id'], ''),
                 data=open(r['output_path'], 'rb').read() if r['output_path'] else None)
            for r in results
        ]
    except Exception:
        st.error('Report generation failed.')
        st.code(traceback.format_exc())

for r in st.session_state.get('generate_results', []):
    title = f"{r['client_id']} - {r['company_name']}"
    if r['error']:
        st.warning(f'{title}: {r["error"]}')
    else:
        st.download_button(
            f'Download {title}',
            data=r['data'],
            file_name=os.path.basename(r['output_path']),
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key=f"dl_{r['client_id']}",
        )
