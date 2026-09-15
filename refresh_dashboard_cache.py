"""
Standalone background job that computes the "All Clients" dashboard
snapshot. Launched as a detached subprocess by dashboard_cache.start_refresh()
so it keeps running independently of the Streamlit session that triggered it.

Run manually with: python refresh_dashboard_cache.py 2024-01-01 2026-08-14

Batches clients (BATCH_SIZE at a time) through the same fast, bounded
IN-clause query the live dashboard uses for explicit client selections -
an unbounded single query across all 5,129 clients was tested and found
impractically slow (see dashboard_cache.py / report_generator.py docstrings).
Progress is written to dashboard_cache/status.json after every batch so the
Streamlit UI can show live progress; the finished result is one timestamped
JSON snapshot under dashboard_cache/.
"""
import glob
import json
import os
import sys
import traceback
from datetime import date, datetime

import pandas as pd

os.chdir(os.path.dirname(os.path.abspath(__file__)))

import report_generator  # noqa: E402
from dashboard_cache import CACHE_DIR, STATUS_FILE  # noqa: E402

BATCH_SIZE = 50
MAX_SNAPSHOTS_KEPT = 14


def _write_status(**fields):
    status = {}
    if os.path.exists(STATUS_FILE):
        try:
            with open(STATUS_FILE, 'r', encoding='utf-8') as f:
                status = json.load(f)
        except (json.JSONDecodeError, OSError):
            status = {}
    status.update(fields)
    tmp_path = STATUS_FILE + '.tmp'
    with open(tmp_path, 'w', encoding='utf-8') as f:
        json.dump(status, f)
    os.replace(tmp_path, STATUS_FILE)


def _prune_old_snapshots():
    snapshots = sorted(glob.glob(os.path.join(CACHE_DIR, 'snapshot_*.json')))
    for path in snapshots[:-MAX_SNAPSHOTS_KEPT]:
        try:
            os.remove(path)
        except OSError:
            pass


def run(date_from, date_to):
    started_at = datetime.now().isoformat(timespec='seconds')
    _write_status(state='running', started_at=started_at, date_from=str(date_from), date_to=str(date_to),
                  progress={'done': 0, 'total': 0}, error=None)

    clients = report_generator.list_available_clients()
    client_ids = [c['client_external_id'] for c in clients]
    batches = [client_ids[i:i + BATCH_SIZE] for i in range(0, len(client_ids), BATCH_SIZE)]
    _write_status(progress={'done': 0, 'total': len(batches)})

    summaries = []
    for i, batch in enumerate(batches, start=1):
        try:
            flat_df, recal_df = report_generator.fetch_clients_data(batch, date_from, date_to)
            summaries.append(report_generator.compute_dashboard_aggregates(flat_df, recal_df))
        except report_generator.NoDataError:
            pass  # no client in this batch had activity in range - not an error
        _write_status(progress={'done': i, 'total': len(batches)})

    combined = pd.concat(summaries, ignore_index=True) if summaries else pd.DataFrame()

    computed_at = datetime.now().isoformat(timespec='seconds')
    snapshot_path = os.path.join(CACHE_DIR, f'snapshot_{computed_at.replace(":", "-")}.json')
    with open(snapshot_path, 'w', encoding='utf-8') as f:
        json.dump({
            'computed_at': computed_at,
            'date_from': str(date_from),
            'date_to': str(date_to),
            'client_count_total': len(client_ids),
            'rows': combined.to_dict(orient='records'),
        }, f)

    _prune_old_snapshots()
    _write_status(state='done', finished_at=computed_at, snapshot_path=snapshot_path, error=None)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Usage: python refresh_dashboard_cache.py <date_from YYYY-MM-DD> <date_to YYYY-MM-DD>')
        sys.exit(1)
    d_from = date.fromisoformat(sys.argv[1])
    d_to = date.fromisoformat(sys.argv[2])
    try:
        run(d_from, d_to)
    except Exception:
        _write_status(state='error', error=traceback.format_exc())
        raise
