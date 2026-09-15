"""
Background-refreshed cache for the Executive Dashboard's "All Clients" view.

Why this exists: an unfiltered "all clients" query was tested directly
against this database and found impractically slow (a lightweight two-table
distinct-client count alone exceeded two minutes; the full TAT query would
be worse) - see report_generator.py / sql_queries.py docstrings. There are
also 5,129 total clients, too many to page through live within a normal
page-load. So the "All Clients" dashboard view reads a snapshot computed by
refresh_dashboard_cache.py running as a detached background process,
kicked off from the UI, rather than querying live on every visit.

Each refresh is written as its own timestamped JSON file under CACHE_DIR, so
the two most recent snapshots are available for the trend/increment
comparison (current vs previous) required by the dashboard - never
comparing against a fabricated "previous" value.
"""
import glob
import json
import os
import subprocess
import sys

CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dashboard_cache')
STATUS_FILE = os.path.join(CACHE_DIR, 'status.json')
REFRESH_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'refresh_dashboard_cache.py')

os.makedirs(CACHE_DIR, exist_ok=True)

# A snapshot's per-client rows must have these keys to be usable by the
# current code. A refresh started before a column rename in
# report_generator.compute_dashboard_aggregates (e.g. total_checks_flat ->
# case_count_flat/check_count_flat) can still be running with the old
# column names loaded in its process, so its output snapshot may predate
# the schema the rest of the app now expects. Rather than crash on a
# missing key, such a snapshot is treated as if it doesn't exist - the UI
# falls back to prompting for a fresh refresh instead.
REQUIRED_ROW_KEYS = {'case_count_flat', 'check_count_flat', 'case_count_recal', 'check_count_recal'}


def _is_compatible(snapshot):
    rows = snapshot.get('rows')
    if not rows:
        return True
    return REQUIRED_ROW_KEYS.issubset(rows[0].keys())


def _snapshot_files():
    return sorted(glob.glob(os.path.join(CACHE_DIR, 'snapshot_*.json')))


def list_snapshots():
    """Returns snapshot metadata (path, computed_at, date_from, date_to),
    oldest first, without loading the (potentially large) per-client rows."""
    snapshots = []
    for path in _snapshot_files():
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            snapshots.append({
                'path': path,
                'computed_at': data.get('computed_at'),
                'date_from': data.get('date_from'),
                'date_to': data.get('date_to'),
                'client_count': len(data.get('rows', [])),
            })
        except (json.JSONDecodeError, OSError):
            continue
    return snapshots


def _compatible_snapshots_newest_first():
    for path in reversed(_snapshot_files()):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            continue
        if _is_compatible(data):
            yield data


def get_latest_snapshot():
    return next(_compatible_snapshots_newest_first(), None)


def get_previous_snapshot():
    gen = _compatible_snapshots_newest_first()
    next(gen, None)
    return next(gen, None)


def get_status():
    if not os.path.exists(STATUS_FILE):
        return {'state': 'idle'}
    try:
        with open(STATUS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {'state': 'idle'}


def is_refresh_running():
    status = get_status()
    return status.get('state') == 'running'


def start_refresh(date_from, date_to):
    """Launches refresh_dashboard_cache.py as a detached background process
    so it keeps running independently of this Streamlit session/rerun cycle.
    Returns False without starting anything if a refresh is already running."""
    if is_refresh_running():
        return False

    creationflags = 0x00000008 | 0x00000200 if os.name == 'nt' else 0  # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    subprocess.Popen(
        [sys.executable, REFRESH_SCRIPT, str(date_from), str(date_to)],
        cwd=os.path.dirname(os.path.abspath(__file__)),
        creationflags=creationflags,
        close_fds=True,
    )
    return True
