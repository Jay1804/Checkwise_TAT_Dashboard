"""
Reads the Flat/Recal case-check data straight from the checkpoint_live MySQL
database (credentials from .env) instead of the S3 CSV extracts. The raw rows
come from sql_queries.QUERY_BASE_DATA; all TAT/ageing/IT-OT arithmetic is done
by tat_logic.py (validated against the original S3 output).

Always filters by client_external_id - an unfiltered scan is impractically slow.
"""
import os

import numpy as np
import pandas as pd
import pymysql
import pymysql.cursors
from dotenv import load_dotenv

import sql_queries
import tat_logic

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'))

REQUIRED_DB_VARS = ['DB_HOST', 'DB_PORT', 'DB_USER', 'DB_PASSWORD', 'DB_NAME']

BASE_COLUMN_ORDER = [
    'case_id', 'case_check_id', 'client_external_id', 'company_name', 'cat', 'cat_tl',
    'account_manager', 'process_name', 'case_ars_no', 'received_date', 'case_status',
    'insuff_fulfill_date', 'check_status', 'check_severity', 'check_closure_date',
    'check_tat_category', 'check_tat', 'reopen_date_with_condition', 'goahead_date_with_condition',
    'max_date', 'unique check name', 'check_ops_name',
]
FLAT_COLUMN_ORDER = BASE_COLUMN_ORDER + [
    'total_sat_sun_holi_flat', 'chk_due_date_flat', 'chk_opt_flat_due_date',
    'ageing_flat', 'ageing_bucket_flat', 'tat_check flat tat ',
]
RECAL_COLUMN_ORDER = BASE_COLUMN_ORDER + [
    'total_sat_sun_holi', 'chk_due_date_recal', 'chk_opt_recal_due_date',
    'ageing', 'ageing_bucket', 'tat_check tat',
]

RAW_DATE_COLS = ['received_date', 'check_closure_date', 'reopen_date_with_condition',
                 'goahead_date_with_condition', 'insuff_fulfill_date']


class NoDataError(RuntimeError):
    pass


def get_connection(connect_timeout=15):
    missing = [n for n in REQUIRED_DB_VARS if not os.environ.get(n)]
    if missing:
        raise RuntimeError(f"Missing required database configuration in .env: {', '.join(missing)}")
    return pymysql.connect(
        host=os.environ['DB_HOST'], port=int(os.environ['DB_PORT']),
        user=os.environ['DB_USER'], password=os.environ['DB_PASSWORD'],
        database=os.environ['DB_NAME'], cursorclass=pymysql.cursors.DictCursor,
        connect_timeout=connect_timeout,
    )


def fetch_flat_recal(client_ids, date_from, date_to):
    """Returns (flat_df, recal_df) for one or more client_external_ids (an int
    or an iterable of ints) in a received_date range, each row tagged with its
    client_external_id, with the same columns as the S3 CSV extracts."""
    if isinstance(client_ids, (int, np.integer)):
        client_ids = [client_ids]
    client_ids = tuple(int(c) for c in client_ids)
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(sql_queries.QUERY_BASE_DATA, {
                'client_external_ids': client_ids,
                'date_from': date_from,
                'date_to': date_to,
            })
            base_rows = cur.fetchall()
            cur.execute(sql_queries.QUERY_HOLIDAYS)
            holiday_rows = cur.fetchall()
    finally:
        conn.close()

    if not base_rows:
        raise NoDataError(
            f'No case-check rows found for client_external_id(s) {list(client_ids)} between {date_from} and {date_to}.'
        )

    base_df = pd.DataFrame(base_rows)
    for col in RAW_DATE_COLS:
        # errors='coerce': legacy MySQL zero-dates (0000-00-00) become NaT.
        base_df[col] = pd.to_datetime(base_df[col], errors='coerce')

    holidays_df = pd.DataFrame(holiday_rows)
    holiday_dates = (pd.to_datetime(holidays_df['holiday_date'], errors='coerce').values
                     if len(holidays_df) else np.array([], dtype='datetime64[ns]'))

    flat_df, recal_df = tat_logic.build_flat_and_recal(base_df, holiday_dates, pd.Timestamp.today().date())
    return flat_df[FLAT_COLUMN_ORDER].reset_index(drop=True), recal_df[RECAL_COLUMN_ORDER].reset_index(drop=True)


def available_clients():
    """[(client_external_id, company_name), ...] for the UI picker."""
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(sql_queries.QUERY_AVAILABLE_CLIENTS)
            return [(r['client_external_id'], r['company_name']) for r in cur.fetchall()]
    finally:
        conn.close()
