"""
Reproduces the checkwise_flat_recal TAT/ageing/IT-OT business logic in pandas.

The original hand-off SQL computed due dates, ageing and IT/OT status via
chains of correlated subqueries and same-SELECT column references against a
`..._for_networkdays` holiday helper table - a pattern MySQL doesn't support
(no same-level SELECT-alias chaining) and a table that doesn't exist in this
database. The live `ec_master_holidays` table only carries two holiday
types: type 1 = Sundays + ad-hoc festival holidays, type 2 = every Saturday.

The formulas below were validated against the original pipeline's own
historical output (checkwise_flat000.csv, case_check_id 182137830: expected
total_sat_sun_holi_flat=4, chk_due_date_flat=2024-01-20, ageing_flat=5,
bucket '3--10', IT) and reproduce it exactly for the 'Working'/'workday'
category, which is the only category with unambiguous live-schema holiday
data. For 'Calendar' and 'WorkingPlusSaturday' the same exclusion-set
mechanism is applied with the category's own exclusion rule (see
_daily_calendar) since no historical sample was available in this
environment to validate those two categories against.
"""
import numpy as np
import pandas as pd

AGEING_BUCKETS = [
    (-3, 2, '0--2'),
    (3, 10, '3--10'),
    (11, 14, '11--14'),
    (15, 20, '15--20'),
    (21, 25, '21--25'),
    (26, 30, '26--30'),
]

WORKING_NAMES = ('Working', 'workday')
CALENDAR_NAMES = ('Calendar', 'calendar')
WORKING_PLUS_SATURDAY_NAMES = ('WorkingPlusSaturday', 'workday-sat')

IT_OVERRIDE_CHECK_STATUSES = ('On Hold', 'Insufficient', 'case closed by client')

COLUMN_NAMES = {
    'flat': dict(
        total='total_sat_sun_holi_flat', due='chk_due_date_flat', opt_due='chk_opt_flat_due_date',
        ageing='ageing_flat', bucket='ageing_bucket_flat', tat_status='tat_check flat tat ',
    ),
    'recal': dict(
        total='total_sat_sun_holi', due='chk_due_date_recal', opt_due='chk_opt_recal_due_date',
        ageing='ageing', bucket='ageing_bucket', tat_status='tat_check tat',
    ),
}


def _daily_calendar(start, end, holiday_dates):
    """One row per calendar day in [start, end] with per-category exclusion
    flags, used to build sorted excluded-date arrays for fast counting."""
    days = pd.date_range(start.normalize(), end.normalize(), freq='D')
    dow = days.dayofweek  # Monday=0 ... Sunday=6
    is_saturday = dow == 5
    is_sunday = dow == 6
    is_holiday = days.isin(holiday_dates)

    excluded_working = is_saturday | is_sunday | is_holiday
    excluded_wps = is_sunday | (is_holiday & ~is_saturday)

    return {
        'Working': np.sort(days[excluded_working].values),
        'Calendar': np.array([], dtype='datetime64[ns]'),
        'WorkingPlusSaturday': np.sort(days[excluded_wps].values),
    }


def _category_key(category):
    if category in WORKING_NAMES:
        return 'Working'
    if category in CALENDAR_NAMES:
        return 'Calendar'
    if category in WORKING_PLUS_SATURDAY_NAMES:
        return 'WorkingPlusSaturday'
    return None


def _count_excluded(sorted_excluded, start, end):
    """Vectorized inclusive-range count of dates present in sorted_excluded,
    for every (start[i], end[i]) pair."""
    start = np.asarray(start, dtype='datetime64[ns]')
    end = np.asarray(end, dtype='datetime64[ns]')
    lo = np.searchsorted(sorted_excluded, start, side='left')
    hi = np.searchsorted(sorted_excluded, end, side='right')
    return (hi - lo).astype('int64')


def _bucket(ageing):
    a = ageing.astype('float64')
    result = pd.Series(None, index=a.index, dtype=object)
    result = result.mask(a > 30, '30 +')
    for lo, hi, label in AGEING_BUCKETS:
        result = result.mask((a >= lo) & (a <= hi), label)
    return result


def compute_max_date(df):
    """NULL-ignoring GREATEST of received_date and the three conditional
    event dates - MySQL's GREATEST() returns NULL if any argument is NULL,
    unlike Postgres/Redshift, so this is done in pandas instead."""
    cols = ['received_date', 'reopen_date_with_condition', 'goahead_date_with_condition', 'insuff_fulfill_date']
    return df[cols].max(axis=1, skipna=True)


def compute_tat_fields(df, baseline_col, holiday_dates, today, variant):
    """Adds due-date, ageing, ageing-bucket and IT/OT columns derived from
    `baseline_col` (received_date for Flat, max_date for Recal)."""
    df = df.copy()
    names = COLUMN_NAMES[variant]
    baseline = df[baseline_col]
    check_tat = pd.to_numeric(df['check_tat'], errors='coerce')
    category_key = df['check_tat_category'].map(_category_key)

    unrecognized = df.loc[category_key.isna() & df['check_tat_category'].notna(), 'check_tat_category'].unique()
    if len(unrecognized):
        raise ValueError(
            f"Unrecognized check_tat_category value(s) not handled by the TAT logic: {list(unrecognized)}"
        )

    closure_or_today = df['check_closure_date'].fillna(pd.Timestamp(today))
    check_tat_safe = check_tat.fillna(0).astype('int64')
    span_start = min(baseline.min(), df['received_date'].min())
    span_end = max(
        closure_or_today.max(),
        pd.Timestamp(today),
        (baseline.fillna(pd.Timestamp(today)) + pd.to_timedelta(check_tat_safe + 30, unit='D')).max(),
    )
    exclusion_sets = _daily_calendar(span_start, span_end, holiday_dates)

    first_pass_due = baseline + pd.to_timedelta(check_tat_safe, unit='D')

    total_holi = pd.Series(0, index=df.index, dtype='int64')
    due_date = pd.Series(pd.NaT, index=df.index)
    ageing = pd.Series(np.nan, index=df.index, dtype='float64')

    for key in ('Working', 'Calendar', 'WorkingPlusSaturday'):
        mask = (category_key == key) & baseline.notna() & check_tat.notna()
        if not mask.any():
            continue
        sorted_excluded = exclusion_sets[key]

        holi_count = _count_excluded(sorted_excluded, baseline[mask], first_pass_due[mask])
        total_holi.loc[mask] = holi_count
        due_date.loc[mask] = baseline[mask] + pd.to_timedelta(check_tat[mask].astype('int64') + holi_count, unit='D')

        window_days = (closure_or_today[mask] - baseline[mask]).dt.days + 1
        excluded_in_window = _count_excluded(sorted_excluded, baseline[mask], closure_or_today[mask])
        # Clamp at 0: a small number of cases have check_closure_date before
        # the baseline date (a data artifact, not a valid window) - COUNT(1)
        # over an empty/reversed date range is 0, never negative.
        ageing.loc[mask] = (window_days - excluded_in_window).clip(lower=0).astype('float64')

    dow = due_date.dt.dayofweek
    opt_due_date = due_date.copy()
    working_mask = category_key == 'Working'
    opt_due_date.loc[working_mask & (dow == 5)] += pd.Timedelta(days=2)
    opt_due_date.loc[working_mask & (dow == 6)] += pd.Timedelta(days=1)
    wps_mask = category_key == 'WorkingPlusSaturday'
    opt_due_date.loc[wps_mask & (dow == 6)] += pd.Timedelta(days=1)

    ageing_bucket = _bucket(ageing)

    it_override = df['check_status'].isin(IT_OVERRIDE_CHECK_STATUSES)
    known = ageing.notna() & check_tat.notna()
    tat_status = pd.Series(None, index=df.index, dtype=object)
    tat_status.loc[known] = np.where(ageing[known].to_numpy() <= check_tat[known].to_numpy(), 'IT', 'OT')
    tat_status.loc[it_override] = 'IT'

    df[names['total']] = total_holi
    df[names['due']] = due_date
    df[names['opt_due']] = opt_due_date
    df[names['ageing']] = ageing.astype('Int64')
    df[names['bucket']] = ageing_bucket
    df[names['tat_status']] = tat_status
    return df


def build_flat_and_recal(base_df, holiday_dates, today):
    """Given the raw per-check dataframe from QUERY_BASE_DATA and the full
    holiday calendar, returns (flat_df, recal_df) matching the exact column
    names/contract build_report.py expects (see checkwise_flat000.csv /
    checkwise_recal000.csv from the original pipeline)."""
    base_df = base_df.copy()
    base_df['max_date'] = compute_max_date(base_df)

    flat_df = compute_tat_fields(base_df, 'received_date', holiday_dates, today, variant='flat')
    recal_df = compute_tat_fields(base_df, 'max_date', holiday_dates, today, variant='recal')

    return flat_df, recal_df
