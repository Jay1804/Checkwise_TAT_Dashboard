# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Streamlit MIS reporting app ("Checkwise Check TAT Analysis") that connects directly to a live
`checkpoint_live` MySQL database, applies TAT (turnaround-time)/ageing/IT-OT business logic, generates
per-client Excel workbooks via Excel COM automation, and emails them to recipients from an uploaded
mapping file. It also has an Executive Dashboard tab with KPIs, Flat-vs-Recal breakdowns, gauges, and
trend indicators.

It replaces an older pipeline (`checkwise_flat_recal`, not in this repo) that pulled the same data from
S3 parquet files and ran SQL written for Redshift/Postgres. That SQL doesn't run on this MySQL database
(different dialect, different table names) - see "Why the SQL/business-logic split looks the way it does"
below. There is no build step, no lint config, and no automated test suite; correctness was validated by
running the pipeline against the live DB and diff-checking output against the original pipeline's
historical CSV exports (see comments in `tat_logic.py`).

## Commands

```bash
pip install -r requirements.txt
streamlit run app.py                                    # the only entry point - do not run other files as the app
python refresh_dashboard_cache.py 2024-01-01 2026-08-14  # manually (re)compute the All-Clients dashboard cache
python db_connect.py                                     # sanity-check the DB connection/credentials
```

Requires Windows + a licensed Excel installation (Excel COM automation is used to build workbooks and
pivot tables - `openpyxl` alone can't create/refresh PivotTables). Credentials/config live in `.env`
(`DB_*`, `SMTP_*`, `EMAIL_SENDER`/`EMAIL_PASSWORD`) - never hardcode these; `db_connect.py` and
`email_sender.py` raise a clear `ConfigError`/`EmailConfigError` if a required var is missing.
`EMAIL_RECEIVERS`/`EMAIL_CC` in `.env` are vestigial and **not read by any code** - per-client recipients
now come exclusively from the mapping file uploaded in the UI (see below).

There's no test suite to run. To sanity-check a change to the TAT/ageing math, compare output against the
original pipeline's `checkwise_flat000.csv`/`checkwise_recal000.csv` if available, or spot-check a known
`case_check_id` by hand (see the worked example in `tat_logic.py`'s docstring).

## Architecture

### Data flow

```
checkpoint_live (MySQL) --sql_queries.py--> raw per-check rows --tat_logic.py--> Flat/Recal dataframes
                                                                                        |
                                                          report_generator.py builds the Excel workbook
                                                          (Excel COM) and/or dashboard aggregates
                                                                                        |
                                                    app.py (Streamlit) <--- dashboard_cache.py (All-Clients)
                                                          |
                                                    email_sender.py (per-client, via recipient_mapping.py)
```

- **`db_connect.py`** - MySQL connection from `.env`. Nothing else touches credentials directly.
- **`sql_queries.py`** - `QUERY_BASE_DATA` fetches one row per case-check (joined/decoded raw fields) for
  an explicit list of `client_external_id`s via an `IN` clause. **Never remove the client filter** - an
  unfiltered scan across all clients was tested directly against this DB and found impractically slow
  (a lightweight two-table count alone exceeded two minutes; there are 5,129 clients total). Also has
  `QUERY_HOLIDAYS`, `QUERY_AVAILABLE_CLIENTS`, `QUERY_CLIENT_DATE_RANGE`.
- **`tat_logic.py`** - pure-pandas business logic: due dates, ageing, ageing buckets, and IT/OT status,
  computed twice per row (once from `received_date` = "Flat", once from a reopen/insufficiency-aware
  `max_date` = "Recal"). This is the one place TAT/IT/OT/ageing are defined - reuse it, don't
  reimplement elsewhere. See "Why the SQL/business-logic split looks the way it does" below for why this
  isn't just SQL.
- **`report_generator.py`** - two responsibilities in one file (intentionally, carried over from the
  original `build_report.py`):
  1. Excel/COM workbook building (`build_client_workbook`, `write_data_sheet`, `build_count_pivot`,
     `with_excel_session`, etc.) - pure presentation logic, no DB dependency. Builds `DATA_Flat`/
     `DATA_Recal` sheets plus `Flat_summary`/`Recal_summary` pivot sheets per client.
  2. Data/aggregation: `fetch_clients_data` (DB fetch + `tat_logic` for an explicit client list),
     `generate_reports_for_clients` (builds all requested clients' workbooks in one shared Excel
     session), `compute_dashboard_aggregates`/`overall_totals` (Case Count = `COUNT(DISTINCT
     case_ars_no)`, Check Count = `COUNT(case_check_id)`, IT/OT %, ageing - always from summed counts,
     never by averaging per-client percentages), `previous_period` (same-length prior window for trend).
- **`dashboard_cache.py` / `refresh_dashboard_cache.py`** - the dashboard's "All Clients" view can't
  query live (same reason as above, at full scale). `refresh_dashboard_cache.py` is launched as a
  **detached background subprocess** (`dashboard_cache.start_refresh`) that batches through every client
  (50 at a time, reusing the same bounded query path) and writes progress to
  `dashboard_cache/status.json` plus a timestamped `dashboard_cache/snapshot_*.json` on completion. The
  two most recent snapshots back the trend/increment comparison. This can take well over an hour for the
  full client base - that's expected, not a bug.
- **`recipient_mapping.py`** - parses the user-uploaded CSV/XLSX (`client_external_id, company_name,
  To_address, CC_address`). A client absent from the mapping is generated but never emailed - addresses
  are never guessed or defaulted.
- **`email_sender.py`** - SMTP send using `.env` infra config + the mapping's per-client To/CC.
- **`app.py`** - Streamlit UI, two pages switched via `st.segmented_control` (see gotcha below):
  Executive Dashboard and Generate & Send Reports. Client-selection scope (none = All Clients / one /
  many) flows into every dashboard section from a single `summary` dataframe - don't fork the filter
  logic per-section.

### Why the SQL/business-logic split looks the way it does

The project inherited SQL written in Redshift/Postgres dialect (`DATEADD`, `DATE_PART('dow', ...)`,
`INTERVAL '2 days'`) referencing tables/columns that don't exist on this MySQL instance, plus a 3rd
holiday type that doesn't exist in the live `ec_master_holidays` table (only type 1 = Sunday+festivals,
type 2 = Saturday). Beyond the dialect gap, the original SQL chained several computed columns within one
`SELECT` list (e.g. a due-date expression referencing an aggregate calculated earlier in the same
`SELECT`) - MySQL doesn't support that. Rather than nesting many subquery levels, `sql_queries.py` fetches
raw/decoded fields only, and all date arithmetic lives in `tat_logic.py` (validated against the original
pipeline's historical output - see its docstring for the worked example and match rate).

### Known gotchas (avoid re-introducing)

- **`st.tabs()` resets to the first tab on every rerun.** `app.py` uses `st.segmented_control` instead
  (its value persists in `session_state` like any other widget) - don't switch this back to `st.tabs()`.
- **`st.download_button` triggers its own rerun.** Generation results must be persisted in
  `st.session_state` (see `generate_results` in `app.py`), not held in a local variable - otherwise
  clicking one download button wipes out the others before they can be clicked.
- **Excel COM is single-process and shared.** `app.py`'s `is_excel_running()` check blocks report
  generation if *any* `EXCEL.EXE` is running (including a report the user has open to inspect) - this is
  intentional, not a bug to "fix" by removing the check. `report_generator.with_excel_session` also
  force-kills a lingering `EXCEL.EXE` if `Quit()` doesn't fully terminate it (a known pywin32 issue).
- **pywin32 bulk datetime writes are timezone-buggy.** `report_generator.normalize_value` converts
  timestamps to Excel serial-date floats before writing via COM, rather than passing datetime objects
  directly, to avoid a silent timezone-shift bug that previously corrupted date columns by up to a day.
- **`FLAT_TAT_COL = 'tat_check flat tat '` has a trailing space.** This isn't a typo - it matches the
  original pipeline's column-naming contract (`checkwise_flat000.csv`), and pivot/aggregation code
  depends on the exact string.
- **This database may be a point-in-time snapshot, not a live replica.** If computed ageing/TAT numbers
  look stale near "today", check whether recent case activity is actually present in `ec_case_master`
  before assuming a logic bug.
