# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Not a software project — a data pipeline for producing per-client "Check TAT Analysis" Excel reports from
case-check data pulled **live from the `checkpoint_live` MySQL database**. There is no build, lint, or test
suite; run it via the Streamlit UI (`streamlit run streamlit_app.py`) or `python build_report.py <client_id>
[YYYY-MM-DD YYYY-MM-DD]`. There is **no S3 dependency** anymore.

Each client's workbook is built **entirely from scratch** on every run — there is no template `.xlsx` to
clone. (The old template-cloning approach kept failing when the template file went missing; suspected
antivirus/EDR reacting to heavy Excel COM automation, never confirmed. Building fresh removes that single
point of failure.)

Dependencies are in `requirements.txt`: `pandas`, `numpy`, `pymysql`, `python-dotenv`, `pywin32`, `streamlit`.
The COM steps require a real, licensed Microsoft Excel on Windows — `openpyxl` alone cannot create/refresh
PivotTables.

## Current pipeline

```
checkpoint_live (MySQL) --db_source.py + sql_queries.py--> raw per-check rows --tat_logic.py--> Flat/Recal dataframes
                                                                                  --build_report.py (Excel COM)--> workbook
                                                                                  <-- streamlit_app.py (UI)
```

- **`db_source.py`** — connects using `.env` (`DB_HOST/PORT/NAME/USER/PASSWORD`; never hardcode). `fetch_flat_recal(client_ids,
  date_from, date_to)` runs `QUERY_BASE_DATA` for the given clients in one query and returns `(flat_df, recal_df)`;
  `available_clients()` feeds the UI picker. **Never remove the client filter** — an unfiltered scan across all
  clients is impractically slow.
- **`sql_queries.py` / `tat_logic.py`** — raw-field SQL and the pandas Flat/Recal due-date / ageing / bucket / IT-OT
  logic. Copied from `Streamlit_app/modules/checkwise_tat_dashboard/vendor/`; keep in sync. The one place TAT
  logic is defined — reuse, don't reimplement. Output matched the old S3 CSVs for client 5613 except a few dozen
  still-open checks whose ageing moves with "today" (plus new cases since the CSV snapshot).
- **`build_report.py`** — `build_reports_for_clients(client_ids, date_from, date_to, progress_callback)` fetches all
  selected clients in one query, then builds one workbook per client in a single shared Excel session (a client
  with no rows or a failed build is reported via `error`, not fatal). Each workbook has:
  - `DATA_Combined` — ONE data tab (Excel Table `tblDataCombined`): Flat and Recal rows joined on
    `case_check_id` (outer join) — shared columns once, then the Flat-only columns, then the Recal-only columns.
    Bold+frozen header, dates formatted `dd-mm-yyyy`.
  - `Flat_summary` / `Recal_summary` — two separate summary tabs, both pivoting off `tblDataCombined` (Flat uses
    the `*_flat` columns, Recal the others). Each has an "IT vs OT by month" pivot (row field is a computed
    `received_month` first-of-month date shown as `mmm'yy`) with manual `IT %`/`OT %` columns incl. Grand Total,
    plus ageing-bucket and severity breakdowns (count + % of row) by check type. Ageing buckets use a fixed
    custom order via `PivotItem.Position` (`AGEING_BUCKET_ORDER`). Each pivot is a styled "card" (merged title
    bar, shaded headers, bold Grand Total, grid borders, tab colors, gridlines off).
- **`streamlit_app.py`** — UI. A searchable multi-select of clients (labels `"<client_external_id> - <company name>"`,
  so type an ID or a name), a date range (filters `received_date`), live log + progress bar, and one download
  button per generated client (results persisted in `st.session_state`, since a download click reruns the
  script). Blocks with an error if `EXCEL.EXE` is already running, since a workbook left open elsewhere can be
  silently closed by the automation.

## Removed legacy code

The old S3 / template-cloning scripts (`run_pipeline.py`, `download_*.py`, `split_and_update.py`,
`convert_to_dynamic_tables.py`, `refresh_pivots.py`, `finalize.py`, `inspect_5613.py`, `fix_pivot_ranges.py`),
the S3 CSV extracts and old generated workbooks were deleted. They contained hardcoded AWS credentials —
rotate those keys if still active.

## Deployment (planned — server not yet available)

Not deployed yet; this will be deployed to a server once it is up and running. Constraints to plan around:

- **Needs Windows + licensed Microsoft Excel** (COM automation via `pywin32`). It will NOT run on Linux, Docker,
  Streamlit Community Cloud or typical PaaS. The target server must be a Windows machine/VM with Excel installed.
- **Single-user / one job at a time.** The app refuses to generate if any `EXCEL.EXE` is running and force-kills a
  lingering one afterwards, so concurrent users would clash. Don't run other Excel work on the same server
  (or under the same session) as the app.
- **Linux alternative if Windows is not possible:** replace the Excel pivots with pre-computed summary tables
  written via `openpyxl`/`xlsxwriter` (static tables instead of refreshable pivots). The sibling project
  `checkwise_tat_dashboard` has a `linux_report.py` that may already do this.
- **Before go-live:**
  - Rotate the DB password (it was shared in chat) and use a read-only DB user.
  - Allowlist the server's IP on the RDS security group (`ab-mum-prod-bridge...ap-south-1.rds.amazonaws.com:3306`).
  - Provide `DB_*` via server environment variables / secrets rather than a checked-in or copied `.env`
    (`.env` is gitignored; `.env.example` lists the names).
  - Put authentication/network restriction (VPN/internal only) in front of the Streamlit app — it has no login.
  - Run as a service (e.g. NSSM/Task Scheduler) with a fixed port, and make sure the service account can launch Excel.
  - Do a full end-to-end run (Excel build included) and compare numbers against a trusted report — the full
    workbook build against the DB source has not been verified yet, only the DB fetch.

## Known gotchas hit during development (avoid re-introducing)

- **pywin32 bulk-array datetime write is timezone-buggy.** Writing a naive Python `datetime`/`pd.Timestamp`
  via a 2D `Range.Value = [[...]]` COM array assignment silently shifts it by the local timezone offset
  (observed: -5:30 for IST), often rolling the date back a full day. Fix in place: `normalize_value()` in
  `build_report.py` converts timestamps to Excel serial-date floats (`(v - pd.Timestamp('1899-12-30')).total_seconds()/86400`)
  instead of passing datetime objects, then an explicit `NumberFormat` is applied to the cells. This bug
  silently corrupted every date column in every workbook for a period during development before being
  caught — if date values ever look off by ~1 day again, check this first.
- **`Application.Quit()` doesn't reliably kill `EXCEL.EXE`.** A lingering COM reference can leave an
  invisible orphan process running indefinitely. `_ensure_excel_process_killed()` in `build_report.py`
  tracks the Excel process's PID (via `win32process.GetWindowThreadProcessId(excel.Hwnd)`) and force-kills
  it via `taskkill` if it's still alive a few seconds after `Quit()`.
- **A pivot field can't be both the column-breakdown field and the count/value field.** Doing so silently
  drops the column split (Excel just shows one aggregate column). Always count a neutral field
  (`case_id`) instead.
- **`xlPercentOfRow` is `6`, not `7`** (`7` is `xlPercentOfColumn`) — easy to get backwards from memory;
  verify against actual output (e.g. does `part/row_total` match what's displayed) if adding new
  percentage pivots.
- **Deleting and re-`Add()`-ing a Table/PivotTable to resize it can wipe header text**, replacing it with
  generic `Column1, Column2...` placeholders. Use `ListObject.Resize(range)` / re-point the pivot cache by
  table name instead of delete+recreate.
- **`Range.Group()` for automatic date grouping (Year>Month hierarchy) was unreliable via COM** (produced
  a broken single group instead of a proper hierarchy). Sidestepped entirely by pre-computing a
  first-of-month date column in pandas and pivoting on that as a plain field — simpler and reliable.
- Merged title-bar cells must be sized/merged **after** the pivot exists (once its actual column width is
  known), and re-merged wider if columns are added afterward (e.g. the IT%/OT% ratio columns) — the merge
  doesn't auto-expand.
