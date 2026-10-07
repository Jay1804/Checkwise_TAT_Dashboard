# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Not a software project — a data pipeline / working folder for producing per-client "Check TAT Analysis"
Excel reports from raw case-check data pulled live from the checkpoint_live MySQL database. There is no build, lint, or test suite; scripts are
run ad hoc with `python <script>.py`, or via the Streamlit UI (`streamlit run streamlit_app.py`).

Each client's workbook is now built **entirely from scratch** out of the live database data on every run —
there is no template `.xlsx` the pipeline depends on or clones. This was a deliberate architecture change:
the old template-cloning approach (see "Legacy / superseded scripts" below) kept failing when the template
file went missing (repeatedly disappeared during this project's development — see git-less history/
conversation for details — with no definitive root cause found; suspected antivirus/EDR reacting to heavy
Excel COM automation, though never confirmed). Building fresh each time removes that single point of failure.

Dependencies are pinned in `requirements.txt`: `pandas`, `numpy`, `pymysql`, `python-dotenv` (pyarrow no longer needed)
`pywin32`, `streamlit`. The COM automation
steps require a real, licensed Microsoft Excel installation on the machine — `openpyxl` alone cannot
create/refresh PivotTables, it only reads/writes the underlying XML.

## Current pipeline

**Data source (default): live MySQL `checkpoint_live`**, credentials in `.env` (`DB_*`, never hardcode). `db_source.py`
runs `sql_queries.QUERY_BASE_DATA` (one client at a time - never remove the client filter, unfiltered scans are
impractically slow) and `tat_logic.py` computes the Flat/Recal due date/ageing/bucket/IT-OT columns in pandas
(copied from `Streamlit_app/modules/checkwise_tat_dashboard/vendor/`; keep in sync). Output matches the legacy S3 CSVs
to within a handful of still-open checks whose ageing moves with "today". There is no S3 dependency anymore: `build_report.py` is self-contained
and no longer imports `run_pipeline.py`. Data goes into ONE `DATA_Combined` tab (Flat + Recal columns joined on `case_check_id`);
`Flat_summary`/`Recal_summary` both pivot off `tblDataCombined`.

- **`build_report.py`** — the pipeline. `main()` downloads the latest data and builds a workbook for every
  `client_external_id` found. `build_reports_for_clients(client_ids, date_from, date_to)` builds one workbook per selected
  client, optionally filtered to a `received_date` range — this is what the Streamlit app calls.
  For each client it creates a **brand-new** workbook via Excel COM (`excel.Workbooks.Add()`) with:
  - `DATA_Flat` / `DATA_Recal` — that client's rows as real Excel Tables (`tblDataFlat`/`tblDataRecal`),
    bold+frozen header row, date columns formatted `dd-mm-yyyy`.
  - `Flat_summary` / `Recal_summary` — for each of Flat and Recal: an "IT vs OT by month" pivot (row field
    is a computed `received_month` column, first-of-month date formatted `mmm'yy` so it displays as
    `Jan'25` but still sorts chronologically) with manual `IT %`/`OT %` ratio columns including the Grand
    Total row (pivots don't support cross-column ratios natively), plus ageing-bucket and
    severity breakdowns (count + % of row) by check type. Ageing buckets are forced into a fixed custom
    order via `PivotItem.Position` (`AGEING_BUCKET_ORDER`), not Excel's default alphabetical order.
    Every pivot section is styled as a self-contained "card": full-width merged/centered title bar,
    shaded bold header rows, bold Grand Total row, full grid borders, sheet tab colors, gridlines off.
- **`streamlit_app.py`** — UI over `build_report.py`. Lets the user type a `client_external_id` and pick a
  date range, shows a live log during generation, and offers a download button for the resulting file. It
  also previews available client IDs / date range from whatever was last downloaded locally (separate
  from actually generating a report, which always re-downloads fresh data). Blocks with an error if
  `EXCEL.EXE` is already running, since a workbook left open elsewhere can get silently closed by the
  automation.

## Removed legacy code

The old S3 / template-cloning scripts (`run_pipeline.py`, `download_*.py`, `split_and_update.py`,
`convert_to_dynamic_tables.py`, `refresh_pivots.py`, `finalize.py`, `inspect_5613.py`, `fix_pivot_ranges.py`)
and the S3 CSV extracts were deleted. Nothing depends on S3 anymore.

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
