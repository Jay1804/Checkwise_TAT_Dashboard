"""
Builds a client's Check TAT Analysis workbook from live database data.

The Excel/pivot-building code below (write_data_sheet, build_count_pivot,
style_pivot_chrome, build_client_workbook, with_excel_session, etc.) is
carried over verbatim from the original checkwise_flat_recal project's
build_report.py - it is pure presentation logic with no dependency on where
the data came from, so it needed no changes. Only the data-acquisition layer
is new: fetch_clients_data() replaces the original S3/parquet download with a
direct database query (sql_queries.py) plus the pandas TAT/ageing
calculations (tat_logic.py) that substitute for the SQL this project
inherited but which cannot run against this database (see tat_logic.py and
sql_queries.py docstrings for why).

Run with: streamlit run app.py
"""
import datetime as dt
import os
import subprocess
import time

import numpy as np
import pandas as pd
import pythoncom
import win32com.client as win32
import win32process

import sql_queries
import tat_logic
from db_connect import get_connection

os.chdir(os.path.dirname(os.path.abspath(__file__)))

OUTPUT_FILE_TEMPLATE = 'Check Wise Check TAT Analysis-{client_id}.xlsx'

FLAT_TAT_COL = 'tat_check flat tat '
RECAL_TAT_COL = 'tat_check tat'
FLAT_AGEING_BUCKET_COL = 'ageing_bucket_flat'
RECAL_AGEING_BUCKET_COL = 'ageing_bucket'
CHECK_NAME_COL = 'unique check name'
SEVERITY_COL = 'check_severity'
MONTH_COL = 'received_month'

FLAT_DATE_COLS = [
    'received_date', 'insuff_fulfill_date', 'check_closure_date',
    'reopen_date_with_condition', 'goahead_date_with_condition', 'max_date',
    'chk_due_date_flat', 'chk_opt_flat_due_date',
]
RECAL_DATE_COLS = [
    'received_date', 'insuff_fulfill_date', 'check_closure_date',
    'reopen_date_with_condition', 'goahead_date_with_condition', 'max_date',
    'chk_due_date_recal', 'chk_opt_recal_due_date',
]

BASE_COLUMN_ORDER = [
    'case_id', 'case_check_id', 'client_external_id', 'company_name', 'cat', 'cat_tl',
    'account_manager', 'process_name', 'case_ars_no', 'received_date', 'case_status',
    'insuff_fulfill_date', 'check_status', 'check_severity', 'check_closure_date',
    'check_tat_category', 'check_tat', 'reopen_date_with_condition', 'goahead_date_with_condition',
    'max_date', 'unique check name', 'check_ops_name',
]
FLAT_COLUMN_ORDER = BASE_COLUMN_ORDER + [
    'total_sat_sun_holi_flat', 'chk_due_date_flat', 'chk_opt_flat_due_date',
    'ageing_flat', 'ageing_bucket_flat', FLAT_TAT_COL,
]
RECAL_COLUMN_ORDER = BASE_COLUMN_ORDER + [
    'total_sat_sun_holi', 'chk_due_date_recal', 'chk_opt_recal_due_date',
    'ageing', 'ageing_bucket', RECAL_TAT_COL,
]

AGEING_BUCKET_ORDER = ['0--2', '3--10', '11--14', '21--25', '15--20', '26--30', '30 +']

xlRowField = 1
xlColumnField = 2
xlDatabase = 1
xlCount = -4112
xlSrcRange = 1
xlYes = 1
xlPercentOfRow = 6
xlEdgeLeft = 7
xlEdgeTop = 8
xlEdgeBottom = 9
xlEdgeRight = 10
xlInsideVertical = 11
xlInsideHorizontal = 12
xlContinuous = 1
xlThin = 2
xlMedium = -4138
xlFreezePanes = 3
xlCenter = -4108
xlLeft = -4131


def excel_color(r, g, b):
    """VBA-style RGB->OLE color conversion expected by Font.Color/Interior.Color."""
    return r + (g << 8) + (b << 16)


COLOR_TITLE_BG = excel_color(31, 73, 125)      # dark blue
COLOR_TITLE_FONT = excel_color(255, 255, 255)  # white
COLOR_HEADER_BG = excel_color(217, 217, 217)   # light gray
COLOR_TOTAL_BG = excel_color(242, 242, 242)    # very light gray
COLOR_BORDER = excel_color(166, 166, 166)      # medium gray
COLOR_TAB_DATA = excel_color(89, 89, 89)       # slate gray tab for raw data sheets
COLOR_TAB_SUMMARY = excel_color(31, 73, 125)   # dark blue tab for summary sheets


MONTH_NUMBER_FORMAT = r"mmm\'yy"  # displays as Jan'25 while the underlying value is a real date, so it still sorts chronologically


def add_month_column(df):
    df = df.copy()
    # First-of-month date (not a string) so the pivot's row field sorts
    # chronologically - the mmm'yy display comes from a number format, below.
    df[MONTH_COL] = pd.to_datetime(df['received_date'].dt.strftime('%Y-%m-01'))
    df[MONTH_COL] = df[MONTH_COL].where(df['received_date'].notna(), None)
    return df


DATE_DISPLAY_FORMAT = 'dd-mm-yyyy'

EXCEL_EPOCH = pd.Timestamp('1899-12-30')


def normalize_value(v):
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, pd.Timestamp):
        # Write as an Excel serial-date float, not a Python datetime - pywin32's
        # bulk 2D-array COM write silently shifts naive datetime objects by the
        # local timezone offset (e.g. -5:30 for IST), which was corrupting every
        # date column by up to a full calendar day. Floats have no timezone to
        # get misinterpreted, so this sidesteps the bug entirely; the cell just
        # needs an explicit date NumberFormat since Excel won't auto-detect a
        # plain float as a date.
        return (v - EXCEL_EPOCH).total_seconds() / 86400.0
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        return None if np.isnan(v) else float(v)
    if isinstance(v, np.bool_):
        return bool(v)
    if pd.isna(v):
        return None
    return v


def dataframe_to_rows(df):
    return [[normalize_value(v) for v in row] for row in df.itertuples(index=False, name=None)]


def write_data_sheet(wb, sheet_name, df, table_name, sheet_index, date_cols=()):
    if sheet_index <= wb.Sheets.Count:
        ws = wb.Sheets(sheet_index)
        ws.Name = sheet_name
    else:
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheet_name

    ncols = len(df.columns)
    header = list(df.columns)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, ncols)).Value = [header]

    nrows = len(df)
    if nrows > 0:
        data = dataframe_to_rows(df)
        ws.Range(ws.Cells(2, 1), ws.Cells(1 + nrows, ncols)).Value = data

    last_row = 1 + nrows
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, ncols))
    lo = ws.ListObjects.Add(xlSrcRange, rng, None, xlYes)
    lo.Name = table_name

    if nrows > 0:
        for col_name in date_cols:
            if col_name in header:
                col_idx = header.index(col_name) + 1
                ws.Range(ws.Cells(2, col_idx), ws.Cells(last_row, col_idx)).NumberFormat = DATE_DISPLAY_FORMAT
        if MONTH_COL in header:
            month_col_idx = header.index(MONTH_COL) + 1
            ws.Range(ws.Cells(2, month_col_idx), ws.Cells(last_row, month_col_idx)).NumberFormat = MONTH_NUMBER_FORMAT

    header_rng = ws.Range(ws.Cells(1, 1), ws.Cells(1, ncols))
    header_rng.Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(last_row, ncols)).Columns.AutoFit()
    ws.Tab.Color = COLOR_TAB_DATA
    ws.Activate()
    ws.Range('A2').Select()
    ws.Application.ActiveWindow.FreezePanes = True
    return ws, last_row


def add_sheet(wb, sheet_name):
    ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = sheet_name
    return ws


def build_count_pivot(wb, ws, anchor_row, anchor_col, table_name, row_field, col_field, value_field,
                       pivot_name, percent_of_row=False, title=None, row_number_format=None):
    if title:
        ws.Cells(anchor_row - 1, anchor_col).Value = title

    pc = wb.PivotCaches().Create(xlDatabase, table_name)
    dest = ws.Cells(anchor_row, anchor_col)
    pt = pc.CreatePivotTable(TableDestination=dest, TableName=pivot_name)

    pt.PivotFields(row_field).Orientation = xlRowField
    if row_number_format:
        pt.PivotFields(row_field).NumberFormat = row_number_format
    pt.PivotFields(col_field).Orientation = xlColumnField
    label = '% of Row Total' if percent_of_row else 'Count'
    field = pt.AddDataField(pt.PivotFields(value_field), label, xlCount)
    if percent_of_row:
        field.Calculation = xlPercentOfRow
        field.NumberFormat = '0.0%'
    else:
        field.NumberFormat = '#,##0'

    style_pivot_chrome(ws, pt, title=title)
    return pt


def style_title_bar(ws, row, first_col, last_col, text_value=None):
    """(Re-)merge a title bar across exactly [first_col, last_col]. If the
    cell is already part of a merge (e.g. we're widening it after adding
    extra columns), unmerge first and carry the existing text forward."""
    anchor = ws.Cells(row, first_col)
    if text_value is None:
        text_value = anchor.Value
        if anchor.MergeCells:
            anchor.MergeArea.UnMerge()
    elif anchor.MergeCells:
        anchor.MergeArea.UnMerge()

    title_rng = ws.Range(ws.Cells(row, first_col), ws.Cells(row, last_col))
    title_rng.Merge()
    title_rng.Value = text_value
    title_rng.Font.Bold = True
    title_rng.Font.Size = 12
    title_rng.Font.Color = COLOR_TITLE_FONT
    title_rng.Interior.Color = COLOR_TITLE_BG
    title_rng.HorizontalAlignment = xlCenter


def style_pivot_chrome(ws, pt, title=None):
    """Give each pivot a self-contained 'card' look: a full-width merged title
    bar, shaded/bold header rows, a bold Grand Total row, and a complete grid
    of borders around and inside the table - then autofit the columns."""
    tbl = pt.TableRange2
    top_row = tbl.Row
    last_row = top_row + tbl.Rows.Count - 1
    first_col = tbl.Column
    last_col = first_col + tbl.Columns.Count - 1
    title_row = top_row - 1

    if title:
        style_title_bar(ws, title_row, first_col, last_col, text_value=title)

    header_rng = ws.Range(ws.Cells(top_row, first_col), ws.Cells(pt.DataBodyRange.Row - 1, last_col))
    header_rng.Font.Bold = True
    header_rng.Interior.Color = COLOR_HEADER_BG
    header_rng.HorizontalAlignment = xlCenter

    total_rng = ws.Range(ws.Cells(last_row, first_col), ws.Cells(last_row, last_col))
    total_rng.Font.Bold = True
    total_rng.Interior.Color = COLOR_TOTAL_BG
    total_rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    total_rng.Borders(xlEdgeTop).Weight = xlThin

    apply_card_border(ws, title_row if title else top_row, top_row, last_row, first_col, last_col)
    ws.Range(ws.Cells(top_row, first_col), ws.Cells(last_row, last_col)).Columns.AutoFit()


def apply_card_border(ws, box_top_row, header_top_row, last_row, first_col, last_col):
    """Thin grid lines inside the table, plus a slightly heavier outer box
    around the whole title+table block, for a self-contained 'card' look."""
    inner = ws.Range(ws.Cells(header_top_row, first_col), ws.Cells(last_row, last_col))
    inner.Borders(xlInsideVertical).LineStyle = xlContinuous
    inner.Borders(xlInsideVertical).Weight = xlThin
    inner.Borders(xlInsideVertical).Color = COLOR_BORDER
    inner.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    inner.Borders(xlInsideHorizontal).Weight = xlThin
    inner.Borders(xlInsideHorizontal).Color = COLOR_BORDER

    box = ws.Range(ws.Cells(box_top_row, first_col), ws.Cells(last_row, last_col))
    for edge in (xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom):
        box.Borders(edge).LineStyle = xlContinuous
        box.Borders(edge).Weight = xlMedium
        box.Borders(edge).Color = COLOR_BORDER


def set_pivot_item_order(pt, field_name, order):
    """Force a pivot field's column items into a specific order (Excel's
    default is alphabetical) via PivotItem.Position. Silently skips any item
    not present in this particular pivot (e.g. a bucket with zero rows in a
    filtered/date-ranged report)."""
    field = pt.PivotFields(field_name)
    for i, item_name in enumerate(order, start=1):
        try:
            field.PivotItems(item_name).Position = i
        except Exception:
            pass


def next_anchor_row(pt, gap=3):
    """Row to start the next stacked pivot at, based on where this one actually
    ended - avoids hardcoding row offsets that could overlap if a pivot's row
    count varies (e.g. a client with a wider date range or more check types)."""
    return pt.TableRange2.Row + pt.TableRange2.Rows.Count + gap


def col_letter(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def add_month_tat_ratio_formulas(ws, pt):
    """Add IT%/OT% ratio columns next to a month x IT/OT count pivot, including
    the Grand Total row (a manual formula, not a pivot-native feature).
    Uses the pivot's real DataBodyRange geometry rather than assumed offsets."""
    data_body = pt.DataBodyRange
    first_data_row = data_body.Row
    last_data_row = first_data_row + data_body.Rows.Count - 1
    first_data_col = data_body.Column
    n_data_cols = data_body.Columns.Count
    if n_data_cols != 3:
        raise ValueError(f'Expected 3 data columns (IT, OT, Grand Total), got {n_data_cols}')
    last_data_col = first_data_col + n_data_cols - 1  # Grand Total column

    header_row = first_data_row - 1
    ratio_col = last_data_col + 1
    it_l = col_letter(first_data_col)       # 'IT' sorts before 'OT' alphabetically
    ot_l = col_letter(first_data_col + 1)
    tot_l = col_letter(last_data_col)

    header_rng = ws.Range(ws.Cells(header_row, ratio_col), ws.Cells(header_row, ratio_col + 1))
    header_rng.Value = [['IT %', 'OT %']]
    header_rng.Font.Bold = True
    header_rng.Interior.Color = COLOR_HEADER_BG
    header_rng.HorizontalAlignment = xlCenter

    for r in range(first_data_row, last_data_row + 1):
        ws.Cells(r, ratio_col).Formula = f'={it_l}{r}/{tot_l}{r}'
        ws.Cells(r, ratio_col + 1).Formula = f'={ot_l}{r}/{tot_l}{r}'

    data_rng = ws.Range(ws.Cells(first_data_row, ratio_col), ws.Cells(last_data_row, ratio_col + 1))
    data_rng.NumberFormat = '0.0%'

    total_rng = ws.Range(ws.Cells(last_data_row, ratio_col), ws.Cells(last_data_row, ratio_col + 1))
    total_rng.Font.Bold = True
    total_rng.Interior.Color = COLOR_TOTAL_BG
    total_rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    total_rng.Borders(xlEdgeTop).Weight = xlThin

    # Widen the title bar and card border to also span the ratio columns, so
    # the whole pivot+ratio section reads as one cohesive box instead of two.
    tbl = pt.TableRange2
    style_title_bar(ws, tbl.Row - 1, tbl.Column, ratio_col + 1)
    apply_card_border(ws, tbl.Row - 1, tbl.Row, last_data_row, tbl.Column, ratio_col + 1)
    ws.Range(ws.Cells(header_row, ratio_col), ws.Cells(last_data_row, ratio_col + 1)).Columns.AutoFit()


def _process_alive(pid):
    result = subprocess.run(['tasklist', '/FI', f'PID eq {pid}'], capture_output=True, text=True)
    return str(pid) in result.stdout


def _ensure_excel_process_killed(pid, wait_seconds=5):
    if pid is None:
        return
    for _ in range(wait_seconds):
        if not _process_alive(pid):
            return
        time.sleep(1)
    if _process_alive(pid):
        subprocess.run(['taskkill', '/PID', str(pid), '/F'], capture_output=True)
        print(f'  Force-killed lingering Excel process (PID {pid})')


def with_excel_session(fn):
    """Runs fn(excel) inside a managed Excel COM session. Guarantees cleanup,
    including force-killing a lingering EXCEL.EXE if Quit() doesn't fully work
    (a known pywin32 issue where a COM reference can keep the process alive)."""
    pythoncom.CoInitialize()
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel_pid = None
        try:
            excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        except Exception:
            pass
        try:
            return fn(excel)
        finally:
            excel.Quit()
            del excel
            _ensure_excel_process_killed(excel_pid)
    finally:
        pythoncom.CoUninitialize()


def finalize_summary_sheet(ws):
    """Sheet-wide polish once every pivot on it is built: tab color, no
    gridlines (the pivots have their own borders now), and one final autofit
    across everything actually used, so column widths reflect the widest
    content across ALL sections rather than whichever pivot autofit last."""
    ws.Tab.Color = COLOR_TAB_SUMMARY
    ws.Activate()
    ws.Application.ActiveWindow.DisplayGridlines = False
    ws.UsedRange.Columns.AutoFit()


def build_client_workbook(excel, client_id, flat_df, recal_df):
    output_path = os.path.abspath(OUTPUT_FILE_TEMPLATE.format(client_id=client_id))
    wb = excel.Workbooks.Add()
    try:
        write_data_sheet(wb, 'DATA_Flat', flat_df, 'tblDataFlat', 1, date_cols=FLAT_DATE_COLS)
        write_data_sheet(wb, 'DATA_Recal', recal_df, 'tblDataRecal', 2, date_cols=RECAL_DATE_COLS)
        print(f'  DATA_Flat/DATA_Recal written: {len(flat_df)}/{len(recal_df)} rows')

        ws_flat = add_sheet(wb, 'Flat_summary')
        row = 3
        pt = build_count_pivot(wb, ws_flat, row, 1, 'tblDataFlat', MONTH_COL, FLAT_TAT_COL, 'case_id',
                                'ITOTPivotFlat', title='IT vs OT by month', row_number_format=MONTH_NUMBER_FORMAT)
        add_month_tat_ratio_formulas(ws_flat, pt)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_flat, row, 1, 'tblDataFlat', CHECK_NAME_COL, FLAT_AGEING_BUCKET_COL, 'case_id',
                                'AgeingCountFlat', title='Ageing by check type (Flat) - Count')
        set_pivot_item_order(pt, FLAT_AGEING_BUCKET_COL, AGEING_BUCKET_ORDER)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_flat, row, 1, 'tblDataFlat', CHECK_NAME_COL, FLAT_AGEING_BUCKET_COL, 'case_id',
                                'AgeingPctFlat', percent_of_row=True, title='Ageing by check type (Flat) - % of row')
        set_pivot_item_order(pt, FLAT_AGEING_BUCKET_COL, AGEING_BUCKET_ORDER)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_flat, row, 1, 'tblDataFlat', CHECK_NAME_COL, SEVERITY_COL, 'case_id',
                                'SeverityCountFlat', title='Severity by check type (Flat) - Count')
        row = next_anchor_row(pt)
        build_count_pivot(wb, ws_flat, row, 1, 'tblDataFlat', CHECK_NAME_COL, SEVERITY_COL, 'case_id',
                           'SeverityPctFlat', percent_of_row=True, title='Severity by check type (Flat) - % of row')
        finalize_summary_sheet(ws_flat)
        print('  Flat_summary: IT/OT + ageing + severity pivots built (Flat)')

        ws_recal = add_sheet(wb, 'Recal_summary')
        row = 3
        pt = build_count_pivot(wb, ws_recal, row, 1, 'tblDataRecal', MONTH_COL, RECAL_TAT_COL, 'case_id',
                                'ITOTPivotRecal', title='IT vs OT by month', row_number_format=MONTH_NUMBER_FORMAT)
        add_month_tat_ratio_formulas(ws_recal, pt)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_recal, row, 1, 'tblDataRecal', CHECK_NAME_COL, RECAL_AGEING_BUCKET_COL, 'case_id',
                                'AgeingCountRecal', title='Ageing by check type (Recal) - Count')
        set_pivot_item_order(pt, RECAL_AGEING_BUCKET_COL, AGEING_BUCKET_ORDER)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_recal, row, 1, 'tblDataRecal', CHECK_NAME_COL, RECAL_AGEING_BUCKET_COL, 'case_id',
                                'AgeingPctRecal', percent_of_row=True, title='Ageing by check type (Recal) - % of row')
        set_pivot_item_order(pt, RECAL_AGEING_BUCKET_COL, AGEING_BUCKET_ORDER)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_recal, row, 1, 'tblDataRecal', CHECK_NAME_COL, SEVERITY_COL, 'case_id',
                                'SeverityCountRecal', title='Severity by check type (Recal) - Count')
        row = next_anchor_row(pt)
        build_count_pivot(wb, ws_recal, row, 1, 'tblDataRecal', CHECK_NAME_COL, SEVERITY_COL, 'case_id',
                           'SeverityPctRecal', percent_of_row=True, title='Severity by check type (Recal) - % of row')
        finalize_summary_sheet(ws_recal)
        print('  Recal_summary: IT/OT + ageing + severity pivots built (Recal)')

        wb.Sheets(1).Activate()
        wb.SaveAs(output_path)
        print(f'  Saved {output_path}')
    finally:
        wb.Close(SaveChanges=False)


class NoDataError(RuntimeError):
    """Raised when the database query returns no rows for the requested
    client/date range."""


def list_available_clients():
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(sql_queries.QUERY_AVAILABLE_CLIENTS)
            return cur.fetchall()
    finally:
        conn.close()


def get_client_date_range(client_external_id):
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(sql_queries.QUERY_CLIENT_DATE_RANGE, {'client_external_id': client_external_id})
            row = cur.fetchone()
            return (row['min_date'], row['max_date']) if row else (None, None)
    finally:
        conn.close()


def fetch_clients_data(client_external_ids, date_from, date_to):
    """Runs the base query (for all requested clients at once, via an IN
    clause - much cheaper than an unfiltered all-clients scan, see
    sql_queries.py) + holiday calendar against the live database and applies
    the checkwise TAT/ageing business logic. Returns (flat_df, recal_df)
    combined across every requested client, each still tagged with its own
    client_external_id. Raises NoDataError only if NONE of the requested
    clients have any rows in range - a client with zero rows among several
    requested is reported by the caller instead, since that's a normal
    per-client outcome, not a failure of the whole batch."""
    client_external_ids = tuple(int(c) for c in client_external_ids)
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(sql_queries.QUERY_BASE_DATA, {
                'client_external_ids': client_external_ids,
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
            f'No case-check rows found for the selected client(s) between {date_from} and {date_to}.'
        )

    base_df = pd.DataFrame(base_rows)
    date_cols = ['received_date', 'check_closure_date', 'reopen_date_with_condition',
                 'goahead_date_with_condition', 'insuff_fulfill_date']
    for col in date_cols:
        base_df[col] = pd.to_datetime(base_df[col])

    holidays_df = pd.DataFrame(holiday_rows)
    holiday_dates = pd.to_datetime(holidays_df['holiday_date']).values if len(holidays_df) else np.array([], dtype='datetime64[ns]')

    flat_df, recal_df = tat_logic.build_flat_and_recal(base_df, holiday_dates, pd.Timestamp.today().date())
    flat_df = add_month_column(flat_df)[FLAT_COLUMN_ORDER + [MONTH_COL]].reset_index(drop=True)
    recal_df = add_month_column(recal_df)[RECAL_COLUMN_ORDER + [MONTH_COL]].reset_index(drop=True)
    return flat_df, recal_df


def generate_reports_for_clients(client_external_ids, date_from, date_to, client_names=None, progress_callback=None):
    """Fetches data for every requested client in one query, then builds one
    workbook per client inside a single shared Excel session (much cheaper
    than opening/closing Excel per client). Returns a list of per-client
    result dicts: {client_external_id, company_name, output_path, error}.
    A client with no rows, or one whose workbook build fails, is reported
    with `error` set rather than aborting the rest of the batch.

    client_names: optional {client_external_id: company_name} for clients
    that may end up with zero rows (so the result still shows a name).
    progress_callback: optional fn(done_count, total_count, company_name)
    called after each client finishes, for UI progress updates."""
    client_names = client_names or {}
    flat_df, recal_df = fetch_clients_data(client_external_ids, date_from, date_to)

    results = []

    def run(excel):
        total = len(client_external_ids)
        for i, client_id in enumerate(client_external_ids, start=1):
            flat_subset = flat_df[flat_df['client_external_id'] == client_id].reset_index(drop=True)
            recal_subset = recal_df[recal_df['client_external_id'] == client_id].reset_index(drop=True)
            company_name = (
                flat_subset['company_name'].iloc[0] if len(flat_subset)
                else client_names.get(client_id, str(client_id))
            )
            result = {'client_external_id': client_id, 'company_name': company_name, 'output_path': None, 'error': None}
            try:
                if len(flat_subset) == 0:
                    raise NoDataError(f'No case-check rows found for {company_name} between {date_from} and {date_to}.')
                print(f'client {client_id}: flat={len(flat_subset)} rows, recal={len(recal_subset)} rows')
                build_client_workbook(excel, client_id, flat_subset, recal_subset)
                result['output_path'] = os.path.abspath(OUTPUT_FILE_TEMPLATE.format(client_id=client_id))
            except Exception as e:
                result['error'] = str(e)
            results.append(result)
            if progress_callback:
                progress_callback(i, total, company_name)

    with_excel_session(run)
    return results


def _aggregate_variant(df, tat_col, ageing_col):
    """Per-client counts for one TAT variant (Flat or Recal). Case Count is
    COUNT(DISTINCT case_ars_no) - a case can have several checks (rows) -
    while Check Count is COUNT(case_check_id), i.e. the row count. IT/OT/
    ageing are per-check, so their percentages are derived from Check Count,
    never from averaging per-client percentages (that would misweight
    small/large clients)."""
    if df.empty:
        return pd.DataFrame(columns=['client_external_id', 'company_name', 'case_count', 'check_count', 'it_count', 'ot_count', 'avg_ageing'])
    return df.groupby(['client_external_id', 'company_name'], as_index=False).agg(
        case_count=('case_ars_no', 'nunique'),
        check_count=('case_check_id', 'count'),
        it_count=(tat_col, lambda s: (s == 'IT').sum()),
        ot_count=(tat_col, lambda s: (s == 'OT').sum()),
        avg_ageing=(ageing_col, 'mean'),
    )


def compute_dashboard_aggregates(flat_df, recal_df):
    """Per-client Flat AND Recal aggregates side by side: Case Count, Check
    Count, IT/OT counts, IT/OT % (from summed counts, not averaged
    per-client %) and average ageing, for each variant independently. This
    is the shared computation used by both the live (explicit client
    selection) dashboard path and the batched all-clients cache builder
    (refresh_dashboard_cache.py)."""
    flat_summary = _aggregate_variant(flat_df, FLAT_TAT_COL, 'ageing_flat')
    recal_summary = _aggregate_variant(recal_df, RECAL_TAT_COL, 'ageing')

    merged = flat_summary.merge(
        recal_summary, on=['client_external_id', 'company_name'], how='outer', suffixes=('_flat', '_recal'),
    )
    count_cols = ['case_count_flat', 'check_count_flat', 'it_count_flat', 'ot_count_flat',
                  'case_count_recal', 'check_count_recal', 'it_count_recal', 'ot_count_recal']
    for col in count_cols:
        merged[col] = merged[col].fillna(0).astype('int64')

    merged['it_pct_flat'] = _safe_pct(merged['it_count_flat'], merged['check_count_flat'])
    merged['ot_pct_flat'] = _safe_pct(merged['ot_count_flat'], merged['check_count_flat'])
    merged['it_pct_recal'] = _safe_pct(merged['it_count_recal'], merged['check_count_recal'])
    merged['ot_pct_recal'] = _safe_pct(merged['ot_count_recal'], merged['check_count_recal'])
    merged['avg_ageing_flat'] = merged['avg_ageing_flat'].round(1)
    merged['avg_ageing_recal'] = merged['avg_ageing_recal'].round(1)

    return merged.sort_values('it_pct_flat', ascending=False).reset_index(drop=True)


def _safe_pct(numerator, denominator):
    return (numerator / denominator.replace(0, pd.NA) * 100).round(1)


def overall_totals(summary):
    """Portfolio-wide totals from a per-client summary dataframe (as returned
    by compute_dashboard_aggregates), computed from summed counts - never by
    averaging per-client percentages, per the report's aggregation rule."""
    totals = {}
    for variant in ('flat', 'recal'):
        case_count = int(summary[f'case_count_{variant}'].sum())
        check_count = int(summary[f'check_count_{variant}'].sum())
        it = int(summary[f'it_count_{variant}'].sum())
        ot = int(summary[f'ot_count_{variant}'].sum())
        totals[variant] = {
            'case_count': case_count,
            'check_count': check_count,
            'it_count': it,
            'ot_count': ot,
            'it_pct': round(it / check_count * 100, 1) if check_count else None,
            'ot_pct': round(ot / check_count * 100, 1) if check_count else None,
            'avg_ageing': round(summary[f'avg_ageing_{variant}'].mean(), 1) if len(summary) else None,
        }
    return totals


def get_dashboard_summary(client_external_ids, date_from, date_to):
    """Live per-client Flat+Recal aggregates for an explicit, bounded list of
    clients - fast because it's bounded (see fetch_clients_data/sql_queries.py
    docstrings for why an unbounded all-clients scan isn't viable on this
    database). Used for the dashboard's "N selected clients" mode; the "All
    Clients" mode instead reads a background-computed cache, see
    dashboard_cache.py."""
    flat_df, recal_df = fetch_clients_data(client_external_ids, date_from, date_to)
    return compute_dashboard_aggregates(flat_df, recal_df)


def previous_period(date_from, date_to):
    """The same-length window immediately preceding [date_from, date_to],
    used for trend/increment comparisons. Returns (prev_from, prev_to) as
    plain date objects."""
    length = date_to - date_from
    prev_to = date_from - dt.timedelta(days=1)
    prev_from = prev_to - length
    return prev_from, prev_to
