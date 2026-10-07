"""
Builds each client's Check TAT Analysis workbook entirely from scratch out of
the live database data - no dependency on any pre-existing template xlsx file.

Data comes live from the checkpoint_live MySQL database (see db_source.py).
Run with: python build_report.py <client_external_id> [YYYY-MM-DD YYYY-MM-DD]
"""
import os
import subprocess
import time

import numpy as np
import pandas as pd
import pythoncom
import win32com.client as win32
import win32process

import db_source

os.chdir(os.path.dirname(os.path.abspath(__file__)))

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
EXCEL_EPOCH = pd.Timestamp('1899-12-30')

OUTPUT_FILE_TEMPLATE = 'Check Wise Check TAT Analysis-{client_id}.xlsx'

FLAT_TAT_COL = 'tat_check flat tat '
RECAL_TAT_COL = 'tat_check tat'
FLAT_AGEING_BUCKET_COL = 'ageing_bucket_flat'
RECAL_AGEING_BUCKET_COL = 'ageing_bucket'
CHECK_NAME_COL = 'unique check name'
SEVERITY_COL = 'check_severity'
MONTH_COL = 'received_month'

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


def normalize_value(v):
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, pd.Timestamp):
        # Excel serial-date float, not a datetime: pywin32's bulk 2D-array COM
        # write shifts naive datetimes by the local timezone offset (see CLAUDE.md).
        return (v - EXCEL_EPOCH).total_seconds() / 86400.0
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        return None if np.isnan(v) else float(v)
    if isinstance(v, np.bool_):
        return bool(v)
    return v


def dataframe_to_rows(df):
    return [[normalize_value(v) for v in row] for row in df.itertuples(index=False, name=None)]


def add_month_column(df):
    df = df.copy()
    # First-of-month date (not a string) so the pivot's row field sorts
    # chronologically - the mmm'yy display comes from a number format, below.
    df[MONTH_COL] = pd.to_datetime(df['received_date'].dt.strftime('%Y-%m-01'))
    df[MONTH_COL] = df[MONTH_COL].where(df['received_date'].notna(), None)
    return df


DATE_DISPLAY_FORMAT = 'dd-mm-yyyy'


def combine_flat_recal(flat_df, recal_df):
    """One row per case_check_id: the columns shared by both extracts once,
    then the Flat-only columns, then the Recal-only columns. Outer join so a
    check present in only one extract is never dropped."""
    shared = [c for c in flat_df.columns if c in recal_df.columns]
    flat_only = [c for c in flat_df.columns if c not in shared]
    recal_only = [c for c in recal_df.columns if c not in shared]
    merged = flat_df[shared + flat_only].merge(
        recal_df[['case_check_id'] + recal_only], on='case_check_id', how='outer')
    # Checks only in Recal have no shared-column values from the Flat side.
    missing = ~merged['case_check_id'].isin(flat_df['case_check_id'])
    if missing.any():
        extra = recal_df.set_index('case_check_id').loc[merged.loc[missing, 'case_check_id'], shared[1:]]
        for c in shared[1:]:
            merged.loc[missing, c] = extra[c].values
    return merged[shared + flat_only + recal_only].reset_index(drop=True)


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
        combined_df = combine_flat_recal(flat_df, recal_df)
        write_data_sheet(wb, 'DATA_Combined', combined_df, 'tblDataCombined', 1,
                         date_cols=sorted(set(FLAT_DATE_COLS) | set(RECAL_DATE_COLS)))
        print(f'  DATA_Combined written: {len(combined_df)} rows (flat={len(flat_df)}, recal={len(recal_df)})')

        ws_flat = add_sheet(wb, 'Flat_summary')
        row = 3
        pt = build_count_pivot(wb, ws_flat, row, 1, 'tblDataCombined', MONTH_COL, FLAT_TAT_COL, 'case_id',
                                'ITOTPivotFlat', title='IT vs OT by month', row_number_format=MONTH_NUMBER_FORMAT)
        add_month_tat_ratio_formulas(ws_flat, pt)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_flat, row, 1, 'tblDataCombined', CHECK_NAME_COL, FLAT_AGEING_BUCKET_COL, 'case_id',
                                'AgeingCountFlat', title='Ageing by check type (Flat) - Count')
        set_pivot_item_order(pt, FLAT_AGEING_BUCKET_COL, AGEING_BUCKET_ORDER)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_flat, row, 1, 'tblDataCombined', CHECK_NAME_COL, FLAT_AGEING_BUCKET_COL, 'case_id',
                                'AgeingPctFlat', percent_of_row=True, title='Ageing by check type (Flat) - % of row')
        set_pivot_item_order(pt, FLAT_AGEING_BUCKET_COL, AGEING_BUCKET_ORDER)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_flat, row, 1, 'tblDataCombined', CHECK_NAME_COL, SEVERITY_COL, 'case_id',
                                'SeverityCountFlat', title='Severity by check type (Flat) - Count')
        row = next_anchor_row(pt)
        build_count_pivot(wb, ws_flat, row, 1, 'tblDataCombined', CHECK_NAME_COL, SEVERITY_COL, 'case_id',
                           'SeverityPctFlat', percent_of_row=True, title='Severity by check type (Flat) - % of row')
        finalize_summary_sheet(ws_flat)
        print('  Flat_summary: IT/OT + ageing + severity pivots built (Flat)')

        ws_recal = add_sheet(wb, 'Recal_summary')
        row = 3
        pt = build_count_pivot(wb, ws_recal, row, 1, 'tblDataCombined', MONTH_COL, RECAL_TAT_COL, 'case_id',
                                'ITOTPivotRecal', title='IT vs OT by month', row_number_format=MONTH_NUMBER_FORMAT)
        add_month_tat_ratio_formulas(ws_recal, pt)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_recal, row, 1, 'tblDataCombined', CHECK_NAME_COL, RECAL_AGEING_BUCKET_COL, 'case_id',
                                'AgeingCountRecal', title='Ageing by check type (Recal) - Count')
        set_pivot_item_order(pt, RECAL_AGEING_BUCKET_COL, AGEING_BUCKET_ORDER)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_recal, row, 1, 'tblDataCombined', CHECK_NAME_COL, RECAL_AGEING_BUCKET_COL, 'case_id',
                                'AgeingPctRecal', percent_of_row=True, title='Ageing by check type (Recal) - % of row')
        set_pivot_item_order(pt, RECAL_AGEING_BUCKET_COL, AGEING_BUCKET_ORDER)
        row = next_anchor_row(pt)
        pt = build_count_pivot(wb, ws_recal, row, 1, 'tblDataCombined', CHECK_NAME_COL, SEVERITY_COL, 'case_id',
                                'SeverityCountRecal', title='Severity by check type (Recal) - Count')
        row = next_anchor_row(pt)
        build_count_pivot(wb, ws_recal, row, 1, 'tblDataCombined', CHECK_NAME_COL, SEVERITY_COL, 'case_id',
                           'SeverityPctRecal', percent_of_row=True, title='Severity by check type (Recal) - % of row')
        finalize_summary_sheet(ws_recal)
        print('  Recal_summary: IT/OT + ageing + severity pivots built (Recal)')

        wb.Sheets(1).Activate()
        wb.SaveAs(output_path)
        print(f'  Saved {output_path}')
    finally:
        wb.Close(SaveChanges=False)


def build_reports_for_clients(client_ids, date_from=None, date_to=None, progress_callback=None):
    """Fetches every requested client in one query, then builds one workbook
    per client inside a single shared Excel session. Returns a list of
    {client_id, output_path, error}; a client with no rows or a failed build
    is reported via `error` instead of aborting the rest.
    progress_callback: optional fn(done_count, total_count, client_id)."""
    from datetime import datetime, time as dtime
    client_ids = [int(c) for c in client_ids]
    start = datetime.combine(date_from, dtime.min) if date_from else datetime(2000, 1, 1)
    end = datetime.combine(date_to, dtime.max.replace(microsecond=0)) if date_to else datetime.now()

    results = []
    try:
        flat_all, recal_all = db_source.fetch_flat_recal(client_ids, start, end)
    except db_source.NoDataError as e:
        return [{'client_id': c, 'output_path': None, 'error': str(e)} for c in client_ids]
    flat_all = add_month_column(flat_all)
    recal_all = add_month_column(recal_all)

    def run(excel):
        for i, cid in enumerate(client_ids, start=1):
            flat_df = flat_all[flat_all['client_external_id'] == cid].reset_index(drop=True)
            recal_df = recal_all[recal_all['client_external_id'] == cid].reset_index(drop=True)
            if len(flat_df) == 0 and len(recal_df) == 0:
                results.append({'client_id': cid, 'output_path': None,
                                'error': 'No rows in the selected date range.'})
            else:
                print(f'client {cid}: flat={len(flat_df)} rows, recal={len(recal_df)} rows')
                try:
                    build_client_workbook(excel, cid, flat_df, recal_df)
                    results.append({'client_id': cid, 'error': None,
                                    'output_path': os.path.abspath(OUTPUT_FILE_TEMPLATE.format(client_id=cid))})
                except Exception as e:
                    results.append({'client_id': cid, 'output_path': None, 'error': f'{type(e).__name__}: {e}'})
            if progress_callback:
                progress_callback(i, len(client_ids), cid)

    with_excel_session(run)
    return results


def main():
    import sys
    from datetime import date
    if len(sys.argv) not in (2, 4):
        sys.exit('Usage: python build_report.py <client_external_id> [YYYY-MM-DD YYYY-MM-DD]')
    client_id = int(sys.argv[1])
    date_from = date.fromisoformat(sys.argv[2]) if len(sys.argv) == 4 else None
    date_to = date.fromisoformat(sys.argv[3]) if len(sys.argv) == 4 else None
    for r in build_reports_for_clients([client_id], date_from, date_to):
        print(r['output_path'] or f"client {r['client_id']} failed: {r['error']}")


if __name__ == '__main__':
    main()
