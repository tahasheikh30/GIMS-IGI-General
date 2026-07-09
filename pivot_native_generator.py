"""
Creates genuine, interactive Excel PivotTables (with a real field list,
drag-and-drop, filter dropdowns, drill-down, and refresh) by driving
Microsoft Excel itself via COM automation.

This is necessary because no Python library (openpyxl, xlsxwriter, etc.)
can write the actual PivotTable/PivotCache object Excel uses - they can only
produce a static table that *looks* like a pivot but has none of the
underlying functionality. Only Excel can create a real one, so this module
opens each client's workbook in a hidden Excel instance, builds the pivot
from the "Raw Data" sheet, and saves.

Requirements: Windows + Microsoft Excel installed (same requirement as the
Outlook COM integration). This cannot be tested outside of Windows + Excel.
"""
import os

try:
    import win32com.client as win32
    from win32com.client import constants as xlc
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


def start_excel():
    """Launch a hidden Excel instance to reuse across many report files
    (much faster than starting Excel fresh for every client)."""
    if not EXCEL_AVAILABLE:
        return None
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def stop_excel(excel_app):
    if excel_app is None:
        return
    try:
        excel_app.Quit()
    except Exception:
        pass


def add_native_pivot_table(filepath, client_name, raw_sheet_name="Raw Data",
                            pivot_sheet_name="Outstanding Balance",
                            row_field="Department", column_field="Year",
                            value_field="O/s", excel_app=None):
    """
    Opens `filepath` (an .xlsx that already has a `raw_sheet_name` sheet
    with the client's transaction rows - e.g. from
    report_generator.write_raw_only_workbook) and inserts a genuine,
    interactive Excel PivotTable on a new sheet, built from that data.

    excel_app: pass an already-running Excel Application COM object to
        reuse across a batch (recommended). If None, a temporary instance
        is started and quit just for this one file.
    """
    if not EXCEL_AVAILABLE:
        raise RuntimeError(
            "pywin32 / Excel COM not available - native PivotTable creation "
            "requires Windows with Microsoft Excel installed."
        )

    own_excel = excel_app is None
    excel = excel_app or start_excel()

    wb = excel.Workbooks.Open(os.path.abspath(filepath))
    try:
        raw_ws = wb.Sheets(raw_sheet_name)
        used = raw_ws.UsedRange
        n_rows = used.Rows.Count
        n_cols = used.Columns.Count
        source_range = raw_ws.Range(raw_ws.Cells(1, 1), raw_ws.Cells(n_rows, n_cols))

        # Remove a stale pivot sheet from a prior run of this same file, if any
        for sh in list(wb.Sheets):
            if sh.Name == pivot_sheet_name:
                sh.Delete()

        pivot_ws = wb.Sheets.Add(Before=raw_ws)
        pivot_ws.Name = pivot_sheet_name
        pivot_ws.Range("A1").Value = "Client:"
        pivot_ws.Range("B1").Value = client_name
        pivot_ws.Range("A1").Font.Bold = True
        pivot_ws.Range("B1").Font.Bold = True

        pivot_cache = wb.PivotCaches().Create(SourceType=xlc.xlDatabase, SourceData=source_range)
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=pivot_ws.Range("A3"),
            TableName="OutstandingBalancePivot",
        )

        pivot_table.PivotFields(row_field).Orientation = xlc.xlRowField
        pivot_table.PivotFields(column_field).Orientation = xlc.xlColumnField
        data_field = pivot_table.AddDataField(
            pivot_table.PivotFields(value_field), f"Sum of {value_field}", xlc.xlSum
        )
        data_field.NumberFormat = "#,##0;(#,##0);-"
        pivot_table.ColumnGrand = True
        pivot_table.RowGrand = True
        try:
            pivot_table.TableStyle2 = "PivotStyleMedium9"
        except Exception:
            pass  # cosmetic only - don't fail the whole report over a style name

        # Excel's default headers say "Row Labels" / "Column Labels" rather
        # than the actual field names - rename them so the sheet reads
        # "Department" / "Year" instead. Using Find() rather than a fixed
        # cell offset so this is robust regardless of exact pivot layout.
        try:
            table_range = pivot_table.TableRange2
            row_label_cell = table_range.Find("Row Labels")
            if row_label_cell is not None:
                row_label_cell.Value = row_field
            col_label_cell = table_range.Find("Column Labels")
            if col_label_cell is not None:
                col_label_cell.Value = column_field
        except Exception:
            pass  # cosmetic only - don't fail the whole report over a header rename

        pivot_ws.Columns("A:F").AutoFit()
        pivot_ws.Move(Before=wb.Sheets(1))  # put the pivot sheet first

        wb.Save()
    except Exception:
        wb.Close(SaveChanges=False)
        raise
    else:
        wb.Close(SaveChanges=False)  # already saved above; avoid the "save changes?" prompt
    finally:
        if own_excel:
            stop_excel(excel)


if __name__ == "__main__":
    # Manual smoke test - only runs meaningfully on Windows with Excel installed.
    import sys
    if not EXCEL_AVAILABLE:
        print("pywin32/Excel not available in this environment - cannot run this test here.")
        sys.exit(0)

    import data_processor as dp
    import report_generator as rg

    raw = dp.load_raw_data("/mnt/user-data/uploads/Karachi_Conventional_Outstanding.xlsx", sheet_name="Sheet1")
    clean, issues = dp.validate_data(raw)
    df, totals = dp.build_pivot(clean)
    q = dp.qualifying_clients(df, totals, threshold=2_000_000)
    client = q.index[0]
    raw_rows = dp.get_client_raw_rows(df, client)

    out_dir = "./native_pivot_test_output"
    fp = rg.write_raw_only_workbook(client, raw_rows, out_dir)
    add_native_pivot_table(fp, client_name=client)
    print("Done:", fp)
