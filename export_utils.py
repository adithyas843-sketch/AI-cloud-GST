"""Professional Excel export helpers for the GST reconciliation workbook."""
from __future__ import annotations

from io import BytesIO
from numbers import Number
from typing import Any

import pandas as pd
import xlsxwriter


SOURCE_HEADER = "#1F4E78"
MATCH_FILL = "#C6EFCE"
WARNING_FILL = "#FFF2CC"
MISMATCH_FILL = "#FFC7CE"

TEMPLATE_DEFINITIONS = {
    "GSTR-1": (
        ["Customer GSTIN", "Invoice no", "Date", "Customer Name", "Item Taxable Value", "IGST", "CGST", "SGST"],
        [["27ABCDE1234F1Z5", "G1-1001", "05-Apr-2026", "Aster Retail", 100000, 18000, 0, 0], ["29PQRSX6789L1Z2", "G1-1002", "08-Apr-2026", "Bluebird Foods", 50000, 0, 4500, 4500]],
    ),
    "Tally": (
        ["GSTIN", "Voucher Number", "Voucher Date", "Party Name", "Taxable Value", "IGST", "CGST", "SGST"],
        [["27ABCDE1234F1Z5", "G1-1001", "05-Apr-2026", "Aster Retail", 100000, 18000, 0, 0], ["29PQRSX6789L1Z2", "G1-1002", "08-Apr-2026", "Bluebird Foods", 50000, 0, 4500, 4500]],
    ),
    "Zoho": (
        ["GSTIN", "Invoice Number", "Invoice Date", "Customer Name", "Taxable Value", "IGST", "CGST", "SGST"],
        [["27ABCDE1234F1Z5", "G1-1001", "05-Apr-2026", "Aster Retail", 100000, 18000, 0, 0], ["29PQRSX6789L1Z2", "G1-1002", "08-Apr-2026", "Bluebird Foods", 50000, 0, 4500, 4500]],
    ),
    "GSTR-3B": (
        ["Month", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Apr-2026", 150000, 18000, 4500, 4500], ["May-2026", 75000, 13500, 0, 0]],
    ),
}


def sample_template_bytes(source_name: str) -> bytes:
    """Create a deploy-safe sample template when a bundled file is unavailable."""
    headers, rows = TEMPLATE_DEFINITIONS[source_name]
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    sheet = workbook.add_worksheet("Sample Data")
    header_format = workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#FFFFFF", "bg_color": SOURCE_HEADER, "border": 1})
    body_format = workbook.add_format({"font_name": "Book Antiqua", "border": 1, "border_color": "#D9D9D9"})
    for col, header in enumerate(headers):
        sheet.write(0, col, header, header_format)
        values = [str(row[col]) for row in rows]
        sheet.set_column(col, col, min(max(len(str(header)), *(len(value) for value in values)) + 2, 32))
    for row_index, row in enumerate(rows, start=1):
        for col, value in enumerate(row):
            _write_value(sheet, row_index, col, value, body_format)
    sheet.autofilter(0, 0, len(rows), len(headers) - 1)
    sheet.freeze_panes(1, 0)
    sheet.hide_gridlines(2)
    workbook.close()
    return output.getvalue()


def _is_blank(value: Any) -> bool:
    return value is None or (not isinstance(value, str) and bool(pd.isna(value)))


def _source_row(frame: pd.DataFrame, gstin: Any, invoice_number: Any) -> int | None:
    """Return the Excel row for the best available source record."""
    if frame is None or frame.empty:
        return None
    gstin_value, invoice_value = str(gstin), str(invoice_number)
    exact = frame[(frame["GSTIN"] == gstin_value) & (frame["Invoice Number"] == invoice_value)]
    by_invoice = frame[frame["Invoice Number"] == invoice_value]
    by_gstin = frame[frame["GSTIN"] == gstin_value]
    candidate = exact if not exact.empty else by_invoice if not by_invoice.empty else by_gstin
    return None if candidate.empty else int(candidate.index[0]) + 2


def _write_value(worksheet, row: int, col: int, value: Any, cell_format) -> None:
    if _is_blank(value):
        worksheet.write_blank(row, col, None, cell_format)
    elif isinstance(value, pd.Timestamp):
        worksheet.write_datetime(row, col, value.to_pydatetime(), cell_format)
    elif isinstance(value, Number) and not isinstance(value, bool):
        worksheet.write_number(row, col, float(value), cell_format)
    else:
        worksheet.write(row, col, str(value), cell_format)


def _column_width(values: pd.Series, heading: Any) -> int:
    """Return a safe display width for any source-data type or column heading."""
    rendered = ["" if _is_blank(value) else str(value) for value in values.head(500).tolist()]
    widest = max([len(str(heading)), *(len(value) for value in rendered)], default=len(str(heading)))
    return min(max(widest + 2, 12), 32)


def _write_source_sheet(workbook, sheet_name: str, source_data: pd.DataFrame, mismatch_rows: set[int], formats: dict) -> None:
    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.hide_gridlines(2)
    headers = list(source_data.columns)
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, formats["header"])
        worksheet.set_column(col, col, _column_width(source_data.iloc[:, col], header))
    for excel_row, (_, row) in enumerate(source_data.iterrows(), start=2):
        row_format = formats["source_mismatch"] if excel_row in mismatch_rows else formats["body"]
        for col, value in enumerate(row.tolist()):
            _write_value(worksheet, excel_row - 1, col, value, row_format)
    if headers:
        worksheet.autofilter(0, 0, max(len(source_data), 1), len(headers) - 1)
    worksheet.freeze_panes(1, 0)


def _reconciliation_row_format(status: Any, mismatch_type: Any, formats: dict):
    if str(status).strip().lower() == "matched":
        return formats["match"]
    mismatch_text = str(mismatch_type).lower()
    if any(token in mismatch_text for token in ("missing", "duplicate", "date mismatch", "partial")):
        return formats["warning"]
    return formats["mismatch"]


def professional_report_bytes(
    gstr1_original: pd.DataFrame,
    books_original: pd.DataFrame,
    gstr1_standardized: pd.DataFrame,
    books_standardized: pd.DataFrame,
    reconciliation: pd.DataFrame,
) -> bytes:
    """Create the three-sheet, audit-friendly GST reconciliation workbook."""
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    formats = {
        "header": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#FFFFFF", "bg_color": SOURCE_HEADER, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#D9E2F3"}),
        "body": workbook.add_format({"font_name": "Book Antiqua", "border": 1, "border_color": "#D9D9D9"}),
        "source_mismatch": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#9C0006", "bg_color": MISMATCH_FILL, "border": 1, "border_color": "#D9D9D9"}),
        "title": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#FFFFFF", "bg_color": SOURCE_HEADER, "align": "center", "valign": "vcenter", "font_size": 14, "border": 1, "border_color": SOURCE_HEADER}),
        "summary_label": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#1F1F1F", "bg_color": "#D9EAF7", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#A6A6A6"}),
        "summary_value": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#1F1F1F", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#A6A6A6", "num_format": "#,##0.0"}),
        "match": workbook.add_format({"font_name": "Book Antiqua", "bg_color": MATCH_FILL, "border": 1, "border_color": "#A6A6A6"}),
        "warning": workbook.add_format({"font_name": "Book Antiqua", "bg_color": WARNING_FILL, "border": 1, "border_color": "#A6A6A6"}),
        "mismatch": workbook.add_format({"font_name": "Book Antiqua", "bg_color": MISMATCH_FILL, "border": 1, "border_color": "#A6A6A6"}),
    }

    gstr1_highlights: set[int] = set()
    books_highlights: set[int] = set()
    for _, record in reconciliation[reconciliation["Status"] != "Matched"].iterrows():
        gstr1_row = _source_row(gstr1_standardized, record["GSTIN"], record["Invoice Number"])
        books_row = _source_row(books_standardized, record["GSTIN"], record["Invoice Number"])
        if gstr1_row:
            gstr1_highlights.add(gstr1_row)
        if books_row:
            books_highlights.add(books_row)

    _write_source_sheet(workbook, "GSTR-1 Data", gstr1_original, gstr1_highlights, formats)
    _write_source_sheet(workbook, "Books Data", books_original, books_highlights, formats)

    sheet = workbook.add_worksheet("Reconciliation")
    sheet.hide_gridlines(2)
    sheet.merge_range(0, 0, 0, 13, "GST Reconciliation Summary", formats["title"])
    total = len(reconciliation)
    matched = int((reconciliation["Status"] == "Matched").sum()) if not reconciliation.empty else 0
    summary_items = [
        ("Total Invoices", total),
        ("Matched", matched),
        ("Mismatches", total - matched),
        ("Match %", round(100 * matched / max(total, 1), 1)),
        ("GST Exposure", float(reconciliation.get("Exposure Amount", pd.Series(dtype=float)).sum())),
    ]
    for index, (label, value) in enumerate(summary_items):
        start_col = index * 2
        sheet.merge_range(1, start_col, 1, start_col + 1, label, formats["summary_label"])
        sheet.merge_range(2, start_col, 2, start_col + 1, value, formats["summary_value"])

    headers = ["GSTIN", "Invoice Number", "Customer Name", "GSTR-1 Row Number", "Open GSTR-1", "Books Row Number", "Open Books", "Status", "Mismatch Type", "Taxable Difference", "IGST Difference", "CGST Difference", "SGST Difference", "Remarks"]
    header_row = 5
    for col, header in enumerate(headers):
        sheet.write(header_row, col, header, formats["header"])
    widths = [18, 18, 24, 16, 16, 16, 16, 12, 34, 18, 15, 15, 15, 40]
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)

    for excel_row, (_, record) in enumerate(reconciliation.iterrows(), start=header_row + 2):
        row_index = excel_row - 1
        row_format = _reconciliation_row_format(record["Status"], record["Mismatch Type"], formats)
        gstr1_row = _source_row(gstr1_standardized, record["GSTIN"], record["Invoice Number"])
        books_row = _source_row(books_standardized, record["GSTIN"], record["Invoice Number"])
        values = [record["GSTIN"], record["Invoice Number"], record.get("Customer", ""), gstr1_row, None, books_row, None, record["Status"], record["Mismatch Type"], record["Taxable Difference"], record["IGST Difference"], record["CGST Difference"], record["SGST Difference"], record["Remarks"]]
        for col, value in enumerate(values):
            if col == 4 and gstr1_row:
                sheet.write_formula(row_index, col, f'=HYPERLINK("#\'GSTR-1 Data\'!A{gstr1_row}","Open GSTR-1")', row_format, "Open GSTR-1")
            elif col == 6 and books_row:
                sheet.write_formula(row_index, col, f'=HYPERLINK("#\'Books Data\'!A{books_row}","Open Books")', row_format, "Open Books")
            else:
                _write_value(sheet, row_index, col, value, row_format)
    if headers:
        sheet.autofilter(header_row, 0, max(header_row + 1, header_row + len(reconciliation)), len(headers) - 1)
    sheet.freeze_panes(header_row + 1, 0)
    workbook.close()
    return output.getvalue()
