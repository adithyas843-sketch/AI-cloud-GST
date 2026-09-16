"""Template and professional Excel export helpers for GST reconciliation."""
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
    "Sales Register": (
        ["Document Type", "GSTIN", "Voucher Number", "Voucher Date", "Customer Name", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Invoice", "27ABCDE1234F1Z5", "INV-1001", "05-Apr-2026", "Aster Retail", 100000, 18000, 0, 0], ["Credit Note", "29PQRSX6789L1Z2", "CN-1002", "08-Apr-2026", "Bluebird Foods", 5000, 900, 0, 0]],
    ),
    "Purchase Register": (
        ["Document Type", "GSTIN", "Voucher Number", "Voucher Date", "Vendor Name", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Invoice", "27ABCDE1234F1Z5", "PINV-1001", "05-Apr-2026", "Aster Supplies", 85000, 15300, 0, 0], ["Debit Note", "29PQRSX6789L1Z2", "DN-1002", "08-Apr-2026", "Bluebird Traders", 5000, 0, 450, 450]],
    ),
    "GSTR-1": (
        ["Document Type", "Customer GSTIN", "Invoice No", "Date", "Customer Name", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Invoice", "27ABCDE1234F1Z5", "INV-1001", "05-Apr-2026", "Aster Retail", 100000, 18000, 0, 0], ["Credit Note", "29PQRSX6789L1Z2", "CN-1002", "08-Apr-2026", "Bluebird Foods", 5000, 900, 0, 0]],
    ),
    "GSTR-2B": (
        ["Document Type", "Supplier GSTIN", "Invoice No", "Invoice Date", "Supplier Name", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Invoice", "27ABCDE1234F1Z5", "PINV-1001", "05-Apr-2026", "Aster Supplies", 85000, 15300, 0, 0], ["Debit Note", "29PQRSX6789L1Z2", "DN-1002", "08-Apr-2026", "Bluebird Traders", 5000, 0, 450, 450]],
    ),
    "GSTR-3B": (
        ["Month", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Apr-2026", 150000, 18000, 4500, 4500], ["May-2026", 75000, 13500, 0, 0]],
    ),
}


def _is_blank(value: Any) -> bool:
    return value is None or (not isinstance(value, str) and bool(pd.isna(value)))


def _write_value(worksheet, row: int, col: int, value: Any, cell_format) -> None:
    if _is_blank(value):
        worksheet.write_blank(row, col, None, cell_format)
    elif isinstance(value, pd.Timestamp):
        worksheet.write_datetime(row, col, value.to_pydatetime(), cell_format)
    elif isinstance(value, Number) and not isinstance(value, bool):
        worksheet.write_number(row, col, float(value), cell_format)
    else:
        worksheet.write(row, col, str(value), cell_format)


def sample_template_bytes(source_name: str) -> bytes:
    """Create an Excel sample template for a specific upload control."""
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


def _column_width(values: pd.Series, heading: Any) -> int:
    rendered = ["" if _is_blank(value) else str(value) for value in values.head(500).tolist()]
    widest = max([len(str(heading)), *(len(value) for value in rendered)], default=len(str(heading)))
    return min(max(widest + 2, 12), 32)


def _source_row(frame: pd.DataFrame, document_type: Any, gstin: Any, invoice_number: Any) -> int | None:
    """Return an Excel row using the Document Type + GSTIN + invoice key."""
    if frame is None or frame.empty:
        return None
    if not {"Document Type", "GSTIN", "Invoice Number"}.issubset(frame.columns):
        return 2
    document_value, gstin_value, invoice_value = str(document_type), str(gstin), str(invoice_number)
    exact = frame[(frame["Document Type"] == document_value) & (frame["GSTIN"] == gstin_value) & (frame["Invoice Number"] == invoice_value)]
    by_document_invoice = frame[(frame["Document Type"] == document_value) & (frame["Invoice Number"] == invoice_value)]
    by_document_gstin = frame[(frame["Document Type"] == document_value) & (frame["GSTIN"] == gstin_value)]
    candidate = exact if not exact.empty else by_document_invoice if not by_document_invoice.empty else by_document_gstin
    return None if candidate.empty else int(candidate.index[0]) + 2


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


def _row_format(status: Any, mismatch_type: Any, formats: dict):
    if str(status).strip().lower() == "matched":
        return formats["match"]
    if any(token in str(mismatch_type).lower() for token in ("missing", "duplicate", "date mismatch", "partial")):
        return formats["warning"]
    return formats["mismatch"]


def reconciliation_export_bytes(source_sheet_name: str, reference_sheet_name: str, source_original: pd.DataFrame, reference_original: pd.DataFrame, source_standardized: pd.DataFrame, reference_standardized: pd.DataFrame, reconciliation: pd.DataFrame) -> bytes:
    """Create a linked, formatted three-sheet reconciliation workbook."""
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    formats = {
        "header": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#FFFFFF", "bg_color": SOURCE_HEADER, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#D9E2F3"}),
        "body": workbook.add_format({"font_name": "Book Antiqua", "border": 1, "border_color": "#D9D9D9"}),
        "source_mismatch": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#9C0006", "bg_color": MISMATCH_FILL, "border": 1, "border_color": "#D9D9D9"}),
        "title": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#FFFFFF", "bg_color": SOURCE_HEADER, "align": "center", "valign": "vcenter", "font_size": 14, "border": 1}),
        "summary_label": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "bg_color": "#D9EAF7", "align": "center", "border": 1, "border_color": "#A6A6A6"}),
        "summary_value": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "align": "center", "border": 1, "border_color": "#A6A6A6", "num_format": "#,##0.0"}),
        "match": workbook.add_format({"font_name": "Book Antiqua", "bg_color": MATCH_FILL, "border": 1, "border_color": "#A6A6A6"}),
        "warning": workbook.add_format({"font_name": "Book Antiqua", "bg_color": WARNING_FILL, "border": 1, "border_color": "#A6A6A6"}),
        "mismatch": workbook.add_format({"font_name": "Book Antiqua", "bold": True, "font_color": "#9C0006", "bg_color": MISMATCH_FILL, "border": 1, "border_color": "#A6A6A6"}),
    }
    source_highlights: set[int] = set()
    reference_highlights: set[int] = set()
    for _, record in reconciliation[reconciliation["Status"] != "Matched"].iterrows():
        source_row = _source_row(source_standardized, record.get("Document Type", "Invoice"), record.get("GSTIN", ""), record.get("Invoice Number", ""))
        reference_row = _source_row(reference_standardized, record.get("Document Type", "Invoice"), record.get("GSTIN", ""), record.get("Invoice Number", ""))
        if source_row:
            source_highlights.add(source_row)
        if reference_row:
            reference_highlights.add(reference_row)
    _write_source_sheet(workbook, source_sheet_name, source_original, source_highlights, formats)
    _write_source_sheet(workbook, reference_sheet_name, reference_original, reference_highlights, formats)

    sheet = workbook.add_worksheet("Reconciliation")
    sheet.hide_gridlines(2)
    headers = ["Document Type", "GSTIN", "Invoice Number", "Customer Name", "Source Row Number", "Open Source Record", "Reference Row Number", "Open Reference Record", "Status", "Mismatch Type", "Taxable Difference", "IGST Difference", "CGST Difference", "SGST Difference", "Remarks"]
    sheet.merge_range(0, 0, 0, len(headers) - 1, f"{source_sheet_name} vs {reference_sheet_name} Reconciliation", formats["title"])
    total = len(reconciliation)
    matched = int((reconciliation["Status"] == "Matched").sum()) if not reconciliation.empty else 0
    summary_items = [("Total Invoices", total), ("Matched", matched), ("Mismatches", total - matched), ("Match %", round(100 * matched / max(total, 1), 1)), ("GST Exposure", float(reconciliation.get("Exposure Amount", pd.Series(dtype=float)).sum()))]
    for index, (label, value) in enumerate(summary_items):
        start = index * 2
        sheet.merge_range(1, start, 1, start + 1, label, formats["summary_label"])
        sheet.merge_range(2, start, 2, start + 1, value, formats["summary_value"])
    header_row = 5
    for col, header in enumerate(headers):
        sheet.write(header_row, col, header, formats["header"])
    widths = [16, 18, 18, 24, 17, 20, 19, 22, 12, 34, 18, 15, 15, 15, 40]
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    for excel_row, (_, record) in enumerate(reconciliation.iterrows(), start=header_row + 2):
        row_index = excel_row - 1
        row_format = _row_format(record["Status"], record["Mismatch Type"], formats)
        document_type = record.get("Document Type", "Invoice")
        source_row = _source_row(source_standardized, document_type, record.get("GSTIN", ""), record.get("Invoice Number", ""))
        reference_row = _source_row(reference_standardized, document_type, record.get("GSTIN", ""), record.get("Invoice Number", ""))
        values = [document_type, record.get("GSTIN", ""), record.get("Invoice Number", ""), record.get("Customer", ""), source_row, None, reference_row, None, record["Status"], record["Mismatch Type"], record.get("Taxable Difference", 0), record.get("IGST Difference", 0), record.get("CGST Difference", 0), record.get("SGST Difference", 0), record.get("Remarks", "")]
        for col, value in enumerate(values):
            if col == 5 and source_row:
                sheet.write_formula(row_index, col, f'=HYPERLINK("#\'{source_sheet_name}\'!A{source_row}","Open Source Record")', row_format, "Open Source Record")
            elif col == 7 and reference_row:
                sheet.write_formula(row_index, col, f'=HYPERLINK("#\'{reference_sheet_name}\'!A{reference_row}","Open Reference Record")', row_format, "Open Reference Record")
            else:
                _write_value(sheet, row_index, col, value, row_format)
    sheet.autofilter(header_row, 0, max(header_row + 1, header_row + len(reconciliation)), len(headers) - 1)
    sheet.freeze_panes(header_row + 1, 0)
    workbook.close()
    return output.getvalue()


def professional_report_bytes(gstr1_original, books_original, gstr1_standardized, books_standardized, reconciliation) -> bytes:
    """Backward-compatible wrapper for the original GSTR-1 vs Books export."""
    return reconciliation_export_bytes("GSTR-1", "Books Data", gstr1_original, books_original, gstr1_standardized, books_standardized, reconciliation)
