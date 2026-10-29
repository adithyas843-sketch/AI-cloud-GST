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
        ["Document Type", "Customer GSTIN", "Invoice no", "Date", "Customer Name", "Item Taxable Value", "IGST", "CGST", "SGST"],
        [["Invoice", "27ABCDE1234F1Z5", "G1-1001", "05-Apr-2026", "Aster Retail", 100000, 18000, 0, 0], ["Credit Note", "29PQRSX6789L1Z2", "CN-1002", "08-Apr-2026", "Bluebird Foods", 50000, 0, 4500, 4500]],
    ),
    "Tally": (
        ["Document Type", "GSTIN", "Voucher Number", "Voucher Date", "Party Name", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Invoice", "27ABCDE1234F1Z5", "G1-1001", "05-Apr-2026", "Aster Retail", 100000, 18000, 0, 0], ["Credit Note", "29PQRSX6789L1Z2", "CN-1002", "08-Apr-2026", "Bluebird Foods", 50000, 0, 4500, 4500]],
    ),
    "Zoho": (
        ["Document Type", "GSTIN", "Invoice Number", "Invoice Date", "Supplier Name", "Taxable Value", "IGST", "CGST", "SGST", "Remarks"],
        [["Invoice", "27ABCDE1234F1Z5", "P-1001", "05-Apr-2026", "Aster Supplies", 100000, 18000, 0, 0, "Purchase invoice"], ["Debit Note", "29PQRSX6789L1Z2", "DN-1002", "08-Apr-2026", "Bluebird Foods", 50000, 0, 4500, 4500, "Debit note"], ["Journal Entry", "27ABCDE1234F1Z5", "JV-1003", "10-Apr-2026", "Input GST adjustment", 0, 0, 900, 900, "Journal entry"]],
    ),
    "GSTR-2B": (
        ["Document Type", "Supplier GSTIN", "Invoice Number", "Invoice Date", "Supplier Name", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Invoice", "27ABCDE1234F1Z5", "P-1001", "05-Apr-2026", "Aster Supplies", 100000, 18000, 0, 0], ["Debit Note", "29PQRSX6789L1Z2", "DN-1002", "08-Apr-2026", "Bluebird Foods", 50000, 0, 4500, 4500]],
    ),
    "GSTR-3B": (
        ["Month", "Taxable Value", "IGST", "CGST", "SGST"],
        [["Apr-2026", 150000, 18000, 4500, 4500], ["May-2026", 75000, 13500, 0, 0]],
    ),
}


def sample_template_bytes(source_name: str) -> bytes:
    """Create a deploy-safe sample template when a bundled file is unavailable."""
    # Accept both UI labels and the legacy source keys used by the existing backend.
    aliases = {"Sales Register": "Tally", "Purchase Register": "Zoho", "GSTR 1": "GSTR-1", "GSTR 2B": "GSTR-2B", "GSTR 3B": "GSTR-3B"}
    source_name = aliases.get(source_name, source_name)
    if source_name not in TEMPLATE_DEFINITIONS:
        raise ValueError(f"Unknown template type: {source_name}")
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
