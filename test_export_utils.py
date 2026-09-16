from io import BytesIO
import zipfile

import pandas as pd

from export_utils import professional_report_bytes, sample_template_bytes
from reconciliation import reconcile, standardize_columns


def _source() -> pd.DataFrame:
    return pd.DataFrame({
        "GSTIN": ["27ABCDE1234F1Z5"],
        "Invoice Number": ["A1"],
        "Invoice Date": ["2026-01-01"],
        "Customer": ["Aster Retail"],
        "Taxable Value": [1000],
        "IGST": [180],
        "CGST": [0],
        "SGST": [0],
    })


def test_export_accepts_numeric_and_duplicate_source_columns():
    gstr1 = _source()
    books = _source()
    books.insert(5, "Taxable Value Copy", 1000)
    books.columns = ["GSTIN", "Invoice Number", "Invoice Date", "Customer", "Taxable Value", "Taxable Value", "IGST", "CGST", "SGST"]
    reconciliation = reconcile({"GSTR-1": gstr1, "Tally": _source()})
    report = professional_report_bytes(gstr1, books, standardize_columns(gstr1), standardize_columns(_source()), reconciliation)
    assert zipfile.is_zipfile(BytesIO(report))


def test_sample_template_fallback_returns_an_excel_workbook():
    assert zipfile.is_zipfile(BytesIO(sample_template_bytes("GSTR-1")))
