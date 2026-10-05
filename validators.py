"""Data quality checks for GST sales-register exports."""
from __future__ import annotations
import re
from datetime import date
import pandas as pd

GSTIN_RE = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")
STATE_CODES = {f"{n:02d}" for n in range(1, 38)} | {"97", "99"}

def valid_gstin(value: object) -> bool:
    """Return whether value has a structurally valid Indian GSTIN."""
    if value is None or pd.isna(value):
        return False
    gstin = str(value).strip().upper()
    return bool(GSTIN_RE.fullmatch(gstin) and gstin[:2] in STATE_CODES)

def validation_report(frame: pd.DataFrame, source: str = "Books") -> pd.DataFrame:
    """Return one audit row per detected data-quality exception."""
    if frame is None or frame.empty:
        return pd.DataFrame(columns=["Source", "Row", "Invoice Number", "GSTIN", "Issue", "Severity", "Details"])
    rows = []
    mandatory = ["Document Type", "Invoice Number", "Invoice Date", "GSTIN"]
    amount_cols = ["Taxable Value", "IGST", "CGST", "SGST"]
    for idx, r in frame.iterrows():
        inv, gstin = r.get("Invoice Number"), r.get("GSTIN")
        def add(issue, severity, details):
            rows.append({"Source": source, "Row": int(idx) + 2, "Invoice Number": inv,
                         "GSTIN": gstin, "Issue": issue, "Severity": severity, "Details": details})
        for col in mandatory:
            if pd.isna(r.get(col)) or str(r.get(col)).strip() in {"", "nan", "None"}:
                add(f"Missing {col}", "High", f"{col} is mandatory.")
        if not pd.isna(gstin) and str(gstin).strip() and not valid_gstin(gstin):
            add("Invalid GSTIN", "High", "GSTIN must be 15 characters with a valid state code.")
        invoice_date = pd.to_datetime(r.get("Invoice Date"), errors="coerce")
        if not pd.isna(invoice_date) and invoice_date.date() > date.today():
            add("Future invoice date", "Medium", "Invoice date cannot be after today.")
        for col in amount_cols:
            value = pd.to_numeric(r.get(col), errors="coerce")
            if not pd.isna(value) and value < 0:
                add(f"Negative {col}", "Medium", f"{col} is negative.")
    dupe_cols = [c for c in ["Document Type", "GSTIN", "Invoice Number"] if c in frame]
    if len(dupe_cols) == 3:
        for idx, r in frame[frame.duplicated(dupe_cols, keep=False)].iterrows():
            rows.append({"Source": source, "Row": int(idx)+2, "Invoice Number": r.get("Invoice Number"),
                         "GSTIN": r.get("GSTIN"), "Issue": "Duplicate invoice", "Severity": "High",
                         "Details": "Document type, GSTIN and invoice number occur more than once."})
    return pd.DataFrame(rows, columns=["Source", "Row", "Invoice Number", "GSTIN", "Issue", "Severity", "Details"])
