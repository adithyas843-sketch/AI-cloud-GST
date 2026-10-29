"""Column mapping, invoice reconciliation and GST risk calculations."""
from __future__ import annotations
import re
import numpy as np
import pandas as pd

ALIASES = {
 "Document Type": ["document type", "document type name", "transaction type", "voucher type", "entry type"],
 "GSTIN": ["gstin", "customer gstin", "supplier gstin", "vendor gstin", "gst number", "gst registration no", "party gstin", "gst identification number"],
 "Invoice Number": ["invoice number", "invoice no", "inv no", "voucher number", "document number", "reference number"],
 "Invoice Date": ["invoice date", "date", "voucher date", "document date"],
 "Taxable Value": ["taxable value", "tax amount", "assessable value", "taxable amount"],
 "IGST": ["igst", "integrated tax"], "CGST": ["cgst", "central tax"], "SGST": ["sgst", "state tax"],
 "Customer": ["customer", "customer name", "party name", "buyer name", "ledger name", "vendor name", "supplier name"]}
REQUIRED = ["Document Type", "GSTIN", "Invoice Number", "Invoice Date", "Taxable Value", "IGST", "CGST", "SGST"]

def _norm(x): return re.sub(r"[^a-z0-9]", "", str(x).lower())
def standardize_columns(raw: pd.DataFrame) -> pd.DataFrame:
    """Map common source headings to the application's canonical schema."""
    lookup = {_norm(c): c for c in raw.columns}; renames = {}
    for canonical, aliases in ALIASES.items():
        found = next((lookup.get(_norm(a)) for a in aliases if _norm(a) in lookup), None)
        if found: renames[found] = canonical
    out = raw.rename(columns=renames).copy()
    for col in REQUIRED:
        if col not in out: out[col] = np.nan
    if "Customer" not in out: out["Customer"] = "Unspecified customer"
    out["Document Type"] = out["Document Type"].fillna("Invoice").astype(str).str.strip().str.title()
    out.loc[out["Document Type"].eq(""), "Document Type"] = "Invoice"
    out["GSTIN"] = out["GSTIN"].fillna("").astype(str).str.strip().str.upper()
    out["Invoice Number"] = out["Invoice Number"].fillna("").astype(str).str.strip().str.upper().str.replace(r"\.0$", "", regex=True)
    out["Invoice Date"] = pd.to_datetime(out["Invoice Date"], errors="coerce")
    for col in ["Taxable Value", "IGST", "CGST", "SGST"]: out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    return out

def read_upload(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    raw = pd.read_csv(uploaded) if name.endswith(".csv") else pd.read_excel(uploaded)
    return standardize_columns(raw)

def _key(df): return df["Document Type"].astype(str) + "|" + df["GSTIN"].astype(str) + "|" + df["Invoice Number"].astype(str)

def _missing_source_reason(base: pd.Series, candidate_source: pd.DataFrame, source_name: str) -> str:
    """Classify a missing composite key using invoice/date candidates in a source."""
    same_document_invoice = candidate_source[
        (candidate_source["Document Type"] == base["Document Type"])
        & (candidate_source["Invoice Number"] == base["Invoice Number"])
    ]
    if not same_document_invoice.empty and (same_document_invoice["GSTIN"] != base["GSTIN"]).any():
        return f"GSTIN mismatch with {source_name}"
    same_gstin_invoice = candidate_source[
        (candidate_source["GSTIN"] == base["GSTIN"])
        & (candidate_source["Invoice Number"] == base["Invoice Number"])
    ]
    if not same_gstin_invoice.empty and (same_gstin_invoice["Document Type"] != base["Document Type"]).any():
        return f"Document type mismatch with {source_name}"
    invoice_date = base["Invoice Date"]
    if not pd.isna(invoice_date):
        same_gstin_date = candidate_source[
            (candidate_source["Document Type"] == base["Document Type"])
            & (candidate_source["GSTIN"] == base["GSTIN"])
            & (candidate_source["Invoice Date"] == invoice_date)
        ]
        if len(same_gstin_date) == 1:
            return f"Invoice number mismatch with {source_name}"
        if len(same_gstin_date) > 1:
            return f"Multiple invoice mappings in {source_name}"
    return f"Missing in {source_name}"

def reconcile(sources: dict[str, pd.DataFrame], tolerance: float = 1.0) -> pd.DataFrame:
    """Outer-join all supplied registers, producing a source-aware exception ledger."""
    active = {name: standardize_columns(df) for name, df in sources.items() if df is not None and not df.empty}
    if not active:
        return pd.DataFrame(columns=["Document Type", "GSTIN", "Invoice Number", "Invoice Date", "Customer", "Status", "Mismatch Type", "Taxable Difference", "IGST Difference", "CGST Difference", "SGST Difference", "Exposure Amount", "Remarks"])
    all_keys = pd.Index([])
    keyed = {}
    for name, df in active.items():
        copy = df.copy(); copy["_key"] = _key(copy)
        # deterministic aggregation also exposes multi-mapping as a mismatch
        grouped = copy.groupby("_key", dropna=False).agg({"Document Type":"first", "GSTIN":"first", "Invoice Number":"first", "Invoice Date":"first", "Customer":"first", "Taxable Value":"sum", "IGST":"sum", "CGST":"sum", "SGST":"sum"})
        grouped["_count"] = copy.groupby("_key").size(); keyed[name] = grouped
        all_keys = all_keys.union(grouped.index)
    records=[]
    reference = "GSTR-1" if "GSTR-1" in keyed else next(iter(keyed))
    for k in all_keys:
        rows={n: d.loc[k] if k in d.index else None for n,d in keyed.items()}
        base=rows.get(reference)
        if base is None:
            base=next(v for v in rows.values() if v is not None)
        reasons=[]
        for name, row in rows.items():
            if row is None: reasons.append(_missing_source_reason(base, keyed[name], name))
            elif row["_count"] > 1: reasons.append(f"Duplicate invoices in {name}")
        comparisons=[r for r in rows.values() if r is not None]
        diffs={c: max([r[c] for r in comparisons])-min([r[c] for r in comparisons]) for c in ["Taxable Value","IGST","CGST","SGST"]}
        for c, diff in diffs.items():
