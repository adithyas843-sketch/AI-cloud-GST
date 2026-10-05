"""Explain reconciliation exceptions using deterministic, audit-friendly language."""
from __future__ import annotations
import pandas as pd

def generate_insights(recon: pd.DataFrame) -> pd.DataFrame:
    columns=["Document Type", "Invoice Number", "GSTIN", "Exception", "Likely Cause", "Recommended Action", "Exposure Amount"]
    if recon.empty or "Status" not in recon:
        return pd.DataFrame(columns=columns)
    rows=[]
    for _, r in recon[recon["Status"]=="Mismatch"].iterrows():
        issue=r["Mismatch Type"]
        if "Missing in GSTR-1" in issue: cause, action="Invoice may not have been included during return filing.","Review sales register and amend GSTR-1 where required."
        elif "GSTIN" in issue: cause, action="Possible customer master or manual-entry error.","Verify customer GSTIN against invoice and master data."
        elif "date" in issue.lower(): cause, action="Posting or invoice-date timing differs across sources.","Compare source documents and reporting period."
        elif "Duplicate" in issue: cause, action="The same invoice appears more than once in a source register.","Confirm whether duplicate entries require reversal."
        else: cause, action="Invoice may have been modified after a return was prepared.","Compare source documents and correct the relevant register or return."
        rows.append({"Document Type":r.get("Document Type", "Invoice"),"Invoice Number":r["Invoice Number"],"GSTIN":r["GSTIN"],"Exception":issue,"Likely Cause":cause,"Recommended Action":action,"Exposure Amount":r["Exposure Amount"]})
    return pd.DataFrame(rows, columns=columns)
