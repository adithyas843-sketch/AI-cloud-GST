# AI GST Compliance Copilot

A Streamlit MVP for finance teams and Chartered Accountants to reconcile GSTR-1 against Tally and Zoho sales registers. It standardizes inconsistent exports, validates invoice data, identifies exceptions, calculates tax exposure, and produces audit-ready Excel reports.

## Features

- XLSX, XLS and CSV upload support for GSTR-1, Tally, Zoho Books and GSTR-3B exports
- Automatic mapping of common GSTIN, invoice, date and tax column aliases
- GSTIN-format, duplicate, missing-field, negative-value and future-date validation
- Invoice matching across sources using GSTIN + invoice number, with date and tax comparisons
- Exposure, health score, risk-customer, mismatch and monthly-trend analysis
- Deterministic AI-style explanations with recommended actions
- Excel exports for reconciliation, exceptions, exposure and management summary

## Installation

Requires Python 3.10 or later.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

On macOS/Linux, activate the virtual environment with `source .venv/bin/activate`.

## Demo data

Generate the four demonstration workbooks in the current directory:

```bash
python sample_data_generator.py
```

Upload `sample_gstr1.xlsx`, `sample_tally.xlsx`, and `sample_zoho.xlsx` in the sidebar, then click **Run Reconciliation**. The data includes duplicate invoices, a missing invoice, a taxable-value difference, GSTIN mismatch and date mismatch.

## Tests

```bash
pytest
```

## Deploy to Streamlit Community Cloud

1. Push this repository to GitHub.
2. In Streamlit Community Cloud, select **Create app**, then choose the repository, branch, and `app.py` entry point.
3. Set the Python version to 3.10+ and deploy. No database, secret, or external service is required.

## Important

This tool provides reconciliation assistance; it does not replace review by a qualified tax professional or statutory return filing controls.
