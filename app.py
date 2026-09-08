from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from reconciliation import (
    CANONICAL_FIELDS,
    generate_management_summary,
    infer_column_mapping,
    read_excel_file,
    reconcile_data,
    standardize_dataframe,
)


st.set_page_config(
    page_title="GST Compliance Copilot",
    page_icon="GST",
    layout="wide",
)

st.markdown(
    """
    <style>
    .stApp {
        background: #f5f7f8;
    }

    .hero {
        background: linear-gradient(120deg, #063b35, #0d6b5d);
        color: white;
        padding: 2rem 2.2rem;
        border-radius: 18px;
        margin-bottom: 1.5rem;
    }

    .hero h1 {
        font-size: 2.4rem;
        margin: 0;
    }

    .hero p {
        color: #d8f3ed;
        margin: 0.6rem 0 0;
        font-size: 1.05rem;
    }

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #dce6e3;
        border-radius: 14px;
        padding: 1rem;
    }

    .section-label {
        color: #0d6b5d;
        font-weight: 700;
        letter-spacing: .08em;
        text-transform: uppercase;
        font-size: .75rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>GST Compliance Copilot</h1>
        <p>Invoice-level reconciliation between GSTR-1 and Tally Sales Register.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Control Panel")
    tolerance = st.number_input(
        "Amount tolerance (Rs)",
        min_value=0.0,
        max_value=1000.0,
        value=1.0,
        step=0.5,
        help="Differences within this amount are treated as matched.",
    )

    st.caption(
        "Upload both Excel files to enable reconciliation. "
        "Column names are auto-mapped using common GSTR-1 and Tally aliases."
    )

st.markdown('<div class="section-label">Source Files</div>', unsafe_allow_html=True)
upload_col_1, upload_col_2 = st.columns(2)

with upload_col_1:
    gstr1_file = st.file_uploader(
        "Upload GSTR-1 Excel file",
        type=["xlsx", "xls"],
        key="gstr1",
    )

with upload_col_2:
    tally_file = st.file_uploader(
        "Upload Tally Sales Register Excel file",
        type=["xlsx", "xls"],
        key="tally",
    )

if not gstr1_file or not tally_file:
    st.info(
        "Upload one GSTR-1 file and one Tally Sales Register file to begin."
    )
    st.stop()

try:
    raw_gstr1 = read_excel_file(gstr1_file)
    raw_tally = read_excel_file(tally_file)
except Exception as error:
    st.error(f"Could not read the Excel files: {error}")
    st.stop()

if raw_gstr1.empty or raw_tally.empty:
    st.error("At least one uploaded workbook contains no usable rows.")
    st.stop()

gstr1_mapping = infer_column_mapping(list(raw_gstr1.columns))
tally_mapping = infer_column_mapping(list(raw_tally.columns))

st.markdown('<div class="section-label">Auto-Mapped Columns</div>', unsafe_allow_html=True)

mapping_rows = []
for field in CANONICAL_FIELDS:
    mapping_rows.append(
        {
            "Canonical field": field.replace("_", " ").title(),
            "GSTR-1 column": gstr1_mapping.get(field) or "Not found",
            "Tally column": tally_mapping.get(field) or "Not found",
        }
    )

st.dataframe(
    pd.DataFrame(mapping_rows),
    use_container_width=True,
    hide_index=True,
)

missing_gstr1 = [
    field
    for field in ["invoice_number", "taxable_value"]
    if not gstr1_mapping.get(field)
]
missing_tally = [
    field
    for field in ["invoice_number", "taxable_value"]
    if not tally_mapping.get(field)
]

if missing_gstr1 or missing_tally:
    if missing_gstr1:
        st.warning(
            "GSTR-1 required columns not detected: "
            + ", ".join(missing_gstr1)
        )
    if missing_tally:
        st.warning(
            "Tally required columns not detected: "
            + ", ".join(missing_tally)
        )

    st.stop()

gstr1 = standardize_dataframe(raw_gstr1, gstr1_mapping)
tally = standardize_dataframe(raw_tally, tally_mapping)

results, summary = reconcile_data(
    gstr1,
    tally,
    tolerance=tolerance,
)

st.markdown('<div class="section-label">GST Health Dashboard</div>', unsafe_allow_html=True)

metric_cols = st.columns(5)

metric_cols[0].metric(
    "Total Invoices",
    f"{summary['total_invoices']:,}",
)

metric_cols[1].metric(
    "Matched Invoices",
    f"{summary['matched_invoices']:,}",
)

metric_cols[2].metric(
    "Mismatches",
    f"{summary['mismatch_invoices']:,}",
)

metric_cols[3].metric(
    "GST Exposure",
    f"Rs {summary['gst_exposure']:,.2f}",
)

metric_cols[4].metric(
    "GST Health Score",
    f"{summary['health_score']:.1f}/100",
)

st.progress(
    min(max(summary["health_score"] / 100, 0.0), 1.0),
    text=f"Reconciliation health: {summary['health_score']:.1f}%",
)

left_col, right_col = st.columns([1, 2])

with left_col:
    st.subheader("Exception Breakdown")

    issue_data = pd.DataFrame(
        [
            {"Issue": label, "Count": count}
            for label, count in summary["issue_counts"].items()
            if count > 0
        ]
    )

    if issue_data.empty:
        st.success("No reconciliation exceptions detected.")
    else:
        st.bar_chart(
            issue_data.set_index("Issue"),
            horizontal=True,
        )

with right_col:
    st.subheader("Management Summary")
    st.text_area(
        "AI-generated management summary",
        value=generate_management_summary(summary),
        height=260,
        label_visibility="collapsed",
    )

st.markdown('<div class="section-label">Invoice-Level Reconciliation</div>', unsafe_allow_html=True)

status_filter = st.multiselect(
    "Filter by status",
    options=["Matched", "Mismatch", "Review"],
    default=["Mismatch", "Review", "Matched"],
)

filtered_results = results[
    results["status"].isin(status_filter)
].copy()

display_columns = [
    "invoice_number",
    "status",
    "issues",
    "gstr1_gstin",
    "tally_gstin",
    "gstr1_taxable_value",
    "tally_taxable_value",
    "gstr1_igst",
    "tally_igst",
    "gstr1_cgst",
    "tally_cgst",
    "gstr1_sgst",
    "tally_sgst",
    "gst_exposure",
]

st.dataframe(
    filtered_results[display_columns],
    use_container_width=True,
    hide_index=True,
    column_config={
        "gstr1_taxable_value": st.column_config.NumberColumn(
            "GSTR-1 Taxable Value",
            format="Rs %.2f",
        ),
        "tally_taxable_value": st.column_config.NumberColumn(
            "Tally Taxable Value",
            format="Rs %.2f",
        ),
        "gstr1_igst": st.column_config.NumberColumn(
            "GSTR-1 IGST",
            format="Rs %.2f",
        ),
        "tally_igst": st.column_config.NumberColumn(
            "Tally IGST",
            format="Rs %.2f",
        ),
        "gstr1_cgst": st.column_config.NumberColumn(
            "GSTR-1 CGST",
            format="Rs %.2f",
        ),
        "tally_cgst": st.column_config.NumberColumn(
            "Tally CGST",
            format="Rs %.2f",
        ),
        "gstr1_sgst": st.column_config.NumberColumn(
            "GSTR-1 SGST",
            format="Rs %.2f",
        ),
        "tally_sgst": st.column_config.NumberColumn(
            "Tally SGST",
            format="Rs %.2f",
        ),
        "gst_exposure": st.column_config.NumberColumn(
            "GST Exposure",
            format="Rs %.2f",
        ),
    },
)

st.subheader("AI Explanation")

if filtered_results.empty:
    st.info("No invoices match the selected filters.")
else:
    selected_invoice = st.selectbox(
        "Select an invoice",
        options=filtered_results["invoice_number"].tolist(),
    )

    selected_row = filtered_results[
        filtered_results["invoice_number"] == selected_invoice
    ].iloc[0]

    if selected_row["status"] == "Matched":
        st.success(
            f"{selected_invoice} matched successfully across both files."
        )
    else:
        st.warning(selected_row["ai_explanation"])

st.subheader("Export")

export_buffer = io.BytesIO()
with pd.ExcelWriter(export_buffer, engine="openpyxl") as writer:
    results.drop(columns=["issue_codes"], errors="ignore").to_excel(
        writer,
        index=False,
        sheet_name="Reconciliation",
    )

    pd.DataFrame(
        [
            {
                "Metric": "Total Invoices",
                "Value": summary["total_invoices"],
            },
            {
                "Metric": "Matched Invoices",
                "Value": summary["matched_invoices"],
            },
            {
                "Metric": "Mismatches",
                "Value": summary["mismatch_invoices"],
            },
            {
                "Metric": "GST Exposure",
                "Value": summary["gst_exposure"],
            },
            {
                "Metric": "GST Health Score",
                "Value": summary["health_score"],
            },
        ]
    ).to_excel(
        writer,
        index=False,
        sheet_name="Summary",
    )

st.download_button(
    "Download Reconciliation Excel",
    data=export_buffer.getvalue(),
    file_name="gst_reconciliation_output.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)

st.download_button(
    "Download Management Summary",
    data=generate_management_summary(summary),
    file_name="gst_management_summary.txt",
    mime="text/plain",
)
