from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from reconciliation import read_upload, reconcile, summary, gstr3b_control
from validators import validation_report
from ai_insights import generate_insights
from export_utils import professional_report_bytes, sample_template_bytes


# ---------------------------------------------------------------------------
# Page setup and minimalist styling
# ---------------------------------------------------------------------------
st.set_page_config(page_title="GST Copilot", page_icon="🧾", layout="wide")

st.markdown(
    """
    <style>
    :root {
        --page: #f7f6f2;
        --ink: #191919;
        --muted: #6f6d67;
        --line: #e5e2da;
        --green: #237a50;
        --amber: #b87512;
        --red: #bf332b;
    }
    .stApp { background: var(--page); color: var(--ink); }
    [data-testid="stHeader"] { background: rgba(247, 246, 242, 0.94); }
    [data-testid="stSidebar"] {
        background: #f0eee8;
        border-right: 1px solid var(--line);
    }
    [data-testid="stSidebar"] > div { padding-top: 1.4rem; }
    .block-container { max-width: 1400px; padding-top: 2rem; padding-bottom: 3rem; }
    h1, h2, h3 { color: var(--ink); letter-spacing: -0.035em; }
    h1 { font-family: Georgia, 'Times New Roman', serif; font-weight: 500; }
    h2, h3 { font-weight: 550; }
    p, label, [data-testid="stCaptionContainer"] { color: var(--muted); }
    [data-testid="stMetric"] {
        background: transparent;
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 16px 18px;
        box-shadow: none;
    }
    [data-testid="stMetricLabel"] { color: var(--muted) !important; font-size: 0.82rem; }
    [data-testid="stMetricValue"] { color: var(--ink) !important; font-size: 1.8rem; }
    [data-testid="stFileUploader"] {
        border: 1px solid var(--line);
        border-radius: 8px;
        background: rgba(255,255,255,0.45);
        padding: 10px;
    }
    .stButton > button, .stDownloadButton > button {
        border-radius: 22px;
        border: 1px solid #c9c5bc;
        box-shadow: none;
        min-height: 2.55rem;
    }
    .stButton > button[kind="primary"] {
        background: var(--ink);
        color: #fff;
        border-color: var(--ink);
    }
    div[data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 8px; }
    hr { border-color: var(--line); }
    .eyebrow {
        text-transform: uppercase;
        letter-spacing: 0.12em;
        font-size: 0.72rem;
        color: var(--muted);
        margin-bottom: 0.3rem;
    }
    .subtle-box {
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 15px 17px;
        background: rgba(255,255,255,0.35);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

TEMPLATE_DIR = Path(__file__).parent / "templates"
TEMPLATES = {
    "GSTR-1": "sample_gstr1.xlsx",
    "Tally": "sample_tally.xlsx",
    "Zoho": "sample_zoho.xlsx",
    "GSTR-3B": "sample_gstr3b.xlsx",
}


def read_original_upload(uploaded) -> pd.DataFrame:
    """Read a pristine copy for the source-data sheets in the Excel export."""
    content = BytesIO(uploaded.getvalue())
    return (
        pd.read_csv(content)
        if uploaded.name.lower().endswith(".csv")
        else pd.read_excel(content)
    )


def get_template_bytes(key: str) -> bytes:
    """Return the checked-in template, or generate one when it is not present."""
    template_path = TEMPLATE_DIR / TEMPLATES[key]
    return template_path.read_bytes() if template_path.exists() else sample_template_bytes(key)


def warning_count(recon_df: pd.DataFrame, validation_df: pd.DataFrame) -> int:
    """Count warning/review records where the existing outputs expose a status."""
    count = 0
    if not recon_df.empty and "Status" in recon_df.columns:
        status = recon_df["Status"].astype(str).str.strip().str.lower()
        count = int(status.isin({"warning", "review", "needs review", "needs_review"}).sum())
    if count == 0 and not validation_df.empty and "Status" in validation_df.columns:
        status = validation_df["Status"].astype(str).str.strip().str.lower()
        count = int(status.isin({"warning", "review", "needs review", "needs_review"}).sum())
    return count


# ---------------------------------------------------------------------------
# Sidebar navigation
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown(
        "<div style='font-family:Georgia,serif;font-size:1.45rem;color:#191919'>GST Copilot</div>"
        "<div style='font-size:.78rem;color:#77746d;margin-top:2px'>Reconciliation workspace</div>",
        unsafe_allow_html=True,
    )
    st.divider()
    page = st.radio(
        "Navigation",
        ["Upload & Reconcile", "Dashboard", "Exceptions", "Insights", "Exports", "Data", "Help"],
        label_visibility="collapsed",
        key="main_navigation",
    )
    st.divider()
    st.caption("Review all reconciliation findings against supporting records before filing.")


# ---------------------------------------------------------------------------
# Main header
# ---------------------------------------------------------------------------
header_left, header_right = st.columns([4, 1.5], vertical_alignment="center")
with header_left:
    st.markdown("<div class='eyebrow'>GST COMPLIANCE WORKSPACE</div>", unsafe_allow_html=True)
    st.title("Reconcile")
    st.caption("Invoice-level matching, exceptions and audit-ready reporting.")
with header_right:
    st.download_button(
        "Download sample templates",
        data=get_template_bytes("GSTR-1"),
        file_name=TEMPLATES["GSTR-1"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="header_sample_template",
    )

st.divider()


# ---------------------------------------------------------------------------
# Upload & Reconcile
# Existing backend source keys are deliberately retained so the current
# reconciliation and export functions continue to work.
# ---------------------------------------------------------------------------
if page == "Upload & Reconcile":
    st.subheader("Source files")
    st.caption("Upload the source files needed for the current reconciliation workflow.")

    source_columns = st.columns(2, gap="large")
    uploads = {}

    source_config = [
        ("GSTR-1", "GSTR-1", "GSTR-1 export"),
        ("Sales Register", "Tally", "Sales register (currently mapped to the Tally source key)"),
        ("Zoho Books Sales Register", "Zoho", "Optional alternate sales-books source"),
        ("GSTR-3B", "GSTR-3B", "GSTR-3B summary/control data"),
    ]

    for index, (label, key, help_text) in enumerate(source_config):
        with source_columns[index % 2]:
            st.markdown(f"**{label}**")
            st.caption(help_text)
            uploads[key] = st.file_uploader(
                f"Upload {label}",
                type=["xlsx", "xls", "csv"],
                key=f"upload_{key}",
                label_visibility="collapsed",
            )
            st.download_button(
                "Download sample template",
                data=get_template_bytes(key),
                file_name=TEMPLATES[key],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"template_{key}",
            )
            st.markdown("")

    st.divider()
    run_col, note_col = st.columns([1, 2], vertical_alignment="center")
    with run_col:
        run = st.button("Run reconciliation", type="primary", use_container_width=True)
    with note_col:
        st.caption("Metrics and reports appear after a successful run. Uploaded files alone will not show dashboard metrics.")

    if run:
        invoice_sources = {}
        gstr3b_raw = None
        source_originals = {}
        validation_frames = []
        progress = st.progress(0, text="Reading source files…")

        for name, uploaded in uploads.items():
            if not uploaded:
                continue
            if name == "GSTR-3B":
                try:
                    gstr3b_raw = (
                        pd.read_csv(uploaded)
                        if uploaded.name.lower().endswith(".csv")
                        else pd.read_excel(uploaded)
                    )
                except Exception as exc:
                    st.error(f"Could not read GSTR-3B: {exc}")
                continue

            try:
                original = read_original_upload(uploaded)
                standardized = read_upload(uploaded)
                invoice_sources[name] = standardized
                source_originals[name] = original
                validation_frames.append(validation_report(standardized, name))
            except Exception as exc:
                st.error(f"Could not read {name}: {exc}")

        progress.progress(55, text="Reconciling invoices…")
        if len(invoice_sources) < 2:
            progress.empty()
            st.error("Upload at least two invoice-level sources (GSTR-1 and a books source) before running reconciliation.")
        else:
            try:
                result = reconcile(invoice_sources)
                validation_result = (
                    pd.concat(validation_frames, ignore_index=True)
                    if validation_frames
                    else pd.DataFrame()
                )
                insights_result = generate_insights(result)
                control_result = gstr3b_control(invoice_sources.get("GSTR-1"), gstr3b_raw)

                # Store all results only after successful processing.
                st.session_state.recon = result
                st.session_state.validations = validation_result
                st.session_state.insights = insights_result
                st.session_state.gstr3b = control_result
                st.session_state.source_originals = source_originals
                st.session_state.invoice_sources = invoice_sources
                st.session_state.reconciliation_completed = True

                progress.progress(100, text="Analysis complete.")
                st.success("Reconciliation completed successfully.")
                st.rerun()
            except Exception as exc:
                progress.empty()
                st.error(f"Reconciliation could not be completed: {exc}")

    if not st.session_state.get("reconciliation_completed", False):
        st.markdown(
            "<div class='subtle-box'><strong>Ready when you are</strong><br>"
            "<span style='color:#77746d'>Upload the required source files, then run reconciliation. "
            "Your metrics will appear after a successful run.</span></div>",
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# Metrics are intentionally hidden until a successful reconciliation exists.
# ---------------------------------------------------------------------------
has_results = (
    st.session_state.get("reconciliation_completed", False)
    and "recon" in st.session_state
    and isinstance(st.session_state.get("recon"), pd.DataFrame)
)

if has_results:
    recon = st.session_state.recon
    validations = st.session_state.get("validations", pd.DataFrame())
    insights = st.session_state.get("insights", pd.DataFrame())
    gstr3b = st.session_state.get("gstr3b", pd.DataFrame())
    s = summary(recon, validations)

    if page in {"Upload & Reconcile", "Dashboard", "Exceptions", "Insights", "Exports"}:
        st.markdown("<div class='eyebrow'>RECONCILIATION RESULTS</div>", unsafe_allow_html=True)
        st.subheader("Summary")
        metric_cols = st.columns(4)
        values = [
            s.get("total", 0),
            s.get("matched", 0),
            warning_count(recon, validations),
            s.get("mismatches", 0),
        ]
        labels = ["Total invoices", "Matched", "Warnings", "Mismatches"]
        for col, label, value in zip(metric_cols, labels, values):
            col.metric(label, f"{value:,}" if isinstance(value, (int, float)) else value)
        st.caption(f"Risk level: **{s.get('risk', 'N/A')}** — {s.get('score_reason', '')}")
        st.divider()

    if page == "Dashboard":
        st.subheader("Reconciliation overview")
        mismatches = recon[recon["Status"].astype(str).str.lower() == "mismatch"] if "Status" in recon.columns else pd.DataFrame()
        left, right = st.columns(2, gap="large")
        with left:
            if not mismatches.empty and "Mismatch Type" in mismatches.columns:
                st.plotly_chart(
                    px.pie(mismatches, names="Mismatch Type", title="Mismatch type distribution"),
                    use_container_width=True,
                )
            else:
                st.info("No mismatch-type chart is available for these results.")
        with right:
            if not mismatches.empty and {"Customer", "Exposure Amount"}.issubset(mismatches.columns):
                risk = (
                    mismatches.groupby("Customer", as_index=False)["Exposure Amount"]
                    .sum()
                    .nlargest(10, "Exposure Amount")
                )
                if not risk.empty:
                    st.plotly_chart(
                        px.bar(risk, x="Customer", y="Exposure Amount", title="Top risk customers"),
                        use_container_width=True,
                    )
            else:
                st.info("No customer exposure chart is available for these results.")
        if {"Invoice Date", "Exposure Amount"}.issubset(recon.columns):
            monthly = recon.copy()
            monthly["Month"] = pd.to_datetime(monthly["Invoice Date"], errors="coerce").dt.to_period("M").astype(str)
            monthly = monthly.dropna(subset=["Month"]).groupby("Month", as_index=False)["Exposure Amount"].sum()
            if not monthly.empty:
                st.plotly_chart(
                    px.line(monthly, x="Month", y="Exposure Amount", markers=True, title="Monthly exposure trend"),
                    use_container_width=True,
                )

    elif page == "Exceptions":
        st.subheader("Exceptions and validation")
        st.markdown("**Invoice reconciliation exceptions**")
        if "Status" in recon.columns:
            exceptions = recon[recon["Status"].astype(str).str.lower().isin(["mismatch", "warning", "review", "needs review", "needs_review"])]
        else:
            exceptions = recon
        st.dataframe(exceptions, use_container_width=True, hide_index=True)
        st.markdown("**Validation report**")
        st.dataframe(validations, use_container_width=True, hide_index=True)

    elif page == "Insights":
        st.subheader("Exception insights")
        st.dataframe(insights, use_container_width=True, hide_index=True)

    elif page == "Exports":
        st.subheader("Exports")
        originals = st.session_state.get("source_originals", {})
        standardized = st.session_state.get("invoice_sources", {})
        books_source = "Tally" if "Tally" in standardized else "Zoho" if "Zoho" in standardized else None
        if "GSTR-1" not in standardized or books_source is None:
            st.info("Upload GSTR-1 and a supported books source to create the professional reconciliation workbook.")
        else:
            try:
                workbook = professional_report_bytes(
                    originals["GSTR-1"],
                    originals[books_source],
                    standardized["GSTR-1"],
                    standardized[books_source],
                    recon,
                )
                st.download_button(
                    "Download professional reconciliation workbook",
                    workbook,
                    "gst_reconciliation_workbook.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"The export workbook could not be prepared: {exc}")

        if not gstr3b.empty:
            st.markdown("**GSTR-1 vs GSTR-3B control check**")
            st.dataframe(gstr3b, use_container_width=True, hide_index=True)

    elif page == "Data":
        st.subheader("Data management")
        st.caption("Clear results or reset the current session.")
        clear_col, results_col, reset_col = st.columns(3)
        with clear_col:
            if st.button("Clear uploaded files", use_container_width=True):
                for key in ["upload_GSTR-1", "upload_Tally", "upload_Zoho", "upload_GSTR-3B"]:
                    st.session_state.pop(key, None)
                st.rerun()
        with results_col:
            if st.button("Clear reconciliation results", use_container_width=True):
                for key in [
                    "recon", "validations", "insights", "gstr3b",
                    "source_originals", "invoice_sources", "reconciliation_completed",
                ]:
                    st.session_state.pop(key, None)
                st.rerun()
        with reset_col:
            if st.button("Reset application", type="primary", use_container_width=True):
                st.session_state.clear()
                st.rerun()

    elif page == "Help":
        st.subheader("Help")
        st.markdown(
            "- Upload GSTR-1 and a supported books register to run invoice-level reconciliation.\n"
            "- Upload GSTR-3B to include the aggregate control check.\n"
            "- Review exceptions and validation findings before relying on the results.\n"
            "- Export a workbook from the Exports section."
        )

elif page in {"Dashboard", "Exceptions", "Insights", "Exports"}:
    st.info("No completed reconciliation is available yet. Open **Upload & Reconcile**, upload your files, and run reconciliation to see results.")
elif page == "Data":
    st.subheader("Data management")
    st.caption("You can clear results or reset the current session after running a reconciliation.")
    if st.button("Reset application", type="primary"):
        st.session_state.clear()
        st.rerun()
elif page == "Help":
    st.subheader("Help")
    st.markdown(
        "- Start in **Upload & Reconcile**.\n"
        "- Upload at least two supported invoice-level sources.\n"
        "- Click **Run reconciliation**.\n"
        "- Review the dashboard, exceptions, insights, and exports after a successful run."
    )
