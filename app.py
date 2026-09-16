from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from ai_insights import generate_insights
from export_utils import reconciliation_export_bytes, sample_template_bytes
from reconciliation import gstr3b_reconciliation, read_upload, reconcile, summary
from storage import clear_reconciliation_results, clear_uploaded_files, load_state, reset_application, save_state, save_uploaded_file
from validators import validation_report


st.set_page_config(page_title="AI GST Compliance Copilot", page_icon="🧾", layout="wide")
st.markdown('<style>[data-testid="stMetric"]{background-color:#1e293b;border:1px solid #334155;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.25)}[data-testid="stMetricLabel"]{color:white!important;font-weight:600}[data-testid="stMetricValue"]{color:white!important;font-weight:700}.block-container{padding-top:1.5rem}</style>', unsafe_allow_html=True)

TEMPLATE_DIR = Path(__file__).parent / "templates"
UPLOAD_CONFIGS = [
    ("Sales Register", "sample_sales_register.xlsx", True),
    ("Purchase Register", "sample_purchase_register.xlsx", True),
    ("GSTR-1", "sample_gstr1.xlsx", True),
    ("GSTR-2B", "sample_gstr2b.xlsx", True),
    ("GSTR-3B", "sample_gstr3b.xlsx", False),
]
PAIR_CONFIGS = {
    "gstr1_sales": ("GSTR-1", "Sales Register", "GSTR-1 VS SALES REGISTER", "gstr1_vs_sales_register.xlsx"),
    "gstr2b_purchase": ("GSTR-2B", "Purchase Register", "GSTR-2B VS PURCHASE REGISTER", "gstr2b_vs_purchase_register.xlsx"),
    "gstr1_3b": ("GSTR-1", "GSTR-3B", "GSTR-1 VS GSTR-3B", "gstr1_vs_gstr3b.xlsx"),
}


def read_original_upload(uploaded) -> pd.DataFrame:
    """Read a pristine copy for the source-data sheets in Excel exports."""
    content = BytesIO(uploaded.getvalue())
    return pd.read_csv(content) if uploaded.name.lower().endswith(".csv") else pd.read_excel(content)


def restore_persisted_state() -> None:
    if st.session_state.get("persistence_checked"):
        return
    saved = load_state()
    if saved:
        for key, value in saved.items():
            st.session_state.setdefault(key, value)
        st.session_state["restored_message"] = True
    st.session_state["persistence_checked"] = True


def clear_session_keys(keys: list[str]) -> None:
    for key in keys:
        st.session_state.pop(key, None)


def data_management() -> None:
    st.subheader("Data Management")
    action = st.session_state.get("pending_data_action")
    if action:
        st.warning(f"Confirm {action.replace('_', ' ')}. This action cannot be undone.")
        confirm, cancel = st.columns(2)
        if confirm.button("Confirm", type="primary", key="confirm_data_action"):
            if action == "clear_uploaded_files":
                clear_uploaded_files()
                clear_session_keys(["source_originals", "invoice_sources", "gstr3b_raw"])
                saved = load_state() or {}
                for key in ["source_originals", "invoice_sources", "gstr3b_raw"]:
                    saved.pop(key, None)
                save_state(saved)
            elif action == "clear_reconciliation_results":
                clear_reconciliation_results()
                clear_session_keys(["recon", "validations", "insights", "pair_reconciliations"])
                save_state({
                    "source_originals": st.session_state.get("source_originals", {}),
                    "invoice_sources": st.session_state.get("invoice_sources", {}),
                    "gstr3b_raw": st.session_state.get("gstr3b_raw"),
                })
            else:
                reset_application()
                clear_session_keys(["source_originals", "invoice_sources", "gstr3b_raw", "recon", "validations", "insights", "pair_reconciliations"])
            st.session_state.pop("pending_data_action", None)
            st.success("Data cleared successfully.")
            st.rerun()
        if cancel.button("Cancel", key="cancel_data_action"):
            st.session_state.pop("pending_data_action", None)
            st.rerun()
    else:
        clear_uploads, clear_results, reset = st.columns(3)
        if clear_uploads.button("Clear Uploaded Files", use_container_width=True):
            st.session_state["pending_data_action"] = "clear_uploaded_files"
            st.rerun()
        if clear_results.button("Clear Reconciliation Results", use_container_width=True):
            st.session_state["pending_data_action"] = "clear_reconciliation_results"
            st.rerun()
        if reset.button("Reset Application", type="primary", use_container_width=True):
            st.session_state["pending_data_action"] = "reset_application"
            st.rerun()


def template_data(source_name: str, filename: str) -> bytes:
    path = TEMPLATE_DIR / filename
    return path.read_bytes() if path.exists() else sample_template_bytes(source_name)


def save_reconciliation_state(source_originals, invoice_sources, gstr3b_raw, pair_reconciliations, validations) -> None:
    primary = next((pair_reconciliations[key] for key in ("gstr1_sales", "gstr2b_purchase", "gstr1_3b") if key in pair_reconciliations), None)
    state = {
        "source_originals": source_originals,
        "invoice_sources": invoice_sources,
        "gstr3b_raw": gstr3b_raw,
        "pair_reconciliations": pair_reconciliations,
        "recon": primary,
        "validations": validations,
        "insights": generate_insights(primary) if primary is not None else pd.DataFrame(),
    }
    save_state(state)
    for key, value in state.items():
        st.session_state[key] = value


restore_persisted_state()
st.title("AI GST Compliance Copilot")
st.caption("Invoice-level GST reconciliation, exception intelligence, and audit-ready reporting.")
data_management()
if st.session_state.pop("restored_message", False):
    st.success("Previously uploaded data and reconciliation results have been restored.")

with st.sidebar:
    st.header("Upload files")
    uploads = {}
    for label, filename, invoice_level in UPLOAD_CONFIGS:
        uploads[label] = st.file_uploader(label, type=["xlsx", "xls", "csv"], key=f"upload_{label}")
        st.download_button("📥 Download Sample Template", data=template_data(label, filename), file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"template_{label}", use_container_width=True)
        st.caption("Accepted formats: XLSX, XLS, CSV")
    run = st.button("Run Reconciliation", type="primary", use_container_width=True)
    st.divider()
    st.subheader("About")
    st.caption("Uploaded files and results are saved locally until you explicitly clear them.")

if run:
    source_originals = dict(st.session_state.get("source_originals", {}))
    invoice_sources = dict(st.session_state.get("invoice_sources", {}))
    gstr3b_raw = st.session_state.get("gstr3b_raw")
    validation_frames = []
    progress = st.progress(0, "Reading source files…")
    for source_name, _, invoice_level in UPLOAD_CONFIGS:
        uploaded = uploads[source_name]
        if not uploaded:
            continue
        try:
            content = uploaded.getvalue()
            save_uploaded_file(source_name, uploaded.name, content)
            original = read_original_upload(uploaded)
            source_originals[source_name] = original
            if invoice_level:
                standardized = read_upload(uploaded)
                invoice_sources[source_name] = standardized
            else:
                gstr3b_raw = original
        except Exception as exc:
            st.error(f"Could not read {source_name}: {exc}")
    for source_name, frame in invoice_sources.items():
        validation_frames.append(validation_report(frame, source_name))
    progress.progress(55, "Reconciling source pairs…")
    pair_reconciliations = {}
    if {"GSTR-1", "Sales Register"}.issubset(invoice_sources):
        pair_reconciliations["gstr1_sales"] = reconcile({"GSTR-1": invoice_sources["GSTR-1"], "Sales Register": invoice_sources["Sales Register"]})
    if {"GSTR-2B", "Purchase Register"}.issubset(invoice_sources):
        pair_reconciliations["gstr2b_purchase"] = reconcile({"GSTR-2B": invoice_sources["GSTR-2B"], "Purchase Register": invoice_sources["Purchase Register"]})
    if "GSTR-1" in invoice_sources and gstr3b_raw is not None and not gstr3b_raw.empty:
        pair_reconciliations["gstr1_3b"] = gstr3b_reconciliation(invoice_sources["GSTR-1"], gstr3b_raw)
    validations = pd.concat(validation_frames, ignore_index=True) if validation_frames else pd.DataFrame()
    if not pair_reconciliations:
        st.error("Upload a valid reconciliation pair: GSTR-1 + Sales Register, GSTR-2B + Purchase Register, or GSTR-1 + GSTR-3B.")
    else:
        save_reconciliation_state(source_originals, invoice_sources, gstr3b_raw, pair_reconciliations, validations)
        progress.progress(100, "Analysis complete.")
        st.success("Reconciliation completed and saved successfully.")

if "recon" not in st.session_state or st.session_state.recon is None:
    st.info("Upload a source pair in the sidebar and select **Run Reconciliation**. Your data will remain available after refresh.")
    st.stop()

recon = st.session_state.recon
validations = st.session_state.get("validations", pd.DataFrame())
insights = st.session_state.get("insights", pd.DataFrame())
summary_data = summary(recon, validations)
metrics = st.columns(6)
for column, label, value in zip(metrics, ["Total Invoices", "Matched", "Mismatches", "Match %", "GST Exposure", "Health Score"], [summary_data["total"], summary_data["matched"], summary_data["mismatches"], f'{summary_data["match_pct"]}%', f'₹{summary_data["exposure"]:,.0f}', f'{summary_data["score"]}/100']):
    column.metric(label, value)
st.caption(f"Risk level: **{summary_data['risk']}** — {summary_data['score_reason']}")

tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "Reconciliation", "Exceptions & Validation", "Exports"])
with tab1:
    left, right = st.columns(2)
    mismatches = recon[recon["Status"] == "Mismatch"]
    with left:
        if not mismatches.empty:
            st.plotly_chart(px.pie(mismatches, names="Mismatch Type", title="Mismatch Type Distribution"), use_container_width=True)
    with right:
        risk = mismatches.groupby("Customer", as_index=False)["Exposure Amount"].sum().nlargest(10, "Exposure Amount")
        if not risk.empty:
            st.plotly_chart(px.bar(risk, x="Customer", y="Exposure Amount", title="Top Risk Customers"), use_container_width=True)
    monthly = recon.assign(Month=pd.to_datetime(recon["Invoice Date"], errors="coerce").dt.to_period("M").astype(str)).groupby("Month", as_index=False)["Exposure Amount"].sum()
    if not monthly.empty:
        st.plotly_chart(px.line(monthly, x="Month", y="Exposure Amount", markers=True, title="Monthly Exposure Trend"), use_container_width=True)
with tab2:
    st.subheader("Invoice reconciliation ledger")
    st.dataframe(recon, use_container_width=True, hide_index=True)
    pair_results = st.session_state.get("pair_reconciliations", {})
    for pair_key, (_, _, title, _) in PAIR_CONFIGS.items():
        if pair_key != "gstr1_sales" and pair_key in pair_results:
            st.subheader(title)
            st.dataframe(pair_results[pair_key], use_container_width=True, hide_index=True)
with tab3:
    st.subheader("AI-style exception guidance")
    st.dataframe(insights, use_container_width=True, hide_index=True)
    st.subheader("Validation report")
    st.dataframe(validations, use_container_width=True, hide_index=True)
with tab4:
    originals = st.session_state.get("source_originals", {})
    standardized = st.session_state.get("invoice_sources", {})
    pair_results = st.session_state.get("pair_reconciliations", {})
    for pair_key, (source_name, reference_name, title, filename) in PAIR_CONFIGS.items():
        reconciliation = pair_results.get(pair_key)
        if reconciliation is None:
            st.info(f"{title}: upload both source files and run reconciliation to enable this export.")
            continue
        try:
            reference_standardized = standardized.get(reference_name, originals.get(reference_name, pd.DataFrame()))
            workbook = reconciliation_export_bytes(source_name, reference_name, originals[source_name], originals[reference_name], standardized.get(source_name, originals[source_name]), reference_standardized, reconciliation)
            st.download_button(f"Download {title}", workbook, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"export_{pair_key}", use_container_width=True)
        except Exception as exc:
            st.error(f"{title} export could not be prepared: {exc}")
