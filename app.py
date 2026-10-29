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

st.set_page_config(page_title="GST Copilot", page_icon="🧾", layout="wide")

# Minimalist visual system
st.markdown("""
<style>
:root{--page:#f7f6f2;--panel:#fffefa;--ink:#191919;--muted:#77746d;--line:#e5e2da;--green:#237a50;--amber:#b87512;--red:#bf332b}
.stApp{background:var(--page);color:var(--ink)}
[data-testid="stHeader"]{background:rgba(247,246,242,.96)}
[data-testid="stSidebar"]{background:#f0eee8;border-right:1px solid var(--line)}
[data-testid="stSidebar"]>div{padding-top:1.2rem}
.block-container{max-width:1450px;padding-top:1.5rem;padding-bottom:2.5rem}
h1,h2,h3{color:var(--ink);letter-spacing:-.035em}h1{font-family:Georgia,'Times New Roman',serif;font-weight:500}
p,[data-testid="stCaptionContainer"]{color:var(--muted)}
[data-testid="stMetric"]{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:15px 17px;box-shadow:none}
[data-testid="stMetricLabel"]{color:var(--muted)!important;font-size:.83rem}
[data-testid="stMetricValue"]{color:var(--ink)!important;font-size:1.8rem}
[data-testid="stFileUploader"]{border:1px solid var(--line);border-radius:9px;background:rgba(255,255,255,.48);padding:8px}
.stButton>button,.stDownloadButton>button{border-radius:22px;border:1px solid #c9c5bc;box-shadow:none;min-height:2.55rem}
.stButton>button[kind="primary"]{background:var(--ink);color:white;border-color:var(--ink)}
hr{border-color:var(--line)}
.eyebrow{text-transform:uppercase;letter-spacing:.12em;font-size:.72rem;color:var(--muted);margin-bottom:.25rem}
.panel{border:1px solid var(--line);border-radius:12px;background:rgba(255,254,250,.78);padding:18px 20px}
.ready{background:#191919;color:white;border-radius:12px;padding:18px 22px}
.ready small{color:#d2d0ca}
div[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:8px}
</style>
""", unsafe_allow_html=True)

TEMPLATE_DIR = Path(__file__).parent / "templates"
TEMPLATES = {
    "Sales Register": ("Tally", "sample_tally.xlsx"),
    "Purchase Register": ("Zoho", "sample_zoho.xlsx"),
    "GSTR-1": ("GSTR-1", "sample_gstr1.xlsx"),
    "GSTR-2B": ("GSTR-2B", "sample_gstr2b.xlsx"),
    "GSTR-3B": ("GSTR-3B", "sample_gstr3b.xlsx"),
}
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_template_bytes(key: str) -> bytes:
    """Get a bundled template if present, otherwise safely generate one."""
    if key not in TEMPLATES:
        raise ValueError(f"Unknown template type: {key}")
    template_key, filename = TEMPLATES[key]
    path = TEMPLATE_DIR / filename
    if path.exists():
        return path.read_bytes()
    return sample_template_bytes(template_key)


def read_original_upload(uploaded) -> pd.DataFrame:
    content = BytesIO(uploaded.getvalue())
    if uploaded.name.lower().endswith(".csv"):
        return pd.read_csv(content)
    return pd.read_excel(content)


def warning_count(recon_df: pd.DataFrame, validation_df: pd.DataFrame) -> int:
    for frame in (recon_df, validation_df):
        if not frame.empty and "Status" in frame.columns:
            status = frame["Status"].astype(str).str.strip().str.lower()
            count = int(status.isin({"warning", "review", "needs review", "needs_review"}).sum())
            if count:
                return count
    return 0


def clear_results() -> None:
    for key in ["recon", "validations", "insights", "gstr3b", "source_originals", "invoice_sources", "reconciliation_completed"]:
        st.session_state.pop(key, None)


has_results = bool(
    st.session_state.get("reconciliation_completed", False)
    and isinstance(st.session_state.get("recon"), pd.DataFrame)
)

# Sidebar stays intentionally minimal until the first successful run.
with st.sidebar:
    st.markdown("<div style='font-family:Georgia,serif;font-size:1.45rem;color:#191919'>GST Copilot</div><div style='font-size:.78rem;color:#77746d;margin-top:2px'>Reconciliation workspace</div>", unsafe_allow_html=True)
    st.divider()
    nav_items = ["Upload & Reconcile"]
    if has_results:
        nav_items += ["Dashboard", "Exception Report", "AI Insights", "Exports", "Data Management", "Help & Guide"]
    page = st.radio("Navigation", nav_items, label_visibility="collapsed", key="main_navigation")
    if has_results:
        st.divider()
        st.caption("Review reconciliation findings against supporting records before filing.")

# Main heading and single template button
head_left, head_right = st.columns([4, 1.5], vertical_alignment="center")
with head_left:
    st.markdown("<div class='eyebrow'>GST COMPLIANCE WORKSPACE</div>", unsafe_allow_html=True)
    st.title("Reconcile")
    st.caption("Invoice-level matching, exceptions and audit-ready reports.")
with head_right:
    if st.button("⇩  Download sample templates", use_container_width=True, key="open_templates"):
        st.session_state["show_templates"] = not st.session_state.get("show_templates", False)
st.divider()

# A right-side template panel appears only when the single button is clicked.
if st.session_state.get("show_templates", False):
    left_area, template_panel = st.columns([2.3, 1], gap="large")
    with template_panel:
        st.markdown("### Sample templates")
        st.caption("Choose only the template you need.")
        for label, (template_key, filename) in TEMPLATES.items():
            st.markdown(f"**{label}**")
            st.download_button(
                f"Download {label} template",
                data=get_template_bytes(label),
                file_name=filename,
                mime=MIME_XLSX,
                key=f"download_template_{label.replace('-', '_').replace(' ', '_')}",
                use_container_width=True,
            )
        if st.button("Close templates", use_container_width=True):
            st.session_state["show_templates"] = False
            st.rerun()
    with left_area:
        st.markdown("<div class='panel'><strong>Templates for your GST workflow</strong><p>Use the sample files to confirm the expected column headings before uploading your own data.</p></div>", unsafe_allow_html=True)

if page == "Upload & Reconcile":
    st.markdown("### Source files")
    st.caption("Upload the registers and GST returns you want to reconcile.")
    descriptions = {
        "Sales Register": "Sales invoices and credit notes",
        "Purchase Register": "Purchase invoices, debit notes and journal entries",
        "GSTR-1": "Outward supplies reported in GSTR-1",
        "GSTR-2B": "Inward supplies and input tax credit",
        "GSTR-3B": "Monthly summary and tax control check",
    }
    uploads = {}
    source_cols = st.columns(2, gap="large")
    for index, label in enumerate(TEMPLATES):
        with source_cols[index % 2]:
            st.markdown(f"**{label}**")
            st.caption(descriptions[label])
            upload_key = f"upload_{label.replace('-', '_').replace(' ', '_')}"
            uploads[label] = st.file_uploader(
                f"Upload {label}", type=["xlsx", "xls", "csv"], key=upload_key,
                label_visibility="collapsed", help="Accepted formats: XLSX, XLS, CSV · maximum 200 MB per file",
            )
            if uploads[label]:
                st.caption(f"Selected: {uploads[label].name}")
    st.divider()
    run_col, status_col = st.columns([1, 2.6], vertical_alignment="center")
    with run_col:
        run = st.button("▶  Run reconciliation", type="primary", use_container_width=True)
    with status_col:
        if not has_results:
            st.caption("Upload the relevant files, then run reconciliation. Metrics and navigation appear after a successful run.")
        else:
            st.caption("A previous successful run is available. Run again to refresh the results with the currently selected files.")

    if run:
        invoice_sources: dict[str, pd.DataFrame] = {}
        source_originals: dict[str, pd.DataFrame] = {}
        validation_frames = []
        gstr3b_raw = None
        errors = []
        progress = st.progress(0, text="Reading source files…")
        # Map the new upload labels onto the current engine's existing source keys.
        source_key_map = {"Sales Register": "Tally", "Purchase Register": "Zoho", "GSTR-1": "GSTR-1", "GSTR-2B": "GSTR-2B"}
        for label, uploaded in uploads.items():
            if uploaded is None:
                continue
            try:
                original = read_original_upload(uploaded)
                if label == "GSTR-3B":
                    gstr3b_raw = original
                    continue
                standardized = read_upload(uploaded)
                source_key = source_key_map[label]
                invoice_sources[source_key] = standardized
                source_originals[source_key] = original
                validation_frames.append(validation_report(standardized, source_key))
            except Exception as exc:
                errors.append(f"{label}: {exc}")
        if errors:
            for error in errors:
                st.error(f"Could not read {error}")
        progress.progress(55, text="Reconciling uploaded records…")
        if len(invoice_sources) < 2:
            progress.empty()
            st.error("Upload at least two invoice-level files to run reconciliation.")
        else:
            try:
                result = reconcile(invoice_sources)
                validation_result = pd.concat(validation_frames, ignore_index=True) if validation_frames else pd.DataFrame()
                insights_result = generate_insights(result)
                control_result = gstr3b_control(invoice_sources.get("GSTR-1"), gstr3b_raw)
                st.session_state.recon = result
                st.session_state.validations = validation_result
                st.session_state.insights = insights_result
                st.session_state.gstr3b = control_result
                st.session_state.source_originals = source_originals
                st.session_state.invoice_sources = invoice_sources
                st.session_state.reconciliation_completed = True
                progress.progress(100, text="Analysis complete")
                st.success("Reconciliation completed successfully.")
                st.rerun()
            except Exception as exc:
                progress.empty()
                st.error(f"Reconciliation could not be completed: {exc}")

    if not has_results:
        st.markdown("<div class='ready'><strong>Ready when you are</strong><br><small>Upload your source files, then run reconciliation to see your results.</small></div>", unsafe_allow_html=True)

# Results and navigation only appear after a successful reconciliation.
has_results = bool(st.session_state.get("reconciliation_completed", False) and isinstance(st.session_state.get("recon"), pd.DataFrame))
if has_results:
    recon = st.session_state.recon
    validations = st.session_state.get("validations", pd.DataFrame())
    insights = st.session_state.get("insights", pd.DataFrame())
    gstr3b = st.session_state.get("gstr3b", pd.DataFrame())
    s = summary(recon, validations)
    st.markdown("<div class='eyebrow'>RECONCILIATION RESULTS</div>", unsafe_allow_html=True)
    st.subheader("Summary")
    metric_cols = st.columns(4)
    metric_values = [s.get("total", 0), s.get("matched", 0), warning_count(recon, validations), s.get("mismatches", 0)]
    for col, label, value in zip(metric_cols, ["Total invoices", "Matched", "Warnings", "Mismatches"], metric_values):
        col.metric(label, f"{value:,}")
    st.divider()

    if page == "Dashboard":
        st.subheader("Reconciliation overview")
        mismatches = recon[recon["Status"].astype(str).str.lower() == "mismatch"] if "Status" in recon.columns else pd.DataFrame()
        c1, c2 = st.columns(2, gap="large")
        with c1:
            if not mismatches.empty and "Mismatch Type" in mismatches.columns:
                st.plotly_chart(px.pie(mismatches, names="Mismatch Type", title="Mismatch type distribution"), use_container_width=True)
            else:
                st.info("No mismatch-type chart is available for these results.")
        with c2:
            if not mismatches.empty and {"Customer", "Exposure Amount"}.issubset(mismatches.columns):
                risk = mismatches.groupby("Customer", as_index=False)["Exposure Amount"].sum().nlargest(10, "Exposure Amount")
                st.plotly_chart(px.bar(risk, x="Customer", y="Exposure Amount", title="Top risk customers"), use_container_width=True)
            else:
                st.info("No customer exposure chart is available for these results.")
        if {"Invoice Date", "Exposure Amount"}.issubset(recon.columns):
            monthly = recon.copy()
            monthly["Month"] = pd.to_datetime(monthly["Invoice Date"], errors="coerce").dt.to_period("M").astype(str)
            monthly = monthly.dropna(subset=["Month"]).groupby("Month", as_index=False)["Exposure Amount"].sum()
            if not monthly.empty:
                st.plotly_chart(px.line(monthly, x="Month", y="Exposure Amount", markers=True, title="Monthly exposure trend"), use_container_width=True)
    elif page == "Exception Report":
        st.subheader("Exceptions and validation")
        if "Status" in recon.columns:
            exceptions = recon[recon["Status"].astype(str).str.lower().isin(["mismatch", "warning", "review", "needs review", "needs_review"])]
        else:
            exceptions = recon
        st.dataframe(exceptions, use_container_width=True, hide_index=True)
        st.markdown("**Validation report**")
        st.dataframe(validations, use_container_width=True, hide_index=True)
    elif page == "AI Insights":
        st.subheader("Exception insights")
        st.dataframe(insights, use_container_width=True, hide_index=True)
    elif page == "Exports":
        st.subheader("Exports")
        originals = st.session_state.get("source_originals", {})
        standardized = st.session_state.get("invoice_sources", {})
        book_key = "Tally" if "Tally" in standardized else "Zoho" if "Zoho" in standardized else None
        if "GSTR-1" in standardized and book_key:
            try:
                report = professional_report_bytes(originals["GSTR-1"], originals[book_key], standardized["GSTR-1"], standardized[book_key], recon)
                st.download_button("Download professional reconciliation workbook", report, "gst_reconciliation_workbook.xlsx", MIME_XLSX, use_container_width=True)
            except Exception as exc:
                st.error(f"The export workbook could not be prepared: {exc}")
        else:
            st.info("Upload GSTR-1 and a sales/purchase register to create the professional reconciliation workbook.")
        if not gstr3b.empty:
            st.markdown("**GSTR-1 vs GSTR-3B control check**")
            st.dataframe(gstr3b, use_container_width=True, hide_index=True)
    elif page == "Data Management":
        st.subheader("Data management")
        st.caption("Clear uploaded files, clear the results, or reset the current application session.")
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("Clear uploaded files", use_container_width=True):
                for label in TEMPLATES:
                    key = f"upload_{label.replace('-', '_').replace(' ', '_')}"
                    st.session_state.pop(key, None)
                st.rerun()
        with c2:
            if st.button("Clear reconciliation results", use_container_width=True):
                clear_results()
                st.rerun()
        with c3:
            if st.button("Reset application", type="primary", use_container_width=True):
                st.session_state.clear()
                st.rerun()
    elif page == "Help & Guide":
        st.subheader("Help & Guide")
        st.markdown("- Upload the source files you want to compare.\n- Click **Run reconciliation**.\n- Review dashboard metrics, exceptions and insights.\n- Download the reconciliation workbook from **Exports**.\n- Always verify discrepancies against source documents before filing GST returns.")
elif page != "Upload & Reconcile":
    st.info("No completed reconciliation is available yet. Open Upload & Reconcile and run the analysis first.")
