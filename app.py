from __future__ import annotations
from io import BytesIO
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st
from reconciliation import read_upload, reconcile, summary, gstr3b_control
from validators import validation_report
from ai_insights import generate_insights
from export_utils import professional_report_bytes

st.set_page_config(page_title="AI GST Compliance Copilot", page_icon="🧾", layout="wide")
st.markdown('<style>[data-testid="stMetric"]{background-color:#1e293b;border:1px solid #334155;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.25)}[data-testid="stMetricLabel"]{color:white!important;font-weight:600}[data-testid="stMetricValue"]{color:white!important;font-weight:700}.block-container{padding-top:1.5rem}</style>', unsafe_allow_html=True)
st.title("AI GST Compliance Copilot")
st.caption("Invoice-level GST reconciliation, exception intelligence, and audit-ready reporting.")

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
    return pd.read_csv(content) if uploaded.name.lower().endswith(".csv") else pd.read_excel(content)

with st.sidebar:
    st.header("Upload files")
    uploads={}
    for label,key in [("GSTR-1 Export","GSTR-1"),("Tally Sales Register","Tally"),("Zoho Books Sales Register","Zoho"),("GSTR-3B Summary","GSTR-3B")]:
        uploads[key]=st.file_uploader(label,type=["xlsx","xls","csv"],key=key)
        template_path = TEMPLATE_DIR / TEMPLATES[key]
        if template_path.exists():
            st.download_button("📥 Download Sample Template", data=template_path.read_bytes(), file_name=TEMPLATES[key], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"template_{key}", use_container_width=True)
    run=st.button("Run Reconciliation",type="primary",use_container_width=True)
    st.divider(); st.subheader("About")
    st.caption("Files stay in memory for the duration of this session. Review all recommendations with supporting documents before filing.")

if run:
    invoice_sources={}; gstr3b_raw=None
    source_originals={}
    validation_frames=[]
    progress=st.progress(0,"Reading source files…")
    for name,up in uploads.items():
        if up and name=="GSTR-3B":
            try: gstr3b_raw=pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
            except Exception as exc: st.error(f"Could not read GSTR-3B: {exc}")
        elif up:
            try:
                original = read_original_upload(up)
                df=read_upload(up); invoice_sources[name]=df; source_originals[name]=original; validation_frames.append(validation_report(df,name))
            except Exception as exc: st.error(f"Could not read {name}: {exc}")
    progress.progress(55,"Reconciling invoices…")
    if len(invoice_sources) < 2:
        st.error("Upload at least two invoice-level sources (GSTR-1, Tally, or Zoho) for invoice reconciliation.")
    else:
        st.session_state.recon=reconcile(invoice_sources)
        st.session_state.validations=pd.concat(validation_frames,ignore_index=True) if validation_frames else pd.DataFrame()
        st.session_state.insights=generate_insights(st.session_state.recon)
        st.session_state.gstr3b=gstr3b_control(invoice_sources.get("GSTR-1"),gstr3b_raw)
        st.session_state.source_originals=source_originals
        st.session_state.invoice_sources=invoice_sources
        progress.progress(100,"Analysis complete."); st.success("Reconciliation completed successfully.")

if "recon" not in st.session_state:
    st.info("Upload source registers in the sidebar and select **Run Reconciliation**. Use `python sample_data_generator.py` to create a demo dataset.")
    st.stop()
recon, validations, insights=st.session_state.recon,st.session_state.validations,st.session_state.insights
gstr3b=st.session_state.get("gstr3b",pd.DataFrame())
s=summary(recon,validations)
cols=st.columns(6)
for c,label,value in zip(cols,["Total Invoices","Matched","Mismatches","Match %","GST Exposure","Health Score"],[s["total"],s["matched"],s["mismatches"],f'{s["match_pct"]}%',f'₹{s["exposure"]:,.0f}',f'{s["score"]}/100']): c.metric(label,value)
st.caption(f"Risk level: **{s['risk']}** — {s['score_reason']}")
tab1,tab2,tab3,tab4=st.tabs(["Dashboard","Reconciliation","Exceptions & Validation","Exports"])
with tab1:
    left,right=st.columns(2); mismatches=recon[recon.Status=="Mismatch"]
    with left:
        if not mismatches.empty: st.plotly_chart(px.pie(mismatches,names="Mismatch Type",title="Mismatch Type Distribution"),use_container_width=True)
    with right:
        risk=mismatches.groupby("Customer",as_index=False)["Exposure Amount"].sum().nlargest(10,"Exposure Amount")
        if not risk.empty: st.plotly_chart(px.bar(risk,x="Customer",y="Exposure Amount",title="Top Risk Customers"),use_container_width=True)
    monthly=recon.assign(Month=pd.to_datetime(recon["Invoice Date"]).dt.to_period("M").astype(str)).groupby("Month",as_index=False)["Exposure Amount"].sum()
    if not monthly.empty: st.plotly_chart(px.line(monthly,x="Month",y="Exposure Amount",markers=True,title="Monthly Exposure Trend"),use_container_width=True)
with tab2:
    st.subheader("Invoice reconciliation ledger"); st.dataframe(recon,use_container_width=True,hide_index=True)
    if not gstr3b.empty:
        st.subheader("GSTR-1 vs GSTR-3B aggregate control"); st.dataframe(gstr3b,use_container_width=True,hide_index=True)
with tab3:
    st.subheader("AI-style exception guidance"); st.dataframe(insights,use_container_width=True,hide_index=True)
    st.subheader("Validation report"); st.dataframe(validations,use_container_width=True,hide_index=True)
with tab4:
    originals=st.session_state.get("source_originals", {})
    standardized=st.session_state.get("invoice_sources", {})
    books_source="Tally" if "Tally" in standardized else "Zoho" if "Zoho" in standardized else None
    if "GSTR-1" not in standardized or books_source is None:
        st.info("Upload GSTR-1 and either Tally or Zoho to create the professional reconciliation workbook.")
    else:
        workbook=professional_report_bytes(originals["GSTR-1"], originals[books_source], standardized["GSTR-1"], standardized[books_source], recon)
        st.download_button("Download Professional Reconciliation Workbook", workbook, "gst_reconciliation_workbook.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
