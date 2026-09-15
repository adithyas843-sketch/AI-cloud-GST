from __future__ import annotations
from io import BytesIO
import pandas as pd
import plotly.express as px
import streamlit as st
from reconciliation import read_upload, reconcile, summary, gstr3b_control
from validators import validation_report
from ai_insights import generate_insights

st.set_page_config(page_title="AI GST Compliance Copilot", page_icon="🧾", layout="wide")
st.markdown("<style>.stMetric{background:#fff;border:1px solid #e6e9ef;border-radius:12px;padding:12px}.block-container{padding-top:1.5rem}</style>",unsafe_allow_html=True)
st.title("AI GST Compliance Copilot")
st.caption("Invoice-level GST reconciliation, exception intelligence, and audit-ready reporting.")

def excel_bytes(sheets: dict[str,pd.DataFrame]):
    out=BytesIO()
    with pd.ExcelWriter(out,engine="xlsxwriter") as writer:
        for name, df in sheets.items(): df.to_excel(writer,sheet_name=name[:31],index=False)
    return out.getvalue()

with st.sidebar:
    st.header("Upload files")
    uploads={}
    for label,key in [("GSTR-1 Export","GSTR-1"),("Tally Sales Register","Tally"),("Zoho Books Sales Register","Zoho"),("GSTR-3B Summary","GSTR-3B")]:
        uploads[key]=st.file_uploader(label,type=["xlsx","xls","csv"],key=key)
    run=st.button("Run Reconciliation",type="primary",use_container_width=True)
    st.divider(); st.subheader("About")
    st.caption("Files stay in memory for the duration of this session. Review all recommendations with supporting documents before filing.")

if run:
    invoice_sources={}; gstr3b_raw=None
    validation_frames=[]
    progress=st.progress(0,"Reading source files…")
    for name,up in uploads.items():
        if up and name=="GSTR-3B":
            try: gstr3b_raw=pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
            except Exception as exc: st.error(f"Could not read GSTR-3B: {exc}")
        elif up:
            try:
                df=read_upload(up); invoice_sources[name]=df; validation_frames.append(validation_report(df,name))
            except Exception as exc: st.error(f"Could not read {name}: {exc}")
    progress.progress(55,"Reconciling invoices…")
    if len(invoice_sources) < 2:
        st.error("Upload at least two invoice-level sources (GSTR-1, Tally, or Zoho) for invoice reconciliation.")
    else:
        st.session_state.recon=reconcile(invoice_sources)
        st.session_state.validations=pd.concat(validation_frames,ignore_index=True) if validation_frames else pd.DataFrame()
        st.session_state.insights=generate_insights(st.session_state.recon)
        st.session_state.gstr3b=gstr3b_control(invoice_sources.get("GSTR-1"),gstr3b_raw)
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
    management=pd.DataFrame([{"Metric":"Total invoices processed","Value":s["total"]},{"Metric":"Matched invoices","Value":s["matched"]},{"Metric":"Mismatch invoices","Value":s["mismatches"]},{"Metric":"Match percentage","Value":s["match_pct"]},{"Metric":"GST exposure","Value":s["exposure"]},{"Metric":"GST health score","Value":s["score"]},{"Metric":"Risk level","Value":s["risk"]}])
    st.download_button("Download Excel Reconciliation Report",excel_bytes({"Reconciliation":recon,"GSTR3B Control":gstr3b,"Validations":validations,"Insights":insights}),"gst_reconciliation_report.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download Exception Report",excel_bytes({"Exceptions":insights}),"gst_exception_report.xlsx")
    st.download_button("Download GST Exposure Report",excel_bytes({"Exposure":recon[recon.Status=="Mismatch"]}),"gst_exposure_report.xlsx")
    st.download_button("Download Management Summary",excel_bytes({"Management Summary":management}),"gst_management_summary.xlsx")
