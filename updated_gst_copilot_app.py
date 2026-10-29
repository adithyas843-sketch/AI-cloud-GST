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
