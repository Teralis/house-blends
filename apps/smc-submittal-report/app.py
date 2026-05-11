import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

import streamlit as st
from generate_weekly_report import analyze_submittals, generate_word_report

st.set_page_config(page_title="NSW Ports Hold Point Report", page_icon="📋")

st.title("NSW Ports Hold Point Report")
st.write("Upload your Procore submittal log export to generate the weekly report.")

uploaded = st.file_uploader("Drop SubmittalLog CSV here", type="csv")

if uploaded:
    with st.spinner("Processing..."):
        try:
            report_data = analyze_submittals(uploaded)
            buf         = generate_word_report(report_data)
        except Exception as e:
            st.error(f"Failed to process file: {e}")
            st.stop()

    week_end = report_data['week_end']
    filename = f"NSW_Ports_Weekly_Report_{week_end.strftime('%Y%m%d')}.docx"

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Hold Points",   report_data['total_hold_points'])
    col2.metric("Open",                report_data['status_counts'].get('Open',   0))
    col3.metric("Activity This Week",  report_data['week_activity_count'])

    st.download_button(
        label="Download Report",
        data=buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
