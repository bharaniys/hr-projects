"""
app.py - Streamlit HR dashboard on top of HR_Dashboard.xlsx

Run:  streamlit run app.py
"""
import pandas as pd
import streamlit as st

from validators import (add_attendance_pct, add_onboarding_pct, kpis,
                        load_data, validate_payroll, MIN_BASIC_PCT)

st.set_page_config(page_title="HR Analytics Dashboard", layout="wide")
st.title("HR Analytics & Employee Records Dashboard")

# ---- data source: uploaded workbook or the default file
uploaded = st.sidebar.file_uploader("Upload HR workbook (.xlsx)", type="xlsx")
path = uploaded if uploaded else "HR_Dashboard.xlsx"
min_basic = st.sidebar.slider("Minimum Basic % of Gross (policy)", 0.30, 0.60, MIN_BASIC_PCT, 0.01)

try:
    data = load_data(path)
except Exception as e:
    st.error(f"Could not read the workbook: {e}")
    st.stop()

emp, can = data["Employees"], data["Candidates"]
att = add_attendance_pct(data["Attendance"])
onb = add_onboarding_pct(data["Onboarding"])
pay = validate_payroll(data["Payroll_Input"], emp, min_basic)
k = kpis(data)

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["Overview", "Hiring", "Attendance", "Onboarding", "Payroll Validation"])

with tab1:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Active headcount", k["active"])
    c2.metric("Resigned", k["resigned"])
    c3.metric("Attrition", f"{k['attrition']:.1%}")
    c4.metric("Hire conversion", f"{k['hire_conversion']:.1%}")
    st.subheader("Headcount by department (active)")
    st.bar_chart(emp[emp["Status"] == "Active"].groupby("Department").size())
    st.subheader("Employee records")
    dept = st.multiselect("Filter department", sorted(emp["Department"].dropna().unique()))
    st.dataframe(emp[emp["Department"].isin(dept)] if dept else emp, use_container_width=True)

with tab2:
    order = ["Applied", "Screening", "Interview", "Offer", "Joined", "Rejected"]
    stage = can["Stage"].value_counts().reindex(order).fillna(0)
    st.subheader("Hiring pipeline")
    st.bar_chart(stage)
    st.subheader("Candidates by source")
    st.bar_chart(can["Source"].value_counts())
    st.dataframe(can, use_container_width=True)

with tab3:
    trend = att.groupby("Month")["Attendance %"].mean()
    st.subheader("Average attendance % by month")
    st.line_chart(trend)
    low = att[att["Attendance %"] < 0.90]
    st.subheader("Employee-months below 90% attendance")
    st.dataframe(low, use_container_width=True)

with tab4:
    st.subheader("Onboarding completion")
    st.metric("Average completion", f"{onb['Completion %'].mean():.0%}")
    st.dataframe(onb, use_container_width=True)
    st.caption("Pending tasks: " + str(int((onb.drop(columns=['Emp ID', 'Employee Name', 'Completion %']) == 'No').sum().sum())))

with tab5:
    flagged = pay[pay["Validation Check"] != "OK"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Rows checked", len(pay))
    c2.metric("Rows flagged", len(flagged))
    c3.metric("Error rate", f"{len(flagged) / len(pay):.1%}" if len(pay) else "0%")
    st.subheader("Flagged rows")
    st.dataframe(flagged, use_container_width=True)
    st.download_button("Download flagged rows (CSV)", flagged.to_csv(index=False),
                       file_name="payroll_flagged.csv", mime="text/csv")
    with st.expander("All payroll rows"):
        st.dataframe(pay, use_container_width=True)
