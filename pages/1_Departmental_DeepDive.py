import sys
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src import ui
from src.data_loader import load_clean_data
from src.visualizers import plot_department_hotspots

st.set_page_config(page_title="Departmental Deep-Dive", page_icon="🏢", layout="wide")

ui.render_header(
    title="Departmental Deep-Dive & Role Vulnerability",
    subtitle="Comparative breakdown of absenteeism volume, overtime friction burden, and job role exposure across business units.",
    badge_text="Screen UI-02"
)

df = load_clean_data()

if df.empty:
    st.error("No data available.")
    st.stop()

# Department selector
selected_dept = st.selectbox(
    "Select Department for In-Depth Role Audit",
    options=["All Departments"] + sorted(df["department_name"].unique().tolist())
)

filtered_df = df if selected_dept == "All Departments" else df[df["department_name"] == selected_dept]

col1, col2 = st.columns([1.5, 1])

with col1:
    st.subheader("Departmental Friction Breakdown")
    fig_dept = plot_department_hotspots(filtered_df)
    st.plotly_chart(fig_dept, use_container_width=True)

with col2:
    st.subheader("Absence by Job Role")
    role_agg = filtered_df.groupby("job_title").agg(
        lost_hours=("absent_hours", "sum"),
        total_loss=("friction_loss_total", "sum"),
        event_count=("event_id", "count")
    ).reset_index().sort_values(by="lost_hours", ascending=True)

    fig_role = px.bar(
        role_agg,
        x="lost_hours",
        y="job_title",
        orientation="h",
        color="total_loss",
        color_continuous_scale="Viridis",
        labels={"lost_hours": "Lost Hours", "job_title": "Job Title", "total_loss": "Financial Loss ($)"}
    )
    fig_role.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=30, b=30)
    )
    st.plotly_chart(fig_role, use_container_width=True)

st.markdown("### Detailed Department Performance Ledger")
dept_table = filtered_df.groupby(["department_name", "job_title"]).agg(
    total_events=("event_id", "count"),
    total_hours=("absent_hours", "sum"),
    avg_hours_per_event=("absent_hours", "mean"),
    direct_loss=("direct_wage_loss", "sum"),
    friction_loss=("friction_loss_total", "sum")
).reset_index()

dept_table["avg_hours_per_event"] = dept_table["avg_hours_per_event"].round(1)
dept_table["direct_loss"] = dept_table["direct_loss"].apply(lambda v: f"${v:,.2f}")
dept_table["friction_loss"] = dept_table["friction_loss"].apply(lambda v: f"${v:,.2f}")

st.dataframe(dept_table, use_container_width=True, hide_index=True)
