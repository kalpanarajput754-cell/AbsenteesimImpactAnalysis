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
from src.stats_engine import run_day_of_week_chisquare

st.set_page_config(page_title="Temporal Absence Patterns", page_icon="📅", layout="wide")

ui.render_header(
    title="Temporal Absence Trends & Weekday Surge Detection",
    subtitle="Detecting cyclical patterns, seasonal surges, and Monday/Friday casual absence clustering.",
    badge_text="Screen UI-05"
)

df = load_clean_data()
if df.empty:
    st.error("No data available.")
    st.stop()

# Chi-Square Analysis
chi_res = run_day_of_week_chisquare(df)

if chi_res["is_significant"]:
    ui.render_insight_box(
        title=f"Statistical Anomaly Detected: Monday/Friday Surge (+{chi_res['surge_pct']:.1f}%)",
        message=f"Chi-Square test (χ² = {chi_res['chi2_stat']:.2f}, p = {chi_res['p_value']:.4e}) conclusively rejects the null hypothesis of equal daily distribution. Absence volume spikes heavily on Mondays and Fridays adjacent to weekends.",
        alert_type="danger"
    )
else:
    ui.render_insight_box(
        title="Uniform Weekday Distribution",
        message="No statistically significant weekday spike detected within the current data filter.",
        alert_type="info"
    )

st.markdown("### ")
col1, col2 = st.columns([1.2, 1.2])

with col1:
    st.subheader("Day-of-Week Absence Volume")
    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    df_counts = pd.DataFrame({
        "Weekday": weekday_order,
        "Absence Events": [chi_res["counts"].get(d, 0) for d in weekday_order]
    })
    fig_days = px.bar(
        df_counts,
        x="Weekday",
        y="Absence Events",
        color="Absence Events",
        color_continuous_scale="Blues",
        text="Absence Events"
    )
    fig_days.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_days, use_container_width=True)

with col2:
    st.subheader("Quarterly Absence Seasonality")
    q_agg = df.groupby("quarter").agg(
        lost_hours=("absent_hours", "sum"),
        total_loss=("friction_loss_total", "sum")
    ).reset_index()

    fig_q = px.bar(
        q_agg,
        x="quarter",
        y="lost_hours",
        color="total_loss",
        color_continuous_scale="Reds",
        labels={"quarter": "Quarter", "lost_hours": "Lost Hours", "total_loss": "Total Loss ($)"},
        text="lost_hours"
    )
    fig_q.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_q, use_container_width=True)
