import sys
import os
from pathlib import Path
import streamlit as st
import pandas as pd

# Path configuration for cloud and local execution
ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src import ui
from src.data_loader import load_clean_data, get_filter_options
from src.metrics import calculate_payroll_loss, calculate_bradford_scores
from src.visualizers import plot_monthly_trend, plot_department_hotspots

# Streamlit Page Setup
st.set_page_config(
    page_title="Workforce Absenteeism Diagnostic",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------
# 1. SIDEBAR FILTERS (Shared via st.session_state)
# -------------------------------------------------------------
st.sidebar.image("https://img.icons8.com/fluency/96/combo-chart.png", width=64)
st.sidebar.title("Diagnostic Filters")
st.sidebar.caption("PRJ-AIA-01 | Executive Controls")

filter_opts = get_filter_options()

# Date range selection
date_range = st.sidebar.date_input(
    "Absence Date Range",
    value=(filter_opts["min_date"], filter_opts["max_date"]),
    min_value=filter_opts["min_date"],
    max_value=filter_opts["max_date"]
)

# Department selection
all_depts = filter_opts["departments"]
selected_depts = st.sidebar.multiselect(
    "Filter Departments",
    options=all_depts,
    default=all_depts
)

# Reason category selection
all_cats = filter_opts["categories"]
selected_cats = st.sidebar.multiselect(
    "Filter Absence Category",
    options=all_cats,
    default=all_cats
)

st.sidebar.divider()
st.sidebar.markdown(
    """
    **Project Metadata:**
    - **Analyst:** Kalpana (2500320140071)
    - **Institution:** ABES Engineering College
    - **Mentor:** Prof. Gunjan Agrawal
    - **Tech Stack:** Python 3.12, SQLite, Plotly, Tailwind CDN
    """
)

# -------------------------------------------------------------
# 2. DATA INGESTION (Cached SQLite Query)
# -------------------------------------------------------------
df_filtered = load_clean_data(
    departments=selected_depts if len(selected_depts) < len(all_depts) else None,
    reason_categories=selected_cats if len(selected_cats) < len(all_cats) else None,
    date_range=date_range if isinstance(date_range, tuple) and len(date_range) == 2 else None
)

# -------------------------------------------------------------
# 3. EXECUTIVE HEADER
# -------------------------------------------------------------
ui.render_header(
    title="Workforce Productivity & Financial Loss Diagnostic",
    subtitle="Interactive enterprise surveillance engine quantifying operational absenteeism disruption, Bradford risk tiers, and overtime friction losses.",
    badge_text="Executive Portal v1.0"
)

# -------------------------------------------------------------
# 4. KPI METRIC GRID
# -------------------------------------------------------------
if df_filtered.empty:
    st.warning("No absence records match the selected sidebar filter criteria. Please broaden your filters.")
else:
    direct_loss, friction_loss = calculate_payroll_loss(df_filtered)
    total_hours = float(df_filtered["absent_hours"].sum())
    
    # Standard workforce denominator: 120 staff * 240 days * 8 hours = 230,400 hours
    total_available_hours = 120 * 240 * 8.0
    absence_rate = (total_hours / total_available_hours) * 100.0 if total_available_hours > 0 else 0.0

    bradford_df = calculate_bradford_scores(df_filtered)
    severe_count = int((bradford_df["bradford_tier"] == "Severe").sum())

    kpi_cards = [
        ("Total Lost Hours", f"{total_hours:,.0f} hrs", "Aggregate productive time lost", "Capacity Deficit", "blue"),
        ("Financial Friction", f"${friction_loss:,.0f}", f"Direct: ${direct_loss:,.0f} + 20% Friction", "Critical Loss", "red"),
        ("Absenteeism Rate", f"{absence_rate:.2f}%", "Proportion of contractual capacity", "Benchmark 3.5%", "amber"),
        ("Chronic Bradford Risk", f"{severe_count} Staff", "Bradford Score B >= 750", "High Disruption", "purple")
    ]
    ui.render_kpi_grid_4(kpi_cards)

    st.markdown("### ")

    # -------------------------------------------------------------
    # 5. VISUAL ANALYTICS SECTION
    # -------------------------------------------------------------
    col_chart1, col_chart2 = st.columns([1.6, 1.2])

    with col_chart1:
        st.subheader("📈 Temporal Absence Trajectory")
        fig_trend = plot_monthly_trend(df_filtered)
        st.plotly_chart(fig_trend, use_container_width=True)

    with col_chart2:
        st.subheader("🏢 Departmental Burden")
        fig_hotspots = plot_department_hotspots(df_filtered)
        st.plotly_chart(fig_hotspots, use_container_width=True)

    # -------------------------------------------------------------
    # 6. EXECUTIVE INSIGHTS & EXPORT
    # -------------------------------------------------------------
    st.markdown("### ")
    col_ins, col_act = st.columns([2, 1])

    with col_ins:
        ui.render_insight_box(
            title="Executive Diagnostic Alert: Operational Overtime Friction",
            message=f"Current filtered scope reflects {total_hours:,.0f} lost productive hours, driving a direct wage depletion of ${direct_loss:,.0f} and an additional ${(friction_loss - direct_loss):,.0f} in operational friction. {severe_count} employees exhibit severe Bradford disruption scores requiring structured HR intervention.",
            alert_type="danger" if severe_count > 5 else "warning"
        )

    with col_act:
        csv_data = df_filtered.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download Filtered Data (CSV)",
            data=csv_data,
            file_name="filtered_absence_analysis.csv",
            mime="text/csv",
            use_container_width=True
        )
