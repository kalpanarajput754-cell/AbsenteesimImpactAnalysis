import sys
from pathlib import Path
import streamlit as st
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src import ui
from src.data_loader import load_clean_data
from src.metrics import calculate_payroll_loss, simulate_savings
from src.visualizers import plot_loss_waterfall

st.set_page_config(page_title="Financial Loss Simulator", page_icon="💰", layout="wide")

ui.render_header(
    title="Financial Loss Modeling & What-If Savings Simulator",
    subtitle="Decomposing payroll wage loss, operational overtime friction, and forecasting ROI from wellness policy interventions.",
    badge_text="Screen UI-06"
)

df = load_clean_data()
if df.empty:
    st.error("No data available.")
    st.stop()

direct_loss, friction_loss = calculate_payroll_loss(df)

col1, col2 = st.columns([1.4, 1.2])

with col1:
    st.subheader("Direct Payroll Loss vs Operational Friction (+20%)")
    fig_waterfall = plot_loss_waterfall(direct_loss, friction_loss)
    st.plotly_chart(fig_waterfall, use_container_width=True)

with col2:
    st.subheader("What-If Absenteeism Reduction Slider")
    st.markdown("Adjust targeted absenteeism reduction percentage to model net organization savings:")
    reduction_pct = st.slider("Targeted Reduction Scenario (%)", min_value=5, max_value=25, value=10, step=1)
    target_scope = st.radio("Intervention Scope", ["Preventable & Lifestyle Absences Only", "Entire Workforce Absences"])

    sim_res = simulate_savings(df, reduction_pct, preventable_only=(target_scope == "Preventable & Lifestyle Absences Only"))

    st.markdown("---")
    sim_kpis = [
        ("Simulated Net Savings", f"${sim_res['total_net_savings']:,.0f}", f"At {reduction_pct}% targeted reduction", "Financial ROI", "green"),
        ("Projected Loss", f"${sim_res['projected_loss']:,.0f}", f"Baseline: ${friction_loss:,.0f}", "Residual", "blue")
    ]
    ui.render_kpi_card(
        title=f"Net Financial Recovery ({reduction_pct}%)",
        value=f"${sim_res['total_net_savings']:,.0f}",
        subtitle=f"Direct: ${sim_res['direct_savings']:,.0f} + Overtime: ${sim_res['friction_savings']:,.0f}",
        trend="ROI Forecast",
        color_theme="green"
    )

    csv_sim = pd.DataFrame([sim_res]).to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Download Financial Simulation Summary (CSV)",
        data=csv_sim,
        file_name="financial_simulation_report.csv",
        mime="text/csv",
        use_container_width=True
    )
