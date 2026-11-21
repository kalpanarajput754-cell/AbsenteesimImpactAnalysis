import sys
from pathlib import Path
import streamlit as st
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src import ui
from src.data_loader import load_clean_data
from src.stats_engine import compute_commute_correlations, run_bmi_anova
from src.visualizers import plot_icd10_treemap, plot_commute_scatter, plot_bmi_box

st.set_page_config(page_title="Health & Lifestyle Determinants", page_icon="🩺", layout="wide")

ui.render_header(
    title="Health Diagnostics & Lifestyle Determinants",
    subtitle="Evaluating ICD-10 clinical classifications, commute distance friction, and BMI variance across workforce cohorts.",
    badge_text="Screen UI-04"
)

df = load_clean_data()
if df.empty:
    st.error("No data available.")
    st.stop()

# 1. ICD-10 Treemap
st.subheader("ICD-10 Clinical Diagnostic Hierarchy")
fig_treemap = plot_icd10_treemap(df)
st.plotly_chart(fig_treemap, use_container_width=True)

st.markdown("### ")
col1, col2 = st.columns([1, 1])

# 2. Commute Burden
with col1:
    st.subheader("Commute Distance vs Absence Duration")
    commute_stats = compute_commute_correlations(df)
    fig_commute = plot_commute_scatter(df, commute_stats)
    st.plotly_chart(fig_commute, use_container_width=True)
    st.caption(
        f"**Correlation Analysis:** Pearson r = {commute_stats['pearson_r']:.3f} (p = {commute_stats['pearson_p']:.4f}) | "
        f"Spearman ρ = {commute_stats['spearman_rho']:.3f} | R² = {commute_stats['r_squared']:.3f}"
    )

# 3. BMI ANOVA
with col2:
    st.subheader("BMI Cohorts vs Absence Duration")
    anova_res = run_bmi_anova(df)
    fig_bmi = plot_bmi_box(df)
    st.plotly_chart(fig_bmi, use_container_width=True)
    st.caption(
        f"**One-Way ANOVA:** F-Statistic = {anova_res['f_stat']:.3f} | p-value = {anova_res['p_value']:.4f} | "
        f"Statistically Significant: {'YES (p < 0.05)' if anova_res['is_significant'] else 'NO'}"
    )
