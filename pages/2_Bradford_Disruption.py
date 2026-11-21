import sys
from pathlib import Path
import streamlit as st
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src import ui
from src.data_loader import load_clean_data
from src.metrics import calculate_bradford_scores
from src.stats_engine import compute_pareto_curve
from src.visualizers import plot_bradford_scatter, plot_pareto_curve

st.set_page_config(page_title="Bradford Disruption & Pareto", page_icon="⚡", layout="wide")

ui.render_header(
    title="Bradford Disruption Index & Pareto 80/20 Concentration",
    subtitle="Mathematically scoring unplanned absence frequency (B = S^2 * D) and isolating the vital few driving workforce disruption.",
    badge_text="Screen UI-03"
)

df = load_clean_data()
if df.empty:
    st.error("No data available.")
    st.stop()

bradford_df = calculate_bradford_scores(df)

# Tier counts
low_c = int((bradford_df["bradford_tier"] == "Low").sum())
med_c = int((bradford_df["bradford_tier"] == "Medium").sum())
high_c = int((bradford_df["bradford_tier"] == "High").sum())
sev_c = int((bradford_df["bradford_tier"] == "Severe").sum())

tier_cards = [
    ("Low Risk (< 50)", f"{low_c} Staff", "Informal line check-in", "Stable", "green"),
    ("Medium Risk (50-249)", f"{med_c} Staff", "Documented supervisory review", "Monitor", "amber"),
    ("High Risk (250-749)", f"{high_c} Staff", "Formal HR attendance counsel", "Alert", "purple"),
    ("Severe Risk (>= 750)", f"{sev_c} Staff", "Disciplinary / PIP intervention", "Critical", "red")
]
ui.render_kpi_grid_4(tier_cards)

st.markdown("### ")
col1, col2 = st.columns([1.5, 1.3])

with col1:
    st.subheader("Bradford Factor Disruption Landscape")
    fig_scatter = plot_bradford_scatter(bradford_df)
    st.plotly_chart(fig_scatter, use_container_width=True)

with col2:
    st.subheader("Pareto 80/20 Lorenz Curve")
    pareto_df = compute_pareto_curve(df)
    fig_pareto = plot_pareto_curve(pareto_df)
    st.plotly_chart(fig_pareto, use_container_width=True)

st.subheader("High & Severe Risk Action Roster (Bradford Score >= 250)")
high_risk_roster = bradford_df[bradford_df["bradford_score"] >= 250].copy()
high_risk_roster["total_cost"] = high_risk_roster["total_cost"].apply(lambda v: f"${v:,.2f}")
high_risk_roster["bradford_score"] = high_risk_roster["bradford_score"].round(0).astype(int)

st.dataframe(high_risk_roster, use_container_width=True, hide_index=True)
