import sys
from pathlib import Path
import streamlit as st
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src import ui
from src.data_loader import load_clean_data, load_raw_table
from src.data_cleaner import audit_dataset_health, validate_schema_assertions

st.set_page_config(page_title="Data Quality & Integrity Audit", page_icon="🛡️", layout="wide")

ui.render_header(
    title="Data Quality Assurance & Schema Integrity Certification",
    subtitle="Auditing completeness rates, primary key uniqueness, foreign key integrity, and dataset dictionary.",
    badge_text="Screen UI-07"
)

df = load_clean_data()
audit = audit_dataset_health(df)
validation = validate_schema_assertions(df)

quality_cards = [
    ("Completeness Score", f"{audit['completeness_score']:.1f}%", "Zero missing critical cells", "Certified", "green"),
    ("Missing Cells", f"{audit['null_cells']}", "Null attributes detected", "Clean", "blue"),
    ("Duplicate Events", f"{audit['duplicate_events']}", "Unique primary keys certified", "Valid", "purple"),
    ("Outlier Events", f"{audit['outlier_events']}", "IQR duration bounds check", "Monitored", "amber")
]
ui.render_kpi_grid_4(quality_cards)

st.markdown("### ")
st.subheader("Relational Tables Inspection")
selected_table = st.selectbox(
    "Select SQLite Entity to Inspect",
    ["dim_departments", "dim_reasons", "dim_employees", "fact_absence_events"]
)

raw_df = load_raw_table(selected_table)
st.write(f"Showing **{len(raw_df)} rows** from `{selected_table}`:")
st.dataframe(raw_df, use_container_width=True, hide_index=True)

csv_raw = raw_df.to_csv(index=False).encode("utf-8")
st.download_button(
    f"📥 Download {selected_table}.csv",
    data=csv_raw,
    file_name=f"{selected_table}.csv",
    mime="text/csv"
)
