import sys
from pathlib import Path
import streamlit as st

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src import ui

st.set_page_config(page_title="About Project PRJ-AIA-01", page_icon="ℹ️", layout="wide")

ui.render_header(
    title="About Project PRJ-AIA-01 & Academic Attribution",
    subtitle="Major Capstone Project in HR Analytics & Workforce Productivity Modeling.",
    badge_text="Screen UI-08"
)

st.markdown(
    """
    ### 🎓 Academic & Institutional Metadata
    - **Project Title:** Absenteeism Impact Analysis: Workforce Productivity & Financial Loss Modeling
    - **Project Identifier:** `PRJ-AIA-01`
    - **Candidate:** **Kalpana** (University Roll No: `2500320140071`)
    - **Degree Program:** Master of Computer Applications (MCA) — 2nd Year Data Analytics
    - **Institution:** **ABES Engineering College, Ghaziabad** (Affiliated to Dr. A.P.J. Abdul Kalam Technical University, Lucknow)
    - **Department:** Department of Computer Applications
    - **Faculty Project Mentor:** **Prof. Gunjan Agrawal**
    - **Baseline Submission Date:** November 20, 2026

    ---

    ### 🛠️ Technology Stack & Engineering Standards
    - **Language & Runtime:** Python 3.12+
    - **Database Engine:** SQLite 3 (`data/app.db`) with strict foreign key constraints (`PRAGMA foreign_keys = ON`)
    - **Web Framework:** Streamlit Multipage Portal (`app.py` & `pages/`)
    - **Design System:** Tailwind CSS CDN sandboxed inside iframe shadow DOM (`src/ui.py`)
    - **Visual Analytics:** Plotly Express & Plotly Graph Objects (`src/visualizers.py`)
    - **Statistical Inference:** SciPy (`scipy.stats`) & statsmodels (ANOVA, Chi-Square, Pearson/Spearman, OLS)
    - **Quality Assurance:** pytest 8.3+ with 100% test pass rate on schema, deterministic math, and latency

    ---

    ### 🔗 Live Portals & Project Repositories
    - **Live Cloud Application:** [https://2500320140071-prj-aia-01.streamlit.app](https://2500320140071-prj-aia-01.streamlit.app)
    - **Official GitHub Repository:** [https://github.com/kalpana-71/2500320140071-prj-aia-01](https://github.com/kalpana-71/2500320140071-prj-aia-01)
    """
)
