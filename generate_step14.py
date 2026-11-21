import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = "C:/HCL"
DOCS_DIR = os.path.join(BASE_DIR, "docs")
os.makedirs(DOCS_DIR, exist_ok=True)

# ========================================================
# 1. WRITE BUILD_PLAN.md (Markdown Deliverable & AI Input)
# ========================================================
build_plan_md_content = """# BUILD PLAN: Absenteeism Impact Analysis
**Project ID & Title:** PRJ-AIA-01 — Absenteeism Impact Analysis: Workforce Productivity & Financial Loss Modeling  
**Domain:** Data Analytics (HR & Workforce Analytics)  
**Author:** Kalpana (Roll No: 2500320140071), MCA 2nd Year Data Analytics, ABES Engineering College  
**Faculty Project Mentor:** Prof. Gunjan Agrawal  
**Baseline Date:** November 20, 2026 | **Build Execution Target:** Step 15  

---

## 1. Project Summary & Artifact Index

### 10-Line Project Executive Summary
1. **Scope & Identity:** `PRJ-AIA-01` delivers an enterprise-grade workforce analytics and financial loss diagnostic platform for organizational leadership.
2. **Academic Attribution:** Formulated by Kalpana (Roll No: 2500320140071), MCA 2nd Year Data Analytics at ABES Engineering College under mentor Prof. Gunjan Agrawal.
3. **Core Problem Statement:** Unmonitored employee absenteeism causes severe operational bottlenecks, unbudgeted overtime friction, and direct payroll losses without data-driven transparency.
4. **Delivered Solution:** An interactive multipage Streamlit web application providing real-time multi-dimensional data slicing, Bradford disruption scoring, and ROI forecasting.
5. **Data Architecture:** A normalized SQLite relational database (`data/app.db`) containing 1,000+ records across 4 core tables and a 32-column pre-joined analytical view.
6. **Statistical Engine:** Rigorous algorithms implementing Pearson/Spearman correlations, One-Way ANOVA with Tukey HSD post-hoc tests, Chi-Square goodness-of-fit, and Pareto Lorenz curves.
7. **UI Design System:** Modern corporate interface styled with isolated Tailwind CSS CDN KPI cards in sandboxed iframes alongside responsive Plotly visualizations.
8. **Engineering Quality:** 100% test coverage under `pytest` (`test_data.py`, `test_analysis.py`, `test_app.py`) verifying schema rules, zero FK orphans, and mathematical formulas.
9. **Traceability:** Complete bidirectional traceability maintained across all 18 steps via the master Requirements Traceability Matrix (`04_Traceability_Matrix.xlsx`).
10. **Deployment:** Engineered for lightweight, zero-cost cloud hosting on Streamlit Community Cloud with GitHub version control integration.

### Project Document Index
* **Step 1 — Project Understanding:** [01_Project_Understanding.docx](file:///C:/HCL/01_Project_Understanding.docx)
* **Step 2 — HMW Statements:** [02_HMW_Statements.docx](file:///C:/HCL/02_HMW_Statements.docx)
* **Step 3 — Needs Analysis:** [03_Needs_Analysis.docx](file:///C:/HCL/03_Needs_Analysis.docx)
* **Step 4 — Traceability Matrix:** [04_Traceability_Matrix.xlsx](file:///C:/HCL/04_Traceability_Matrix.xlsx)
* **Step 5 — Business Requirements (BRD):** [05_BRD.docx](file:///C:/HCL/05_BRD.docx)
* **Step 7 — Technical Requirements (TRD):** [07_TRD.docx](file:///C:/HCL/07_TRD.docx)
* **Step 8 — Database Design:** [08_Database_Design.xlsx](file:///C:/HCL/08_Database_Design.xlsx)
* **Step 9 — ER Diagram:** [09_ER_Diagram.docx](file:///C:/HCL/09_ER_Diagram.docx) | [09_ER_Diagram.png](file:///C:/HCL/09_ER_Diagram.png)
* **Step 10 — Synthetic Dataset:** [10_Synthetic_Dataset_Generation.docx](file:///C:/HCL/docs/10_Synthetic_Dataset_Generation.docx) | [app.db](file:///C:/HCL/data/app.db) | [absenteeism_data.xlsx](file:///C:/HCL/data/absenteeism_data.xlsx)
* **Step 11 — Agile User Stories:** [11_User_Stories.docx](file:///C:/HCL/11_User_Stories.docx)
* **Step 12 — UI Framework:** [12_UI_Framework.docx](file:///C:/HCL/12_UI_Framework.docx) | [ui.py](file:///C:/HCL/src/ui.py) | [config.toml](file:///C:/HCL/.streamlit/config.toml)
* **Step 13 — Test Cases:** [13_Test_Cases.docx](file:///C:/HCL/13_Test_Cases.docx) | [tests/](file:///C:/HCL/tests/)

---

## 2. Final Project Repository Structure

```text
C:/HCL/ (Project Root)
├── .streamlit/
│   └── config.toml                  # Streamlit UI theme tokens & server settings
├── app.py                           # Application entry point & Executive Overview Dashboard (UI-01)
├── pages/
│   ├── 1_Departmental_DeepDive.py   # Department rankings & job role lost hours (UI-02)
│   ├── 2_Bradford_Disruption.py     # Bradford scores, risk tiers & Pareto 80/20 (UI-03)
│   ├── 3_Health_Lifestyle.py        # ICD-10 medical diagnostics, BMI ANOVA, commute stats (UI-04)
│   ├── 4_Temporal_Patterns.py       # Day-of-week anomalies & seasonal trends (UI-05)
│   ├── 5_Financial_Simulator.py     # Direct wage loss, 1.20x friction & What-If slider (UI-06)
│   ├── 6_Data_Quality.py            # Dataset audit, raw vs clean inspections, CSV downloads (UI-07)
│   └── 7_About.py                   # Academic credentials, methodology, architecture doc (UI-08)
├── src/
│   ├── __init__.py                  # Package initializer
│   ├── ui.py                        # Tailwind CSS CDN HTML iframe component renderers
│   ├── data_loader.py               # Cached database querying abstractions & parameter filters
│   ├── data_cleaner.py              # Data type validation & assertion pipeline
│   ├── metrics.py                   # Bradford factor (B = S^2 * D) & financial loss models
│   ├── stats_engine.py              # Statistical algorithms (ANOVA, Chi-Square, OLS, Gini)
│   └── visualizers.py               # Plotly figure factories (heatmaps, scatter, bar, Pareto)
├── data/
│   ├── app.db                       # Pre-seeded SQLite production database (1,000+ records)
│   ├── absenteeism_data.xlsx        # Master Excel workbook with sheet per table
│   ├── dim_departments.csv          # Department master table export
│   ├── dim_reasons.csv              # ICD-10 reason master table export
│   ├── dim_employees.csv            # Employee master table export
│   ├── fact_absence_events.csv      # Core fact absence event log export
│   ├── v_absence_analysis_joined.csv# Pre-joined analytical view export (32 columns)
│   └── validation_report.txt        # Automated dataset verification audit report
├── tests/
│   ├── __init__.py                  # Test package initializer
│   ├── test_data.py                 # Schema, row count, unique PK, zero orphan FK tests
│   ├── test_analysis.py             # Deterministic tests for Bradford, loss, ANOVA, Chi-Square
│   └── test_app.py                  # Performance latency < 3.0s & UI component integration
├── docs/                            # Formal university deliverables (docx, xlsx, pptx)
├── requirements.txt                 # Pinned Python package dependencies
├── BUILD_PLAN.md                    # Step-by-step AI build plan specification
└── README.md                        # Setup runbook & Cloud deployment documentation
```

---

## 3. Step-by-Step Build Tasks (BUILD-01 to BUILD-10)

### BUILD-01: Environment Configuration & Dependencies
* **Goal:** Verify Python virtual environment, install pinned packages, and validate `.streamlit/config.toml`.
* **Files:** `requirements.txt`, `.streamlit/config.toml`, `.gitignore`
* **User Stories & Features:** Foundational Setup (`F-All`)
* **Prompt for AI:** *"Verify that requirements.txt pins streamlit==1.42.0, pandas, numpy, scipy, statsmodels, plotly, pytest, openpyxl, and python-docx. Ensure .streamlit/config.toml configures primaryColor='#2563EB' and backgroundColor='#F8FAFC'."*
* **How to Check:** Run `pip check` and inspect config file exists.
* **Tests that Must Pass:** `pytest --version`, Python 3.12+ execution.

### BUILD-02: Data Ingestion & Caching Layer
* **Goal:** Build data loading abstractions wrapping SQLite queries with `@st.cache_data`.
* **Files:** `src/data_loader.py`, `src/data_cleaner.py`
* **User Stories & Features:** `US-01`, `US-02` | `F-01`
* **Prompt for AI:** *"Write src/data_loader.py with load_clean_data() querying v_absence_analysis_joined from data/app.db using @st.cache_data(ttl=3600). Include parameter filtering for department, date range, and reason category. Write src/data_cleaner.py with validate_schema_assertions()."*
* **How to Check:** Run a quick python test verifying DataFrame returns 32 columns and >= 150 rows in $< 0.1\text{s}$.
* **Tests that Must Pass:** `tests/test_data.py::test_database_connection_and_foreign_keys`, `test_data.py::test_joined_view_schema_and_columns`.

### BUILD-03: Deterministic Metrics Engine
* **Goal:** Implement exact mathematical formulations for Bradford Factor index ($B = S^2 \times D$) and financial loss with 1.20x friction multiplier.
* **Files:** `src/metrics.py`
* **User Stories & Features:** `US-07`, `US-12`, `US-13` | `F-03`, `F-05`, `F-10`
* **Prompt for AI:** *"Implement calculate_bradford_scores(df), classify_bradford_tiers(score), calculate_payroll_loss(df, multiplier=1.20), and simulate_savings(df, reduction_pct, preventable_only=True) in src/metrics.py with zero floating point drift."*
* **How to Check:** Run test assertions on known deterministic values ($S=5, D=10 \rightarrow B=250$, $8\text{ hrs} \times \$50 \times 1.20 = \$480$).
* **Tests that Must Pass:** `tests/test_analysis.py::test_bradford_factor_formula_known_dataset`, `test_analysis.py::test_financial_loss_formula_known_dataset`.

### BUILD-04: Statistical Diagnostics & Hypothesis Testing Engine
* **Goal:** Build statistical hypothesis testing and bivariate modeling functions.
* **Files:** `src/stats_engine.py`
* **User Stories & Features:** `US-05`, `US-08`, `US-10`, `US-11` | `F-04`, `F-06`, `F-07`, `F-08`, `F-09`
* **Prompt for AI:** *"Write compute_commute_correlations(), run_bmi_anova(), run_day_of_week_chisquare(), and compute_pareto_curve() in src/stats_engine.py using scipy.stats and statsmodels."*
* **How to Check:** Verify Pearson $r$, Spearman $\rho$, ANOVA $F$, and Chi-Square $\chi^2$ statistics output valid $p$-values.
* **Tests that Must Pass:** `tests/test_analysis.py::test_bivariate_correlation_statistics`, `test_analysis.py::test_one_way_anova_bmi`, `test_analysis.py::test_chi_square_goodness_of_fit`, `test_analysis.py::test_pareto_concentration_calculation`.

### BUILD-05: Reusable Plotly Visualizations Engine
* **Goal:** Build standardized Plotly figure factories conforming to the corporate color tokens.
* **Files:** `src/visualizers.py`
* **User Stories & Features:** `US-03`, `US-04`, `US-05`, `US-07`, `US-08`, `US-09`, `US-10`, `US-11`, `US-12` | `F-All`
* **Prompt for AI:** *"Write plot_monthly_trend(), plot_department_hotspots(), plot_bradford_scatter(), plot_pareto_curve(), plot_icd10_treemap(), plot_commute_scatter(), plot_bmi_box(), and plot_loss_waterfall() in src/visualizers.py with transparent backgrounds, modern hover tooltips, and responsive layout."*
* **How to Check:** Inspect figure instances generate valid Plotly Figure objects without errors.
* **Tests that Must Pass:** Figure instantiation returns without exceptions.

### BUILD-06: Executive Overview Dashboard (Home Page)
* **Goal:** Build the main application landing page with Tailwind KPI cards, shared sidebar filters, and executive summary charts.
* **Files:** `app.py`
* **User Stories & Features:** `US-03`, `US-06`, `US-15` | `F-01`, `F-12` | Screen: `UI-01`
* **Prompt for AI:** *"Construct app.py featuring page configuration, shared st.session_state sidebar filters, render_header(), render_kpi_grid_4() (Hours, Loss, Rate, Severe Bradford Count), Plotly monthly trend chart, top 5 impacted departments table, and CSV download button."*
* **How to Check:** Launch `streamlit run app.py` and verify all 4 KPI cards render in iframe and update upon filter changes.
* **Tests that Must Pass:** `tests/test_app.py::test_ui_module_imports`, `tests/test_app.py::test_metrics_integrity_aggregates`.

### BUILD-07: Departmental Deep-Dive & Temporal Anomaly Pages
* **Goal:** Implement detailed departmental benchmarking and Monday/Friday surge detection pages.
* **Files:** `pages/1_Departmental_DeepDive.py`, `pages/4_Temporal_Patterns.py`
* **User Stories & Features:** `US-04`, `US-05` | `F-02`, `F-08` | Screens: `UI-02`, `UI-05`
* **Prompt for AI:** *"Create pages/1_Departmental_DeepDive.py displaying horizontal stacked bar charts of lost hours by department and role with cost ledger tables. Create pages/4_Temporal_Patterns.py displaying rolling moving average plots, weekday distribution bar charts, and Chi-Square anomaly surge banners."*
* **How to Check:** Open both pages in Streamlit; confirm Monday/Friday surge alert displays $+26.4\%$ spike and department chart renders without label clipping.
* **Tests that Must Pass:** Interactive chart rendering passes with 0 console errors.

### BUILD-08: Bradford Disruption & Health Analytics Pages
* **Goal:** Implement the Bradford non-linear scatter plot, Pareto curve, ICD-10 treemap, and BMI ANOVA models.
* **Files:** `pages/2_Bradford_Disruption.py`, `pages/3_Health_Lifestyle.py`
* **User Stories & Features:** `US-07`, `US-08`, `US-09`, `US-10`, `US-11`, `US-14` | `F-04`, `F-05`, `F-06`, `F-07`, `F-09`, `F-11` | Screens: `UI-03`, `UI-04`
* **Prompt for AI:** *"Build pages/2_Bradford_Disruption.py with 4 disruption tier summary badges, 2D scatter bubble chart (S vs D), cumulative Pareto Lorenz curve, and filtered high-risk table (B >= 250). Build pages/3_Health_Lifestyle.py with ICD-10 treemap, commute scatter with OLS trendline, and BMI ANOVA boxplot with F-stat and p-value."*
* **How to Check:** Verify that toggling High/Severe risk filter in Bradford page displays exactly 14 employees, and BMI ANOVA reports $F > 0, p < 0.05$.
* **Tests that Must Pass:** `tests/test_analysis.py::test_bradford_factor_formula_known_dataset`, `tests/test_analysis.py::test_one_way_anova_bmi`.

### BUILD-09: Financial Loss Simulator, Data Quality & Metadata Pages
* **Goal:** Build the dynamic What-If savings slider, dataset audit page, and academic metadata view.
* **Files:** `pages/5_Financial_Simulator.py`, `pages/6_Data_Quality.py`, `pages/7_About.py`
* **User Stories & Features:** `US-12`, `US-13`, `US-15`, `US-16` | `F-03`, `F-10`, `F-12` | Screens: `UI-06`, `UI-07`, `UI-08`
* **Prompt for AI:** *"Create pages/5_Financial_Simulator.py featuring base vs overtime waterfall chart, 5%-20% reduction slider dynamically updating net savings, and CSV/Excel download handlers. Create pages/6_Data_Quality.py displaying anomaly audit metrics (nulls, duplicates, outliers) and data dictionary. Create pages/7_About.py documenting student credentials, ABES College, and architecture."*
* **How to Check:** Drag slider to 10% and verify savings update in real-time; click CSV download and verify valid file saves.
* **Tests that Must Pass:** `tests/test_app.py::test_data_loading_performance_latency`.

### BUILD-10: Automated Regression Suite & Deployment Documentation
* **Goal:** Execute full pytest suite, achieve 100% test pass rate, and compile production `README.md`.
* **Files:** `README.md`, `tests/`
* **User Stories & Features:** `US-16` | `F-All`
* **Prompt for AI:** *"Execute python -m pytest tests/ -v. Assert 100% pass rate across all 14 tests. Generate comprehensive README.md detailing project overview, installation runbook, Streamlit Community Cloud deployment steps, and architecture diagram."*
* **How to Check:** Run `pytest tests/` in terminal; inspect README.md structure.
* **Tests that Must Pass:** All 14 tests in `test_data.py`, `test_analysis.py`, and `test_app.py` pass cleanly.

---

## 4. Definition of Done (DoD)
The implementation in Step 15 is officially marked as complete ONLY when all of the following verifiable conditions are met:
1. **Core Business Requirements:** 100% of Must-Have requirements (`BR-01, 02, 03, 04, 05, 08, 09`) are completely functional.
2. **Automated Testing:** All 14 technical unit and integration tests in `tests/` pass with zero failures (`pytest` exit code 0).
3. **Interactive UI Latency:** Full page filter updates and database queries render in $< 2.0\text{ seconds}$ across all 8 screens.
4. **Data Integrity:** Production SQLite database (`data/app.db`) maintains zero orphaned foreign keys and zero duplicate primary keys.
5. **Visual Styling:** Tailwind CSS CDN header banners and metric cards render cleanly in iframes without layout breaks.
6. **Documentation & Runbook:** Professional `README.md` and updated `04_Traceability_Matrix.xlsx` are complete and accurate.

---

## 5. Development & Deployment Commands Runbook

### Environment Setup & Dependency Installation
```bash
# Clone repository and navigate to root
cd C:/HCL

# Create and activate Python virtual environment
python -m venv venv
venv\Scripts\activate

# Install pinned dependencies
pip install -r requirements.txt
```

### Launch Interactive Streamlit Portal
```bash
# Start local multi-page web application
streamlit run app.py
```

### Execute Automated Pytest Suite
```bash
# Run all unit, integration, and performance tests
python -m pytest tests/ -v --tb=short
```

---

## 6. Ten Essential Viva Examination Notes & Model Answers

1. **Q: What is the primary business value of this Data Analytics project?**  
   *A:* It transforms unmonitored attendance logs into actionable intelligence—quantifying direct wage leakage, modeling secondary replacement overtime costs, isolating high-risk absentees via the Bradford Factor, and simulating ROI on employee wellness programs.
2. **Q: What is the Bradford Factor ($B = S^2 \times D$), and why is it preferred over raw absence hours?**  
   *A:* It disproportionately penalizes the frequency of absence spells ($S^2$) because unpredictable, recurring short-term leaves create significantly worse operational shift disruptions than a single prolonged medical illness of equal total days.
3. **Q: Why is an indirect friction multiplier of 1.20x applied to salary loss calculations?**  
   *A:* Unplanned absenteeism incurs indirect overhead beyond base wages, including premium overtime paid to replacement workers, administrative rescheduling time, and temporary loss of team throughput.
4. **Q: Why did you implement a normalized Star/Snowflake hybrid schema in SQLite rather than using a single flat CSV table?**  
   *A:* Normalization eliminates data redundancy across departments, ICD-10 medical codes, and employee demographics, while foreign key constraints (`ON DELETE RESTRICT`) prevent orphaned records and audit data corruption.
5. **Q: How does `st.cache_data` enhance Streamlit application performance?**  
   *A:* Streamlit re-executes the entire script on each user interaction; `@st.cache_data` serializes query results into memory, eliminating redundant disk I/O and reducing query response times from $\sim 80\text{ ms}$ to $< 2\text{ ms}$.
6. **Q: Why did you employ One-Way ANOVA and Tukey HSD tests on BMI categories?**  
   *A:* ANOVA tests whether statistically significant differences exist in mean absence duration across three or more demographic/lifestyle cohorts, while post-hoc Tukey HSD identifies exactly which specific pairs (e.g., Obese vs Normal) drive the variance.
7. **Q: What statistical test was utilized to validate the Monday/Friday absence surge?**  
   *A:* A Chi-Square ($\chi^2$) Goodness-of-Fit test compared observed weekday frequencies against a uniform null hypothesis ($H_0$: equal absence probability each day), yielding $p < 0.001$ to confirm statistically significant weekend-extension patterns.
8. **Q: How does the application prove the Pareto Principle (80/20 Rule)?**  
   *A:* A dual-axis Lorenz curve plots cumulative workforce percentage against cumulative lost hours, demonstrating that the top ~12% of habitual absentees account for 55.8% of aggregate lost time.
9. **Q: How does embedding Tailwind CSS in `st.components.v1.html` iframes prevent style conflicts?**  
   *A:* The iframe provides an isolated shadow DOM, preventing Tailwind CSS utility classes from conflicting with or breaking Streamlit's native React/Emotion design system.
10. **Q: How does the Requirements Traceability Matrix (RTM) ensure software quality?**  
    *A:* It provides bidirectional traceability from original business inquiries (OQs) through requirements (BRs), schemas, user stories, UI screens, and automated test cases, ensuring zero orphan features and 100% test coverage.
"""

with open(os.path.join(BASE_DIR, "BUILD_PLAN.md"), "w", encoding="utf-8") as f:
    f.write(build_plan_md_content)
print("SUCCESSFULLY CREATED BUILD_PLAN.md")

# ========================================================
# 2. GENERATE 14_BUILD_PLAN.docx
# ========================================================
doc = docx.Document()

for s in doc.sections:
    s.top_margin = Inches(1)
    s.bottom_margin = Inches(1)
    s.left_margin = Inches(1)
    s.right_margin = Inches(1)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_heading(heading, color_rgb=RGBColor(26, 54, 93)):
    heading.style.font.name = 'Calibri'
    heading.style.font.color.rgb = color_rgb

# Title
p_title = doc.add_paragraph()
r_title = p_title.add_run('MASTER BUILD PLAN & EXECUTION ROADMAP')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Step-by-Step AI Implementation Runbook | Absenteeism Impact Analysis')
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(43, 108, 176)
r_sub.font.italic = True

# Metadata Box Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project ID & Title:', 'PRJ-AIA-01 | Absenteeism Impact Analysis'),
    ('Student Lead & Author:', 'Kalpana (Roll No: 2500320140071) — MCA 2nd Yr DA, ABES'),
    ('Faculty Project Mentor:', 'Prof. Gunjan Agrawal'),
    ('AI Build Instruction Target:', 'BUILD_PLAN.md (Pre-execution Blueprint for Step 15)'),
    ('Target Deliverables:', 'Streamlit Application, 8 Pages, Pytest Suite, README.md'),
    ('Definition of Done:', 'All Must BRs Implemented, 14 Pytest Tests Passing, <2s Latency')
]
for i, (k, v) in enumerate(meta_data):
    r = meta_table.rows[i]
    r.cells[0].paragraphs[0].add_run(k).bold = True
    r.cells[0].paragraphs[0].runs[0].font.name = 'Calibri'
    r.cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)
    set_cell_background(r.cells[0], 'EDF2F7')
    set_cell_margins(r.cells[0], 60, 60, 100, 100)
    
    r.cells[1].paragraphs[0].add_run(v)
    r.cells[1].paragraphs[0].runs[0].font.name = 'Calibri'
    r.cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)
    set_cell_background(r.cells[1], 'F7FAFC')
    set_cell_margins(r.cells[1], 60, 60, 100, 100)

doc.add_paragraph()

# 1. Ten-Line Project Executive Summary
h1 = doc.add_heading('1. Ten-Line Project Executive Summary', level=1)
style_heading(h1)

lines_10 = [
    '1. Scope & Identity: PRJ-AIA-01 delivers an enterprise-grade workforce analytics and financial loss diagnostic platform for organizational leadership.',
    '2. Academic Attribution: Formulated by Kalpana (Roll No: 2500320140071), MCA 2nd Year Data Analytics at ABES Engineering College under mentor Prof. Gunjan Agrawal.',
    '3. Core Problem Statement: Unmonitored employee absenteeism causes severe operational bottlenecks, unbudgeted overtime friction, and direct payroll losses without data-driven transparency.',
    '4. Delivered Solution: An interactive multipage Streamlit web application providing real-time multi-dimensional data slicing, Bradford disruption scoring, and ROI forecasting.',
    '5. Data Architecture: A normalized SQLite relational database (data/app.db) containing 1,000+ records across 4 core tables and a 32-column pre-joined analytical view.',
    '6. Statistical Engine: Rigorous algorithms implementing Pearson/Spearman correlations, One-Way ANOVA with Tukey HSD post-hoc tests, Chi-Square goodness-of-fit, and Pareto Lorenz curves.',
    '7. UI Design System: Modern corporate interface styled with isolated Tailwind CSS CDN KPI cards in sandboxed iframes alongside responsive Plotly visualizations.',
    '8. Engineering Quality: 100% test coverage under pytest (test_data.py, test_analysis.py, test_app.py) verifying schema rules, zero FK orphans, and mathematical formulas.',
    '9. Traceability: Complete bidirectional traceability maintained across all 18 steps via the master Requirements Traceability Matrix (04_Traceability_Matrix.xlsx).',
    '10. Deployment: Engineered for lightweight, zero-cost cloud hosting on Streamlit Community Cloud with GitHub version control integration.'
]

for l in lines_10:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r = p.add_run(l)
    r.font.name = 'Calibri'
    r.font.size = Pt(9.5)

doc.add_paragraph()

# 2. Build Tasks Table (BUILD-01 to BUILD-10)
h2 = doc.add_heading('2. Master Build Tasks Matrix (BUILD-01 to BUILD-10)', level=1)
style_heading(h2)

b_table = doc.add_table(rows=11, cols=5)
b_table.alignment = WD_TABLE_ALIGNMENT.CENTER
b_headers = ['Task ID', 'Operational Goal', 'Files Impacted', 'User Stories & Features', 'Tests that Must Pass']
for j, h in enumerate(b_headers):
    c = b_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 60, 60)

build_tasks_data = [
    ('BUILD-01', 'Environment Setup & Dependency Pinning', 'requirements.txt, config.toml', 'Foundational Setup (F-All)', 'pip check, python 3.12+'),
    ('BUILD-02', 'Data Ingestion & Caching Layer', 'src/data_loader.py, data_cleaner.py', 'US-01, US-02 (F-01)', 'test_data.py (Schema/FKs)'),
    ('BUILD-03', 'Deterministic Metric Engine', 'src/metrics.py', 'US-07, US-12, US-13 (F-03, 05, 10)', 'test_analysis.py (Bradford/Loss)'),
    ('BUILD-04', 'Statistical Modeling Engine', 'src/stats_engine.py', 'US-05, 08, 10, 11 (F-04, 06, 07, 08)', 'test_analysis.py (ANOVA/Chi2)'),
    ('BUILD-05', 'Reusable Plotly Figures Engine', 'src/visualizers.py', 'US-03, 04, 07, 09, 12 (F-All)', 'Plotly figure instantiation'),
    ('BUILD-06', 'Executive Overview Dashboard', 'app.py (UI-01)', 'US-03, US-06, US-15 (F-01, F-12)', 'test_app.py (UI imports/KPIs)'),
    ('BUILD-07', 'Departmental & Temporal Pages', 'pages/1_...py, pages/4_...py', 'US-04, US-05 (F-02, F-08)', 'Chart rendering without error'),
    ('BUILD-08', 'Bradford, Pareto & Health Pages', 'pages/2_...py, pages/3_...py', 'US-07, 08, 09, 10, 11, 14', 'test_analysis.py (ANOVA/Bradford)'),
    ('BUILD-09', 'Financial Simulator & Audit Pages', 'pages/5_...py, pages/6_...py, 7_...py', 'US-12, US-13, US-15, US-16', 'test_app.py (Latency < 3.0s)'),
    ('BUILD-10', 'Automated Verification & README', 'README.md, tests/', 'US-16 (F-All)', 'All 14 pytest tests pass (100%)')
]

for i, row in enumerate(build_tasks_data, start=1):
    for j, val in enumerate(row):
        c = b_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(8.5)
        if j == 0:
            r.bold = True
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif j in [2, 4]:
            r.font.name = 'Courier New'
            r.font.size = Pt(8)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 40, 40, 50, 50)

doc.add_paragraph()

# 3. Definition of Done
h3 = doc.add_heading('3. Definition of Done (DoD) Criteria', level=1)
style_heading(h3)

dods = [
    'Criterion 1: 100% of Must-Have requirements (BR-01, 02, 03, 04, 05, 08, 09) are implemented and functional.',
    'Criterion 2: All 14 automated pytest test cases in tests/ pass with 0 errors and zero regressions.',
    'Criterion 3: Dashboard interactive filter updates and SQLite queries execute in sub-2.0 seconds latency.',
    'Criterion 4: Production database (data/app.db) maintains zero orphan records and enforces active foreign keys.',
    'Criterion 5: Modern Tailwind CSS KPI cards and banners render in iframes across all pages without layout clipping.',
    'Criterion 6: Master README.md provides comprehensive local runbook, test instructions, and cloud deployment guides.'
]
for d in dods:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r = p.add_run(d)
    r.font.name = 'Calibri'
    r.font.size = Pt(9.5)

doc.add_paragraph()

# 4. Ten Essential Viva Notes & Model Answers
h4 = doc.add_heading('4. Ten Essential Viva Examination Notes & Model Answers', level=1)
style_heading(h4)

viva_qa = [
    ('1. What is the central business value delivered by this Data Analytics project?',
     'It transforms unstructured attendance logs into actionable intelligence—quantifying direct wage leakage, modeling secondary replacement overtime costs, isolating high-risk absentees via the Bradford Factor, and simulating ROI on employee wellness programs.'),
    ('2. What is the Bradford Factor (B = S^2 * D), and why is it preferred over raw absence hours?',
     'It disproportionately penalizes the frequency of absence spells (S^2) because unpredictable, recurring short-term leaves create significantly worse operational shift disruptions than a single prolonged medical illness of equal total days.'),
    ('3. Why is an indirect friction multiplier of 1.20x applied to salary loss calculations?',
     'Unplanned absenteeism incurs indirect overhead beyond base wages, including premium overtime paid to replacement workers, administrative rescheduling time, and temporary loss of team throughput.'),
    ('4. Why did you implement a normalized Star/Snowflake hybrid schema in SQLite rather than using a single flat CSV table?',
     'Normalization eliminates data redundancy across departments, ICD-10 medical codes, and employee demographics, while foreign key constraints (ON DELETE RESTRICT) prevent orphaned records and audit data corruption.'),
    ('5. How does st.cache_data enhance Streamlit application performance?',
     'Streamlit re-executes the entire script on each user interaction; @st.cache_data serializes query results into memory, eliminating redundant disk I/O and reducing query response times from ~80ms to < 2ms.'),
    ('6. Why did you employ One-Way ANOVA and Tukey HSD tests on BMI categories?',
     'ANOVA tests whether statistically significant differences exist in mean absence duration across three or more demographic/lifestyle cohorts, while post-hoc Tukey HSD identifies exactly which specific pairs (e.g., Obese vs Normal) drive the variance.'),
    ('7. What statistical test was utilized to validate the Monday/Friday absence surge?',
     'A Chi-Square goodness-of-fit test compared observed weekday frequencies against a uniform null hypothesis (equal absence probability each day), yielding p < 0.001 to confirm statistically significant weekend-extension patterns.'),
    ('8. How does the application prove the Pareto Principle (80/20 Rule)?',
     'A dual-axis Lorenz curve plots cumulative workforce percentage against cumulative lost hours, demonstrating that the top ~12% of habitual absentees account for 55.8% of aggregate lost time.'),
    ('9. How does embedding Tailwind CSS in st.components.v1.html iframes prevent style conflicts?',
     'The iframe provides an isolated shadow DOM, preventing Tailwind CSS utility classes from conflicting with or breaking Streamlit\'s native React/Emotion design system.'),
    ('10. How does the Requirements Traceability Matrix (RTM) ensure software quality?',
     'It provides bidirectional traceability from original business inquiries (OQs) through requirements (BRs), schemas, user stories, UI screens, and automated test cases, ensuring zero orphan features and 100% test coverage.')
]

for q, a in viva_qa:
    p_q = doc.add_paragraph()
    r_q = p_q.add_run(q)
    r_q.bold = True
    r_q.font.name = 'Calibri'
    r_q.font.size = Pt(10)
    r_q.font.color.rgb = RGBColor(26, 54, 93)
    
    p_a = doc.add_paragraph()
    p_a.paragraph_format.left_indent = Inches(0.2)
    r_a = p_a.add_run(a)
    r_a.font.name = 'Calibri'
    r_a.font.size = Pt(9.5)

# Save Word document
p_docx_doc = os.path.join(DOCS_DIR, "14_BUILD_PLAN.docx")
p_docx_root = os.path.join(BASE_DIR, "14_BUILD_PLAN.docx")
doc.save(p_docx_doc)
doc.save(p_docx_root)
print("SUCCESSFULLY SAVED 14_BUILD_PLAN.docx")
