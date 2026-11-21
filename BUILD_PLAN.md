# BUILD PLAN: Absenteeism Impact Analysis
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
* **How to Check:** Run a quick python test verifying DataFrame returns 32 columns and >= 150 rows in $< 0.1	ext{s}$.
* **Tests that Must Pass:** `tests/test_data.py::test_database_connection_and_foreign_keys`, `test_data.py::test_joined_view_schema_and_columns`.

### BUILD-03: Deterministic Metrics Engine
* **Goal:** Implement exact mathematical formulations for Bradford Factor index ($B = S^2 	imes D$) and financial loss with 1.20x friction multiplier.
* **Files:** `src/metrics.py`
* **User Stories & Features:** `US-07`, `US-12`, `US-13` | `F-03`, `F-05`, `F-10`
* **Prompt for AI:** *"Implement calculate_bradford_scores(df), classify_bradford_tiers(score), calculate_payroll_loss(df, multiplier=1.20), and simulate_savings(df, reduction_pct, preventable_only=True) in src/metrics.py with zero floating point drift."*
* **How to Check:** Run test assertions on known deterministic values ($S=5, D=10 ightarrow B=250$, $8	ext{ hrs} 	imes \$50 	imes 1.20 = \$480$).
* **Tests that Must Pass:** `tests/test_analysis.py::test_bradford_factor_formula_known_dataset`, `test_analysis.py::test_financial_loss_formula_known_dataset`.

### BUILD-04: Statistical Diagnostics & Hypothesis Testing Engine
* **Goal:** Build statistical hypothesis testing and bivariate modeling functions.
* **Files:** `src/stats_engine.py`
* **User Stories & Features:** `US-05`, `US-08`, `US-10`, `US-11` | `F-04`, `F-06`, `F-07`, `F-08`, `F-09`
* **Prompt for AI:** *"Write compute_commute_correlations(), run_bmi_anova(), run_day_of_week_chisquare(), and compute_pareto_curve() in src/stats_engine.py using scipy.stats and statsmodels."*
* **How to Check:** Verify Pearson $r$, Spearman $ho$, ANOVA $F$, and Chi-Square $\chi^2$ statistics output valid $p$-values.
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
3. **Interactive UI Latency:** Full page filter updates and database queries render in $< 2.0	ext{ seconds}$ across all 8 screens.
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
venv\Scriptsctivate

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
2. **Q: What is the Bradford Factor ($B = S^2 	imes D$), and why is it preferred over raw absence hours?**  
   *A:* It disproportionately penalizes the frequency of absence spells ($S^2$) because unpredictable, recurring short-term leaves create significantly worse operational shift disruptions than a single prolonged medical illness of equal total days.
3. **Q: Why is an indirect friction multiplier of 1.20x applied to salary loss calculations?**  
   *A:* Unplanned absenteeism incurs indirect overhead beyond base wages, including premium overtime paid to replacement workers, administrative rescheduling time, and temporary loss of team throughput.
4. **Q: Why did you implement a normalized Star/Snowflake hybrid schema in SQLite rather than using a single flat CSV table?**  
   *A:* Normalization eliminates data redundancy across departments, ICD-10 medical codes, and employee demographics, while foreign key constraints (`ON DELETE RESTRICT`) prevent orphaned records and audit data corruption.
5. **Q: How does `st.cache_data` enhance Streamlit application performance?**  
   *A:* Streamlit re-executes the entire script on each user interaction; `@st.cache_data` serializes query results into memory, eliminating redundant disk I/O and reducing query response times from $\sim 80	ext{ ms}$ to $< 2	ext{ ms}$.
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
