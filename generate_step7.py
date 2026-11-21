import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

doc = docx.Document()

# 1-inch margins
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

# ==================== COVER / TITLE ====================
p_title = doc.add_paragraph()
r_title = p_title.add_run('TECHNICAL REQUIREMENTS DOCUMENT (TRD)')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(24)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('System Architecture & Engineering Specification | Absenteeism Impact Analysis')
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(43, 108, 176)
r_sub.font.italic = True

# Metadata Box Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Document Reference:', 'TRD-PRJ-AIA-01-v1.0'),
    ('Project ID & Title:', 'PRJ-AIA-01 | Absenteeism Impact Analysis'),
    ('Student Author:', 'Kalpana (Roll No: 2500320140071) — MCA 2nd Yr DA, ABES'),
    ('Faculty Project Mentor:', 'Prof. Gunjan Agrawal'),
    ('Target Architecture:', 'Python 3.12, Streamlit Multipage, SQLite, Tailwind CDN, Plotly'),
    ('Deployment Target:', 'Streamlit Community Cloud + GitHub CI/CD Pipeline')
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

# 1. Architecture Pipeline (Mermaid)
h1 = doc.add_heading('1. System Architecture & End-to-End Pipeline', level=1)
style_heading(h1)

p_arch_desc = doc.add_paragraph()
p_arch_desc.add_run('The application follows a decoupled multi-tier architecture spanning ingestion, validation, analytical computing, and interactive presentation:')
p_arch_desc.runs[0].font.name = 'Calibri'
p_arch_desc.runs[0].font.size = Pt(10)

p_mermaid = doc.add_paragraph()
p_mermaid.paragraph_format.left_indent = Inches(0.2)
r_m = p_mermaid.add_run('''[Mermaid Architecture Diagram]
flowchart TD
    subgraph Data_Layer ["Data Tier (SQLite)"]
        A[("data/app.db\n(5 Relational Tables)")]
    end

    subgraph Ingestion_Layer ["Ingestion & Caching (src/data_loader.py)"]
        B["SQL Query Interface\n(Read-Only URI)"]
        C["st.cache_data Engine\n(TTL = 3600s, Zero-Copy)"]
    end

    subgraph Cleaning_Layer ["Validation & Transformation (src/data_cleaner.py)"]
        D["Schema Assertion Engine\n(Type Casting, Range Verification)"]
        E["Feature Derivation\n(Spell Grouping, Duration Hours)"]
    end

    subgraph Analytics_Layer ["Analytical & Statistical Tier (src/metrics.py & stats_engine.py)"]
        F["Bradford Factor Calculator\n(B = S^2 * D, Disruption Tiers)"]
        G["Financial Loss Engine\n(Wage * Hours * 1.20x Friction)"]
        H["Statistical Modeling\n(ANOVA, Pearson/Spearman, Chi-Square, Pareto)"]
    end

    subgraph UI_Presentation ["Presentation Tier (Streamlit Multipage + Tailwind)"]
        I["app.py\n(Executive Overview & KPI Cards)"]
        J["1_Departmental_DeepDive.py"]
        K["2_Bradford_Disruption.py"]
        L["3_Health_Lifestyle.py"]
        M["4_Temporal_Patterns.py"]
        N["5_Financial_Simulator.py"]
        O["st.components.v1.html\n(Tailwind CSS CDN KPI Cards)"]
        P["Plotly Visualizations\n(Interactive Sub-second Charts)"]
    end

    A --> B --> C --> D --> E
    E --> F & G & H
    F & G & H --> I & J & K & L & M & N
    I & J & K & L & M & N --> O & P''')
r_m.font.name = 'Courier New'
r_m.font.size = Pt(8.5)
r_m.font.color.rgb = RGBColor(43, 108, 176)

doc.add_paragraph()

# 2. Folder Structure
h2 = doc.add_heading('2. Project Repository & Folder Structure', level=1)
style_heading(h2)

p_tree = doc.add_paragraph()
r_tree = p_tree.add_run('''C:/HCL/ (Project Root)
├── .streamlit/
│   └── config.toml                  # UI theme colors, font families, telemetry disable
├── app.py                           # Application entrypoint & Executive Summary Dashboard
├── pages/
│   ├── 1_Departmental_DeepDive.py   # Department rankings, job role lost hours (BR-02)
│   ├── 2_Bradford_Disruption.py     # Bradford scoring, risk tiering, Pareto 80/20 (BR-05, BR-09)
│   ├── 3_Health_Lifestyle.py        # Medical ICD-10, BMI, habits, commute stats (BR-04, BR-06, BR-07)
│   ├── 4_Temporal_Patterns.py       # Day-of-week anomalies, holiday spikes (BR-08, BR-01)
│   └── 5_Financial_Simulator.py     # Direct wage loss, 1.20x friction, What-If slider (BR-03, BR-10)
├── src/
│   ├── __init__.py                  # Package initialization
│   ├── database.py                  # SQLite connection factory & DDL execution
│   ├── data_loader.py               # st.cache_data query abstractions & parameter filters
│   ├── data_cleaner.py              # Data type validation, missing-value assertions
│   ├── metrics.py                   # Deterministic metrics (Bradford B=S^2*D, payroll loss)
│   ├── stats_engine.py              # Statistical algorithms (ANOVA, Chi-Square, OLS, Gini)
│   ├── ui_components.py             # Tailwind CDN HTML iframe card generators
│   └── visualizers.py               # Plotly chart factories (heatmaps, scatter, bar, Pareto)
├── data/
│   ├── app.db                       # Pre-seeded SQLite production database (1,000+ records)
│   └── schema.sql                   # Relational DDL script with constraints and indexes
├── tests/
│   ├── __init__.py                  # Test suite runner
│   ├── test_database.py             # Schema constraint & foreign key integrity tests
│   ├── test_metrics.py              # Deterministic unit tests for B=S^2*D & financial loss
│   ├── test_stats_engine.py         # Statistical verification of ANOVA, Pearson, Chi-Square
│   └── test_app_ui.py               # Streamlit AppTest headless integration tests
├── docs/                            # 18 ABES University Deliverables (docx, xlsx, pptx)
│   ├── 01_Project_Understanding.docx
│   ├── 02_HMW_Statements.docx
│   ├── 03_Needs_Analysis.docx
│   ├── 04_Traceability_Matrix.xlsx
│   ├── 05_BRD.docx
│   └── 07_TRD.docx
├── requirements.txt                 # Pinned Python package dependencies
├── README.md                        # Setup, execution runbook & Cloud deployment guide
└── .gitignore                       # Git exclusions (cache, pycache, temporary artifacts)''')
r_tree.font.name = 'Courier New'
r_tree.font.size = Pt(8.5)

doc.add_paragraph()

# 3. Functional Requirements (FR-01 to FR-12) & Modules
h3 = doc.add_heading('3. Functional Requirements Matrix & Module Mapping', level=1)
style_heading(h3)

fr_table = doc.add_table(rows=13, cols=5)
fr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
fr_headers = ['FR ID', 'Functional System Capability', 'Linked BR', 'Target Python Module', 'Target UI View']
for j, h in enumerate(fr_headers):
    c = fr_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

fr_data = [
    ('FR-01', 'Calculate and plot monthly rolling absenteeism rate (%) against contracted hours.', 'BR-01', 'src/metrics.py', 'app.py, 4_Temporal_Patterns.py'),
    ('FR-02', 'Group, sort, and display total lost hours and incidence rates across departments and roles.', 'BR-02', 'src/metrics.py', '1_Departmental_DeepDive.py'),
    ('FR-03', 'Compute base salary loss and apply 1.20x friction multiplier across departments.', 'BR-03', 'src/metrics.py', 'app.py, 5_Financial_Simulator.py'),
    ('FR-04', 'Aggregate and plot absence hours across 21 ICD-10 medical chapters and tenure groups.', 'BR-04', 'src/stats_engine.py', '3_Health_Lifestyle.py'),
    ('FR-05', 'Compute Bradford score B = S^2 * D per employee and assign risk tiers (Low/Med/High/Sev).', 'BR-05', 'src/metrics.py', '2_Bradford_Disruption.py'),
    ('FR-06', 'Compute Pearson r, Spearman rho, and OLS regression between commute distance and absences.', 'BR-06', 'src/stats_engine.py', '3_Health_Lifestyle.py'),
    ('FR-07', 'Execute one-way ANOVA across BMI categories and plot mean absence differences.', 'BR-07', 'src/stats_engine.py', '3_Health_Lifestyle.py'),
    ('FR-08', 'Execute Chi-Square test on day-of-week distribution and flag Monday/Friday anomalies.', 'BR-08', 'src/stats_engine.py', '4_Temporal_Patterns.py'),
    ('FR-09', 'Construct cumulative Pareto concentration curve and compute Gini coefficient.', 'BR-09', 'src/metrics.py', '2_Bradford_Disruption.py'),
    ('FR-10', 'Provide dynamic slider (5%-20%) recalculating projected net annual payroll savings.', 'BR-10', 'src/metrics.py', '5_Financial_Simulator.py'),
    ('FR-11', 'Render sortable, filterable table of high-risk employees (Bradford >= 250).', 'BR-11', 'src/data_loader.py', '2_Bradford_Disruption.py'),
    ('FR-12', 'Generate downloadable CSV/Excel summary datasets containing filtered KPI rollups.', 'BR-12', 'src/ui_components.py', 'app.py, 5_Financial_Simulator.py')
]

for i, row in enumerate(fr_data, start=1):
    for j, val in enumerate(row):
        c = fr_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        if j == 0:
            r.bold = True
        elif j == 2:
            r.bold = True
            r.font.color.rgb = RGBColor(197, 48, 48)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

doc.add_paragraph()

# 4. Data Layer Architecture
h4 = doc.add_heading('4. Data Layer Architecture & Caching Strategy', level=1)
style_heading(h4)

p_dl = doc.add_paragraph()
p_dl.add_run('The application employs an embedded SQLite database engine (data/app.db). To ensure enterprise-grade speed and minimize disk I/O, database interactions follow strict optimization rules:')
p_dl.runs[0].font.name = 'Calibri'
p_dl.runs[0].font.size = Pt(10)

dl_points = [
    ('Connection Mode:', 'The application establishes read-only connection URIs: sqlite3.connect("file:data/app.db?mode=ro", uri=True) to prevent concurrency locks and accidental mutations during multi-user web browsing.'),
    ('In-Memory Query Caching:', 'All analytical queries are wrapped with Streamlit\'s @st.cache_data(ttl=3600, show_spinner=False). The query result is serialized into an in-memory Apache Arrow / Pandas buffer, reducing subsequent filter query latency from 80ms to < 2ms.'),
    ('Data Validation Pipeline:', 'Upon initial loading, src/data_cleaner.py validates schema assertions: verifies that employee_id is non-null, absent_hours is positive and bounded (0 < hrs <= 40 per spell), and foreign keys resolve without orphaned records.')
]
for title, desc in dl_points:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(9.5)
    r2 = p.add_run(desc)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(9.5)

doc.add_paragraph()

# 5. Statistical Methodology for Each OQ
h5 = doc.add_heading('5. Statistical & Algorithmic Methodology per OQ', level=1)
style_heading(h5)

stat_table = doc.add_table(rows=11, cols=4)
stat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
st_headers = ['OQ ID', 'Operational Inquiry', 'Mathematical / Statistical Algorithm', 'Python Library Function']
for j, h in enumerate(st_headers):
    c = stat_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

stat_methods = [
    ('OQ-01', 'Monthly absence rate trending', 'Time-series aggregation; Rate = (Sum(Absent_Hours) / Sum(Contracted_Hours)) * 100 with 3-month rolling MA.', 'pandas.DataFrame.rolling()'),
    ('OQ-02', 'Department & role hotspots', 'Grouped multi-level aggregation; Inter-Quartile Range (IQR) outlier fence (Q3 + 1.5 * IQR); Z-score normalization.', 'scipy.stats.zscore()'),
    ('OQ-03', 'Direct wage loss & replacement', 'Deterministic cost accounting: Loss = Sum(Hours * Wage) * 1.20; Bootstrapped 95% Confidence Interval.', 'numpy.sum(), statsmodels.stats'),
    ('OQ-04', 'Medical ICD-10 breakdown', 'Non-parametric Kruskal-Wallis H-test across tenure/age brackets; Discrete frequency distribution.', 'scipy.stats.kruskal()'),
    ('OQ-05', 'Bradford Factor disruption', 'Deterministic formula: B = (S^2) * D over 52-week rolling window; Binning into 4 ordinal risk categories.', 'pandas.cut(), numpy.square()'),
    ('OQ-06', 'Commute distance correlation', 'Bivariate Pearson correlation (r), Spearman rank correlation (rho), and Ordinary Least Squares (OLS) slope.', 'scipy.stats.pearsonr(), statsmodels.api.OLS()'),
    ('OQ-07', 'Biometric (BMI) ANOVA', 'One-way Analysis of Variance (ANOVA) across BMI bins; Post-hoc Tukey Honestly Significant Difference (HSD) test.', 'scipy.stats.f_oneway(), statsmodels.stats.multicomp'),
    ('OQ-08', 'Day-of-week anomaly detection', 'Chi-Square Goodness-of-Fit test (chi2) against uniform distribution (null hypothesis: equal absence probability).', 'scipy.stats.chisquare()'),
    ('OQ-09', 'Pareto 80/20 concentration', 'Lorenz curve integration; Cumulative percentage calculation; Gini Coefficient G = (A) / (A + B).', 'numpy.cumsum(), scipy.integrate.trapezoid()'),
    ('OQ-10', 'What-If financial simulation', 'Linear sensitivity model: Projected Savings = Total Loss * (Delta% / 100) * Preventable Fraction.', 'numpy.linspace(), interactive Plotly callback')
]

for i, row in enumerate(stat_methods, start=1):
    for j, val in enumerate(row):
        c = stat_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        if j == 0:
            r.bold = True
        elif j == 3:
            r.font.name = 'Courier New'
            r.font.size = Pt(8.5)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

doc.add_paragraph()

# 6. UI Engineering & Tailwind CDN Integration
h6 = doc.add_heading('6. UI Framework & Tailwind CSS CDN Integration', level=1)
style_heading(h6)

p_ui = doc.add_paragraph()
p_ui.add_run('To achieve executive presentation quality without heavy front-end node build steps, the UI integrates modern Tailwind CSS via Streamlit\'s HTML component iframe (st.components.v1.html):')
p_ui.runs[0].font.name = 'Calibri'
p_ui.runs[0].font.size = Pt(10)

ui_cards = [
    ('Tailwind CDN Isolation:', 'Metric cards and hero banners are injected via st.components.v1.html("""<script src="https://cdn.tailwindcss.com"></script><div class="bg-white p-6 rounded-xl shadow-md border border-slate-100 flex items-center justify-between">...</div>""", height=130). This provides glassmorphic styling, responsive flexbox layout, and custom typography without CSS leaking into Streamlit native controls.'),
    ('Color Palette & Visual Tokens:', 'Primary Navy (#1E3A8A / bg-blue-900), Accent Blue (#2563EB / bg-blue-600), Alert Red (#DC2626 / bg-red-600 for High/Severe Bradford), Success Green (#16A34A / bg-green-600 for low rates), Slate Neutral (#F8FAFC / bg-slate-50).'),
    ('Interactive Plotly Visualizations:', 'All charts render with plotly.graph_objects / plotly.express using transparent backgrounds (paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"), interactive hover templates, and custom color sequences aligning with Tailwind tokens.')
]
for title, desc in ui_cards:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(9.5)
    r2 = p.add_run(desc)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(9.5)

doc.add_paragraph()

# 7. Performance, Security & Error Handling
h7 = doc.add_heading('7. Performance, Security, and Error Handling', level=1)
style_heading(h7)

sec_table = doc.add_table(rows=4, cols=3)
sec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sc_headers = ['Engineering Dimension', 'Standard / Target Specification', 'Implementation Mechanism']
for j, h in enumerate(sc_headers):
    c = sec_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

sec_data = [
    ('Performance & Latency', 'Sub-2.0s filter response; memory usage < 800 MB on cloud.', 'st.cache_data for SQL queries, vectorization in NumPy/Pandas, pre-aggregated database indexes on foreign keys.'),
    ('Defensive Error Handling', 'Zero uncaught exceptions; elegant fallback state on empty filters.', 'Defensive try/except wrappers around database loading; st.warning("No records found for active filter selection") UI fallbacks.'),
    ('Data Security & Privacy', 'Zero PII ingestion; SQL injection prevention; read-only DB access.', 'Parametrized SQL queries (sqlite3 placeholders ?); strictly synthetic data; read-only SQLite URI in production.')
]

for i, row in enumerate(sec_data, start=1):
    for j, val in enumerate(row):
        c = sec_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

doc.add_paragraph()

# 8. requirements.txt (Pinned)
h8 = doc.add_heading('8. Pinned Dependency Specification (requirements.txt)', level=1)
style_heading(h8)

p_reqs = doc.add_paragraph()
r_rq = p_reqs.add_run('''# Core Web Application Framework
streamlit==1.42.0

# Numerical & Data Manipulation
pandas==2.2.3
numpy==2.1.3

# Statistical Analysis & Hypothesis Testing
scipy==1.14.1
statsmodels==0.14.4

# Interactive Visualizations
plotly==6.0.0

# Automated Testing Framework
pytest==8.3.4

# Excel & Word Artifact Generation
openpyxl==3.1.5
python-docx==1.2.0''')
r_rq.font.name = 'Courier New'
r_rq.font.size = Pt(9)
r_rq.font.color.rgb = RGBColor(26, 54, 93)

doc.add_paragraph()

# 9. Deployment Requirements
h9 = doc.add_heading('9. Cloud Deployment Specification', level=1)
style_heading(h9)

p_dep = doc.add_paragraph()
p_dep.add_run('The project is engineered for zero-cost public hosting on Streamlit Community Cloud:')
p_dep.runs[0].font.name = 'Calibri'
p_dep.runs[0].font.size = Pt(10)

dep_items = [
    ('Repository Host:', 'Public GitHub repository containing app.py, pages/, src/, data/app.db, and requirements.txt.'),
    ('Deployment Target:', 'Streamlit Community Cloud connected directly to the main branch via GitHub Webhook.'),
    ('Python Runtime:', 'Python 3.12 64-bit Linux runtime environment provided by Streamlit Cloud.'),
    ('Zero External Credentials:', 'Because SQLite is embedded directly in data/app.db inside the git repository, no external cloud database credentials or secrets management are required.'),
    ('Resource Limits:', 'Max RAM: 1.0 GB; Max CPU: 1 vCPU; Disk: 1.0 GB transient storage (Current footprint: ~45 MB total).')
]
for title, desc in dep_items:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(9.5)
    r2 = p.add_run(desc)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(9.5)

doc.add_paragraph()

# 10. Technical Risks & Mitigation
h10 = doc.add_heading('10. Technical Risk Assessment & Mitigation', level=1)
style_heading(h10)

tr_table = doc.add_table(rows=6, cols=5)
tr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
tr_headers = ['Risk ID', 'Technical Risk Description', 'Severity', 'Likelihood', 'Engineering Mitigation Strategy']
for j, h in enumerate(tr_headers):
    c = tr_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

tech_risks = [
    ('TR-01', 'Streamlit session reload causes slow chart re-computation on every filter widget change.', 'High', 'Medium', 'Apply @st.cache_data on database queries; pre-calculate Bradford rolling sums inside SQLite views.'),
    ('TR-02', 'Tailwind iframe height mismatch causes vertical scrollbars on executive KPI cards.', 'Low', 'Low', 'Set exact container pixel height (height=130) and apply overflow-hidden in Tailwind markup.'),
    ('TR-03', 'SQLite file locking conflicts if multiple concurrent users trigger write operations.', 'Medium', 'Low', 'Keep database connection strictly read-only (?mode=ro) in Streamlit runtime; write operations occur offline during Step 10 ETL.'),
    ('TR-04', 'Outlier values in synthetic dataset skew ANOVA F-statistics and Pearson correlation r.', 'Medium', 'Medium', 'Implement robust statistics alongside parametric tests (e.g., Spearman rank correlation alongside Pearson; median alongside mean).'),
    ('TR-05', 'Cloud deployment fails due to unpinned dependency version incompatibilities.', 'High', 'Low', 'Pin exact package versions in requirements.txt (e.g., streamlit==1.42.0, pandas==2.2.3, plotly==6.0.0).')
]

for i, row in enumerate(tech_risks, start=1):
    for j, val in enumerate(row):
        c = tr_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        if j == 0:
            r.bold = True
        elif j in [2, 3]:
            r.bold = True
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

# Save document to C:\HCL\07_TRD.docx
os.makedirs('C:/HCL', exist_ok=True)
p_trd = 'C:/HCL/07_TRD.docx'
doc.save(p_trd)
print('SUCCESSFULLY SAVED:', p_trd)
