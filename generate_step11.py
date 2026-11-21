import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import openpyxl
from openpyxl.styles import Font, Alignment
import os

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
r_title = p_title.add_run('AGILE USER STORIES & ACCEPTANCE CRITERIA')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Sprint Planning & Product Backlog Specification | Absenteeism Impact Analysis')
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(43, 108, 176)
r_sub.font.italic = True

# Metadata Box Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Document Reference:', 'US-PRJ-AIA-01-v1.0'),
    ('Project ID & Title:', 'PRJ-AIA-01 | Absenteeism Impact Analysis'),
    ('Student Product Owner:', 'Kalpana (Roll No: 2500320140071) — MCA 2nd Yr DA, ABES'),
    ('Faculty Project Mentor:', 'Prof. Gunjan Agrawal'),
    ('Agile Methodology:', 'Scrum Framework (2-Week Iterative Sprint Execution)'),
    ('Total User Stories / Points:', '16 Stories | 41 Total Story Points (Fibonacci Scale 1-5)')
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

# 1. Epics Overview
h1 = doc.add_heading('1. Product Epics Overview', level=1)
style_heading(h1)

p_ep_intro = doc.add_paragraph()
p_ep_intro.add_run('The product backlog is decomposed into five thematic Epics spanning the engineering lifecycle:')
p_ep_intro.runs[0].font.name = 'Calibri'
p_ep_intro.runs[0].font.size = Pt(10)

epics_info = [
    ('EPIC-01: Data Layer & Persistence (Data):', 'Covers SQLite schema creation, referential constraints, indexing, and the unified analytical SQL view.'),
    ('EPIC-02: Executive UI & Departmental Dashboards (Dashboard):', 'Covers the Streamlit multipage application framework, Tailwind KPI cards, and departmental drill-down charts.'),
    ('EPIC-03: Diagnostic Statistics & Behavioral Modeling (Analysis):', 'Encompasses Bradford Factor disruption scoring, Pareto concentration curves, commute correlations, and BMI ANOVA models.'),
    ('EPIC-04: Financial Loss Accounting & ROI Simulation (Insights):', 'Delivers direct wage loss calculation, replacement overtime friction multipliers, and What-If scenario forecasting.'),
    ('EPIC-05: Quality Engineering & Regression Testing (Quality):', 'Implements automated pytest suites, Streamlit AppTest integration validations, and cloud deployment pipelines.')
]
for title, desc in epics_info:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = RGBColor(26, 54, 93)
    r2 = p.add_run(desc)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(9.5)

doc.add_paragraph()

# 2. User Stories Master Matrix
h2 = doc.add_heading('2. User Stories Backlog Table (US-01 to US-16)', level=1)
style_heading(h2)

stories_table = doc.add_table(rows=17, cols=7)
stories_table.alignment = WD_TABLE_ALIGNMENT.CENTER
st_headers = ['US ID', 'Persona', 'User Story Statement (As a... I want... So that...)', 'Feature ID', 'BR ID', 'Priority', 'Points']
for j, h in enumerate(st_headers):
    c = stories_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 60, 60)

user_stories_data = [
    ('US-01', 'Data Analyst', 'As a Data Analyst, I want an initialized SQLite database with active foreign keys and indexes so that queries execute with referential integrity.', 'F-01', 'BR-01', 'Must', '3'),
    ('US-02', 'Data Analyst', 'As a Data Analyst, I want a pre-joined analytical view (v_absence_analysis_joined) with 32 attributes so that application queries do not require repetitive multi-table joins.', 'F-01', 'BR-01, BR-03', 'Must', '2'),
    ('US-03', 'HR Operations Manager', 'As an HR Operations Manager, I want an executive summary dashboard with modern Tailwind KPI cards displaying total lost hours, gross financial loss, and absence rate so that I can immediately gauge workforce health.', 'F-01, F-12', 'BR-01, BR-12', 'Must', '3'),
    ('US-04', 'Department Line Manager', 'As a Department Line Manager, I want a ranked visual comparison of absence hours and incidents across departments and job profiles so that I can detect understaffed operational bottlenecks.', 'F-02', 'BR-02', 'Must', '3'),
    ('US-05', 'Department Line Manager', 'As a Department Line Manager, I want an anomaly chart highlighting Monday and Friday absence surges so that I can identify habitual weekend extensions and adjust shift staffing.', 'F-08', 'BR-08', 'Must', '3'),
    ('US-06', 'HR Operations Manager', 'As an HR Operations Manager, I want an interactive sidebar filter (department, date range, medical category) so that all dashboard views update in sub-2 seconds for targeted investigation.', 'F-01', 'BR-01, BR-02', 'Must', '2'),
    ('US-07', 'HR Operations Manager', 'As an HR Operations Manager, I want automated calculation and risk tiering of the Bradford Factor (B = S^2 * D) so that I can objectively identify disruptive chronic short-term casual absentees.', 'F-05', 'BR-05', 'Must', '3'),
    ('US-08', 'HR Operations Manager', 'As an HR Operations Manager, I want a cumulative Pareto 80/20 concentration curve so that I can demonstrate whether a small minority of employees accounts for the majority of lost company hours.', 'F-09', 'BR-09', 'Must', '3'),
    ('US-09', 'Occupational Health Lead', 'As an Occupational Health Lead, I want a diagnostic breakdown of absence hours across 21 ICD-10 medical chapters and employee age cohorts so that I can deploy preventative wellness interventions.', 'F-04', 'BR-04', 'Must', '3'),
    ('US-10', 'Occupational Health Lead', 'As an Occupational Health Lead, I want scatter plots and correlation statistics between commute distance/transit cost and absence frequency so that I can evaluate the health benefits of hybrid work policies.', 'F-06', 'BR-06', 'Should', '3'),
    ('US-11', 'Occupational Health Lead', 'As an Occupational Health Lead, I want an ANOVA comparison of absence duration across BMI categories (Normal, Overweight, Obese) so that I can target nutritional and ergonomic campaigns.', 'F-07', 'BR-07', 'Should', '3'),
    ('US-12', 'CFO / Finance Controller', 'As a CFO, I want a breakdown of direct payroll loss and secondary 1.20x replacement overtime expenses per cost center so that I can audit financial leakage.', 'F-03', 'BR-03', 'Must', '3'),
    ('US-13', 'CFO / CHRO', 'As an Executive Sponsor, I want an interactive What-If cost savings simulator with an adjustable 5% to 20% reduction slider so that I can evaluate the projected ROI of proposed wellness budgets.', 'F-10', 'BR-10', 'Should', '3'),
    ('US-14', 'HR Operations Manager', 'As an HR Operations Manager, I want to filter and inspect a dedicated roster of high-risk employees (Bradford >= 250) so that I can initiate formal attendance review meetings.', 'F-11', 'BR-11', 'Could', '2'),
    ('US-15', 'HR Analyst', 'As an HR Analyst, I want to export filtered KPI summary tables and departmental rankings to downloadable CSV and Excel files so that I can present offline findings in executive reviews.', 'F-12', 'BR-12', 'Could', '2'),
    ('US-16', 'QA Engineer', 'As a QA Engineer, I want automated pytest test suites validating database schema rules, metric algorithms, and Streamlit AppTest UI so that regression defects are caught before deployment.', 'F-All', 'BR-01 to BR-12', 'Must', '3')
]

for i, row in enumerate(user_stories_data, start=1):
    for j, val in enumerate(row):
        c = stories_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(8.5)
        if j in [0, 6]:
            r.bold = True
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif j == 5:
            r.bold = True
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if val == 'Must':
                r.font.color.rgb = RGBColor(197, 48, 48)
            elif val == 'Should':
                r.font.color.rgb = RGBColor(43, 108, 176)
            elif val == 'Could':
                r.font.color.rgb = RGBColor(40, 167, 69)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 40, 40, 50, 50)

doc.add_paragraph()

# 3. Acceptance Criteria (Given / When / Then)
h3 = doc.add_heading('3. Acceptance Criteria per Story (Given / When / Then)', level=1)
style_heading(h3)

criteria_details = [
    ('US-01: SQLite Schema & Foreign Key Initialization',
     'Given the data directory is initialized,\n'
     'When src/database.py executes schema.sql with PRAGMA foreign_keys = ON,\n'
     'Then 4 tables and 4 indexes are created, and attempting to insert an invalid department_id raises an sqlite3.IntegrityError.'),

    ('US-02: Unified Analytical SQL View Creation',
     'Given the SQLite database is populated with production records,\n'
     'When a query executes SELECT * FROM v_absence_analysis_joined,\n'
     'Then 32 enriched attributes are returned, and direct_wage_loss matches absent_hours * hourly_wage with zero null values.'),

    ('US-03: Executive Overview Dashboard & Tailwind KPI Cards',
     'Given the user launches the Streamlit application,\n'
     'When the home page app.py renders,\n'
     'Then 4 isolated Tailwind iframe cards display Total Lost Hours (3,656.0 hrs), Total Financial Loss ($212,586.26), Overall Absence Rate (1.92%), and High-Risk Bradford Employee Count (14) with responsive layout.'),

    ('US-04: Departmental Hotspots & Role Rankings',
     'Given the user navigates to pages/1_Departmental_DeepDive.py,\n'
     'When the departmental analytics tab renders,\n'
     'Then an interactive Plotly horizontal bar chart ranks departments by total lost hours, with Engineering and Operations highlighted at the top.'),

    ('US-05: Monday / Friday Anomaly Detection',
     'Given the user navigates to pages/4_Temporal_Patterns.py,\n'
     'When the day-of-week distribution chart loads,\n'
     'Then Monday and Friday absence bars show a statistically significant surge (+26.4%) over the Tuesday-Thursday average, with a Chi-Square test indicator showing p < 0.05.'),

    ('US-06: Interactive Sidebar Filtering',
     'Given any analytical dashboard page is active,\n'
     'When the user selects "Engineering" in the department filter,\n'
     'Then all metric cards, data tables, and Plotly charts update in < 2.0 seconds displaying exclusively Engineering records.'),

    ('US-07: Bradford Factor Disruption Scoring & Risk Tiering',
     'Given the user opens pages/2_Bradford_Disruption.py,\n'
     'When the Bradford scoring engine evaluates all 120 employees,\n'
     'Then every employee is classified into Low (<50), Medium (50-249), High (250-749), or Severe (>=750) tiers, with the severe tier highlighting high-frequency short-term absentees.'),

    ('US-08: Pareto 80/20 Concentration Curve',
     'Given the Bradford analytics view is active,\n'
     'When the user views the Pareto distribution tab,\n'
     'Then a cumulative percentage Lorenz curve renders demonstrating that the top ~12% of employees account for 55.8% of aggregate lost hours.'),

    ('US-09: ICD-10 Medical Reason Chapter Breakdown',
     'Given the user navigates to pages/3_Health_Lifestyle.py,\n'
     'When the medical analysis tab renders,\n'
     'Then a treemap / bar chart displays hours grouped into 21 ICD-10 chapters and 7 administrative categories, with Musculoskeletal (M00-M99) and Respiratory (J00-J99) displayed as dominant causes.'),

    ('US-10: Commute Distance & Transit Expenditure Correlation',
     'Given the commute analysis view is loaded,\n'
     'When the user inspects the commute scatter plot,\n'
     'Then an OLS regression trendline displays alongside calculated Pearson r and Spearman rho coefficients demonstrating higher absence frequency among employees commuting > 30 km.'),

    ('US-11: Biometric BMI ANOVA Hypothesis Testing',
     'Given the biometric health view is active,\n'
     'When the user initiates the BMI ANOVA comparison,\n'
     'Then a boxplot renders across Underweight, Normal, Overweight, and Obese cohorts, with F-statistic and p-value reported indicating statistically significant variance in mean absence duration.'),

    ('US-12: Direct Wage Loss & 1.20x Overtime Accounting',
     'Given the user opens pages/5_Financial_Simulator.py,\n'
     'When the financial breakdown renders,\n'
     'Then the interface displays both direct base payroll loss ($177,155.36) and total loss with the 1.20x friction multiplier ($212,586.26), itemized by departmental cost center.'),

    ('US-13: Interactive What-If Cost Mitigation Simulator',
     'Given the financial simulator tab is open,\n'
     'When the user drags the reduction slider from 0% to 15%,\n'
     'Then the projected annual cost savings metric recalculates in real-time displaying estimated savings of > $25,000 against preventable medical absence categories.'),

    ('US-14: High-Risk Employee Roster Filtering',
     'Given the user is on the Bradford Factor page,\n'
     'When the user toggles the "Show High & Severe Risk Only (B >= 250)" filter,\n'
     'Then the employee table filters to display exactly the 14 flagged individuals with their badge ID, name, spell count, days absent, and Bradford score.'),

    ('US-15: Executive Summary Report Export',
     'Given the user applies a set of filters on the dashboard,\n'
     'When the user clicks the "Download Executive Summary (CSV)" button,\n'
     'Then a clean, formatted CSV file containing the active aggregated KPIs and departmental breakdown downloads immediately.'),

    ('US-16: Automated Pytest Regression Suite',
     'Given developer runs pytest tests/ in the terminal,\n'
     'When test suites execute across test_database.py, test_metrics.py, test_stats_engine.py, and test_app_ui.py,\n'
     'Then all test cases pass with 0 errors and >= 90% code coverage.')
]

for us_title, gwt in criteria_details:
    p_u = doc.add_paragraph()
    r_u = p_u.add_run(us_title)
    r_u.bold = True
    r_u.font.name = 'Calibri'
    r_u.font.size = Pt(10.5)
    r_u.font.color.rgb = RGBColor(26, 54, 93)
    
    p_gwt = doc.add_paragraph()
    p_gwt.paragraph_format.left_indent = Inches(0.2)
    r_gwt = p_gwt.add_run(gwt)
    r_gwt.font.name = 'Calibri'
    r_gwt.font.size = Pt(9.5)

doc.add_paragraph()

# 4. Two-Week Sprint Plan
h4 = doc.add_heading('4. Two-Week Agile Sprint Execution Plan', level=1)
style_heading(h4)

sprint_table = doc.add_table(rows=11, cols=4)
sprint_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sp_headers = ['Sprint Day', 'Target User Stories Scheduled', 'Epic & Focus Area', 'Deliverable Milestone']
for j, h in enumerate(sp_headers):
    c = sprint_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

sprint_plan_data = [
    ('Day 1 (Sprint 1)', 'US-01, US-02 (5 pts)', 'Epic 1: Data', 'SQLite DDL, tables, indexes, and v_absence_analysis_joined operational.'),
    ('Day 2 (Sprint 1)', 'US-03, US-06 (5 pts)', 'Epic 2: Dashboard', 'app.py skeleton with Tailwind KPI cards and interactive filter sidebar.'),
    ('Day 3 (Sprint 1)', 'US-04, US-05 (6 pts)', 'Epic 2: Dashboard', 'pages/1_Departmental_DeepDive.py & pages/4_Temporal_Patterns.py built.'),
    ('Day 4 (Sprint 1)', 'US-07, US-08 (6 pts)', 'Epic 3: Analysis', 'pages/2_Bradford_Disruption.py with B = S^2 * D and Pareto 80/20 curves.'),
    ('Day 5 (Sprint 1)', 'US-14 (2 pts) + Review', 'Epic 3: Analysis', 'High-risk employee threshold filtering; Sprint 1 Demonstration & Review.'),
    ('Day 6 (Sprint 2)', 'US-09, US-10 (6 pts)', 'Epic 3: Analysis', 'pages/3_Health_Lifestyle.py with ICD-10 breakdown & commute correlation.'),
    ('Day 7 (Sprint 2)', 'US-11 (3 pts)', 'Epic 3: Analysis', 'One-way ANOVA on BMI categories and post-hoc Tukey statistical modeling.'),
    ('Day 8 (Sprint 2)', 'US-12, US-13 (6 pts)', 'Epic 4: Insights', 'pages/5_Financial_Simulator.py with 1.20x loss & interactive What-If slider.'),
    ('Day 9 (Sprint 2)', 'US-15 (2 pts)', 'Epic 4: Insights', 'Exportable executive summary CSV/Excel download handlers implemented.'),
    ('Day 10 (Sprint 2)', 'US-16 (3 pts) + Release', 'Epic 5: Quality', 'Automated pytest test suite execution (>=90% coverage); Cloud Deployment.')
]

for i, row in enumerate(sprint_plan_data, start=1):
    for j, val in enumerate(row):
        c = sprint_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        if j in [0, 1]:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

# Save Word documents
os.makedirs('C:/HCL/docs', exist_ok=True)
p_us_docs = 'C:/HCL/docs/11_User_Stories.docx'
p_us_root = 'C:/HCL/11_User_Stories.docx'
doc.save(p_us_docs)
doc.save(p_us_root)
print('SUCCESSFULLY SAVED USER STORIES DOCX:', p_us_docs)

# ========================================================
# UPDATE RTM WITH USER STORY IDs
# ========================================================
rtm_path = 'C:/HCL/04_Traceability_Matrix.xlsx'
wb_rtm = openpyxl.load_workbook(rtm_path)
ws_rtm = wb_rtm['Traceability_Matrix']

# Mapping of BR ID to User Story ID
us_mapping = {
    'BR-01': 'US-01, US-02, US-03, US-06',
    'BR-12': 'US-03, US-15',
    'BR-02': 'US-04, US-06',
    'BR-11': 'US-14',
    'BR-03': 'US-02, US-12',
    'BR-04': 'US-09',
    'BR-05': 'US-07',
    'BR-06': 'US-10',
    'BR-07': 'US-11',
    'BR-08': 'US-05',
    'BR-09': 'US-08',
    'BR-10': 'US-13'
}

for row in range(2, ws_rtm.max_row + 1):
    br_id = ws_rtm.cell(row=row, column=3).value
    if br_id in us_mapping:
        c_us = ws_rtm.cell(row=row, column=10) # Column 10: User Story
        c_us.value = us_mapping[br_id]
        c_us.font = Font(name='Courier New', size=8.5, color='1A365D')
        c_us.alignment = Alignment(horizontal='center', vertical='center')

# Update Coverage_Analysis sheet in RTM
ws_cov = wb_rtm['Coverage_Analysis']
for r in range(1, ws_cov.max_row + 1):
    v = ws_cov.cell(row=r, column=1).value
    if v and 'Step 11' in str(v):
        c_stat = ws_cov.cell(row=r, column=4)
        c_stat.value = 'COMPLETED (Mapped)'
        c_stat.font = Font(name='Calibri', size=9.5, bold=True, color='28A745')

wb_rtm.save(rtm_path)
print('SUCCESSFULLY UPDATED RTM WITH USER STORIES:', rtm_path)
