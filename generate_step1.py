import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

doc = docx.Document()

# Set standard margins
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

# Title Block
p_title = doc.add_paragraph()
r_title = p_title.add_run('STEP 1: PROJECT UNDERSTANDING')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Absenteeism Impact Analysis: Workforce Productivity & Financial Loss Modeling')
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(43, 108, 176)
r_sub.font.italic = True

# Metadata Box Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project ID & Title:', 'PRJ-AIA-01 | Absenteeism Impact Analysis'),
    ('Domain:', 'Data Analytics (Workforce & HR Analytics)'),
    ('Student Name & Roll No:', 'Kalpana | Roll No: 2500320140071'),
    ('Academic Context:', 'MCA 2nd Year Data Analytics, ABES Engineering College'),
    ('Faculty Mentor:', 'Prof. Gunjan Agrawal'),
    ('Target Deliverable:', 'Python + Streamlit + SQLite Application Architecture')
]
for i, (k, v) in enumerate(meta_data):
    r = meta_table.rows[i]
    r.cells[0].paragraphs[0].add_run(k).bold = True
    r.cells[0].paragraphs[0].runs[0].font.name = 'Calibri'
    r.cells[0].paragraphs[0].runs[0].font.size = Pt(10)
    set_cell_background(r.cells[0], 'EDF2F7')
    set_cell_margins(r.cells[0], 60, 60, 100, 100)
    
    r.cells[1].paragraphs[0].add_run(v)
    r.cells[1].paragraphs[0].runs[0].font.name = 'Calibri'
    r.cells[1].paragraphs[0].runs[0].font.size = Pt(10)
    set_cell_background(r.cells[1], 'F7FAFC')
    set_cell_margins(r.cells[1], 60, 60, 100, 100)

doc.add_paragraph()

# 1. Problem Statement
h1 = doc.add_heading('1. Problem Statement', level=1)
h1.style.font.color.rgb = RGBColor(26, 54, 93)
p_prob = doc.add_paragraph()
p_prob.paragraph_format.left_indent = Inches(0.2)
r_prob = p_prob.add_run('"Unplanned employee absenteeism creates unmonitored operational bottlenecks, inflates direct payroll losses and replacement overtime costs, and impairs departmental productivity due to a lack of centralized, data-driven visibility into absenteeism drivers, employee risk profiles, and financial impact."')
r_prob.font.name = 'Calibri'
r_prob.font.size = Pt(11)
r_prob.font.italic = True
r_prob.font.bold = True

# 2. Intended Audience & Users
h2 = doc.add_heading('2. Intended Audience & Users', level=1)
h2.style.font.color.rgb = RGBColor(26, 54, 93)

aud_table = doc.add_table(rows=5, cols=4)
aud_table.alignment = WD_TABLE_ALIGNMENT.CENTER
aud_headers = ['User Role', 'Target Persona', 'Access Frequency', 'Primary Goal & Interaction']
for j, h in enumerate(aud_headers):
    c = aud_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 100, 100)

aud_data = [
    ('HR Operations Manager', 'Operational HR Lead', 'Bi-weekly / Monthly', 'Monitors enterprise absence trends, enforces leave compliance, and tracks high-risk Bradford Factor scores.'),
    ('Department Line Managers', 'Engineering, Sales, Ops Leads', 'Weekly / Ad-hoc', 'Assesses immediate team availability, plans shift backfills, and buffers project deadlines against absence risk.'),
    ('Chief Financial Officer (CFO)', 'Executive Leadership', 'Monthly / Quarterly', 'Reviews direct salary leakage, excess overtime costs, and models ROI on preventative wellness investments.'),
    ('Workforce Data Analyst', 'Primary System User (Kalpana)', 'Daily / Bi-weekly', 'Conducts statistical exploratory analysis, monitors correlation factors, validates data integrity, and updates ETL.')
]

for i, row in enumerate(aud_data, start=1):
    for j, val in enumerate(row):
        c = aud_table.rows[i].cells[j]
        c.paragraphs[0].add_run(val).font.size = Pt(9.5)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 60, 60, 80, 80)

doc.add_paragraph()

# 3. Stakeholder Analysis
h3 = doc.add_heading('3. Stakeholder Analysis', level=1)
h3.style.font.color.rgb = RGBColor(26, 54, 93)

sh_table = doc.add_table(rows=7, cols=6)
sh_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sh_headers = ['SH ID', 'Stakeholder', 'Role', 'Interest', 'Influence', 'Decision Made with Application']
for j, h in enumerate(sh_headers):
    c = sh_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 80, 80)

sh_data = [
    ('SH-01', 'Chief HR Officer (CHRO)', 'Executive Sponsor', 'High', 'High', 'Formulates organization-wide leave policies, hybrid flex policies, and budget for health programs.'),
    ('SH-02', 'HR Operations Manager', 'Business Process Owner', 'High', 'High', 'Implements attendance governance policies, triggers HR review meetings for chronic absentees.'),
    ('SH-03', 'Department Line Managers', 'Operational Supervisor', 'High', 'Medium', 'Adjusts project delivery timelines, manages cross-training, and re-allocates emergency shift loads.'),
    ('SH-04', 'CFO / Finance Controller', 'Fiscal Stakeholder', 'Medium', 'High', 'Sanctions departmental overtime budgets and audits financial loss attributable to unworked hours.'),
    ('SH-05', 'Occupational Health Lead', 'Workforce Wellness Officer', 'High', 'Medium', 'Designs targeted ergonomic, medical, or mental health interventions for critical hotspot departments.'),
    ('SH-06', 'Internal Data Analyst', 'Technical Lead & Builder', 'High', 'Medium', 'Optimizes analytical models, enhances Streamlit dashboard views, and maintains SQLite database integrity.')
]

for i, row in enumerate(sh_data, start=1):
    for j, val in enumerate(row):
        c = sh_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.size = Pt(9.5)
        if j in [3, 4]:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 60, 60, 80, 80)

doc.add_paragraph()

# 4. Operational Questions (OQs) & Data Viability Verification
h4 = doc.add_heading('4. Operational Questions (OQs) & Data Viability Verification', level=1)
h4.style.font.color.rgb = RGBColor(26, 54, 93)

oq_table = doc.add_table(rows=11, cols=4)
oq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
oq_headers = ['OQ ID', 'Operational Question to be Answered', 'Data Viability & Underlying Attributes', 'Answerable?']
for j, h in enumerate(oq_headers):
    c = oq_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 80, 80)

oqs = [
    ('OQ-01', 'What is the gross organizational absenteeism rate, and how does it fluctuate month-over-month?', 'Employee attendance timestamps, total hours absent, total contracted hours across periods.', 'YES'),
    ('OQ-02', 'Which departments and specific job roles contribute most heavily to lost productive hours?', 'Department ID, Department Name, Job Title, Absence Hours per occurrence.', 'YES'),
    ('OQ-03', 'What is the direct monetary wage loss incurred per department based on individual hourly compensation?', 'Employee Hourly Wage, Hours Absent, Department Categorization, Overtime Multiplier.', 'YES'),
    ('OQ-04', 'What are the top primary medical and non-medical reasons driving workforce absenteeism?', 'ICD-10 Reason Codes (e.g., Musculoskeletal, Dental, Consultations, Stress/Burnout).', 'YES'),
    ('OQ-05', 'Which individual employees generate high operational friction as flagged by the Bradford Factor index?', 'Employee ID, Spell Count (S = distinct instances), Total Absent Days (D), B = S^2 * D.', 'YES'),
    ('OQ-06', 'Is there a measurable correlation between employee commute distance / transit cost and absence frequency?', 'Distance to Work (km), Daily Transportation Cost, Attendance Date, Day of Week.', 'YES'),
    ('OQ-07', 'How do lifestyle and demographic factors (BMI, Age, Smoker/Drinker status, Dependents) correlate with absence duration?', 'Age, BMI Index, Smoker Flag, Drinker Flag, Number of Dependents, Total Absence Duration.', 'YES'),
    ('OQ-08', 'What temporal patterns emerge regarding day-of-the-week (e.g., Monday/Friday spikes) and seasonal cycles?', 'Calendar Date, Day of Week (Monday to Friday), Month of Year, Quarter.', 'YES'),
    ('OQ-09', 'Does the cumulative absence distribution exhibit a Pareto principle (80/20 rule) concentration?', 'Ranked cumulative percentage of hours absent versus cumulative percentage of employees.', 'YES'),
    ('OQ-10', 'What projected financial savings can be achieved under simulated 5% to 20% absence reduction scenarios?', 'Interactive baseline cost aggregation, scenario percentage reduction slider, net cost delta.', 'YES')
]

for i, row in enumerate(oqs, start=1):
    for j, val in enumerate(row):
        c = oq_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.size = Pt(9.5)
        if j == 3:
            r.bold = True
            r.font.color.rgb = RGBColor(40, 167, 69)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 60, 60, 80, 80)

doc.add_paragraph()

# 5. SMART Objectives
h5 = doc.add_heading('5. Project Objectives (SMART)', level=1)
h5.style.font.color.rgb = RGBColor(26, 54, 93)

objs = [
    ('OBJ-01 (Specific & Measurable):', 'Design and implement a normalized SQLite relational database by Step 8 containing 1,000+ synthetic records across 5 core entities with zero referential integrity violations.'),
    ('OBJ-02 (Measurable & Achievable):', 'Build an interactive, modular Streamlit web application providing under-2-second query response times for dynamic multi-parameter filtering (by department, date, demographic, and reason category).'),
    ('OBJ-03 (Relevant & Specific):', 'Implement automated calculation and tier classification for the Bradford Factor index across 100% of employee profiles to categorize operational disruption into Low, Medium, High, and Severe categories.'),
    ('OBJ-04 (Realistic & Time-bound):', 'Model total direct payroll leakage and replacement cost impact with interactive Plotly visual charts and dynamic "What-If" cost mitigation simulation capabilities by Step 15.'),
    ('OBJ-05 (Measurable & Quality-driven):', 'Achieve a minimum 90% pytest unit and functional test coverage across data transformation, metric calculation, and database transaction modules prior to Step 16 deployment.')
]

for title, desc in objs:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(10)
    r2 = p.add_run(desc)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(10)

# 6. Project Scope
h6 = doc.add_heading('6. Project Scope: In-Scope vs. Out-of-Scope', level=1)
h6.style.font.color.rgb = RGBColor(26, 54, 93)

scope_table = doc.add_table(rows=6, cols=2)
scope_table.alignment = WD_TABLE_ALIGNMENT.CENTER
scope_headers = ['In-Scope (Delivered)', 'Out-of-Scope (Excluded Boundaries)']
for j, h in enumerate(scope_headers):
    c = scope_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 100, 100)

scope_data = [
    ('Synthetic enterprise workforce dataset generation (1,000+ realistic records mirroring demographics, absences, wages, and reasons).', 'Ingestion of real personal identifiable information (PII) or confidential corporate payroll databases.'),
    ('Relational schema design and automated SQLite database initialization with referential integrity constraints.', 'Enterprise distributed databases (Oracle, PostgreSQL clusters, Apache Cassandra).'),
    ('Comprehensive Exploratory Data Analysis (EDA), statistical correlation, distribution modeling, and Bradford Factor scoring.', 'Integration with physical biometric IoT fingerprint/face hardware or RFID turnstiles.'),
    ('Interactive Streamlit dashboard with Tailwind-inspired visual hierarchy, KPI stat cards, and dynamic Plotly charts.', 'Automated direct bank payroll salary disbursement or legally binding disciplinary actions.'),
    ('Pytest automated test suites, deployment to Streamlit Community Cloud, and complete ABES-compliant documentation.', 'Complex Deep Learning computer vision or neural network natural language interfaces.')
]

for i, (ins, outs) in enumerate(scope_data, start=1):
    c1 = scope_table.rows[i].cells[0]
    c2 = scope_table.rows[i].cells[1]
    c1.paragraphs[0].add_run(ins).font.size = Pt(9.5)
    c2.paragraphs[0].add_run(outs).font.size = Pt(9.5)
    bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
    set_cell_background(c1, bg)
    set_cell_background(c2, bg)
    set_cell_margins(c1, 60, 60, 80, 80)
    set_cell_margins(c2, 60, 60, 80, 80)

doc.add_paragraph()

# 7. Assumptions & Constraints
h7 = doc.add_heading('7. Assumptions and Constraints', level=1)
h7.style.font.color.rgb = RGBColor(26, 54, 93)

ac_table = doc.add_table(rows=7, cols=3)
ac_table.alignment = WD_TABLE_ALIGNMENT.CENTER
ac_headers = ['ID', 'Classification', 'Description & Rationale']
for j, h in enumerate(ac_headers):
    c = ac_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 100, 100)

acs = [
    ('A-01', 'Assumption', 'Standard work schedules assume 8 productive hours per day, 5 days per week (40 hours/week, 2,080 standard hours/year).'),
    ('A-02', 'Assumption', 'Financial impact calculations include base hourly wage loss plus a standardized 20% friction multiplier representing productivity disruption and replacement overtime overhead.'),
    ('A-03', 'Assumption', 'Medical absence reasons are mapped to standardized categories based on the International Classification of Diseases (ICD-10) adapted from empirical workforce research.'),
    ('C-01', 'Constraint', 'Technology stack is restricted to open-source Python, Streamlit, Plotly, Pandas, SQLite3, and Pytest to ensure zero licensing expenses.'),
    ('C-02', 'Constraint', 'Application runtime memory footprint must not exceed 1.0 GB to adhere to free-tier hosting limits on Streamlit Community Cloud.'),
    ('C-03', 'Constraint', 'All developmental phases and academic deliverables must strictly align with the MCA 2nd Year academic submission schedule at ABES Engineering College.')
]

for i, (cid, ctype, cdesc) in enumerate(acs, start=1):
    c0 = ac_table.rows[i].cells[0]
    c1 = ac_table.rows[i].cells[1]
    c2 = ac_table.rows[i].cells[2]
    c0.paragraphs[0].add_run(cid).bold = True
    c1.paragraphs[0].add_run(ctype)
    c2.paragraphs[0].add_run(cdesc)
    for c in [c0, c1, c2]:
        c.paragraphs[0].runs[0].font.size = Pt(9.5)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 60, 60, 80, 80)

doc.add_paragraph()

# 8. Measurable Success Criteria
h8 = doc.add_heading('8. Measurable Success Criteria', level=1)
h8.style.font.color.rgb = RGBColor(26, 54, 93)

crit_table = doc.add_table(rows=6, cols=3)
crit_table.alignment = WD_TABLE_ALIGNMENT.CENTER
crit_headers = ['Metric ID', 'Success Evaluation Metric', 'Target Threshold for Acceptance']
for j, h in enumerate(crit_headers):
    c = crit_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 100, 100)

crits = [
    ('SC-01', 'OQ Functional Coverage', '100% of defined Operational Questions (OQ-01 to OQ-10) answered with dedicated dynamic charts and statistical summaries.'),
    ('SC-02', 'UI & Query Response Latency', 'Under 2.0 seconds latency for complete dashboard filter refreshes across 1,000+ records.'),
    ('SC-03', 'Automated Test Suite Coverage', '>= 90% pytest code coverage with 0 failing test cases across all business logic and calculations.'),
    ('SC-04', 'Cloud Deployment Accessibility', '100% functional live web URL on Streamlit Community Cloud with public accessibility and 0 crash errors.'),
    ('SC-05', 'Documentation & Academic Rigor', 'Complete set of 18 ABES MCA standard deliverables formatted to university specifications and ready for viva defense.')
]

for i, (mid, metric, target) in enumerate(crits, start=1):
    c0 = crit_table.rows[i].cells[0]
    c1 = crit_table.rows[i].cells[1]
    c2 = crit_table.rows[i].cells[2]
    c0.paragraphs[0].add_run(mid).bold = True
    c1.paragraphs[0].add_run(metric)
    c2.paragraphs[0].add_run(target)
    for c in [c0, c1, c2]:
        c.paragraphs[0].runs[0].font.size = Pt(9.5)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 60, 60, 80, 80)

# Save document
os.makedirs('C:/HCL', exist_ok=True)
p1 = 'C:/HCL/01_name_project_understa.docx'
p2 = 'C:/HCL/01_Project_Understanding.docx'
doc.save(p1)
doc.save(p2)
print('SUCCESS')
