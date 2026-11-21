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

# ==================== COVER / HEADER ====================
p_title = doc.add_paragraph()
r_title = p_title.add_run('BUSINESS REQUIREMENTS DOCUMENT (BRD)')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(24)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Project: Absenteeism Impact Analysis — Workforce Productivity & Financial Loss Modeling')
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(43, 108, 176)
r_sub.font.italic = True

# 1. Document Control Table
h1 = doc.add_heading('1. Document Control', level=1)
style_heading(h1)

ctrl_table = doc.add_table(rows=7, cols=2)
ctrl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
ctrl_data = [
    ('Document Reference:', 'BRD-PRJ-AIA-01-v1.0'),
    ('Document Version:', 'Version 1.0 (Formal Baseline)'),
    ('Primary Author:', 'Kalpana (Roll No: 2500320140071) — MCA 2nd Yr DA'),
    ('Institution & Department:', 'Department of Computer Applications, ABES Engineering College'),
    ('Faculty Mentor & Reviewer:', 'Prof. Gunjan Agrawal'),
    ('Baseline Date:', 'November 20, 2026'),
    ('Document Status:', 'Approved for Technical & Architectural Formulation')
]
for i, (k, v) in enumerate(ctrl_data):
    r = ctrl_table.rows[i]
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

# 2. Executive Summary
h2 = doc.add_heading('2. Executive Summary', level=1)
style_heading(h2)
p_exec = doc.add_paragraph()
p_exec.add_run('Employee absenteeism represents one of the largest unmonitored operational liabilities in enterprise workforce management, contributing directly to idle payroll expenditure, replacement overtime friction, and departmental delivery bottlenecks. The Absenteeism Impact Analysis system (PRJ-AIA-01) delivers an end-to-end analytical solution designed to transform unstructured attendance logs into actionable business intelligence. By aggregating absence incidence, categorizing medical and lifestyle determinants, automating Bradford Factor disruption scoring, and dynamically modeling bottom-line financial losses, the platform empowers HR executives, operational managers, and finance leaders to transition from reactive administrative tracking to proactive, data-grounded workforce optimization.')
p_exec.runs[0].font.name = 'Calibri'
p_exec.runs[0].font.size = Pt(10)

# 3. Business Problem and Opportunity
h3 = doc.add_heading('3. Business Problem and Opportunity', level=1)
style_heading(h3)

p_prob = doc.add_paragraph()
r_pr_title = p_prob.add_run('The Problem: ')
r_pr_title.bold = True
r_pr_title.font.color.rgb = RGBColor(197, 48, 48)
p_prob.add_run('Currently, absenteeism is treated as an isolated HR record rather than a strategic operational metric. Organizations lack centralized visibility into how frequent short-term leaves disrupt production schedules compared to isolated long-term medical absences. Furthermore, financial controllers possess no granular model to quantify the compounding fiscal impact of direct salary waste and the secondary overtime costs paid to replacement staff.')

p_opp = doc.add_paragraph()
r_op_title = p_opp.add_run('The Opportunity: ')
r_op_title.bold = True
r_op_title.font.color.rgb = RGBColor(40, 167, 69)
p_opp.add_run('Implementing a dedicated diagnostic workforce analytics platform allows the organization to identify departmental hotspots, detect Friday/Monday casual absence spikes, isolate high-risk disruption profiles using the Bradford Factor index, and simulate prospective ROI on targeted employee wellness and ergonomic programs.')

# 4. Objectives and Success Metrics
h4 = doc.add_heading('4. Objectives and Success Metrics (KPIs)', level=1)
style_heading(h4)

obj_table = doc.add_table(rows=6, cols=3)
obj_table.alignment = WD_TABLE_ALIGNMENT.CENTER
obj_headers = ['Objective ID', 'Formal Business Objective', 'Target KPI / Acceptance Threshold']
for j, h in enumerate(obj_headers):
    c = obj_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

objs_data = [
    ('OBJ-01', 'Consolidate 1,000+ synthetic records into a normalized SQLite relational schema.', '100% data integrity with 0 orphaned records or referential constraint errors.'),
    ('OBJ-02', 'Deliver interactive web-based exploratory data slice and dice capability.', 'Sub-2-second query and visual rendering latency across all filter dimensions.'),
    ('OBJ-03', 'Automate Bradford Factor scoring and disruption tier classification for all employees.', '100% automated calculation of B = S^2 * D with 4 standardized risk tiers.'),
    ('OBJ-04', 'Model direct salary leakage and secondary overtime replacement friction.', 'Dynamic Plotly financial impact charts and interactive 5%-20% savings scenario simulator.'),
    ('OBJ-05', 'Establish verifiable software quality through automated regression testing.', 'Minimum 90% pytest code coverage across calculation, ETL, and data access layers.')
]

for i, row in enumerate(objs_data, start=1):
    for j, val in enumerate(row):
        c = obj_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 70, 70)

doc.add_paragraph()

# 5. Stakeholders - RACI Matrix
h5 = doc.add_heading('5. Stakeholders RACI Matrix', level=1)
style_heading(h5)

p_raci = doc.add_paragraph()
p_raci.add_run('The RACI matrix defines governance responsibilities across project lifecycle decisions: Responsible (R), Accountable (A), Consulted (C), Informed (I).')
p_raci.runs[0].font.name = 'Calibri'
p_raci.runs[0].font.size = Pt(10)

raci_table = doc.add_table(rows=7, cols=7)
raci_table.alignment = WD_TABLE_ALIGNMENT.CENTER
raci_headers = ['Project Activity / Decision', 'SH-01 (CHRO)', 'SH-02 (HR Ops)', 'SH-03 (Line Mgr)', 'SH-04 (CFO)', 'SH-05 (OH&S)', 'SH-06 (Analyst)']
for j, h in enumerate(raci_headers):
    c = raci_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 60, 60)

raci_data = [
    ('Requirements Formulation & Sign-Off', 'Accountable', 'Responsible', 'Consulted', 'Consulted', 'Consulted', 'Responsible'),
    ('Bradford Threshold & Tier Calibration', 'Informed', 'Accountable', 'Consulted', 'Informed', 'Consulted', 'Responsible'),
    ('Financial Multiplier (1.20x) Approval', 'Consulted', 'Informed', 'Informed', 'Accountable', 'Informed', 'Responsible'),
    ('Medical & Reason Taxonomy Mapping', 'Informed', 'Consulted', 'Informed', 'Informed', 'Accountable', 'Responsible'),
    ('Application Validation & UAT Testing', 'Informed', 'Accountable', 'Responsible', 'Consulted', 'Consulted', 'Responsible'),
    ('Policy Interventions & Budget Deployment', 'Accountable', 'Responsible', 'Consulted', 'Responsible', 'Consulted', 'Informed')
]

for i, row in enumerate(raci_data, start=1):
    for j, val in enumerate(row):
        c = raci_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        if j > 0:
            r.bold = True
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if val == 'Accountable':
                r.font.color.rgb = RGBColor(197, 48, 48)
            elif val == 'Responsible':
                r.font.color.rgb = RGBColor(43, 108, 176)
            elif val == 'Consulted':
                r.font.color.rgb = RGBColor(40, 167, 69)
            elif val == 'Informed':
                r.font.color.rgb = RGBColor(113, 128, 150)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

doc.add_paragraph()

# 6. Scope Boundaries
h6 = doc.add_heading('6. Scope: In-Scope vs. Out-of-Scope', level=1)
style_heading(h6)

sc_table = doc.add_table(rows=6, cols=2)
sc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sc_headers = ['In-Scope (Mandated Deliverables)', 'Out-of-Scope (Explicit Exclusions)']
for j, h in enumerate(sc_headers):
    c = sc_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

sc_data = [
    ('Synthesis and cleaning of 1,000+ realistic enterprise HR records.', 'Collection or storage of real employee PII or production payroll databases.'),
    ('Normalized SQLite database schema design with automated DDL generation.', 'Multi-cluster distributed databases (Oracle RAC, Cassandra, CockroachDB).'),
    ('Calculation of monthly rates, Bradford factor scores, and Pareto curves.', 'Physical IoT hardware integration (RFID biometric readers, facial scanning).'),
    ('Interactive Streamlit dashboard featuring multi-dimensional filtering & Plotly.', 'Automated banking integration for salary deductions or punitive payroll clawbacks.'),
    ('Pytest automated test suite and Streamlit Community Cloud hosting.', 'Deep learning computer vision or neural natural language text generation.')
]

for i, (ins, outs) in enumerate(sc_data, start=1):
    c1 = sc_table.rows[i].cells[0]
    c2 = sc_table.rows[i].cells[1]
    c1.paragraphs[0].add_run(ins).font.size = Pt(9.5)
    c2.paragraphs[0].add_run(outs).font.size = Pt(9.5)
    bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
    set_cell_background(c1, bg)
    set_cell_background(c2, bg)
    set_cell_margins(c1, 50, 50, 70, 70)

doc.add_paragraph()

# 7. Business Requirements (BR-01 to BR-12)
h7 = doc.add_heading('7. Detailed Business Requirements (BR-01 to BR-12)', level=1)
style_heading(h7)

br_details = [
    ('BR-01', 'Enterprise Absence Rate Trend Analysis', 'HMW-01', 'Must',
     'The business needs to aggregate, monitor, and trend organizational absenteeism rates across monthly, quarterly, and annual horizons so that HR leadership can forecast cyclical workforce deficits.',
     'Provides macroeconomic visibility into seasonal workforce contractions to optimize temporary hiring and annual leave approvals.',
     'The system must calculate absenteeism rate as (Total Absent Hours / Total Contracted Work Hours) * 100 with zero mathematical variance and support date-range slicing.'),

    ('BR-02', 'Departmental and Job Profile Hotspot Ranking', 'HMW-02', 'Must',
     'The business needs to categorize and rank total lost productive hours and absence incidence rates across departments and job roles so that operational line managers can identify operational bottlenecks.',
     'Enables targeted cross-training and emergency backfill planning in teams suffering from disproportionately high absence rates.',
     'The system must generate dynamic rankings of departments and job roles displaying total lost hours, incident counts, and variance from company average.'),

    ('BR-03', 'Direct Wage Loss and Replacement Expense Quantification', 'HMW-03', 'Must',
     'The business needs to compute direct wage losses and replacement overtime expenses attributable to absent hours across business units so that finance executives can audit payroll leakage.',
     'Transforms raw absence hours into exact monetary figures required by finance controllers to allocate contingency labor budgets.',
     'Calculations must compute direct wage loss (Base Hourly Pay * Hours Absent) and apply an audited 1.20x friction factor for replacement overtime with zero arithmetic error.'),

    ('BR-04', 'Medical & Non-Medical Taxonomy Classification', 'HMW-04', 'Must',
     'The business needs to categorize absence incidents using standardized medical (ICD-10) and non-medical classifications across age and tenure brackets so that occupational health teams can implement wellness initiatives.',
     'Prevents punitive disciplinary responses to genuine medical ailments and guides targeted ergonomic and mental health investments.',
     'Absence records must be segmented into 21 ICD-10 medical classifications and 7 administrative categories, correlated with age cohorts and tenure brackets.'),

    ('BR-05', 'Bradford Factor Disruption Scoring & Tiering', 'HMW-05', 'Must',
     'The business needs to compute the Bradford Factor index for every employee and classify scores into standardized operational disruption tiers so that HR operations can identify disruptive chronic short-term absentees.',
     'Disproportionately weights frequent short-term absence spells (S^2 * D) which cause severe operational friction compared to single long-term illness.',
     'System must compute B = S^2 * D over rolling 52-week windows and categorize each employee into Low (<50), Medium (50-249), High (250-749), or Severe (>=750) tiers.'),

    ('BR-06', 'Commute Distance & Transit Expenditure Correlation', 'HMW-06', 'Should',
     'The business needs to analyze the statistical relationship between employee commute distance / transit expenses and absence frequency so that HR leadership can evaluate hybrid work arrangements.',
     'Empirically demonstrates whether transportation fatigue directly drives incidental absence, supporting remote work policies.',
     'System must compute Pearson/Spearman correlation coefficients and scatter distributions between commute distance (km), transit costs, and absence incident counts.'),

    ('BR-07', 'Biometric & Lifestyle Determinant Correlation', 'HMW-07', 'Should',
     'The business needs to evaluate correlations between employee biometric and lifestyle factors (BMI, smoking, drinking, and family dependents) and absence duration so that wellness officers can target preventive health programs.',
     'Identifies addressable physiological and lifestyle drivers of absenteeism to design subsidized nutrition, smoking cessation, and wellness programs.',
     'System must segment mean absence duration across standard BMI categories (Underweight, Normal, Overweight, Obese) and habit groupings with statistical summaries.'),

    ('BR-08', 'Temporal Day-of-Week & Seasonal Anomaly Detection', 'HMW-08', 'Must',
     'The business needs to identify day-of-the-week patterns and holiday proximity trends in unplanned absence occurrences so that line managers can detect habitual weekend extension patterns.',
     'Uncovers systematic patterns of incidental absenteeism (e.g., Monday/Friday surges) to reinforce shift attendance governance.',
     'System must plot absence frequency across Monday through Friday and flag occurrences within 24 hours of recognized public holidays with statistical deviation metrics.'),

    ('BR-09', 'Pareto 80/20 Concentration Modeling', 'HMW-09', 'Must',
     'The business needs to visualize the cumulative concentration curve of absenteeism across the employee population so that HR and line managers can determine whether a small minority of employees accounts for the majority of lost productivity.',
     'Validates whether managerial attention should be directed toward a critical 20% habitual cohort rather than across the entire workforce.',
     'System must generate a cumulative percentage Lorenz/Pareto distribution showing the percentage of total lost hours driven by deciles of the employee base.'),

    ('BR-10', 'Interactive Cost Mitigation Scenario Simulator', 'HMW-10', 'Should',
     'The business needs to simulate hypothetical financial savings under adjustable percentage reductions (5% to 20%) in preventable absenteeism so that executive leadership can evaluate the projected ROI of proposed wellness interventions.',
     'Provides CFO and CHRO with real-time financial justification for capital allocations toward corporate wellness and ergonomic programs.',
     'System must provide an interactive slider allowing users to adjust reduction percentages from 5% to 20% and display instant projected annual net payroll savings.'),

    ('BR-11', 'Threshold-Based High-Risk Employee Filtering', 'HMW-02, HMW-05', 'Could',
     'The business needs to identify and filter high-risk employees whose absence metrics exceed organizational threshold benchmarks so that HR and department managers can initiate standardized review protocols.',
     'Provides operational supervisors with an actionable roster for quarterly attendance counseling and support interviews.',
     'System must allow filtering of employee rosters by user-defined Bradford score thresholds (e.g., B >= 250) and minimum total absent hours.'),

    ('BR-12', 'Executive Summary Report Export', 'HMW-01, HMW-03', 'Could',
     'The business needs to export standardized workforce absenteeism and financial impact summary reports so that HR and finance teams can share findings during quarterly executive reviews.',
     'Facilitates offline executive briefing and academic defense presentations with verifiable summary statistics.',
     'System must allow downloading filtered KPI summaries, departmental rankings, and cost totals in structured CSV/Excel format.')
]

for br_id, title, hmw, pri, desc, rat, acc in br_details:
    p_b = doc.add_paragraph()
    r_bid = p_b.add_run(f'[{br_id}] {title}')
    r_bid.bold = True
    r_bid.font.name = 'Calibri'
    r_bid.font.size = Pt(11)
    r_bid.font.color.rgb = RGBColor(26, 54, 93)
    
    table_b = doc.add_table(rows=5, cols=2)
    table_b.alignment = WD_TABLE_ALIGNMENT.CENTER
    b_rows = [
        ('Priority & Source:', f'{pri} Priority | Origin: {hmw}'),
        ('Description:', desc),
        ('Business Rationale:', rat),
        ('Acceptance Criteria:', acc),
        ('Audit Traceability:', f'Mapped in RTM as {br_id} with Status: Identified')
    ]
    for r_idx, (k, val) in enumerate(b_rows):
        r_c = table_b.rows[r_idx]
        r_c.cells[0].paragraphs[0].add_run(k).bold = True
        r_c.cells[0].paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_background(r_c.cells[0], 'EDF2F7')
        set_cell_margins(r_c.cells[0], 40, 40, 70, 70)
        
        r_c.cells[1].paragraphs[0].add_run(val).font.size = Pt(9)
        set_cell_background(r_c.cells[1], 'FFFFFF')
        set_cell_margins(r_c.cells[1], 40, 40, 70, 70)
    doc.add_paragraph()

# 8. Business Rules
h8 = doc.add_heading('8. Business Rules Specification', level=1)
style_heading(h8)

rules_table = doc.add_table(rows=7, cols=3)
rules_table.alignment = WD_TABLE_ALIGNMENT.CENTER
r_headers = ['Rule ID', 'Business Rule Statement', 'Operational Constraint / Mathematical Formula']
for j, h in enumerate(r_headers):
    c = rules_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

b_rules = [
    ('RUL-01', 'Standard Working Hours & Contracted Time', 'Standard schedule = 8.0 productive hours per workday, 5 workdays per week (40.0 hrs/week, 2,080 annual base hours per FTE).'),
    ('RUL-02', 'Total Financial Loss Model', 'Financial Loss = (Hours Absent * Base Hourly Wage) * 1.20 Friction Multiplier (incorporating 20% overtime and disruption overhead).'),
    ('RUL-03', 'Bradford Factor Index Formula', 'B = (S^2) * D, where S = number of distinct absence spells (instances) and D = total cumulative absent workdays in a 52-week period.'),
    ('RUL-04', 'Bradford Disruption Risk Tiers', 'Low Risk: B < 50; Medium Risk: 50 <= B <= 249; High Risk: 250 <= B <= 749; Severe Disruption: B >= 750 (triggers HR review).'),
    ('RUL-05', 'ICD-10 Reason Classification', 'Absence reasons must map to standardized ICD-10 medical chapters (e.g., Diseases of Musculoskeletal, Respiratory, Mental) plus 7 administrative non-medical categories.'),
    ('RUL-06', 'Pareto Concentration Standard', 'The analysis evaluates whether the top 20% ranked employees account for >= 60-80% of aggregate lost working hours.')
]

for i, row in enumerate(b_rules, start=1):
    for j, val in enumerate(row):
        c = rules_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 70, 70)

doc.add_paragraph()

# 9. Data Requirements (What, Source, Refresh)
h9 = doc.add_heading('9. Data Requirements Specification', level=1)
style_heading(h9)

data_table = doc.add_table(rows=6, cols=4)
data_table.alignment = WD_TABLE_ALIGNMENT.CENTER
dt_headers = ['Data Entity / Domain', 'Core Required Attributes', 'Source Origin', 'Refresh Cadence']
for j, h in enumerate(dt_headers):
    c = data_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

d_entities = [
    ('Employee Demographics', 'Employee ID, Age, Service Tenure, BMI, Smoker (Y/N), Drinker (Y/N), Children Count, Pet Count, Commute Distance (km).', 'Synthetic Enterprise Generator (Step 10)', 'Static Batch Load / Seed Initialization'),
    ('Department & Cost Centers', 'Department ID, Department Name, Department Cost Center Code, Standard Hourly Pay Scale.', 'Organizational Baseline Setup', 'Static Reference Entity'),
    ('Absence Event Records', 'Event ID, Employee ID, Absence Date, Reason Code, Absent Hours, Day of Week, Month, Season.', 'Synthetic Event Pipeline (Step 10)', 'Periodic Incremental / Batch Refresh'),
    ('Medical & Reason Codes', 'Reason Code (0-28), Category Type (ICD-10 Medical vs Non-Medical), Reason Description, Preventable Flag.', 'WHO ICD-10 Classification Taxonomy', 'Static Master Reference'),
    ('Financial Parameters', 'Base Hourly Wage, Overtime Friction Multiplier (1.20x), Simulated Savings Slider Bounds (5% - 20%).', 'Finance Policy Model (Step 5/Step 7)', 'Dynamic Runtime User Input')
]

for i, row in enumerate(d_entities, start=1):
    for j, val in enumerate(row):
        c = data_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 70, 70)

doc.add_paragraph()

# 10. Assumptions, Constraints, Dependencies
h10 = doc.add_heading('10. Assumptions, Constraints, and Dependencies', level=1)
style_heading(h10)

acd_table = doc.add_table(rows=7, cols=3)
acd_table.alignment = WD_TABLE_ALIGNMENT.CENTER
acd_headers = ['Type', 'Item ID', 'Description & Operational Impact']
for j, h in enumerate(acd_headers):
    c = acd_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

acds = [
    ('Assumption', 'A-01', 'Employees work 8 hours per day, 40 hours per week across a 5-day schedule without weekend shifts.'),
    ('Assumption', 'A-02', 'A standard 20% overtime premium factor represents replacement labor friction without distorting base wage calculations.'),
    ('Constraint', 'C-01', 'Entire software stack is strictly open-source (Python, Streamlit, Plotly, SQLite3, Pytest) to guarantee zero licensing cost.'),
    ('Constraint', 'C-02', 'System memory footprint must not exceed 1.0 GB to comply with Streamlit Community Cloud hosting limits.'),
    ('Dependency', 'DEP-01', 'Step 8 relational SQLite schema must successfully normalize and index employee records prior to dataset generation.'),
    ('Dependency', 'DEP-02', 'Step 10 synthetic dataset must satisfy statistical distributions mirroring real workforce absenteeism studies.')
]

for i, row in enumerate(acds, start=1):
    for j, val in enumerate(row):
        c = acd_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j in [0, 1]:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 70, 70)

doc.add_paragraph()

# 11. Risks and Mitigation Matrix
h11 = doc.add_heading('11. Risk Management & Mitigation Matrix', level=1)
style_heading(h11)

risk_table = doc.add_table(rows=6, cols=5)
risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
rk_headers = ['Risk ID', 'Risk Description', 'Probability', 'Impact', 'Mitigation Strategy']
for j, h in enumerate(rk_headers):
    c = risk_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

risks = [
    ('RSK-01', 'Synthetic data generation may produce unrealistic, uniform, or collinear absence patterns.', 'Medium', 'High', 'Embed empirical probability distributions (gamma/poisson for absence frequency) matching published UC Irvine benchmarks.'),
    ('RSK-02', 'Calculation latency for Bradford Factor and multi-table joins on large recordsets exceeds 2.0s.', 'Low', 'Medium', 'Implement SQLite indexing on foreign keys (employee_id, date) and cache analytical aggregations using Streamlit @st.cache_data.'),
    ('RSK-03', 'Free-tier cloud deployment memory exhaustion during heavy Plotly visual rendering.', 'Low', 'High', 'Downsample rendering payloads, limit client-side trace complexity, and profile memory footprint with psutil during Step 15.'),
    ('RSK-04', 'Requirements creep diverting team from core data analytics into building complex IAM/auth.', 'Medium', 'High', 'Strictly enforce Out-of-Scope boundaries established in BRD; rely on Streamlit built-in session state and cloud password locks.'),
    ('RSK-05', 'Referential integrity corruption during bulk data insertion into SQLite tables.', 'Low', 'High', 'Enforce PRAGMA foreign_keys = ON in SQLite and execute automated pytest integrity tests verifying zero foreign key orphans.')
]

for i, row in enumerate(risks, start=1):
    for j, val in enumerate(row):
        c = risk_table.rows[i].cells[j]
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

doc.add_paragraph()

# 12. Glossary of Terms
h12 = doc.add_heading('12. Glossary of Terms', level=1)
style_heading(h12)

glossary_table = doc.add_table(rows=9, cols=2)
glossary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
gl_headers = ['Term / Acronym', 'Definition & Context in this Project']
for j, h in enumerate(gl_headers):
    c = glossary_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

gl_terms = [
    ('Absenteeism', 'The habitual or unplanned failure of an employee to report for scheduled productive working hours, excluding authorized paid annual vacations.'),
    ('Bradford Factor (B)', 'A mathematical index (B = S^2 * D) measuring workforce disruption, where S is the number of absence spells and D is total absent days over a 52-week period.'),
    ('Absence Spell (S)', 'An unbroken, continuous instance of unplanned absence of one or more consecutive days.'),
    ('Friction Multiplier', 'A financial coefficient (1.20x) applied to base wages to incorporate the indirect costs of supervisor rescheduling, administrative handling, and premium overtime paid to replacement staff.'),
    ('ICD-10 Taxonomy', 'International Statistical Classification of Diseases and Related Health Problems (10th Revision), utilized to classify medical conditions systematically.'),
    ('Pareto Principle (80/20)', 'The empirical observation that approximately 80% of lost organizational hours are generated by approximately 20% of the employee cohort.'),
    ('RTM', 'Requirements Traceability Matrix: a living audit artifact tracing business inquiries (OQs) to requirements, architecture, stories, and test cases.'),
    ('MoSCoW', 'A prioritization framework categorizing requirements into Must Have, Should Have, Could Have, and Won\'t Have to manage project scope.')
]

for i, (term, defn) in enumerate(gl_terms, start=1):
    c1 = glossary_table.rows[i].cells[0]
    c2 = glossary_table.rows[i].cells[1]
    c1.paragraphs[0].add_run(term).bold = True
    c1.paragraphs[0].runs[0].font.size = Pt(9.5)
    c2.paragraphs[0].add_run(defn).font.size = Pt(9.5)
    bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
    set_cell_background(c1, bg)
    set_cell_background(c2, bg)
    set_cell_margins(c1, 50, 50, 70, 70)

doc.add_paragraph()

# 13. Formal Sign-Off & Approvals
h13 = doc.add_heading('13. Formal Sign-Off & Approval Block', level=1)
style_heading(h13)

p_sign = doc.add_paragraph()
p_sign.add_run('By signing below, the undersigned stakeholders confirm that this Business Requirements Document (BRD v1.0) completely and accurately reflects the functional and operational requirements for the Absenteeism Impact Analysis system. The engineering team is authorized to proceed with architectural formulation and Product Requirements Document (PRD) creation.')
p_sign.runs[0].font.name = 'Calibri'
p_sign.runs[0].font.size = Pt(10)

sign_table = doc.add_table(rows=4, cols=4)
sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sn_headers = ['Role & Authority', 'Signatory Name', 'Department / Institution', 'Approval Signature & Date']
for j, h in enumerate(sn_headers):
    c = sign_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

sign_data = [
    ('Primary Student Author', 'Kalpana (Roll No: 2500320140071)', 'MCA 2nd Yr DA, ABES', '[Signed Digitally] 20-Nov-2026'),
    ('Faculty Project Mentor', 'Prof. Gunjan Agrawal', 'Dept. of Computer Applications, ABES', '[Approved] 20-Nov-2026'),
    ('Academic Review Committee', 'ABES MCA Evaluation Committee', 'ABES Engineering College', '[Reviewed] Pending Viva Defense')
]

for i, row in enumerate(sign_data, start=1):
    for j, val in enumerate(row):
        c = sign_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j in [0, 1]:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 70, 70)

# Save document to C:\HCL\05_BRD.docx
os.makedirs('C:/HCL', exist_ok=True)
p_brd = 'C:/HCL/05_BRD.docx'
doc.save(p_brd)
print('SUCCESSFULLY SAVED:', p_brd)
