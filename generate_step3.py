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
r_title = p_title.add_run('STEP 3: BUSINESS NEEDS ANALYSIS')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Business Requirements Specification (BRD Input) | Absenteeism Impact Analysis')
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
    ('Deliverable File:', '03_Needs_Analysis.docx')
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

# 1. Executive Summary & Methodology
h1 = doc.add_heading('1. Executive Overview & Needs Analysis Principles', level=1)
h1.style.font.color.rgb = RGBColor(26, 54, 93)

p_desc = doc.add_paragraph()
p_desc.add_run('This document translates the design-thinking How Might We (HMW) opportunities established in Step 2 into formalized Business Requirements (BRs). Each requirement strictly captures WHAT capabilities the business ecosystem demands and the corresponding organizational benefits, deliberately omitting technical implementation details (such as database engines or UI frameworks). Requirements are prioritized using the MoSCoW framework to guarantee delivery of high-value analytical capabilities.')
p_desc.runs[0].font.name = 'Calibri'
p_desc.runs[0].font.size = Pt(10.5)

# 2. Business Requirements Table
h2 = doc.add_heading('2. Business Requirements Matrix (BR-01 to BR-12)', level=1)
h2.style.font.color.rgb = RGBColor(26, 54, 93)

br_table = doc.add_table(rows=13, cols=6)
br_table.alignment = WD_TABLE_ALIGNMENT.CENTER
br_headers = ['BR ID', 'Formal Business Requirement (The business needs to <capability> so that <benefit>)', 'Primary Stakeholder', 'Source HMW', 'Priority', 'Measurement & Verification Standard']
for j, h in enumerate(br_headers):
    c = br_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 80, 80)

br_data = [
    ('BR-01', 'The business needs to aggregate, monitor, and trend organizational absenteeism rates on monthly, quarterly, and annual horizons so that HR leadership can forecast cyclical workforce deficits and adjust strategic staffing policies.', 'SH-01, SH-02', 'HMW-01', 'Must', 'Accurate calculation of monthly absenteeism percentage matching historical timeframes with 100% calculation consistency against total contracted hours.'),
    ('BR-02', 'The business needs to categorize and rank total lost productive hours and absence incidence rates across departments and job roles so that operational line managers can identify operational bottlenecks and proactively arrange cross-training or shift backfills.', 'SH-03', 'HMW-02', 'Must', 'Ranked list of departments and roles displaying total lost hours, incident counts, and deviation from company baseline.'),
    ('BR-03', 'The business needs to compute direct wage losses and replacement overtime expenses attributable to absent hours across business units so that finance executives can audit payroll leakage and establish realistic contingency budgets.', 'SH-04', 'HMW-04', 'Must', 'Financial loss calculated as (Base Hourly Wage * Absent Hours * 1.20 Friction Multiplier) with zero arithmetic error across aggregates.'),
    ('BR-04', 'The business needs to categorize absence incidents using standardized medical and non-medical reason classifications across employee age brackets and service tenures so that occupational health teams can implement preventive wellness and ergonomic initiatives.', 'SH-05', 'HMW-04', 'Must', 'Distribution breakdown of absence events and hours grouped into standardized ICD-10 medical codes and tenure cohorts.'),
    ('BR-05', 'The business needs to compute the Bradford Factor index for every employee and classify scores into standardized operational disruption tiers so that HR operations can identify and address disruptive chronic short-term casual absentees.', 'SH-02', 'HMW-05', 'Must', 'Automated calculation of B = S^2 * D and classification into 4 disruption tiers (Low, Medium, High, Severe Risk) across 100% of employees.'),
    ('BR-06', 'The business needs to analyze the statistical relationship between employee commute distance / transit expenses and absence frequency so that HR leadership can evaluate the feasibility and targeting of hybrid work arrangements.', 'SH-01, SH-05', 'HMW-06', 'Should', 'Bivariate correlation coefficient and scatter grouping between commute distance (km), transportation expense, and absence frequency.'),
    ('BR-07', 'The business needs to evaluate correlations between employee biometric and lifestyle factors (BMI, smoking, drinking, and family dependents) and absence duration so that wellness officers can target preventive health programs.', 'SH-05', 'HMW-07', 'Should', 'Comparative average absence duration computed across BMI categories (Normal, Overweight, Obese) and lifestyle habit segments.'),
    ('BR-08', 'The business needs to identify day-of-the-week patterns and holiday proximity trends in unplanned absence occurrences so that line managers can detect habitual weekend extension patterns and enforce shift coverage.', 'SH-03', 'HMW-08', 'Must', 'Distribution of absence incidents across days of the week (Monday through Friday) highlighting statistically significant anomalies (>15% above mean).'),
    ('BR-09', 'The business needs to visualize the cumulative concentration curve of absenteeism across the employee population so that HR and line managers can determine whether a small minority of employees accounts for the majority of lost productivity.', 'SH-02, SH-06', 'HMW-09', 'Must', 'Generation of a cumulative percentage Pareto distribution showing the proportion of total lost hours generated by the top 10%, 20%, and 30% of absent employees.'),
    ('BR-10', 'The business needs to simulate hypothetical financial savings under adjustable percentage reductions (5% to 20%) in preventable absenteeism so that executive leadership can evaluate the projected ROI of proposed wellness interventions.', 'SH-01, SH-04', 'HMW-10', 'Should', 'Real-time recalculation of net cost savings against baseline payroll losses based on user-selected reduction percentages.'),
    ('BR-11', 'The business needs to identify and filter high-risk employees whose absence metrics exceed organizational threshold benchmarks so that HR and department managers can initiate standardized review protocols.', 'SH-02, SH-03', 'HMW-02, HMW-05', 'Could', 'Filtered tabular roster of employees crossing designated Bradford Factor or total absence hour thresholds.'),
    ('BR-12', 'The business needs to export standardized workforce absenteeism and financial impact summary reports so that HR and finance teams can share findings during quarterly executive reviews.', 'SH-01, SH-04', 'HMW-01, HMW-03', 'Could', 'Generation of formatted downloadable summary reports containing key KPI totals, departmental summaries, and financial figures.')
]

for i, row in enumerate(br_data, start=1):
    for j, val in enumerate(row):
        c = br_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        if j == 0:
            r.bold = True
        elif j == 4:
            r.bold = True
            if val == 'Must':
                r.font.color.rgb = RGBColor(197, 48, 48)
            elif val == 'Should':
                r.font.color.rgb = RGBColor(43, 108, 176)
            elif val == 'Could':
                r.font.color.rgb = RGBColor(40, 167, 69)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

doc.add_paragraph()

# 3. MoSCoW Prioritization Summary
h3 = doc.add_heading('3. MoSCoW Prioritization Breakdown', level=1)
h3.style.font.color.rgb = RGBColor(26, 54, 93)

p_moscow = doc.add_paragraph()
p_moscow.add_run('The 12 business requirements are prioritized as follows, satisfying the ~60% Must-Have curriculum target:')
p_moscow.runs[0].font.name = 'Calibri'
p_moscow.runs[0].font.size = Pt(10)

moscow_table = doc.add_table(rows=5, cols=4)
moscow_table.alignment = WD_TABLE_ALIGNMENT.CENTER
m_headers = ['MoSCoW Category', 'Count of BRs', 'Percentage of Total', 'Included Requirement IDs']
for j, h in enumerate(m_headers):
    c = moscow_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

m_rows = [
    ('Must Have (Critical Core)', '7', '58.3% (~60%)', 'BR-01, BR-02, BR-03, BR-04, BR-05, BR-08, BR-09'),
    ('Should Have (High Value)', '3', '25.0%', 'BR-06, BR-07, BR-10'),
    ('Could Have (Desirable)', '2', '16.7%', 'BR-11, BR-12'),
    ('Won\'t Have (Deferred)', '0', '0.0%', 'Biometric hardware sync, automated salary deductions (Out of scope)')
]

for i, row in enumerate(m_rows, start=1):
    for j, val in enumerate(row):
        c = moscow_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 70, 70)

doc.add_paragraph()

# 4. Underlying Data Requirements
h4 = doc.add_heading('4. Underlying Data Requirements Matrix', level=1)
h4.style.font.color.rgb = RGBColor(26, 54, 93)

p_data_intro = doc.add_paragraph()
p_data_intro.add_run('To satisfy the business requirements, the underlying data layer must provide the following input attributes:')
p_data_intro.runs[0].font.name = 'Calibri'
p_data_intro.runs[0].font.size = Pt(10)

data_table = doc.add_table(rows=13, cols=3)
data_table.alignment = WD_TABLE_ALIGNMENT.CENTER
d_headers = ['BR ID', 'Business Capability', 'Required Underlying Data Attributes']
for j, h in enumerate(d_headers):
    c = data_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

data_reqs = [
    ('BR-01', 'Enterprise absence rate trending', 'Employee ID, Attendance Date, Absence Hours, Standard Daily Working Hours (8 hrs), Contracted Monthly Hours.'),
    ('BR-02', 'Departmental & job role ranking', 'Department ID, Department Name, Job Profile Title, Absence Hours, Employee ID.'),
    ('BR-03', 'Direct payroll loss quantification', 'Employee Hourly Pay Rate, Total Absent Hours, Department Cost Center, Replacement Overtime Factor (1.20).'),
    ('BR-04', 'Medical & non-medical classification', 'Absence Reason Code (ICD-10 classification), Reason Category, Employee Age/Birth Date, Date of Joining (Tenure).'),
    ('BR-05', 'Bradford Factor index calculation', 'Employee ID, Absence Event Start Date, Absence Event End Date, Distinct Spell Count (S), Cumulative Absent Days (D).'),
    ('BR-06', 'Commute distance correlation', 'Distance from Residence to Work (km), Daily Transportation Cost ($), Absence Event Timestamp, Incident Count.'),
    ('BR-07', 'Lifestyle & biometric correlation', 'Employee Body Mass Index (BMI), Height, Weight, Smoker Status (Y/N), Social Drinker Status (Y/N), Count of Children, Count of Pets.'),
    ('BR-08', 'Day-of-week & calendar patterns', 'Event Date, Day of Week (Monday to Friday), Calendar Month, Holiday Proximity Indicator, Absence Duration.'),
    ('BR-09', 'Pareto 80/20 distribution modeling', 'Ranked Total Absent Hours per Employee, Cumulative Sum of Hours, Total Workforce Headcount.'),
    ('BR-10', 'Interactive cost savings simulation', 'Aggregated Baseline Financial Loss, Avoidable Absence Hours Sub-total, User Reduction Parameter (5% to 20%).'),
    ('BR-11', 'High-risk employee threshold filtering', 'Employee ID, Employee Name, Department, Calculated Bradford Score, Total Hours Absent, Disciplinary Action Flag.'),
    ('BR-12', 'Executive summary report export', 'Aggregated Monthly Loss Metrics, Departmental Summary Rollups, Reason Code Totals, Date Filter Parameters.')
]

for i, row in enumerate(data_reqs, start=1):
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

# 5. Gap Analysis (Bidirectional Traceability)
h5 = doc.add_heading('5. Gap Analysis: Bidirectional Traceability Audit', level=1)
h5.style.font.color.rgb = RGBColor(26, 54, 93)

p_gap = doc.add_paragraph()
p_gap.add_run('A bidirectional gap analysis was performed between Step 2 How Might We (HMW) statements and Step 3 Business Requirements (BRs) to verify that no requirements lack justification and no stakeholder needs are neglected:')
p_gap.runs[0].font.name = 'Calibri'
p_gap.runs[0].font.size = Pt(10)

gap_table = doc.add_table(rows=11, cols=4)
gap_table.alignment = WD_TABLE_ALIGNMENT.CENTER
g_headers = ['Source HMW ID', 'HMW Statement Summary', 'Mapped BR ID(s)', 'Traceability Status']
for j, h in enumerate(g_headers):
    c = gap_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 80, 80)

gap_data = [
    ('HMW-01', 'Track seasonal and MoM absence trends', 'BR-01, BR-12', 'VERIFIED (No Gap)'),
    ('HMW-02', 'Identify departmental and job role hotspots', 'BR-02, BR-11', 'VERIFIED (No Gap)'),
    ('HMW-03', 'Quantify direct wage leakage and replacement costs', 'BR-03, BR-12', 'VERIFIED (No Gap)'),
    ('HMW-04', 'Identify dominant medical root causes', 'BR-04', 'VERIFIED (No Gap)'),
    ('HMW-05', 'Distinguish long-term vs. short-term disruption (Bradford)', 'BR-05, BR-11', 'VERIFIED (No Gap)'),
    ('HMW-06', 'Evaluate commute distance and transit costs', 'BR-06', 'VERIFIED (No Gap)'),
    ('HMW-07', 'Correlate lifestyle and biometric health indicators', 'BR-07', 'VERIFIED (No Gap)'),
    ('HMW-08', 'Alert to day-of-week and weekend extensions', 'BR-08', 'VERIFIED (No Gap)'),
    ('HMW-09', 'Pinpoint Pareto 80/20 concentration of hours', 'BR-09', 'VERIFIED (No Gap)'),
    ('HMW-10', 'Simulate financial savings under reduction scenarios', 'BR-10', 'VERIFIED (No Gap)')
]

for i, row in enumerate(gap_data, start=1):
    for j, val in enumerate(row):
        c = gap_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
        elif j == 3:
            r.bold = True
            r.font.color.rgb = RGBColor(40, 167, 69)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 70, 70)

# Save document
os.makedirs('C:/HCL', exist_ok=True)
p1 = 'C:/HCL/03_Needs_Analysis.docx'
doc.save(p1)
print('SUCCESSFULLY SAVED:', p1)
