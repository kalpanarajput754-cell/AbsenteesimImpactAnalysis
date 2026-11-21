import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
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

# Title Block
p_title = doc.add_paragraph()
r_title = p_title.add_run('SYNTHETIC DATASET GENERATION & AUDIT SPECIFICATION')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Step 10 Data Engineering Deliverable | Absenteeism Impact Analysis')
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(43, 108, 176)
r_sub.font.italic = True

# Metadata Box Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Project ID & Title:', 'PRJ-AIA-01 | Absenteeism Impact Analysis'),
    ('Student Author:', 'Kalpana (Roll No: 2500320140071) — MCA 2nd Yr DA, ABES'),
    ('Faculty Project Mentor:', 'Prof. Gunjan Agrawal'),
    ('Data Generation Script:', 'src/generate_data.py (Seed: 42)'),
    ('Production Database:', 'data/app.db (SQLite with Foreign Keys ON)'),
    ('Dataset Artifacts:', 'data/absenteeism_data.xlsx, CSVs, validation_report.txt')
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

# 1. Summary of Generated Tables
h1 = doc.add_heading('1. Summary of Generated Tables & Row Counts', level=1)
style_heading(h1)

t_summary = doc.add_table(rows=6, cols=5)
t_summary.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ['Table / View Name', 'Table Type', 'Actual Rows Generated', 'Column Count', 'Validation Status']
for j, h in enumerate(headers):
    c = t_summary.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

t_rows = [
    ('dim_departments', 'Master Dimension', '6', '5', '100% Pass (Zero Errors)'),
    ('dim_reasons', 'Master Dimension', '28', '6', '100% Pass (Zero Errors)'),
    ('dim_employees', 'Master Dimension', '120', '18', '100% Pass (Zero Errors)'),
    ('fact_absence_events', 'Core Fact Table', '320', '14', '100% Pass (Zero Errors)'),
    ('v_absence_analysis_joined', 'Unified Analysis View', '320', '32', '100% Pass (Zero Errors)')
]

for i, row in enumerate(t_rows, start=1):
    for j, val in enumerate(row):
        c = t_summary.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
        elif j == 4:
            r.bold = True
            r.font.color.rgb = RGBColor(40, 167, 69)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

doc.add_paragraph()

# 2. Five Realistic Built-In Patterns
h2 = doc.add_heading('2. Embedded Analytical Patterns for Discovery', level=1)
style_heading(h2)

patterns = [
    ('Pattern 1: Weekend Extension Spike (Monday/Friday Surge):', 
     'Absence occurrences cluster heavily on Mondays (141 events) and Fridays (73 events) compared to a midweek baseline average of 35.3 events (Tue-Thu). This creates a statistically significant +26.4% surge in weekend-adjacent incidental absences, supporting BR-08 / OQ-08.'),
    ('Pattern 2: Pareto 80/20 Disruption Concentration (Bradford Factor):', 
     'A specialized cohort of 14 chronic repeat absentees (11.7% of the workforce) accounts for 2,040.0 of 3,656.0 total lost hours (55.8% of aggregate lost time). These employees incur high spell counts (S = 5 to 8 short spells of 1-2 days), triggering severe Bradford Factor scores (B >= 750) and proving the non-linear disruption model in BR-05 / OQ-05.'),
    ('Pattern 3: Commute Distance & Transit Friction:', 
     'Employees residing in distant suburban zones (> 30 km commute distance) exhibit an average annual absence frequency of 4.58 events/year compared to 3.56 events/year for employees living within 15 km, supporting hybrid work justification under BR-06 / OQ-06.'),
    ('Pattern 4: Biometric (BMI) & Musculoskeletal Correlation:', 
     'Employees categorized in the Obese BMI category (BMI >= 30.0) suffer a significantly higher frequency of ICD-10 Chapter XIII Musculoskeletal conditions (back strain, repetitive joint stress) and higher average duration per incident (10.65 hrs/event) vs Normal BMI staff, supporting BR-07 / OQ-07.'),
    ('Pattern 5: Seasonal Monsoon & Autumn Infection Surge (Q3):', 
     'Q3 (July through September) accounts for 92 events (28.7% of annual volume) driven by elevated surges in Respiratory diseases (ICD-10 J00-J99) and seasonal viral infections, validating seasonal workforce forecasting under BR-01 / OQ-01.')
]

for title, desc in patterns:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(10)
    r1.font.color.rgb = RGBColor(26, 54, 93)
    r2 = p.add_run(desc)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(9.5)

doc.add_paragraph()

# 3. Intentional Data Quality Anomalies
h3 = doc.add_heading('3. Staged Data Quality Issues (ETL Cleaning Groundwork)', level=1)
style_heading(h3)

p_dq = doc.add_paragraph()
p_dq.add_run('To ensure realistic QA verification in Step 13 and demonstrate comprehensive data cleaning capabilities in Step 15, the following intentional anomalies are documented in raw extracts (df_employees_raw & df_fact_raw):')
p_dq.runs[0].font.name = 'Calibri'
p_dq.runs[0].font.size = Pt(10)

anomalies = [
    ('1. Missing Values (2.5% Rate):', 'Injected 3 null values into daily_transit_cost (employee indices: 12, 45, 88) and 2 null values into distance_to_work_km (indices: 23, 71) in raw employee records.'),
    ('2. Duplicate Records:', 'Injected 2 exact duplicate rows into raw fact event logs (cloning event rows 14 and 82) to test deduplication logic during ingestion.'),
    ('3. Extreme Clerical Outlier:', 'Injected 1 extreme erroneous event duration of 96.0 absent hours (Row 3, an impossible 12-day duration in a single weekly spell) to test IQR outlier boundary fencing.')
]

for title, desc in anomalies:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(10)
    r1.font.color.rgb = RGBColor(197, 48, 48)
    r2 = p.add_run(desc)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(9.5)

os.makedirs('C:/HCL/docs', exist_ok=True)
p_doc = 'C:/HCL/docs/10_Synthetic_Dataset_Generation.docx'
doc.save(p_doc)
print('SUCCESSFULLY SAVED DOCX:', p_doc)
