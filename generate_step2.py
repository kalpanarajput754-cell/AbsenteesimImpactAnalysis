import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

doc = docx.Document()

# Set standard 1-inch margins
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
r_title = p_title.add_run('STEP 2: HOW MIGHT WE (HMW) STATEMENTS')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Design-Thinking Problem Reframing | Absenteeism Impact Analysis')
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
    ('Deliverable File:', '02_HMW_Statements.docx')
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

# 1. Stakeholder Pain Points
h1 = doc.add_heading('1. Stakeholder Pain Points (Design-Thinking Empathy Map)', level=1)
h1.style.font.color.rgb = RGBColor(26, 54, 93)

p_intro = doc.add_paragraph()
p_intro.add_run('To formulate grounded How Might We (HMW) design statements, we examine the operational friction, cognitive burdens, and analytical blind spots experienced by each project stakeholder identified in Step 1:')
p_intro.runs[0].font.name = 'Calibri'
p_intro.runs[0].font.size = Pt(10.5)

stakeholders_pains = [
    ('SH-01: Chief Human Resources Officer (CHRO)', [
        'Lacks macroscopic visibility into enterprise-wide seasonal absenteeism trends, leaving executive policy updates reactive rather than proactive.',
        'Struggles to substantiate budget requests for corporate wellness programs, flexible working models, and ergonomics with quantifiable financial returns.'
    ]),
    ('SH-02: HR Operations Manager', [
        'Cannot easily differentiate between unavoidable prolonged illness and chronic, disruptive short-term casual absences.',
        'Spends extensive manual effort collating siloed attendance spreadsheets across departments, delaying monthly governance reporting.',
        'Lacks automated early warning alerts to address borderline attendance policy violations before they escalate into disciplinary crises.'
    ]),
    ('SH-03: Department Line Managers', [
        'Unplanned absences clustering around Mondays and Fridays severely disrupt operational shift schedules and client delivery timelines.',
        'Lacks departmental transparency into absence patterns, making it difficult to allocate cross-training resources or schedule overtime coverage in advance.'
    ]),
    ('SH-04: Chief Financial Officer (CFO) / Controller', [
        'Absences trigger double-jeopardy financial leakage: unworked base salary paid out plus emergency overtime compensation paid to covering staff.',
        'No mechanism exists to simulate projected cost savings from wellness interventions before committing capital expenditure.'
    ]),
    ('SH-05: Occupational Health & Safety (OH&S) Lead', [
        'Unable to systematically diagnose whether absences are driven by medical ailments (ICD-10 codes) versus ergonomic strain or lifestyle factors across demographic cohorts.',
        'Lacks empirical validation to determine whether long commute distances and transport costs drive employee fatigue and absenteeism.'
    ]),
    ('SH-06: Internal Data Analyst (Kalpana)', [
        'Manual metric calculation (e.g., Bradford Factor, rolling absenteeism rates) from disparate raw logs is error-prone and labor-intensive.',
        'Difficult to clearly communicate statistical concepts (such as the Pareto 80/20 distribution of hours lost) without dynamic, interactive data visualizers.'
    ])
]

for sh_title, pains in stakeholders_pains:
    p_sh = doc.add_paragraph()
    r = p_sh.add_run(sh_title)
    r.bold = True
    r.font.name = 'Calibri'
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(43, 108, 176)
    for pain in pains:
        p_pt = doc.add_paragraph(style='List Bullet')
        r_pt = p_pt.add_run(pain)
        r_pt.font.name = 'Calibri'
        r_pt.font.size = Pt(10)

doc.add_paragraph()

# 2. HMW Reframing Matrix Table
h2 = doc.add_heading('2. How Might We (HMW) Statements Matrix', level=1)
h2.style.font.color.rgb = RGBColor(26, 54, 93)

p_rule = doc.add_paragraph()
p_rule.add_run('Each HMW statement reframes a specific stakeholder pain point into an open, actionable inquiry without prematurely prescribing technical solutions, maintaining strict traceability back to Step 1 Operational Questions (OQs):')
p_rule.runs[0].font.name = 'Calibri'
p_rule.runs[0].font.size = Pt(10.5)

hmw_table = doc.add_table(rows=11, cols=5)
hmw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
hmw_headers = ['HMW ID', 'How Might We ...?', 'Primary Stakeholder', 'Core Pain Point Addressed', 'Linked OQ']
for j, h in enumerate(hmw_headers):
    c = hmw_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 80, 80)

hmws = [
    ('HMW-01', 'How might we enable HR leadership to track seasonal and month-over-month absence trends across the organization to anticipate workforce shortages?', 'SH-01, SH-02', 'Absence tracking is reactive; lack of visibility into cyclical trends.', 'OQ-01'),
    ('HMW-02 [PRIORITY]', 'How might we help department line managers identify which teams and roles experience the highest concentration of lost hours so they can plan coverage?', 'SH-03', 'Unplanned absenteeism leaves shift teams understaffed and unprepared.', 'OQ-02'),
    ('HMW-03 [PRIORITY]', 'How might we quantify direct salary leakage and overtime replacement expenses for finance leadership across each business unit?', 'SH-04', 'Absences incur hidden costs with no granular financial accounting per department.', 'OQ-03'),
    ('HMW-04 [PRIORITY]', 'How might we help the occupational health team identify dominant medical root causes across different age cohorts and service tenures?', 'SH-05', 'Medical absence reasons are unclassified, preventing preventative health actions.', 'OQ-04'),
    ('HMW-05 [PRIORITY]', 'How might we distinguish between unavoidable long-term medical absences and disruptive repeated short-term absence patterns for operational HR?', 'SH-02', 'High-frequency short absences inflict higher disruption than single long leaves.', 'OQ-05'),
    ('HMW-06', 'How might we evaluate the degree to which employee commute distance and transit costs drive incidental day-off occurrences?', 'SH-01, SH-05', 'Lack of evidence on whether travel burden and commute fatigue trigger absenteeism.', 'OQ-06'),
    ('HMW-07', 'How might we help wellness officers correlate lifestyle and demographic indicators (BMI, social habits, family size) with prolonged absence duration?', 'SH-05', 'Inability to target health programs to high-risk demographic/lifestyle cohorts.', 'OQ-07'),
    ('HMW-08', 'How might we alert team supervisors to recurring day-of-week absence patterns (such as Monday/Friday spikes) before project deadlines are jeopardized?', 'SH-03', 'Incidental absence patterns around weekends compromise team sprint delivery.', 'OQ-08'),
    ('HMW-09', 'How might we pinpoint whether a small minority of habitual absentees account for the overwhelming majority of lost company hours?', 'SH-02, SH-06', 'HR cannot verify if absenteeism follows an 80/20 Pareto distribution.', 'OQ-09'),
    ('HMW-10 [PRIORITY]', 'How might we enable finance and HR executives to simulate potential payroll cost savings under varying scenarios of absenteeism reduction?', 'SH-01, SH-04', 'No mechanism to calculate projected ROI for proposed wellness investments.', 'OQ-10')
]

for i, row in enumerate(hmws, start=1):
    for j, val in enumerate(row):
        c = hmw_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
            if 'PRIORITY' in val:
                r.font.color.rgb = RGBColor(197, 48, 48)
        elif j == 4:
            r.bold = True
            r.font.color.rgb = RGBColor(43, 108, 176)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 60, 60, 80, 80)

doc.add_paragraph()

# 3. Top 5 Priority HMW Statements
h3 = doc.add_heading('3. Top 5 Priority HMW Statements & Strategic Rationale', level=1)
h3.style.font.color.rgb = RGBColor(26, 54, 93)

priorities = [
    ('1. HMW-05 (Bradford Disruption Differentiation):', 'Differentiating frequent short-term absence spells from long-term medical conditions via Bradford Factor scoring is critical because unpredictable incidental absences inflict the highest immediate operational turmoil on team productivity.'),
    ('2. HMW-03 (Granular Financial Impact Quantification):', 'Translating lost hours into direct wage leakage plus replacement overtime expense converts abstract attendance counts into concrete monetary figures required for CFO decision-making.'),
    ('3. HMW-02 (Departmental & Role Hotspot Identification):', 'Pinpointing specific department and job profile bottlenecks enables operational line managers to reallocate capacity, buffer deliverables, and arrange targeted backfills.'),
    ('4. HMW-10 (Dynamic Cost Mitigation Simulation):', 'Equipping executive decision-makers with interactive What-If scenario simulations proves the fiscal viability and ROI of absenteeism reduction programs before funds are allocated.'),
    ('5. HMW-04 (Medical & Physiological Root-Cause Diagnosis):', 'Categorizing ICD-10 medical conditions against demographic groups shifts the organization from punitive disciplinary measures to high-impact preventative wellness interventions.')
]

for title, rationale in priorities:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(10)
    r1.font.color.rgb = RGBColor(197, 48, 48)
    r2 = p.add_run(rationale)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(10)

doc.add_paragraph()

# 4. Operational Questions Coverage Verification
h4 = doc.add_heading('4. Traceability & Operational Question Coverage Verification', level=1)
h4.style.font.color.rgb = RGBColor(26, 54, 93)

p_cov = doc.add_paragraph()
p_cov.add_run('Every Operational Question established in Step 1 is mapped to at least one How Might We statement, ensuring zero orphan questions and complete requirements continuity:')
p_cov.runs[0].font.name = 'Calibri'
p_cov.runs[0].font.size = Pt(10.5)

cov_table = doc.add_table(rows=11, cols=4)
cov_table.alignment = WD_TABLE_ALIGNMENT.CENTER
cov_headers = ['OQ ID', 'Operational Question Summary', 'Mapped HMW Statement(s)', 'Coverage Status']
for j, h in enumerate(cov_headers):
    c = cov_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 80, 80, 80, 80)

cov_data = [
    ('OQ-01', 'Organizational absenteeism rate and temporal trend', 'HMW-01', '100% COVERED'),
    ('OQ-02', 'Departmental and job profile lost-hour hotspots', 'HMW-02', '100% COVERED'),
    ('OQ-03', 'Direct wage loss and replacement overtime expense', 'HMW-03', '100% COVERED'),
    ('OQ-04', 'Medical (ICD-10) and non-medical reason distributions', 'HMW-04', '100% COVERED'),
    ('OQ-05', 'Bradford Factor operational friction index calculation', 'HMW-05', '100% COVERED'),
    ('OQ-06', 'Commute distance and transit cost correlation', 'HMW-06', '100% COVERED'),
    ('OQ-07', 'Demographic and lifestyle correlation (BMI, age, habits)', 'HMW-07', '100% COVERED'),
    ('OQ-08', 'Day-of-week and seasonal absence occurrence patterns', 'HMW-08', '100% COVERED'),
    ('OQ-09', 'Pareto 80/20 distribution of absence concentration', 'HMW-09', '100% COVERED'),
    ('OQ-10', 'Interactive financial savings scenario simulation', 'HMW-10', '100% COVERED')
]

for i, row in enumerate(cov_data, start=1):
    for j, val in enumerate(row):
        c = cov_table.rows[i].cells[j]
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
        set_cell_margins(c, 60, 60, 80, 80)

# Save document
os.makedirs('C:/HCL', exist_ok=True)
p1 = 'C:/HCL/02_HMW_Statements.docx'
doc.save(p1)
print('SUCCESSFULLY SAVED:', p1)
