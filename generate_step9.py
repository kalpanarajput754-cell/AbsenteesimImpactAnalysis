import matplotlib.pyplot as plt
import matplotlib.patches as patches
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

# ========================================================
# 1. GENERATE HIGH-RESOLUTION ER DIAGRAM PNG
# ========================================================
fig, ax = plt.subplots(figsize=(15, 10), dpi=300)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

# Background
fig.patch.set_facecolor('#F8FAFC')
ax.set_facecolor('#F8FAFC')

# Title
ax.text(50, 96, "ENTITY RELATIONSHIP (ER) DIAGRAM — LOGICAL SCHEMA", 
        fontsize=16, fontweight='bold', ha='center', va='center', color='#1A365D', fontfamily='sans-serif')
ax.text(50, 93, "PRJ-AIA-01: Absenteeism Impact Analysis | Relational Star/Snowflake Hybrid", 
        fontsize=11, fontstyle='italic', ha='center', va='center', color='#4A5568', fontfamily='sans-serif')

def draw_entity(ax, x, y, width, height, title, columns):
    # Shadow
    rect_shadow = patches.FancyBboxPatch((x+0.6, y-0.6), width, height,
                                         boxstyle="round,pad=0.2,rounding_size=0.8",
                                         linewidth=0, facecolor='#CBD5E0', alpha=0.5, zorder=2)
    ax.add_patch(rect_shadow)
    # Box
    rect = patches.FancyBboxPatch((x, y), width, height,
                                  boxstyle="round,pad=0.2,rounding_size=0.8",
                                  linewidth=1.5, edgecolor='#1E3A8A', facecolor='#FFFFFF', zorder=3)
    ax.add_patch(rect)
    # Header
    header_height = 3.8
    rect_header = patches.FancyBboxPatch((x, y + height - header_height), width, header_height,
                                         boxstyle="round,pad=0.2,rounding_size=0.8",
                                         linewidth=0, facecolor='#1A365D', zorder=4)
    ax.add_patch(rect_header)
    ax.text(x + width/2, y + height - header_height/2, title, 
            fontsize=10, fontweight='bold', ha='center', va='center', color='#FFFFFF', fontfamily='sans-serif', zorder=5)
    
    # Columns
    line_y = y + height - header_height - 1.8
    for col_type, col_name, key_type in columns:
        prefix = ""
        color = "#2D3748"
        weight = 'normal'
        if key_type == 'PK':
            prefix = "[PK] "
            color = "#C53030"
            weight = 'bold'
        elif key_type == 'FK':
            prefix = "[FK] "
            color = "#2B6CB0"
            weight = 'bold'
            
        ax.text(x + 1.2, line_y, f"{prefix}{col_name}", fontsize=7.5, fontweight=weight, color=color, va='center', fontfamily='monospace', zorder=5)
        ax.text(x + width - 1.2, line_y, col_type, fontsize=7.5, color='#718096', ha='right', va='center', fontfamily='sans-serif', zorder=5)
        line_y -= 1.8

# Entity 1: dim_departments (Top Left)
cols_dept = [
    ("INTEGER", "department_id", "PK"),
    ("TEXT", "department_name", ""),
    ("TEXT", "cost_center_code", ""),
    ("REAL", "standard_hourly_rate", ""),
    ("INTEGER", "fte_headcount", "")
]
draw_entity(ax, 5, 62, 26, 16, "dim_departments", cols_dept)

# Entity 2: dim_reasons (Top Right)
cols_reasons = [
    ("INTEGER", "reason_code", "PK"),
    ("TEXT", "category_type", ""),
    ("TEXT", "icd_chapter_code", ""),
    ("TEXT", "reason_description", ""),
    ("INTEGER", "is_preventable", ""),
    ("TEXT", "risk_level", "")
]
draw_entity(ax, 69, 60, 26, 18, "dim_reasons", cols_reasons)

# Entity 3: dim_employees (Bottom Left)
cols_emp = [
    ("INTEGER", "employee_id", "PK"),
    ("TEXT", "first_name", ""),
    ("TEXT", "last_name", ""),
    ("INTEGER", "department_id", "FK"),
    ("TEXT", "job_title", ""),
    ("REAL", "hourly_wage", ""),
    ("INTEGER", "age", ""),
    ("REAL", "service_tenure_years", ""),
    ("REAL", "distance_to_work_km", ""),
    ("REAL", "daily_transit_cost", ""),
    ("REAL", "bmi", ""),
    ("TEXT", "bmi_category", ""),
    ("INTEGER", "is_smoker", ""),
    ("INTEGER", "is_drinker", ""),
    ("INTEGER", "children_count", ""),
    ("INTEGER", "pet_count", ""),
    ("REAL", "contracted_weekly_hours", ""),
    ("TEXT", "hire_date", "")
]
draw_entity(ax, 5, 8, 26, 42, "dim_employees", cols_emp)

# Entity 4: fact_absence_events (Center / Bottom Right)
cols_fact = [
    ("INTEGER", "event_id", "PK"),
    ("INTEGER", "employee_id", "FK"),
    ("INTEGER", "reason_code", "FK"),
    ("TEXT", "absence_date", ""),
    ("TEXT", "absence_end_date", ""),
    ("TEXT", "day_of_week", ""),
    ("INTEGER", "month_of_year", ""),
    ("TEXT", "quarter", ""),
    ("INTEGER", "year", ""),
    ("REAL", "absent_hours", ""),
    ("REAL", "absent_days", ""),
    ("INTEGER", "is_holiday_adjacent", ""),
    ("INTEGER", "spell_id", ""),
    ("INTEGER", "is_unplanned", "")
]
draw_entity(ax, 42, 16, 29, 34, "fact_absence_events", cols_fact)

# Relationship Lines & Labels
# REL-01: dim_departments -> dim_employees (1:N)
ax.annotate("", xy=(18, 50), xytext=(18, 62),
            arrowprops=dict(arrowstyle="-|>", color="#1A365D", lw=2, mutation_scale=15), zorder=6)
ax.text(19, 56, "1 : N\n'employs'", fontsize=8, fontweight='bold', color="#2B6CB0", va='center', zorder=7)

# REL-02: dim_employees -> fact_absence_events (1:N)
ax.annotate("", xy=(42, 33), xytext=(31, 33),
            arrowprops=dict(arrowstyle="-|>", color="#1A365D", lw=2, mutation_scale=15), zorder=6)
ax.text(36.5, 35, "1 : N\n'incurs'", fontsize=8, fontweight='bold', color="#2B6CB0", ha='center', va='center', zorder=7)

# REL-03: dim_reasons -> fact_absence_events (1:N)
ax.annotate("", xy=(58, 50), xytext=(78, 60),
            arrowprops=dict(arrowstyle="-|>", color="#1A365D", lw=2, mutation_scale=15), zorder=6)
ax.text(71, 53, "1 : N\n'categorizes'", fontsize=8, fontweight='bold', color="#2B6CB0", ha='center', va='center', zorder=7)

# Legend Box
rect_leg = patches.FancyBboxPatch((72, 10), 23, 16, boxstyle="round,pad=0.2,rounding_size=0.5",
                                 linewidth=1, edgecolor='#CBD5E0', facecolor='#FFFFFF', zorder=3)
ax.add_patch(rect_leg)
ax.text(74, 23.5, "SCHEMA DESIGN LEGEND", fontsize=8.5, fontweight='bold', color='#1A365D')
ax.text(74, 20.5, "[PK] Primary Key (Unique Identifier)", fontsize=7.5, color='#C53030', fontweight='bold')
ax.text(74, 18.0, "[FK] Foreign Key (Referential Constraint)", fontsize=7.5, color='#2B6CB0', fontweight='bold')
ax.text(74, 15.5, "->  1-to-Many Directional Cardinality", fontsize=7.5, color='#2D3748')
ax.text(74, 13.0, "DB: SQLite 3 | App Engine: Streamlit", fontsize=7.5, color='#718096', fontstyle='italic')

plt.tight_layout()
os.makedirs("C:/HCL", exist_ok=True)
png_path = "C:/HCL/09_ER_Diagram.png"
plt.savefig(png_path, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
plt.close()
print("SUCCESSFULLY SAVED PNG:", png_path)

# ========================================================
# 2. GENERATE COMPREHENSIVE WORD DOCUMENT (09_ER_Diagram.docx)
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
r_title = p_title.add_run('ENTITY RELATIONSHIP (ER) DIAGRAM & LOGICAL SCHEMA')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Relational Data Modeling & Architectural Critique | Absenteeism Impact Analysis')
r_sub.font.name = 'Calibri'
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = RGBColor(43, 108, 176)
r_sub.font.italic = True

# Metadata Box Table
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ('Document Reference:', 'ERD-PRJ-AIA-01-v1.0'),
    ('Project ID & Title:', 'PRJ-AIA-01 | Absenteeism Impact Analysis'),
    ('Author / Student:', 'Kalpana (Roll No: 2500320140071) — MCA 2nd Yr DA, ABES'),
    ('Faculty Project Mentor:', 'Prof. Gunjan Agrawal'),
    ('Associated Graphic Asset:', '09_ER_Diagram.png (High-Resolution 300 DPI)'),
    ('Design Baseline:', 'Derived from Step 8 Database Design Specification')
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

# 1. Mermaid erDiagram Code
h1 = doc.add_heading('1. Formal Mermaid erDiagram Specification', level=1)
style_heading(h1)

p_mm_desc = doc.add_paragraph()
p_mm_desc.add_run('The logical schema below represents the complete data model, including table names, columns, data types, primary/foreign key designations, explicit cardinalities, and active verb relationship labels:')
p_mm_desc.runs[0].font.name = 'Calibri'
p_mm_desc.runs[0].font.size = Pt(10)

mermaid_code = '''erDiagram
    dim_departments ||--o{ dim_employees : "employs"
    dim_employees ||--o{ fact_absence_events : "incurs"
    dim_reasons ||--o{ fact_absence_events : "categorizes"

    dim_departments {
        int department_id PK
        text department_name
        text cost_center_code
        real standard_hourly_rate
        int fte_headcount
    }

    dim_reasons {
        int reason_code PK
        text category_type
        text icd_chapter_code
        text reason_description
        int is_preventable
        text risk_level
    }

    dim_employees {
        int employee_id PK
        text first_name
        text last_name
        int department_id FK
        text job_title
        real hourly_wage
        int age
        real service_tenure_years
        real distance_to_work_km
        real daily_transit_cost
        real bmi
        text bmi_category
        int is_smoker
        int is_drinker
        int children_count
        int pet_count
        real contracted_weekly_hours
        text hire_date
    }

    fact_absence_events {
        int event_id PK
        int employee_id FK
        int reason_code FK
        text absence_date
        text absence_end_date
        text day_of_week
        int month_of_year
        text quarter
        int year
        real absent_hours
        real absent_days
        int is_holiday_adjacent
        int spell_id
        int is_unplanned
    }'''

p_m_box = doc.add_paragraph()
p_m_box.paragraph_format.left_indent = Inches(0.2)
r_m_code = p_m_box.add_run(mermaid_code)
r_m_code.font.name = 'Courier New'
r_m_code.font.size = Pt(8.5)
r_m_code.font.color.rgb = RGBColor(26, 54, 93)

doc.add_paragraph()

# 2. Visual ER Diagram (Embedded Image)
h2 = doc.add_heading('2. High-Resolution Visual ER Diagram', level=1)
style_heading(h2)

p_img_desc = doc.add_paragraph()
p_img_desc.add_run('The following diagram illustrates the relational layout, cardinality connectors, and entity attributes:')
p_img_desc.runs[0].font.name = 'Calibri'
p_img_desc.runs[0].font.size = Pt(10)

p_pic = doc.add_paragraph()
p_pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_pic = p_pic.add_run()
run_pic.add_picture(png_path, width=Inches(6.2))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_cap = p_cap.add_run('Figure 9.1: Logical Entity-Relationship Diagram for Absenteeism Impact Analysis')
r_cap.font.name = 'Calibri'
r_cap.font.size = Pt(9)
r_cap.font.italic = True
r_cap.font.color.rgb = RGBColor(113, 128, 150)

doc.add_paragraph()

# 3. Plain-English Relationship Statements
h3 = doc.add_heading('3. Plain-English Relationship Narratives', level=1)
style_heading(h3)

rel_narratives = [
    ('Relationship 1 (dim_departments -> dim_employees):', 
     '\"One department employs zero, one, or many employees; every employee must belong to exactly one valid department.\" (1 : N cardinality, enforced via department_id).'),
    ('Relationship 2 (dim_employees -> fact_absence_events):', 
     '\"One employee incurs zero, one, or many discrete absence events over time; every absence event must be incurred by exactly one registered employee.\" (1 : N cardinality, enforced via employee_id).'),
    ('Relationship 3 (dim_reasons -> fact_absence_events):', 
     '\"One standardized absence reason categorizes zero, one, or many absence events; every recorded absence event must be categorized under exactly one classified reason code.\" (1 : N cardinality, enforced via reason_code).')
]

for title, stmt in rel_narratives:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(title + ' ')
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(10)
    r1.font.color.rgb = RGBColor(43, 108, 176)
    r2 = p.add_run(stmt)
    r2.font.name = 'Calibri'
    r2.font.size = Pt(10)

doc.add_paragraph()

# 4. Business Rules Enforced by the Design
h4 = doc.add_heading('4. Business Rules Enforced by the Schema Design', level=1)
style_heading(h4)

rules_data = [
    ('RUL-01: Referential Protection Against Orphan Records:', 
     'All foreign keys are configured with ON DELETE RESTRICT. A department cannot be dropped if active employees are assigned to it, and an employee record cannot be removed if historical absence events reference it. This protects audit integrity for past financial years.'),
    ('RUL-02: Bounded Operational Event Durations:', 
     'The fact table enforces CHECK(absent_hours > 0 AND absent_hours <= 40.0) and CHECK(absent_days > 0 AND absent_days <= 5.0). This prevents clerical data corruption (such as logging negative hours or an individual weekly event exceeding standard full-time capacity).'),
    ('RUL-03: Physiological Data Sanity Limits:', 
     'The employee dimension enforces CHECK(age BETWEEN 18 AND 70) and CHECK(bmi BETWEEN 10.0 AND 60.0). This guarantees that statistical regressions between health determinants and absence duration operate on valid clinical bounds.'),
    ('RUL-04: Controlled Vocabulary Enumerations:', 
     'Domain values are strictly bounded via SQL CHECK constraints: category_type IN (\'Medical\', \'Non-Medical\'), day_of_week IN (\'Monday\', \'Tuesday\', \'Wednesday\', \'Thursday\', \'Friday\'), and bmi_category IN (\'Underweight\', \'Normal\', \'Overweight\', \'Obese\').'),
    ('RUL-05: Non-Null Accounting Identifiers:', 
     'Fiscal attributes cost_center_code, department_name, and standard_hourly_rate carry NOT NULL and UNIQUE constraints, preventing unallocated salary leakage in finance reporting.')
]

for title, desc in rules_data:
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

# 5. Design Problems Identified & Engineering Fixes
h5 = doc.add_heading('5. Architectural Critique: Design Problems & Solutions', level=1)
style_heading(h5)

p_crit_intro = doc.add_paragraph()
p_crit_intro.add_run('A critical peer-review of early draft schemas identified four potential architectural anti-patterns, which were resolved in the final design:')
p_crit_intro.runs[0].font.name = 'Calibri'
p_crit_intro.runs[0].font.size = Pt(10)

critique_table = doc.add_table(rows=5, cols=4)
critique_table.alignment = WD_TABLE_ALIGNMENT.CENTER
cr_headers = ['Issue ID', 'Potential Architectural Anti-Pattern', 'Operational Risk / Problem', 'Architectural Fix Implemented in Step 8/9']
for j, h in enumerate(cr_headers):
    c = critique_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

critiques = [
    ('ISSUE-01', 'Direct Many-to-Many between Employee and Reason Master', 'Employees experience multiple medical reasons over time, and reasons apply to multiple employees. Modeling this as a direct relationship creates M:N ambiguity.', 'Resolved by introducing fact_absence_events as an associative fact table with foreign keys to both dim_employees and dim_reasons, converting M:N into two clean 1:N relationships.'),
    ('ISSUE-02', 'Redundant Storage of Computed Financial Loss in Fact Tables', 'Storing direct_wage_loss ($) and friction_loss ($) directly in fact_absence_events violates 3NF and risks data inconsistency if hourly wages are revised.', 'Resolved by omitting computed financial metrics from the base fact table and materializing them dynamically inside the analytical view v_absence_analysis_joined: ROUND(f.absent_hours * e.hourly_wage * 1.20, 2).'),
    ('ISSUE-03', 'Transitive Dependency in BMI Categorization', 'Storing height, weight, bmi, and bmi_category creates transitive dependencies (bmi_category depends on bmi, which depends on height/weight).', 'Resolved by standardizing on pre-calculated bmi and validated bmi_category in dim_employees while omitting raw height/weight columns to streamline analytic queries.'),
    ('ISSUE-04', 'Absence Spell Grouping & Bradford Computation Complexity', 'Computing the Bradford Factor requires knowing the count of distinct absence spells (S). Deriving spells on the fly via complex SQL date-lag window functions introduces latency.', 'Resolved by introducing an explicit spell_id attribute in fact_absence_events during ETL, enabling instantaneous COUNT(DISTINCT spell_id) aggregations for sub-second UI rendering.')
]

for i, row in enumerate(critiques, start=1):
    for j, val in enumerate(row):
        c = critique_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(9)
        if j == 0:
            r.bold = True
            r.font.color.rgb = RGBColor(197, 48, 48)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 50, 50, 60, 60)

# Save Word document
docx_path = "C:/HCL/09_ER_Diagram.docx"
doc.save(docx_path)
print("SUCCESSFULLY SAVED DOCX:", docx_path)
