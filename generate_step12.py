import os
import openpyxl
from openpyxl.styles import Font, Alignment
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = "C:/HCL"
STREAMLIT_DIR = os.path.join(BASE_DIR, ".streamlit")
SRC_DIR = os.path.join(BASE_DIR, "src")
DOCS_DIR = os.path.join(BASE_DIR, "docs")
os.makedirs(STREAMLIT_DIR, exist_ok=True)
os.makedirs(SRC_DIR, exist_ok=True)
os.makedirs(DOCS_DIR, exist_ok=True)

# ========================================================
# 1. CREATE .streamlit/config.toml
# ========================================================
config_toml_content = """[theme]
primaryColor = "#2563EB"
backgroundColor = "#F8FAFC"
secondaryBackgroundColor = "#FFFFFF"
textColor = "#0F172A"
font = "sans serif"

[server]
headless = true
enableCORS = false
enableXsrfProtection = true
maxUploadSize = 50

[browser]
gatherUsageStats = false
"""
with open(os.path.join(STREAMLIT_DIR, "config.toml"), "w", encoding="utf-8") as f:
    f.write(config_toml_content)
print("SUCCESSFULLY CREATED .streamlit/config.toml")

# ========================================================
# 2. CREATE src/ui.py (Tailwind CDN Components via st.iframe)
# ========================================================
ui_py_content = """import streamlit as st
import streamlit.components.v1 as components

TAILWIND_CDN = '<script src="https://cdn.tailwindcss.com"></script>'

def render_header(title: str, subtitle: str, badge_text: str = "Workforce Analytics v1.0"):
    \"\"\"Renders a modern, responsive corporate header banner using Tailwind CSS.\"\"\"
    html_code = f\"\"\"
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <style>body {{ margin: 0; padding: 0; background-color: transparent; font-family: system-ui, -apple-system, sans-serif; }}</style>
    </head>
    <body>
        <div class="bg-gradient-to-r from-blue-900 via-blue-800 to-indigo-900 text-white p-6 rounded-2xl shadow-lg border border-blue-700/50 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <div>
                <div class="flex items-center gap-2 mb-2">
                    <span class="bg-blue-600/60 text-blue-200 text-xs font-semibold px-3 py-1 rounded-full border border-blue-400/30 uppercase tracking-wider">{badge_text}</span>
                    <span class="text-blue-300 text-xs">PRJ-AIA-01 | ABES MCA 2nd Yr</span>
                </div>
                <h1 class="text-2xl md:text-3xl font-extrabold tracking-tight text-white">{title}</h1>
                <p class="text-blue-200 text-sm mt-1 max-w-2xl">{subtitle}</p>
            </div>
            <div class="bg-blue-800/40 border border-blue-400/20 px-4 py-2 rounded-xl text-right shrink-0">
                <span class="text-xs text-blue-300 block font-medium">Author / Lead Analyst</span>
                <span class="text-sm font-bold text-white block">Kalpana (2500320140071)</span>
            </div>
        </div>
    </body>
    </html>
    \"\"\"
    components.html(html_code, height=135)

def render_kpi_card(title: str, value: str, subtitle: str, trend: str = "", color_theme: str = "blue"):
    \"\"\"Renders a single Tailwind CSS metric card.\"\"\"
    color_map = {
        "blue": ("border-blue-500", "text-blue-600", "bg-blue-50"),
        "green": ("border-emerald-500", "text-emerald-600", "bg-emerald-50"),
        "red": ("border-rose-500", "text-rose-600", "bg-rose-50"),
        "purple": ("border-purple-500", "text-purple-600", "bg-purple-50"),
        "amber": ("border-amber-500", "text-amber-600", "bg-amber-50")
    }
    border_c, text_c, bg_c = color_map.get(color_theme, color_map["blue"])
    
    html_code = f\"\"\"
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <style>body {{ margin: 0; padding: 0; background-color: transparent; font-family: system-ui, -apple-system, sans-serif; }}</style>
    </head>
    <body>
        <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200/80 hover:shadow-md transition-shadow relative overflow-hidden border-l-4 {border_c}">
            <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">{title}</span>
                <span class="text-xs font-bold px-2 py-0.5 rounded-md {bg_c} {text_c}">{trend}</span>
            </div>
            <div class="text-3xl font-black text-slate-800 tracking-tight">{value}</div>
            <p class="text-xs text-slate-500 mt-2 font-medium">{subtitle}</p>
        </div>
    </body>
    </html>
    \"\"\"
    components.html(html_code, height=125)

def render_kpi_grid_4(cards_data):
    \"\"\"Renders 4 metric cards side by side in one responsive iframe container.\"\"\"
    card_htmls = ""
    color_map = {
        "blue": ("border-blue-500", "text-blue-600", "bg-blue-50"),
        "green": ("border-emerald-500", "text-emerald-600", "bg-emerald-50"),
        "red": ("border-rose-500", "text-rose-600", "bg-rose-50"),
        "purple": ("border-purple-500", "text-purple-600", "bg-purple-50"),
        "amber": ("border-amber-500", "text-amber-600", "bg-amber-50")
    }
    for title, val, sub, trend, c_theme in cards_data:
        border_c, text_c, bg_c = color_map.get(c_theme, color_map["blue"])
        card_htmls += f\"\"\"
        <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200/80 hover:shadow-md transition-shadow border-l-4 {border_c} flex-1 min-w-[220px]">
            <div class="flex items-center justify-between mb-1.5">
                <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">{title}</span>
                <span class="text-xs font-bold px-2 py-0.5 rounded-md {bg_c} {text_c}">{trend}</span>
            </div>
            <div class="text-2xl md:text-3xl font-black text-slate-800 tracking-tight">{val}</div>
            <p class="text-xs text-slate-500 mt-1.5 font-medium">{sub}</p>
        </div>
        \"\"\"
        
    full_html = f\"\"\"
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <style>body {{ margin: 0; padding: 0; background-color: transparent; font-family: system-ui, -apple-system, sans-serif; }}</style>
    </head>
    <body>
        <div class="flex flex-wrap md:flex-nowrap gap-4 w-full">
            {card_htmls}
        </div>
    </body>
    </html>
    \"\"\"
    components.html(full_html, height=130)

def render_insight_box(title: str, message: str, alert_type: str = "info"):
    \"\"\"Renders a styled callout insight box.\"\"\"
    type_map = {
        "info": ("bg-blue-50", "border-blue-300", "text-blue-900", "text-blue-700"),
        "warning": ("bg-amber-50", "border-amber-300", "text-amber-900", "text-amber-700"),
        "danger": ("bg-rose-50", "border-rose-300", "text-rose-900", "text-rose-700"),
        "success": ("bg-emerald-50", "border-emerald-300", "text-emerald-900", "text-emerald-700")
    }
    bg, border, text_title, text_body = type_map.get(alert_type, type_map["info"])
    
    html_code = f\"\"\"
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <style>body {{ margin: 0; padding: 0; background-color: transparent; font-family: system-ui, -apple-system, sans-serif; }}</style>
    </head>
    <body>
        <div class="{bg} border-l-4 {border} p-4 rounded-r-xl shadow-sm">
            <div class="flex items-center gap-2">
                <span class="font-bold text-sm {text_title}">{title}</span>
            </div>
            <p class="text-xs {text_body} mt-1 leading-relaxed">{message}</p>
        </div>
    </body>
    </html>
    \"\"\"
    components.html(html_code, height=95)
"""
with open(os.path.join(SRC_DIR, "ui.py"), "w", encoding="utf-8") as f:
    f.write(ui_py_content)
print("SUCCESSFULLY CREATED src/ui.py")

# ========================================================
# 3. GENERATE 12_UI_Framework.docx
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
r_title = p_title.add_run('UI FRAMEWORK & WIREFRAME SPECIFICATION')
r_title.font.name = 'Calibri'
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = RGBColor(26, 54, 93)

p_sub = doc.add_paragraph()
r_sub = p_sub.add_run('Streamlit Multipage Architecture + Tailwind CSS CDN Design System')
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
    ('UI Architecture:', 'Streamlit 1.42+ Multipage Navigation, Plotly 6.0, Tailwind CDN'),
    ('Design Tokens:', 'Corporate Slate / Deep Blue / Emerald / Rose Alert Tiers'),
    ('Deliverable File:', '12_UI_Framework.docx (accompanied by src/ui.py and .streamlit/config.toml)')
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

# 1. Screen Inventory (UI-01 to UI-08)
h1 = doc.add_heading('1. Screen Inventory & Component Mapping (UI-01 to UI-08)', level=1)
style_heading(h1)

screen_table = doc.add_table(rows=9, cols=6)
screen_table.alignment = WD_TABLE_ALIGNMENT.CENTER
sc_headers = ['UI ID', 'Screen Name', 'Core Purpose', 'Key UI Components', 'Active Filters', 'Linked Stories']
for j, h in enumerate(sc_headers):
    c = screen_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 60, 60)

screens_data = [
    ('UI-01', 'Executive Overview (app.py)', 'High-level macro health check & executive KPIs', 'Tailwind banner, 4 KPI cards, monthly trend chart, top 5 absent departments table', 'Date Range, Department Multiselect', 'US-03, US-06'),
    ('UI-02', 'Departmental Deep-Dive (pages/1_...)', 'Department and job role lost hours benchmarking', 'Stat pills, horizontal bar chart (dept lost hours), role-based breakdown, cost ledger', 'Dept Dropdown, Role Filter', 'US-04, US-06'),
    ('UI-03', 'Bradford & Pareto Insights (pages/2_...)', 'Non-linear disruption scoring & 80/20 Pareto analysis', 'Disruption tier badges, Bradford scatter plot (S vs D), Lorenz curve, review roster table', 'Tier Multiselect, Min Spell Slider', 'US-07, US-08, US-14'),
    ('UI-04', 'Health & Lifestyle EDA (pages/3_...)', 'ICD-10 medical diagnostics, BMI ANOVA, commute stats', 'ICD-10 treemap, BMI ANOVA boxplot with F-stat, commute scatter plot with OLS trendline', 'Medical Toggle, BMI Filter', 'US-09, US-10, US-11'),
    ('UI-05', 'Temporal Trends & Spikes (pages/4_...)', 'Seasonal patterns & Monday/Friday surge anomaly', 'Rolling 3-month MA line chart, day-of-week bar chart with Chi-Square stats, holiday surges', 'Quarter Multiselect, Year', 'US-01, US-05'),
    ('UI-06', 'Financial Loss Simulator (pages/5_...)', 'Direct wage loss, 1.20x friction, What-If ROI slider', 'Cost waterfall chart, 5%-20% reduction slider, projected net savings card, CSV export button', 'Preventable Toggle, Slider', 'US-12, US-13, US-15'),
    ('UI-07', 'Data Quality & Validation (pages/6_...)', 'Dataset integrity audit, raw vs clean comparisons', 'Integrity metric cards, outlier/null anomaly summary, raw CSV table viewer', 'Table Selector', 'US-01, US-15, US-16'),
    ('UI-08', 'About & System Metadata (pages/7_...)', 'Academic credentials, methodology, architecture doc', 'ABES badge card, architecture pipeline diagram, data dictionary download link', 'None', 'Documentation')
]

for i, row in enumerate(screens_data, start=1):
    for j, val in enumerate(row):
        c = screen_table.rows[i].cells[j]
        r = c.paragraphs[0].add_run(val)
        r.font.name = 'Calibri'
        r.font.size = Pt(8.5)
        if j in [0, 5]:
            r.bold = True
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 40, 40, 50, 50)

doc.add_paragraph()

# 2. ASCII Wireframes for Primary Screens
h2 = doc.add_heading('2. ASCII Wireframes for Primary Screens', level=1)
style_heading(h2)

wireframes = [
    ('UI-01: Executive Overview Dashboard (app.py)', '''+--------------------------------------------------------------------------------------------------+
|  [SIDEBAR FILTERS]           |  [TAILWIND HEADER BANNER: Absenteeism Impact Analysis]            |
|  - Date Range: Jan-Dec 2025  |  Sub-banner: Workforce Health & Financial Leakage Overview        |
|  - Department: [All v]       +-------------------------------------------------------------------+
|  - Category: [Medical v]     |  [KPI CARD 1]      [KPI CARD 2]      [KPI CARD 3]      [KPI CARD 4]     |
|                              |  Lost Hours        Total Cost ($)    Absence Rate      Severe Bradford  |
|  [RESET FILTERS BUTTON]      |  3,656.0 hrs       $212,586.26       1.92%             14 Employees     |
|                              +-------------------------------------------------------------------+
|                              |  [CHART: Monthly Absence Trend with 3-Month Moving Average (Plotly)]|
|                              |  Line plot: Contracted vs Lost Hours month-by-month               |
|                              +-------------------------------------------------------------------+
|                              |  [TABLE: Top 5 Impacted Departments and Cost Centers]             |
|                              |  Dept Name | Headcount | Hours Lost | Gross Loss ($) | Severity    |
+--------------------------------------------------------------------------------------------------+'''),

    ('UI-03: Bradford Disruption & Pareto Analysis (pages/2_Bradford_Disruption.py)', '''+--------------------------------------------------------------------------------------------------+
|  [SIDEBAR FILTERS]           |  [HEADER: Bradford Disruption & Pareto 80/20 Analysis]             |
|  - Risk Tier: [Low/Med/High] |  Formula: B = S^2 * D | Disproportionate Penalty on Short Spells   |
|  - Min Spell Count: [ 2  ]   +-------------------------------------------------------------------+
|                              |  [DISRUPTION TIER BADGES: Low (<50) | Med (50-249) | High | Severe]   |
|                              +-------------------------------------------------------------------+
|                              |  [CHART A: Bradford Scatter Plot] |  [CHART B: Pareto Lorenz Curve]|
|                              |  X: Spell Count (S)               |  X: Cumulative % of Employees  |
|                              |  Y: Total Days Absent (D)         |  Y: Cumulative % of Lost Hours |
|                              |  Color: Risk Tier | Size: Bradford |  Callout: Top 12% = 55.8% Hours |
|                              +-------------------------------------------------------------------+
|                              |  [ACTIONABLE ROSTER: Flagged High-Risk Employees (B >= 250)]       |
|                              |  Badge ID | Employee Name | Dept | Spells (S) | Days (D) | Score (B)|
+--------------------------------------------------------------------------------------------------+'''),

    ('UI-06: Financial Loss Accounting & What-If Simulator (pages/5_Financial_Simulator.py)', '''+--------------------------------------------------------------------------------------------------+
|  [SIDEBAR FILTERS]           |  [HEADER: Financial Loss Modeling & What-If ROI Simulator]        |
|  - Multiplier: [1.20x v]     |  Wage Leakage + 20% Overtime Overhead Compensation                |
|  - Dept Focus: [All v]       +-------------------------------------------------------------------+
|                              |  [FINANCIAL KPI 1]        [FINANCIAL KPI 2]     [SIMULATED SAVINGS] |
|                              |  Base Payroll Loss        Total Cost (+20%)     Target Savings      |
|                              |  $177,155.36              $212,586.26           +$21,258.63 (at 10%)|
|                              +-------------------------------------------------------------------+
|                              |  [INTERACTIVE SLIDER: Simulated Preventable Reduction (5% to 20%)] |
|                              |  [==== 10% ====--------------------------------------------------] |
|                              +-------------------------------------------------------------------+
|                              |  [CHART: Departmental Financial Impact Waterfall (Base vs Overtime)]|
|                              |  Stacked Bars: Engineering ($78k), Operations ($64k), Sales ($42k)  |
|                              +-------------------------------------------------------------------+
|                              |  [BUTTON: Download Executive Financial Summary Report (CSV / XLSX)] |
+--------------------------------------------------------------------------------------------------+''')
]

for title, ascii_wire in wireframes:
    p_t = doc.add_paragraph()
    r_t = p_t.add_run(title)
    r_t.bold = True
    r_t.font.name = 'Calibri'
    r_t.font.size = Pt(11)
    r_t.font.color.rgb = RGBColor(26, 54, 93)
    
    p_w = doc.add_paragraph()
    p_w.paragraph_format.left_indent = Inches(0.1)
    r_w = p_w.add_run(ascii_wire)
    r_w.font.name = 'Courier New'
    r_w.font.size = Pt(7.5)
    r_w.font.color.rgb = RGBColor(43, 108, 176)

doc.add_paragraph()

# 3. Best Chart for Each OQ
h3 = doc.add_heading('3. Best Chart Justification per Operational Question', level=1)
style_heading(h3)

oq_charts = [
    ('OQ-01 (Rate Trend)', 'Multi-Line Time-Series + 3-Month Moving Average', 
     'Multi-line plot shows exact monthly variance against total contracted capacity, while a smoothed 3-month moving average clarifies long-term trend trajectories.'),
    ('OQ-02 (Hotspots)', 'Horizontal Stacked Bar Chart (Sorted Descending)', 
     'Horizontal orientation permits long job-role and department titles without truncation, while stacked segments clearly illustrate the proportion of lost hours by role.'),
    ('OQ-03 (Financial Loss)', 'Waterfall / Stacked Cost Bar Chart', 
     'Visualizes how base hourly wage loss sequentially compounds with the 20% overtime friction multiplier into total fiscal impact per cost center.'),
    ('OQ-04 (Medical ICD-10)', 'Hierarchical Treemap / Nested Sunburst', 
     'Naturally represents the two-tier medical taxonomy (Medical Chapter -> Specific Condition) with tile areas scaled to total lost hours.'),
    ('OQ-05 (Bradford Factor)', '2D Scatter Bubble Plot (X: Spells, Y: Days, Size: Bradford Score)', 
     'Instantly isolates habitual short-term casual absentees (far right) from single long-term medical patients (top left), visually demonstrating non-linear B = S^2 * D dynamics.'),
    ('OQ-06 (Commute Correlation)', 'Scatter Plot with OLS Linear Regression Trendline', 
     'Displays individual data points alongside Pearson r and Spearman rho correlation slopes, demonstrating how distance > 30 km drives incidental absence rates.'),
    ('OQ-07 (BMI ANOVA)', 'Categorical Box Plot with Overlaid Jitter & ANOVA F-Statistic', 
     'Displays median, IQR, and statistical variance across Underweight, Normal, Overweight, and Obese categories, accompanied by post-hoc Tukey HSD markers.'),
    ('OQ-08 (Day-of-Week)', 'Grouped Column Chart with Baseline Average Line & Chi-Square Badge', 
     'Contrasts Monday and Friday absence surges directly against midweek averages with statistical significance p-value flags.'),
    ('OQ-09 (Pareto 80/20)', 'Dual-Axis Pareto Chart (Bar + Cumulative Percentage Line)', 
     'The definitive standard for demonstrating whether a critical 12-20% employee cohort accounts for the majority (55-80%) of lost organization time.'),
    ('OQ-10 (What-If Simulation)', 'Interactive Bullet / Gauge Chart with Sensitivity Area Plot', 
     'Provides immediate visual feedback as the user drags the 5%-20% reduction slider, displaying current cost against simulated savings goals.')
]

chart_table = doc.add_table(rows=11, cols=3)
chart_table.alignment = WD_TABLE_ALIGNMENT.CENTER
ct_headers = ['OQ ID & Inquiry', 'Selected Best Chart Type', 'Technical & Cognitive Justification']
for j, h in enumerate(ct_headers):
    c = chart_table.rows[0].cells[j]
    run = c.paragraphs[0].add_run(h)
    run.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(c, '2B6CB0')
    set_cell_margins(c, 70, 70, 70, 70)

for i, (oq, chart, just) in enumerate(oq_charts, start=1):
    c0 = chart_table.rows[i].cells[0]
    c1 = chart_table.rows[i].cells[1]
    c2 = chart_table.rows[i].cells[2]
    c0.paragraphs[0].add_run(oq).bold = True
    c1.paragraphs[0].add_run(chart)
    c2.paragraphs[0].add_run(just)
    for c in [c0, c1, c2]:
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(9)
        bg = 'FFFFFF' if i % 2 != 0 else 'F7FAFC'
        set_cell_background(c, bg)
        set_cell_margins(c, 40, 40, 50, 50)

# Save Word document
p_ui_doc = os.path.join(DOCS_DIR, "12_UI_Framework.docx")
p_ui_root = os.path.join(BASE_DIR, "12_UI_Framework.docx")
doc.save(p_ui_doc)
doc.save(p_ui_root)
print("SUCCESSFULLY SAVED 12_UI_Framework.docx")

# ========================================================
# 4. UPDATE RTM WITH UI SCREENS
# ========================================================
rtm_path = "C:/HCL/04_Traceability_Matrix.xlsx"
wb_rtm = openpyxl.load_workbook(rtm_path)
ws_rtm = wb_rtm["Traceability_Matrix"]

ui_mapping = {
    "BR-01": "UI-01, UI-05",
    "BR-12": "UI-01, UI-06, UI-07",
    "BR-02": "UI-01, UI-02",
    "BR-11": "UI-03",
    "BR-03": "UI-01, UI-06",
    "BR-04": "UI-04",
    "BR-05": "UI-03",
    "BR-06": "UI-04",
    "BR-07": "UI-04",
    "BR-08": "UI-05",
    "BR-09": "UI-03",
    "BR-10": "UI-06"
}

for row in range(2, ws_rtm.max_row + 1):
    br_id = ws_rtm.cell(row=row, column=3).value
    if br_id in ui_mapping:
        c_ui = ws_rtm.cell(row=row, column=11) # Column 11: UI Screen
        c_ui.value = ui_mapping[br_id]
        c_ui.font = Font(name="Courier New", size=8.5, color="1A365D")
        c_ui.alignment = Alignment(horizontal="center", vertical="center")

# Update Coverage_Analysis sheet in RTM
ws_cov = wb_rtm["Coverage_Analysis"]
for r in range(1, ws_cov.max_row + 1):
    v = ws_cov.cell(row=r, column=1).value
    if v and "Step 12" in str(v):
        c_stat = ws_cov.cell(row=r, column=4)
        c_stat.value = "COMPLETED (Mapped)"
        c_stat.font = Font(name="Calibri", size=9.5, bold=True, color="28A745")

wb_rtm.save(rtm_path)
print("SUCCESSFULLY UPDATED RTM WITH UI SCREENS:", rtm_path)
