import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

wb = openpyxl.Workbook()

# Style configurations
header_fill_navy = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
header_fill_blue = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
alt_row_fill = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
header_font_white = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
regular_font = Font(name="Calibri", size=9.5)
bold_font = Font(name="Calibri", size=9.5, bold=True)
code_font = Font(name="Courier New", size=9)
thin_border = Border(
    left=Side(style='thin', color='CBD5E0'),
    right=Side(style='thin', color='CBD5E0'),
    top=Side(style='thin', color='CBD5E0'),
    bottom=Side(style='thin', color='CBD5E0')
)

def style_table_header(ws, row_idx, num_cols, fill=header_fill_navy):
    ws.row_dimensions[row_idx].height = 26
    for c in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = fill
        cell.font = header_font_white
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def apply_row_styles(ws, start_row, end_row, num_cols, center_cols=[]):
    for r in range(start_row, end_row + 1):
        ws.row_dimensions[r].height = 20
        is_alt = (r % 2 == 0)
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.font = regular_font
            if is_alt:
                cell.fill = alt_row_fill
            if c in center_cols:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

# ==================== SHEET 1: OVERVIEW ====================
ws_ov = wb.active
ws_ov.title = "Overview"
ws_ov.views.sheetView[0].showGridLines = True

ws_ov.merge_cells("A1:E1")
title_cell = ws_ov["A1"]
title_cell.value = "RELATIONAL DATABASE DESIGN SPECIFICATION"
title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
title_cell.fill = header_fill_navy
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws_ov.row_dimensions[1].height = 32

meta_info = [
    ("Project Title:", "PRJ-AIA-01 | Absenteeism Impact Analysis: Productivity & Financial Loss Modeling"),
    ("Author / Student:", "Kalpana (Roll No: 2500320140071) — MCA 2nd Year Data Analytics, ABES"),
    ("Faculty Mentor:", "Prof. Gunjan Agrawal — Dept. of Computer Applications"),
    ("Database Management System:", "SQLite 3 (Embedded SQL Engine, ACID-compliant, Zero-Config)"),
    ("File Location:", "data/app.db (Repository Root)"),
    ("Connection Protocol:", "sqlite3.connect('file:data/app.db?mode=ro', uri=True)"),
    ("Schema Design Paradigm:", "Normalized Star/Snowflake hybrid with 3 Dimensions, 1 Fact, and 1 Analysis View"),
    ("Referential Integrity:", "Enforced via PRAGMA foreign_keys = ON; Cascade-restricted deletions"),
    ("Estimated Storage Size:", "~2.4 MB (data/app.db pre-seeded with 1,000+ attendance events)")
]

ws_ov.append([])
for k, v in meta_info:
    ws_ov.append([k, v])
    r_idx = ws_ov.max_row
    ws_ov.row_dimensions[r_idx].height = 20
    ws_ov.cell(row=r_idx, column=1).font = bold_font
    ws_ov.cell(row=r_idx, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws_ov.cell(row=r_idx, column=2).font = regular_font
    ws_ov.cell(row=r_idx, column=2).alignment = Alignment(horizontal="left", vertical="center")

ws_ov.column_dimensions["A"].width = 28
ws_ov.column_dimensions["B"].width = 75

# ==================== SHEET 2: TABLES SUMMARY ====================
ws_ts = wb.create_sheet(title="Tables_Summary")
ws_ts.views.sheetView[0].showGridLines = True

ts_headers = ["Table Name", "Table Type", "Granularity (One Row Per...)", "Primary Key", "Foreign Keys", "Column Count", "Planned Rows", "Core Purpose in Analytics"]
ws_ts.append(ts_headers)
style_table_header(ws_ts, 1, len(ts_headers), header_fill_navy)

ts_rows = [
    ("dim_departments", "Dimension (Master)", "One row per organizational department", "department_id", "None", 5, "6 rows", "Departmental cost centers, base salary budgets, and headcount baselines (BR-02, BR-03)."),
    ("dim_reasons", "Dimension (Master)", "One row per standardized absence reason code", "reason_code", "None", 6, "28 rows", "21 ICD-10 medical chapters and 7 administrative non-medical absence categories (BR-04)."),
    ("dim_employees", "Dimension (Master)", "One row per corporate employee profile", "employee_id", "department_id -> dim_departments", 18, "100-150 rows", "Employee demographics, hourly compensation, transit burden, and lifestyle indicators (BR-06, BR-07)."),
    ("fact_absence_events", "Fact Table (Core)", "One row per individual absence event / spell", "event_id", "employee_id -> dim_employees, reason_code -> dim_reasons", 14, "250-500 rows", "Atomic absence event timestamps, duration hours, spell sequences, and calendar patterns (BR-01, BR-05, BR-08)."),
    ("v_absence_analysis_joined", "Analytical SQL View", "One row per enriched absence event with financial metrics", "event_id", "Inherited from base tables", 32, "Dynamically Computed", "Pre-joined analytical dataset powering Streamlit views, Bradford scoring, and financial simulations.")
]

for row_item in ts_rows:
    ws_ts.append(row_item)
apply_row_styles(ws_ts, 2, len(ts_rows) + 1, len(ts_headers), center_cols=[2, 4, 6, 7])

ts_widths = [26, 20, 42, 16, 32, 14, 16, 50]
for c_idx, w in enumerate(ts_widths, start=1):
    ws_ts.column_dimensions[get_column_letter(c_idx)].width = w

# ==================== SHEET 3: DATA DICTIONARIES ====================
def create_dict_sheet(sheet_title, table_name, dict_data):
    ws_d = wb.create_sheet(title=sheet_title)
    ws_d.views.sheetView[0].showGridLines = True
    
    # Title
    ws_d.merge_cells("A1:G1")
    t = ws_d["A1"]
    t.value = f"DATA DICTIONARY: {table_name}"
    t.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    t.fill = header_fill_navy
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws_d.row_dimensions[1].height = 28
    
    headers = ["Column Name", "Data Type", "Key", "Nullable", "Allowed Values / Range", "Column Description & Business Meaning", "OQ / BR Supported"]
    ws_d.append(headers)
    style_table_header(ws_d, 2, len(headers), header_fill_blue)
    
    for row in dict_data:
        ws_d.append(row)
    apply_row_styles(ws_d, 3, len(dict_data) + 2, len(headers), center_cols=[2, 3, 4, 7])
    
    widths = [26, 16, 10, 12, 28, 48, 18]
    for c_idx, w in enumerate(widths, start=1):
        ws_d.column_dimensions[get_column_letter(c_idx)].width = w

# Dict 1: dim_departments
d_dept = [
    ("department_id", "INTEGER", "PK", "No", "1 to 6", "Unique surrogate identifier for the operational department", "OQ-02 / BR-02"),
    ("department_name", "TEXT", "None", "No", "Engineering, Sales, Operations, HR, Finance, Customer Support", "Formal departmental naming banner", "OQ-02 / BR-02"),
    ("cost_center_code", "TEXT", "UNIQUE", "No", "CC-ENG-101, CC-SLS-102, CC-OPS-103, etc.", "Standardized accounting ledger code for fiscal auditing", "OQ-03 / BR-03"),
    ("standard_hourly_rate", "REAL", "None", "No", "25.00 to 75.00", "Baseline departmental hourly labor rate ($)", "OQ-03 / BR-03"),
    ("fte_headcount", "INTEGER", "None", "No", ">= 1", "Total contracted full-time equivalent workforce in department", "OQ-01 / BR-01")
]
create_dict_sheet("Dict_Dim_Departments", "dim_departments", d_dept)

# Dict 2: dim_reasons
d_reasons = [
    ("reason_code", "INTEGER", "PK", "No", "0 to 28", "Unique numerical reason code mapped to ICD-10 taxonomy", "OQ-04 / BR-04"),
    ("category_type", "TEXT", "None", "No", "'Medical', 'Non-Medical'", "High-level classification separating illness from administrative leave", "OQ-04 / BR-04"),
    ("icd_chapter_code", "TEXT", "None", "Yes", "M00-M99 (Musculoskeletal), J00-J99 (Respiratory), F00-F99 (Mental), None", "WHO ICD-10 chapter reference code", "OQ-04 / BR-04"),
    ("reason_description", "TEXT", "None", "No", "Free text: e.g. Musculoskeletal disorders, Dental consultation, Physiotherapy", "Full clinical or administrative description of the absence cause", "OQ-04 / BR-04"),
    ("is_preventable", "INTEGER", "None", "No", "0 (Unavoidable) or 1 (Preventable via Ergonomics/Wellness)", "Flag indicating whether targeted corporate wellness can mitigate the absence", "OQ-10 / BR-10"),
    ("risk_level", "TEXT", "None", "No", "'Low', 'Medium', 'High'", "Clinical severity level of the health condition", "OQ-04 / BR-04")
]
create_dict_sheet("Dict_Dim_Reasons", "dim_reasons", d_reasons)

# Dict 3: dim_employees
d_emp = [
    ("employee_id", "INTEGER", "PK", "No", "1001 to 1150", "Unique corporate employee badge identifier", "OQ-05 / BR-05"),
    ("first_name", "TEXT", "None", "No", "Synthetic string names", "Employee given name (strictly synthetic)", "All"),
    ("last_name", "TEXT", "None", "No", "Synthetic string names", "Employee surname (strictly synthetic)", "All"),
    ("department_id", "INTEGER", "FK", "No", "1 to 6 (references dim_departments)", "Department mapping key", "OQ-02 / BR-02"),
    ("job_title", "TEXT", "None", "No", "Software Eng, Sales Executive, Logistics Specialist, etc.", "Designated organizational role", "OQ-02 / BR-02"),
    ("hourly_wage", "REAL", "None", "No", "18.50 to 95.00", "Base hourly compensation rate ($/hr)", "OQ-03 / BR-03"),
    ("age", "INTEGER", "None", "No", "21 to 65", "Employee chronological age in years", "OQ-07 / BR-07"),
    ("service_tenure_years", "REAL", "None", "No", "0.5 to 30.0", "Completed years of service at the firm", "OQ-04 / BR-04"),
    ("distance_to_work_km", "REAL", "None", "No", "1.0 to 65.0", "One-way commute distance from residence to workplace", "OQ-06 / BR-06"),
    ("daily_transit_cost", "REAL", "None", "No", "5.00 to 50.00", "Estimated daily transportation expenditure ($)", "OQ-06 / BR-06"),
    ("bmi", "REAL", "None", "No", "17.0 to 42.0", "Body Mass Index calculated as weight(kg) / height(m)^2", "OQ-07 / BR-07"),
    ("bmi_category", "TEXT", "None", "No", "'Underweight', 'Normal', 'Overweight', 'Obese'", "Standard WHO BMI category binning", "OQ-07 / BR-07"),
    ("is_smoker", "INTEGER", "None", "No", "0 (Non-smoker) or 1 (Active smoker)", "Social smoking lifestyle indicator", "OQ-07 / BR-07"),
    ("is_drinker", "INTEGER", "None", "No", "0 (Teetotaler) or 1 (Social drinker)", "Alcohol consumption lifestyle indicator", "OQ-07 / BR-07"),
    ("children_count", "INTEGER", "None", "No", "0 to 5", "Number of dependent children in household", "OQ-07 / BR-07"),
    ("pet_count", "INTEGER", "None", "No", "0 to 4", "Number of domestic pets requiring care", "OQ-07 / BR-07"),
    ("contracted_weekly_hours", "REAL", "None", "No", "40.0 (Standard Full-Time)", "Contracted work commitment per week", "OQ-01 / BR-01"),
    ("hire_date", "TEXT", "None", "No", "YYYY-MM-DD format", "Date of employment commencement", "OQ-04 / BR-04")
]
create_dict_sheet("Dict_Dim_Employees", "dim_employees", d_emp)

# Dict 4: fact_absence_events
d_fact = [
    ("event_id", "INTEGER", "PK", "No", "1 to 500+", "Unique surrogate key for each discrete absence occurrence", "All OQs"),
    ("employee_id", "INTEGER", "FK", "No", "References dim_employees.employee_id", "Employee who logged the absence event", "OQ-05 / BR-05"),
    ("reason_code", "INTEGER", "FK", "No", "References dim_reasons.reason_code", "Primary classified reason for the absence", "OQ-04 / BR-04"),
    ("absence_date", "TEXT", "None", "No", "YYYY-MM-DD format", "Starting calendar date of the absence event", "OQ-01 / BR-01"),
    ("absence_end_date", "TEXT", "None", "No", "YYYY-MM-DD format", "Ending calendar date of the absence event", "OQ-05 / BR-05"),
    ("day_of_week", "TEXT", "None", "No", "'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'", "Calendar weekday of the absence occurrence", "OQ-08 / BR-08"),
    ("month_of_year", "INTEGER", "None", "No", "1 to 12", "Calendar month integer (1 = Jan, 12 = Dec)", "OQ-01 / BR-01"),
    ("quarter", "TEXT", "None", "No", "'Q1', 'Q2', 'Q3', 'Q4'", "Fiscal quarter of the absence event", "OQ-01 / BR-01"),
    ("year", "INTEGER", "None", "No", "2025 to 2026", "Calendar year timestamp", "OQ-01 / BR-01"),
    ("absent_hours", "REAL", "None", "No", "1.0 to 40.0", "Total working hours lost during this absence event", "OQ-02 / BR-02"),
    ("absent_days", "REAL", "None", "No", "0.125 to 5.0 (calculated as absent_hours / 8.0)", "Total equivalent workdays lost during event", "OQ-05 / BR-05"),
    ("is_holiday_adjacent", "INTEGER", "None", "No", "0 or 1", "Flag indicating occurrence adjacent to a weekend or public holiday", "OQ-08 / BR-08"),
    ("spell_id", "INTEGER", "None", "No", "Sequential integer per employee", "Absence spell sequence number used to compute Bradford Factor (S)", "OQ-05 / BR-05"),
    ("is_unplanned", "INTEGER", "None", "No", "0 (Scheduled) or 1 (Emergency Unplanned)", "Indicator of sudden short-notice absence notice", "OQ-05 / BR-05")
]
create_dict_sheet("Dict_Fact_Absence_Events", "fact_absence_events", d_fact)

# Dict 5: v_absence_analysis_joined (32 columns)
d_view = [
    ("event_id", "INTEGER", "View Key", "No", "1 to 500+", "Surrogate identifier from fact table", "All OQs"),
    ("employee_id", "INTEGER", "FK Ref", "No", "1001 to 1150", "Employee identifier", "OQ-05 / BR-05"),
    ("employee_full_name", "TEXT", "Computed", "No", "String (first_name || ' ' || last_name)", "Concatenated full employee name", "All"),
    ("department_id", "INTEGER", "FK Ref", "No", "1 to 6", "Department identifier", "OQ-02 / BR-02"),
    ("department_name", "TEXT", "Lookup", "No", "Engineering, Sales, etc.", "Department name string", "OQ-02 / BR-02"),
    ("cost_center_code", "TEXT", "Lookup", "No", "CC-ENG-101, etc.", "Department ledger code", "OQ-03 / BR-03"),
    ("job_title", "TEXT", "Lookup", "No", "Role title string", "Employee operational job role", "OQ-02 / BR-02"),
    ("hourly_wage", "REAL", "Lookup", "No", "18.50 to 95.00", "Base hourly wage ($/hr)", "OQ-03 / BR-03"),
    ("age", "INTEGER", "Lookup", "No", "21 to 65", "Employee age", "OQ-07 / BR-07"),
    ("service_tenure_years", "REAL", "Lookup", "No", "0.5 to 30.0", "Years of service tenure", "OQ-04 / BR-04"),
    ("distance_to_work_km", "REAL", "Lookup", "No", "1.0 to 65.0", "Commute distance (km)", "OQ-06 / BR-06"),
    ("daily_transit_cost", "REAL", "Lookup", "No", "5.00 to 50.00", "Daily commute transit expense ($)", "OQ-06 / BR-06"),
    ("bmi", "REAL", "Lookup", "No", "17.0 to 42.0", "Body Mass Index score", "OQ-07 / BR-07"),
    ("bmi_category", "TEXT", "Lookup", "No", "Normal, Overweight, Obese", "WHO BMI classification", "OQ-07 / BR-07"),
    ("is_smoker", "INTEGER", "Lookup", "No", "0 or 1", "Smoker flag", "OQ-07 / BR-07"),
    ("is_drinker", "INTEGER", "Lookup", "No", "0 or 1", "Drinker flag", "OQ-07 / BR-07"),
    ("children_count", "INTEGER", "Lookup", "No", "0 to 5", "Count of dependent children", "OQ-07 / BR-07"),
    ("pet_count", "INTEGER", "Lookup", "No", "0 to 4", "Count of domestic pets", "OQ-07 / BR-07"),
    ("reason_code", "INTEGER", "FK Ref", "No", "0 to 28", "Reason numerical key", "OQ-04 / BR-04"),
    ("category_type", "TEXT", "Lookup", "No", "'Medical' / 'Non-Medical'", "Reason high-level category", "OQ-04 / BR-04"),
    ("icd_chapter_code", "TEXT", "Lookup", "Yes", "M00-M99, etc.", "ICD-10 Chapter classification", "OQ-04 / BR-04"),
    ("reason_description", "TEXT", "Lookup", "No", "Full reason text", "Reason diagnosis or administrative description", "OQ-04 / BR-04"),
    ("is_preventable", "INTEGER", "Lookup", "No", "0 or 1", "Preventable absence indicator", "OQ-10 / BR-10"),
    ("absence_date", "TEXT", "Timestamp", "No", "YYYY-MM-DD", "Date absence initiated", "OQ-01 / BR-01"),
    ("day_of_week", "TEXT", "Lookup", "No", "Monday to Friday", "Weekday of absence", "OQ-08 / BR-08"),
    ("month_of_year", "INTEGER", "Lookup", "No", "1 to 12", "Calendar month integer", "OQ-01 / BR-01"),
    ("quarter", "TEXT", "Lookup", "No", "Q1 to Q4", "Fiscal quarter", "OQ-01 / BR-01"),
    ("year", "INTEGER", "Lookup", "No", "2025 to 2026", "Calendar year", "OQ-01 / BR-01"),
    ("absent_hours", "REAL", "Fact", "No", "1.0 to 40.0", "Lost productive hours during event", "OQ-02 / BR-02"),
    ("absent_days", "REAL", "Computed", "No", "absent_hours / 8.0", "Lost working days equivalent", "OQ-05 / BR-05"),
    ("direct_wage_loss", "REAL", "Computed", "No", "absent_hours * hourly_wage", "Direct payroll salary loss ($)", "OQ-03 / BR-03"),
    ("friction_loss_total", "REAL", "Computed", "No", "direct_wage_loss * 1.20", "Total financial impact including 20% overtime friction ($)", "OQ-03 / BR-03")
]
create_dict_sheet("Dict_View_Joined_Analysis", "v_absence_analysis_joined", d_view)

# ==================== SHEET 4: RELATIONSHIPS ====================
ws_rel = wb.create_sheet(title="Relationships")
ws_rel.views.sheetView[0].showGridLines = True

rel_headers = ["Relationship ID", "Parent Table (Primary Key)", "Child Table (Foreign Key)", "Join Key Column", "Cardinality", "Integrity Rule (ON DELETE / UPDATE)", "Business Semantics"]
ws_rel.append(rel_headers)
style_table_header(ws_rel, 1, len(rel_headers), header_fill_navy)

rel_rows = [
    ("REL-01", "dim_departments (department_id)", "dim_employees (department_id)", "department_id", "1 to Many (1:N)", "ON DELETE RESTRICT, ON UPDATE CASCADE", "Every employee must belong to exactly one valid department; departments with employees cannot be deleted."),
    ("REL-02", "dim_employees (employee_id)", "fact_absence_events (employee_id)", "employee_id", "1 to Many (1:N)", "ON DELETE RESTRICT, ON UPDATE CASCADE", "An employee can log multiple absence events; absence events cannot exist without an active employee."),
    ("REL-03", "dim_reasons (reason_code)", "fact_absence_events (reason_code)", "reason_code", "1 to Many (1:N)", "ON DELETE RESTRICT, ON UPDATE CASCADE", "Every absence event must reference a classified ICD-10 or administrative reason code.")
]

for row_item in rel_rows:
    ws_rel.append(row_item)
apply_row_styles(ws_rel, 2, len(rel_rows) + 1, len(rel_headers), center_cols=[1, 4, 5, 6])

rel_widths = [16, 32, 32, 20, 18, 38, 48]
for c_idx, w in enumerate(rel_widths, start=1):
    ws_rel.column_dimensions[get_column_letter(c_idx)].width = w

# ==================== SHEET 5: KPIS & MEASURES ====================
ws_kpi = wb.create_sheet(title="KPIs_Measures")
ws_kpi.views.sheetView[0].showGridLines = True

kpi_headers = ["KPI / Measure Name", "Mathematical / Business Formula", "SQL Aggregation Expression", "Unit", "Benchmark / Threshold", "Linked OQ & BR"]
ws_kpi.append(kpi_headers)
style_table_header(ws_kpi, 1, len(kpi_headers), header_fill_navy)

kpi_rows = [
    ("Absenteeism Rate (%)", "(Total Lost Hours / Total Contracted Work Hours) * 100", "ROUND((SUM(absent_hours) / (COUNT(DISTINCT employee_id) * 160.0)) * 100, 2)", "Percentage (%)", "Industry Benchmark: 2.5% - 3.5%", "OQ-01 / BR-01"),
    ("Direct Wage Loss ($)", "Sum of (Absent Hours * Employee Base Hourly Wage)", "SUM(absent_hours * hourly_wage)", "Currency (USD $)", "Budget Variance Tracking", "OQ-03 / BR-03"),
    ("Total Financial Loss ($)", "Direct Wage Loss * 1.20 (Overtime Replacement Multiplier)", "SUM(absent_hours * hourly_wage * 1.20)", "Currency (USD $)", "CFO Contingency Allocation", "OQ-03 / BR-03"),
    ("Bradford Factor Index (B)", "B = (S^2) * D over rolling 52-week period", "POWER(COUNT(DISTINCT spell_id), 2) * SUM(absent_days)", "Index Score", "Low < 50, Med 50-249, High 250-749, Sev >= 750", "OQ-05 / BR-05"),
    ("Departmental Loss Share (%)", "(Department Lost Hours / Total Company Lost Hours) * 100", "ROUND((SUM(dept_hours) / SUM(total_hours)) * 100, 2)", "Percentage (%)", "Alert if any single dept > 35%", "OQ-02 / BR-02"),
    ("Weekend Extension Surge (%)", "(Monday & Friday Absences / Midweek Absences) - 1.0", "ROUND(((mon_fri_cnt / 2.0) / (midweek_cnt / 3.0) - 1.0) * 100, 1)", "Percentage Delta (%)", "Flag if surge > +15.0%", "OQ-08 / BR-08"),
    ("Pareto Concentration Ratio", "Top 20% Employee Hours / Total Company Hours", "SUM(top_20_pct_hours) / SUM(all_hours)", "Ratio / Percentage", "Pareto Target: ~60% - 80%", "OQ-09 / BR-09"),
    ("Simulated Cost Savings ($)", "Total Preventable Financial Loss * (Reduction% / 100)", "SUM(preventable_loss) * (slider_val / 100.0)", "Currency (USD $)", "Target 10% Reduction: > $25,000", "OQ-10 / BR-10")
]

for row_item in kpi_rows:
    ws_kpi.append(row_item)
apply_row_styles(ws_kpi, 2, len(kpi_rows) + 1, len(kpi_headers), center_cols=[4, 6])

kpi_widths = [26, 44, 48, 18, 36, 18]
for c_idx, w in enumerate(kpi_widths, start=1):
    ws_kpi.column_dimensions[get_column_letter(c_idx)].width = w

# ==================== SHEET 6: SQLITE DDL ====================
ws_ddl = wb.create_sheet(title="SQLite_DDL")
ws_ddl.views.sheetView[0].showGridLines = True

ws_ddl.merge_cells("A1:C1")
t_ddl = ws_ddl["A1"]
t_ddl.value = "EXECUTABLE SQLITE DDL SCHEMA SCRIPT (data/schema.sql)"
t_ddl.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
t_ddl.fill = header_fill_navy
t_ddl.alignment = Alignment(horizontal="center", vertical="center")
ws_ddl.row_dimensions[1].height = 28

ddl_headers = ["Line No", "SQL Statement / Clause", "Object Created & Architectural Constraints"]
ws_ddl.append(ddl_headers)
style_table_header(ws_ddl, 2, len(ddl_headers), header_fill_blue)

ddl_script_lines = [
    (1, "-- Enable Foreign Key Referential Integrity Enforcement", "System PRAGMA setting"),
    (2, "PRAGMA foreign_keys = ON;", "Enforces FK constraints across all operations"),
    (3, "", ""),
    (4, "-- ========================================================", "Dimension Table 1"),
    (5, "-- Table: dim_departments", "Master Department Cost Centers"),
    (6, "-- ========================================================", ""),
    (7, "CREATE TABLE IF NOT EXISTS dim_departments (", "Table definition start"),
    (8, "    department_id INTEGER PRIMARY KEY AUTOINCREMENT,", "Surrogate PK"),
    (9, "    department_name TEXT NOT NULL UNIQUE,", "Department name with unique constraint"),
    (10, "    cost_center_code TEXT NOT NULL UNIQUE,", "Ledger code for accounting audit"),
    (11, "    standard_hourly_rate REAL NOT NULL CHECK(standard_hourly_rate > 0),", "Rate validation check (> 0)"),
    (12, "    fte_headcount INTEGER NOT NULL CHECK(fte_headcount > 0)", "Headcount validation check (> 0)"),
    (13, ");", "Table definition end"),
    (14, "", ""),
    (15, "-- ========================================================", "Dimension Table 2"),
    (16, "-- Table: dim_reasons", "Standardized ICD-10 & Administrative Reasons"),
    (17, "-- ========================================================", ""),
    (18, "CREATE TABLE IF NOT EXISTS dim_reasons (", "Table definition start"),
    (19, "    reason_code INTEGER PRIMARY KEY,", "Natural PK (0 to 28)"),
    (20, "    category_type TEXT NOT NULL CHECK(category_type IN ('Medical', 'Non-Medical')),", "Enumeration check constraint"),
    (21, "    icd_chapter_code TEXT,", "Optional ICD-10 chapter reference"),
    (22, "    reason_description TEXT NOT NULL,", "Detailed reason narrative"),
    (23, "    is_preventable INTEGER NOT NULL CHECK(is_preventable IN (0, 1)),", "Binary flag check (0 or 1)"),
    (24, "    risk_level TEXT NOT NULL CHECK(risk_level IN ('Low', 'Medium', 'High'))", "Risk level check constraint"),
    (25, ");", "Table definition end"),
    (26, "", ""),
    (27, "-- ========================================================", "Dimension Table 3"),
    (28, "-- Table: dim_employees", "Corporate Employee Profiles & Biometrics"),
    (29, "-- ========================================================", ""),
    (30, "CREATE TABLE IF NOT EXISTS dim_employees (", "Table definition start"),
    (31, "    employee_id INTEGER PRIMARY KEY,", "Corporate employee badge PK"),
    (32, "    first_name TEXT NOT NULL,", "Synthetic first name"),
    (33, "    last_name TEXT NOT NULL,", "Synthetic surname"),
    (34, "    department_id INTEGER NOT NULL,", "FK referencing dim_departments"),
    (35, "    job_title TEXT NOT NULL,", "Operational job role"),
    (36, "    hourly_wage REAL NOT NULL CHECK(hourly_wage > 0),", "Hourly wage check constraint (> 0)"),
    (37, "    age INTEGER NOT NULL CHECK(age BETWEEN 18 AND 70),", "Age validation range (18 to 70)"),
    (38, "    service_tenure_years REAL NOT NULL CHECK(service_tenure_years >= 0),", "Tenure non-negative check"),
    (39, "    distance_to_work_km REAL NOT NULL CHECK(distance_to_work_km >= 0),", "Commute distance non-negative check"),
    (40, "    daily_transit_cost REAL NOT NULL CHECK(daily_transit_cost >= 0),", "Transit cost non-negative check"),
    (41, "    bmi REAL NOT NULL CHECK(bmi BETWEEN 10.0 AND 60.0),", "Physiological BMI range constraint"),
    (42, "    bmi_category TEXT NOT NULL CHECK(bmi_category IN ('Underweight', 'Normal', 'Overweight', 'Obese')),", "WHO classification check"),
    (43, "    is_smoker INTEGER NOT NULL CHECK(is_smoker IN (0, 1)),", "Binary smoker flag"),
    (44, "    is_drinker INTEGER NOT NULL CHECK(is_drinker IN (0, 1)),", "Binary drinker flag"),
    (45, "    children_count INTEGER NOT NULL CHECK(children_count >= 0),", "Children count non-negative check"),
    (46, "    pet_count INTEGER NOT NULL CHECK(pet_count >= 0),", "Pet count non-negative check"),
    (47, "    contracted_weekly_hours REAL NOT NULL DEFAULT 40.0 CHECK(contracted_weekly_hours > 0),", "Default 40.0 hrs constraint"),
    (48, "    hire_date TEXT NOT NULL,", "ISO-8601 Date string (YYYY-MM-DD)"),
    (49, "    FOREIGN KEY (department_id) REFERENCES dim_departments(department_id) ON DELETE RESTRICT ON UPDATE CASCADE", "FK Constraint with Restricted Deletion"),
    (50, ");", "Table definition end"),
    (51, "", ""),
    (52, "-- ========================================================", "Fact Table"),
    (53, "-- Table: fact_absence_events", "Core Fact Table Logging Discrete Absences"),
    (54, "-- ========================================================", ""),
    (55, "CREATE TABLE IF NOT EXISTS fact_absence_events (", "Table definition start"),
    (56, "    event_id INTEGER PRIMARY KEY AUTOINCREMENT,", "Surrogate fact event PK"),
    (57, "    employee_id INTEGER NOT NULL,", "FK referencing dim_employees"),
    (58, "    reason_code INTEGER NOT NULL,", "FK referencing dim_reasons"),
    (59, "    absence_date TEXT NOT NULL,", "Start date timestamp (YYYY-MM-DD)"),
    (60, "    absence_end_date TEXT NOT NULL,", "End date timestamp (YYYY-MM-DD)"),
    (61, "    day_of_week TEXT NOT NULL CHECK(day_of_week IN ('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')),", "Weekday check constraint"),
    (62, "    month_of_year INTEGER NOT NULL CHECK(month_of_year BETWEEN 1 AND 12),", "Month integer constraint (1 to 12)"),
    (63, "    quarter TEXT NOT NULL CHECK(quarter IN ('Q1', 'Q2', 'Q3', 'Q4')),", "Quarter check constraint"),
    (64, "    year INTEGER NOT NULL CHECK(year >= 2020),", "Year integer constraint"),
    (65, "    absent_hours REAL NOT NULL CHECK(absent_hours > 0 AND absent_hours <= 40.0),", "Event duration bounds check"),
    (66, "    absent_days REAL NOT NULL CHECK(absent_days > 0 AND absent_days <= 5.0),", "Event days bounds check"),
    (67, "    is_holiday_adjacent INTEGER NOT NULL CHECK(is_holiday_adjacent IN (0, 1)),", "Binary holiday proximity flag"),
    (68, "    spell_id INTEGER NOT NULL CHECK(spell_id > 0),", "Sequential spell sequence counter"),
    (69, "    is_unplanned INTEGER NOT NULL CHECK(is_unplanned IN (0, 1)),", "Binary unplanned absence flag"),
    (70, "    FOREIGN KEY (employee_id) REFERENCES dim_employees(employee_id) ON DELETE RESTRICT ON UPDATE CASCADE,", "FK Constraint to Employee"),
    (71, "    FOREIGN KEY (reason_code) REFERENCES dim_reasons(reason_code) ON DELETE RESTRICT ON UPDATE CASCADE", "FK Constraint to Reason Master"),
    (72, ");", "Table definition end"),
    (73, "", ""),
    (74, "-- ========================================================", "Performance Optimization"),
    (75, "-- B-Tree Indexes for Sub-second Analytical Queries", "Index Definitions"),
    (76, "-- ========================================================", ""),
    (77, "CREATE INDEX IF NOT EXISTS idx_absence_employee ON fact_absence_events(employee_id);", "Speeds up employee group by & Bradford sums"),
    (78, "CREATE INDEX IF NOT EXISTS idx_absence_date ON fact_absence_events(absence_date);", "Speeds up time-series date filtering"),
    (79, "CREATE INDEX IF NOT EXISTS idx_absence_reason ON fact_absence_events(reason_code);", "Speeds up ICD-10 medical aggregation"),
    (80, "CREATE INDEX IF NOT EXISTS idx_employee_dept ON dim_employees(department_id);", "Speeds up departmental multi-table joins"),
    (81, "", ""),
    (82, "-- ========================================================", "Unified Analysis View"),
    (83, "-- View: v_absence_analysis_joined (32 Enriched Columns)", "Analytical View Definition"),
    (84, "-- ========================================================", ""),
    (85, "CREATE VIEW IF NOT EXISTS v_absence_analysis_joined AS", "View creation start"),
    (86, "SELECT", ""),
    (87, "    f.event_id,", "Fact Event Key"),
    (88, "    f.employee_id,", "Employee Foreign Key"),
    (89, "    (e.first_name || ' ' || e.last_name) AS employee_full_name,", "Computed Full Name"),
    (90, "    e.department_id,", "Department Foreign Key"),
    (91, "    d.department_name,", "Department Name Lookup"),
    (92, "    d.cost_center_code,", "Ledger Cost Center"),
    (93, "    e.job_title,", "Job Title Role"),
    (94, "    e.hourly_wage,", "Base Wage ($/hr)"),
    (95, "    e.age,", "Employee Age"),
    (96, "    e.service_tenure_years,", "Service Tenure (Yrs)"),
    (97, "    e.distance_to_work_km,", "Commute Distance (km)"),
    (98, "    e.daily_transit_cost,", "Daily Transit Cost ($)"),
    (99, "    e.bmi,", "Body Mass Index"),
    (100, "    e.bmi_category,", "BMI Category"),
    (101, "    e.is_smoker,", "Smoker Indicator"),
    (102, "    e.is_drinker,", "Drinker Indicator"),
    (103, "    e.children_count,", "Children Count"),
    (104, "    e.pet_count,", "Pet Count"),
    (105, "    f.reason_code,", "Reason Foreign Key"),
    (106, "    r.category_type,", "Medical vs Non-Medical"),
    (107, "    r.icd_chapter_code,", "ICD-10 Chapter Code"),
    (108, "    r.reason_description,", "Reason Narrative"),
    (109, "    r.is_preventable,", "Preventable Wellness Flag"),
    (110, "    f.absence_date,", "Absence Start Date"),
    (111, "    f.day_of_week,", "Weekday String"),
    (112, "    f.month_of_year,", "Calendar Month (1-12)"),
    (113, "    f.quarter,", "Fiscal Quarter (Q1-Q4)"),
    (114, "    f.year,", "Calendar Year"),
    (115, "    f.absent_hours,", "Lost Productive Hours"),
    (116, "    f.absent_days,", "Lost Productive Days"),
    (117, "    ROUND(f.absent_hours * e.hourly_wage, 2) AS direct_wage_loss,", "Direct Salary Leakage ($)"),
    (118, "    ROUND(f.absent_hours * e.hourly_wage * 1.20, 2) AS friction_loss_total", "Total Loss ($) with 1.20x Overtime Factor"),
    (119, "FROM fact_absence_events f", "Primary Fact Join Table"),
    (120, "JOIN dim_employees e ON f.employee_id = e.employee_id", "Inner Join to Employee Dimension"),
    (121, "JOIN dim_departments d ON e.department_id = d.department_id", "Inner Join to Department Dimension"),
    (122, "JOIN dim_reasons r ON f.reason_code = r.reason_code;", "Inner Join to Reason Dimension")
]

for l_no, sql_stmt, comment in ddl_script_lines:
    ws_ddl.append([l_no if l_no != "" else "", sql_stmt, comment])
    r_idx = ws_ddl.max_row
    ws_ddl.row_dimensions[r_idx].height = 18
    ws_ddl.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws_ddl.cell(row=r_idx, column=1).font = regular_font
    ws_ddl.cell(row=r_idx, column=2).font = code_font
    ws_ddl.cell(row=r_idx, column=3).font = regular_font
    ws_ddl.cell(row=r_idx, column=3).font = Font(name="Calibri", size=9, italic=True, color="718096")

ws_ddl.column_dimensions["A"].width = 10
ws_ddl.column_dimensions["B"].width = 85
ws_ddl.column_dimensions["C"].width = 45

# Save 08_Database_Design.xlsx
p_db_design = "C:/HCL/08_Database_Design.xlsx"
wb.save(p_db_design)
print("SUCCESSFULLY SAVED:", p_db_design)

# ==================== UPDATE RTM WITH DB TABLE.COLUMN ====================
rtm_path = "C:/HCL/04_Traceability_Matrix.xlsx"
wb_rtm = openpyxl.load_workbook(rtm_path)
ws_rtm = wb_rtm["Traceability_Matrix"]

# Map of BR ID to DB table.column
db_mapping = {
    "BR-01": "fact_absence_events.absent_hours, dim_employees.contracted_weekly_hours, fact_absence_events.month_of_year",
    "BR-12": "v_absence_analysis_joined.*, fact_absence_events.absent_hours, v_absence_analysis_joined.friction_loss_total",
    "BR-02": "dim_departments.department_name, dim_employees.job_title, fact_absence_events.absent_hours",
    "BR-11": "dim_employees.employee_id, dim_departments.department_name, v_absence_analysis_joined.friction_loss_total",
    "BR-03": "dim_employees.hourly_wage, fact_absence_events.absent_hours, v_absence_analysis_joined.friction_loss_total",
    "BR-04": "dim_reasons.reason_code, dim_reasons.icd_chapter_code, dim_employees.age, dim_employees.service_tenure_years",
    "BR-05": "fact_absence_events.spell_id, fact_absence_events.absent_days, dim_employees.employee_id",
    "BR-06": "dim_employees.distance_to_work_km, dim_employees.daily_transit_cost, fact_absence_events.event_id",
    "BR-07": "dim_employees.bmi, dim_employees.is_smoker, dim_employees.is_drinker, dim_employees.children_count",
    "BR-08": "fact_absence_events.day_of_week, fact_absence_events.is_holiday_adjacent, fact_absence_events.month_of_year",
    "BR-09": "dim_employees.employee_id, fact_absence_events.absent_hours, v_absence_analysis_joined.absent_hours",
    "BR-10": "dim_reasons.is_preventable, v_absence_analysis_joined.friction_loss_total, dim_employees.hourly_wage"
}

# Update column 9 (DB table.column) for each BR in RTM
for row in range(2, ws_rtm.max_row + 1):
    br_id = ws_rtm.cell(row=row, column=3).value
    if br_id in db_mapping:
        c_db = ws_rtm.cell(row=row, column=9)
        c_db.value = db_mapping[br_id]
        c_db.font = Font(name="Courier New", size=8.5, color="1A365D")
        c_db.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

# Update Coverage_Analysis sheet in RTM
ws_cov = wb_rtm["Coverage_Analysis"]
for r in range(1, ws_cov.max_row + 1):
    v = ws_cov.cell(row=r, column=1).value
    if v and "Step 8" in str(v):
        c_stat = ws_cov.cell(row=r, column=4)
        c_stat.value = "COMPLETED (Mapped)"
        c_stat.font = Font(name="Calibri", size=9.5, bold=True, color="28A745")

wb_rtm.save(rtm_path)
print("SUCCESSFULLY UPDATED RTM:", rtm_path)
