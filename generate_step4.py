import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

wb = openpyxl.Workbook()

# Sheet 1: RTM
ws_rtm = wb.active
ws_rtm.title = "Traceability_Matrix"
ws_rtm.views.sheetView[0].showGridLines = True

headers = [
    "OQ", "HMW", "Req ID (BR)", "Requirement", "Stakeholder (SH)", 
    "Priority", "Status", "Feature", "DB table.column", 
    "User Story", "UI Screen", "Test Case (TC-F / TC-T)"
]

header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
thin_border = Border(
    left=Side(style='thin', color='CBD5E0'),
    right=Side(style='thin', color='CBD5E0'),
    top=Side(style='thin', color='CBD5E0'),
    bottom=Side(style='thin', color='CBD5E0')
)

ws_rtm.append(headers)
for col_idx in range(1, len(headers) + 1):
    cell = ws_rtm.cell(row=1, column=col_idx)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

rows_data = [
    ("OQ-01", "HMW-01", "BR-01", "The business needs to aggregate, monitor, and trend organizational absenteeism rates on monthly, quarterly, and annual horizons so that HR leadership can forecast cyclical workforce deficits and adjust strategic staffing policies.", "SH-01, SH-02", "Must", "Identified", "", "", "", "", ""),
    ("OQ-01", "HMW-01", "BR-12", "The business needs to export standardized workforce absenteeism and financial impact summary reports so that HR and finance teams can share findings during quarterly executive reviews.", "SH-01, SH-04", "Could", "Identified", "", "", "", "", ""),
    ("OQ-02", "HMW-02", "BR-02", "The business needs to categorize and rank total lost productive hours and absence incidence rates across departments and job roles so that operational line managers can identify operational bottlenecks and proactively arrange cross-training or shift backfills.", "SH-03", "Must", "Identified", "", "", "", "", ""),
    ("OQ-02", "HMW-02", "BR-11", "The business needs to identify and filter high-risk employees whose absence metrics exceed organizational threshold benchmarks so that HR and department managers can initiate standardized review protocols.", "SH-02, SH-03", "Could", "Identified", "", "", "", "", ""),
    ("OQ-03", "HMW-03", "BR-03", "The business needs to compute direct wage losses and replacement overtime expenses attributable to absent hours across business units so that finance executives can audit payroll leakage and establish realistic contingency budgets.", "SH-04", "Must", "Identified", "", "", "", "", ""),
    ("OQ-04", "HMW-04", "BR-04", "The business needs to categorize absence incidents using standardized medical and non-medical reason classifications across employee age brackets and service tenures so that occupational health teams can implement preventive wellness and ergonomic initiatives.", "SH-05", "Must", "Identified", "", "", "", "", ""),
    ("OQ-05", "HMW-05", "BR-05", "The business needs to compute the Bradford Factor index for every employee and classify scores into standardized operational disruption tiers so that HR operations can identify and address disruptive chronic short-term casual absentees.", "SH-02", "Must", "Identified", "", "", "", "", ""),
    ("OQ-06", "HMW-06", "BR-06", "The business needs to analyze the statistical relationship between employee commute distance / transit expenses and absence frequency so that HR leadership can evaluate the feasibility and targeting of hybrid work arrangements.", "SH-01, SH-05", "Should", "Identified", "", "", "", "", ""),
    ("OQ-07", "HMW-07", "BR-07", "The business needs to evaluate correlations between employee biometric and lifestyle factors (BMI, smoking, drinking, and family dependents) and absence duration so that wellness officers can target preventive health programs.", "SH-05", "Should", "Identified", "", "", "", "", ""),
    ("OQ-08", "HMW-08", "BR-08", "The business needs to identify day-of-the-week patterns and holiday proximity trends in unplanned absence occurrences so that line managers can detect habitual weekend extension patterns and enforce shift coverage.", "SH-03", "Must", "Identified", "", "", "", "", ""),
    ("OQ-09", "HMW-09", "BR-09", "The business needs to visualize the cumulative concentration curve of absenteeism across the employee population so that HR and line managers can determine whether a small minority of employees accounts for the majority of lost productivity.", "SH-02, SH-06", "Must", "Identified", "", "", "", "", ""),
    ("OQ-10", "HMW-10", "BR-10", "The business needs to simulate hypothetical financial savings under adjustable percentage reductions (5% to 20%) in preventable absenteeism so that executive leadership can evaluate the projected ROI of proposed wellness interventions.", "SH-01, SH-04", "Should", "Identified", "", "", "", "")
]

row_alt_fill = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
must_font = Font(name="Calibri", size=10, bold=True, color="C53030")
should_font = Font(name="Calibri", size=10, bold=True, color="2B6CB0")
could_font = Font(name="Calibri", size=10, bold=True, color="28A745")
status_font = Font(name="Calibri", size=10, bold=True, color="4A5568")

for row_idx, row_item in enumerate(rows_data, start=2):
    ws_rtm.append(row_item)
    is_alt = (row_idx % 2 == 0)
    for col_idx in range(1, len(row_item) + 1):
        cell = ws_rtm.cell(row=row_idx, column=col_idx)
        cell.border = thin_border
        if is_alt:
            cell.fill = row_alt_fill
        
        # Alignments
        if col_idx in [1, 2, 3, 5, 6, 7]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx == 4:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # Priority font color
        if col_idx == 6:
            if cell.value == "Must":
                cell.font = must_font
            elif cell.value == "Should":
                cell.font = should_font
            elif cell.value == "Could":
                cell.font = could_font
        elif col_idx == 7:
            cell.font = status_font
        else:
            cell.font = Font(name="Calibri", size=10)

# Set auto-filter on RTM sheet
ws_rtm.auto_filter.ref = f"A1:L{len(rows_data) + 1}"

# Set row heights
ws_rtm.row_dimensions[1].height = 28
for r in range(2, len(rows_data) + 2):
    ws_rtm.row_dimensions[r].height = 42

# Column widths
col_widths = {
    1: 10,   # OQ
    2: 12,   # HMW
    3: 14,   # Req ID (BR)
    4: 48,   # Requirement
    5: 16,   # Stakeholder
    6: 12,   # Priority
    7: 14,   # Status
    8: 18,   # Feature (Step 6)
    9: 22,   # DB table.column (Step 8)
    10: 18,  # User Story (Step 11)
    11: 18,  # UI Screen (Step 12)
    12: 24   # Test Case (Step 13)
}
for col_idx, width in col_widths.items():
    ws_rtm.column_dimensions[get_column_letter(col_idx)].width = width

# Sheet 2: Coverage_Analysis
ws_cov = wb.create_sheet(title="Coverage_Analysis")
ws_cov.views.sheetView[0].showGridLines = True

# Title
ws_cov.merge_cells("A1:E1")
title_cell = ws_cov["A1"]
title_cell.value = "REQUIREMENTS TRACEABILITY COVERAGE & GAP ANALYSIS"
title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
title_cell.fill = header_fill
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws_cov.row_dimensions[1].height = 32

# Section 1: Executive KPI Metrics
ws_cov.append([])
ws_cov.append(["Metric Category", "Total Count", "Mapped / Covered", "Orphan / Uncovered", "Coverage Rate"])
ws_cov.row_dimensions[3].height = 24
for col_idx in range(1, 6):
    c = ws_cov.cell(row=3, column=col_idx)
    c.fill = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center")

cov_metrics = [
    ("Operational Questions (OQ)", 10, 10, 0, "100.0%"),
    ("How Might We Statements (HMW)", 10, 10, 0, "100.0%"),
    ("Business Requirements (BR)", 12, 12, 0, "100.0%"),
    ("HMWs without BR", 0, 0, 0, "0 (Zero Gaps)"),
    ("BRs without HMW", 0, 0, 0, "0 (Zero Gaps)"),
    ("OQs Not Covered", 0, 0, 0, "0 (Zero Gaps)")
]

for row_idx, (cat, tot, mapped, uncov, rate) in enumerate(cov_metrics, start=4):
    ws_cov.append([cat, tot, mapped, uncov, rate])
    ws_cov.row_dimensions[row_idx].height = 20
    for col_idx in range(1, 6):
        c = ws_cov.cell(row=row_idx, column=col_idx)
        c.border = thin_border
        c.font = Font(name="Calibri", size=9.5)
        if col_idx in [2, 3, 4, 5]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
        if col_idx == 5:
            c.font = Font(name="Calibri", size=9.5, bold=True, color="28A745")

# Section 2: MoSCoW Allocation Analysis
ws_cov.append([])
ws_cov.append(["MoSCoW Priority", "Count of Requirements", "Percentage Share", "Curriculum Standard", "Compliance Status"])
header_row_2 = ws_cov.max_row
ws_cov.row_dimensions[header_row_2].height = 24
for col_idx in range(1, 6):
    c = ws_cov.cell(row=header_row_2, column=col_idx)
    c.fill = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center")

moscow_kpis = [
    ("Must Have", 7, "58.3%", "~60% Must-Have Target", "MET (Optimal)"),
    ("Should Have", 3, "25.0%", "~20-25% Target", "MET (Optimal)"),
    ("Could Have", 2, "16.7%", "~15-20% Target", "MET (Optimal)"),
    ("Won't Have", 0, "0.0%", "Deferred / Out of Scope", "MET (Optimal)")
]

for cat, cnt, pct, std, stat in moscow_kpis:
    ws_cov.append([cat, cnt, pct, std, stat])
    r_idx = ws_cov.max_row
    ws_cov.row_dimensions[r_idx].height = 20
    for col_idx in range(1, 6):
        c = ws_cov.cell(row=r_idx, column=col_idx)
        c.border = thin_border
        c.font = Font(name="Calibri", size=9.5)
        if col_idx in [2, 3, 4, 5]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
        if col_idx == 5:
            c.font = Font(name="Calibri", size=9.5, bold=True, color="28A745")

# Section 3: Traceability Lifecycle Roadmap
ws_cov.append([])
ws_cov.append(["Lifecycle Step", "Target Artifact", "RTM Columns Populated in that Step", "Current Step Status"])
header_row_3 = ws_cov.max_row
ws_cov.row_dimensions[header_row_3].height = 24
for col_idx in range(1, 5):
    c = ws_cov.cell(row=header_row_3, column=col_idx)
    c.fill = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center")

lifecycle_steps = [
    ("Step 4 (Current)", "Requirements Traceability Matrix", "OQ, HMW, Req ID (BR), Requirement, Stakeholder, Priority, Status", "COMPLETED (Identified)"),
    ("Step 6 (PRD)", "Product Requirements Document", "Feature (F-01 to F-xx)", "PENDING (Step 6)"),
    ("Step 8 (Database)", "Database Design Specification", "DB table.column mapping", "PENDING (Step 8)"),
    ("Step 11 (Agile)", "User Stories & Acceptance Criteria", "User Story (US-01 to US-xx)", "PENDING (Step 11)"),
    ("Step 12 (UI)", "UI Framework & Wireframes", "UI Screen (SCR-01 to SCR-xx)", "PENDING (Step 12)"),
    ("Step 13 (QA)", "Test Cases Specification", "Test Case (TC-F / TC-T)", "PENDING (Step 13)")
]

for step, art, cols, stat in lifecycle_steps:
    ws_cov.append([step, art, cols, stat])
    r_idx = ws_cov.max_row
    ws_cov.row_dimensions[r_idx].height = 20
    for col_idx in range(1, 5):
        c = ws_cov.cell(row=r_idx, column=col_idx)
        c.border = thin_border
        c.font = Font(name="Calibri", size=9.5)
        if col_idx in [1, 4]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
        if "COMPLETED" in stat:
            c.font = Font(name="Calibri", size=9.5, bold=True, color="28A745")
        else:
            c.font = Font(name="Calibri", size=9.5, italic=True, color="718096")

# Adjust column widths for Sheet 2
cov_col_widths = {1: 30, 2: 24, 3: 45, 4: 26, 5: 22}
for c_idx, w in cov_col_widths.items():
    ws_cov.column_dimensions[get_column_letter(c_idx)].width = w

# Save to C:\HCL\04_Traceability_Matrix.xlsx
os.makedirs("C:/HCL", exist_ok=True)
p_xlsx = "C:/HCL/04_Traceability_Matrix.xlsx"
wb.save(p_xlsx)
print("SUCCESSFULLY SAVED:", p_xlsx)
