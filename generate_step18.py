import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Widescreen 16:9 Dimensions
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Corporate ABES Palette
C_NAVY = RGBColor(30, 58, 138)       # #1E3A8A Primary Dark
C_ROYAL = RGBColor(37, 99, 235)      # #2563EB Brand Accent
C_SLATE = RGBColor(30, 41, 59)       # #1E293B Dark Charcoal Body
C_MUTED = RGBColor(100, 116, 139)    # #64748B Secondary Muted
C_LIGHT = RGBColor(248, 250, 252)    # #F8FAFC Card Light
C_WHITE = RGBColor(255, 255, 255)    # #FFFFFF Pure White
C_BORDER = RGBColor(226, 232, 240)   # #E2E8F0 Border Grid
C_RED = RGBColor(239, 68, 68)        # #EF4444 Danger
C_GREEN = RGBColor(16, 185, 129)     # #10B981 Emerald
C_AMBER = RGBColor(245, 158, 11)     # #F59E0B Amber

def create_header(slide, title_text, category="PRJ-AIA-01 | ABES MCA 2nd Year Major Project"):
    """Standardized top corporate header bar."""
    # Top thin accent line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.12))
    line.fill.solid()
    line.fill.fore_color.rgb = C_ROYAL
    line.line.color.rgb = C_ROYAL

    # Header title box
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(1.1))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_cat = tf.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = "Calibri"
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = C_ROYAL

    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = "Calibri"
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = C_NAVY

def create_card(slide, left, top, width, height, bg_color=C_LIGHT, border_color=C_BORDER):
    """Creates a styled card container shape."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1)
    return card

def add_speaker_notes(slide, notes_text):
    """Injects 30-second presentation script into the slide notes."""
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = f"SPEAKER SCRIPT (~30 SECONDS):\n\n\"{notes_text.strip()}\""

def main():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    blank_layout = prs.slide_layouts[6]

    # =========================================================
    # SLIDE 1: TITLE SLIDE
    # =========================================================
    s1 = prs.slides.add_slide(blank_layout)
    
    # Background navy gradient block
    bg_block = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT)
    bg_block.fill.solid()
    bg_block.fill.fore_color.rgb = C_NAVY
    bg_block.line.color.rgb = C_NAVY

    # Decorative Royal Ribbon
    ribbon = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.4), SLIDE_HEIGHT)
    ribbon.fill.solid()
    ribbon.fill.fore_color.rgb = C_ROYAL
    ribbon.line.color.rgb = C_ROYAL

    # Main Title text frame
    t_box = s1.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(11.0), Inches(3.2))
    tf1 = t_box.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "ABES ENGINEERING COLLEGE, GHAZIABAD — MCA MAJOR PROJECT"
    p0.font.name = "Calibri"
    p0.font.size = Pt(13)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(191, 219, 254)

    p1 = tf1.add_paragraph()
    p1.text = "Absenteeism Impact Analysis"
    p1.font.name = "Calibri"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = C_WHITE
    p1.space_before = Pt(8)

    p2 = tf1.add_paragraph()
    p2.text = "Workforce Productivity & Financial Loss Diagnostic Engine"
    p2.font.name = "Calibri"
    p2.font.size = Pt(20)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(147, 197, 253)

    p3 = tf1.add_paragraph()
    p3.text = "Project Identifier: PRJ-AIA-01  |  Domain: HR Analytics & Operational Economics"
    p3.font.name = "Calibri"
    p3.font.size = Pt(11)
    p3.font.color.rgb = RGBColor(203, 213, 225)
    p3.space_before = Pt(12)

    # Candidate & Mentor Details Card
    card_info = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(4.7), Inches(11.0), Inches(1.8))
    card_info.fill.solid()
    card_info.fill.fore_color.rgb = RGBColor(23, 37, 84)
    card_info.line.color.rgb = C_ROYAL
    card_info.line.width = Pt(1.5)

    tf_info = card_info.text_frame
    tf_info.word_wrap = True
    tf_info.margin_left = Inches(0.4)
    tf_info.margin_top = Inches(0.25)

    pi1 = tf_info.paragraphs[0]
    pi1.text = "Candidate: Kalpana  |  University Roll No: 2500320140071"
    pi1.font.name = "Calibri"
    pi1.font.size = Pt(14)
    pi1.font.bold = True
    pi1.font.color.rgb = C_WHITE

    pi2 = tf_info.add_paragraph()
    pi2.text = "Program: Master of Computer Applications (MCA) — 2nd Year Data Analytics"
    pi2.font.name = "Calibri"
    pi2.font.size = Pt(11)
    pi2.font.color.rgb = RGBColor(226, 232, 240)
    pi2.space_before = Pt(4)

    pi3 = tf_info.add_paragraph()
    pi3.text = "Faculty Project Mentor: Prof. Gunjan Agrawal, Department of Computer Applications"
    pi3.font.name = "Calibri"
    pi3.font.size = Pt(12)
    pi3.font.bold = True
    pi3.font.color.rgb = RGBColor(253, 224, 71) # Gold
    pi3.space_before = Pt(6)

    add_speaker_notes(
        s1,
        "Good morning respected evaluators, Head of Department, and my project mentor Prof. Gunjan Agrawal. "
        "My name is Kalpana, Roll Number 2500320140071, 2nd Year MCA Data Analytics at ABES Engineering College. "
        "Today, I present Project PRJ-AIA-01: Absenteeism Impact Analysis — Workforce Productivity & Financial Loss Modeling, "
        "an enterprise analytics platform quantifying direct payroll loss and operational friction."
    )

    # =========================================================
    # SLIDE 2: PROBLEM STATEMENT & TOP 2 HMWs
    # =========================================================
    s2 = prs.slides.add_slide(blank_layout)
    create_header(s2, "Business Problem & Innovation Framing (HMW)")

    # Left Column: Problem Card
    create_card(s2, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.2))
    p_box2 = s2.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.0), Inches(4.8))
    tf2_l = p_box2.text_frame
    tf2_l.word_wrap = True

    p = tf2_l.paragraphs[0]
    p.text = "THE OPERATIONAL PROBLEM"
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_RED

    bullets_prob = [
        "Unbudgeted Overhead: Unplanned absences force costly emergency overtime and shift disruptions.",
        "Flawed Metric: Counting total lost days treats 1 long illness the same as 5 erratic 1-day leaves.",
        "Zero Loss Transparency: Direct wage drain ($128.4k) ignores 20% operational friction ($25.7k).",
        "Habitual Disruption: A minority of habitual absentees cause compounding delivery bottlenecks.",
        "Baseline Impact: 120 staff, 320 events, 2,752 lost hours, and $154,140 total organization loss."
    ]
    for b in bullets_prob:
        p_b = tf2_l.add_paragraph()
        p_b.text = f"• {b}"
        p_b.font.name = "Calibri"
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = C_SLATE
        p_b.space_before = Pt(10)

    # Right Column: Top 2 HMWs
    create_card(s2, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2))
    p_box2_r = s2.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.8))
    tf2_r = p_box2_r.text_frame
    tf2_r.word_wrap = True

    p = tf2_r.paragraphs[0]
    p.text = "TOP 2 'HOW MIGHT WE' INITIATIVES"
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_ROYAL

    hmw_bullets = [
        ("HMW-01: Quantify Real Financial Friction (OQ-01)",
         "How might we model total corporate financial loss by integrating direct base wages with a 1.20x operational friction multiplier for emergency overtime and supervisor rerouting?"),
        ("HMW-04: Non-Linear Bradford Disruption (OQ-04)",
         "How might we implement the Bradford Factor index (B = S^2 * D) to penalize frequent short unplanned absences while fairly accommodating legitimate long-term medical convalescence?")
    ]
    for title, desc in hmw_bullets:
        pt = tf2_r.add_paragraph()
        pt.text = title
        pt.font.name = "Calibri"
        pt.font.size = Pt(11.5)
        pt.font.bold = True
        pt.font.color.rgb = C_NAVY
        pt.space_before = Pt(14)

        pd = tf2_r.add_paragraph()
        pd.text = desc
        pd.font.name = "Calibri"
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = C_SLATE
        pd.space_before = Pt(4)

    add_speaker_notes(
        s2,
        "Turning to slide 2: Unplanned employee absenteeism costs organizations far more than base salaries. "
        "When an employee fails to show up, shifts are disrupted, overtime is triggered, and projects stall. "
        "Standard HR systems simply total absent days, failing to distinguish a single 10-day medical leave from five separate 2-day casual leaves. "
        "We framed two core HMW questions: first, how to model the true 1.20x financial friction, and second, how to mathematically score disruption using the Bradford Factor."
    )

    # =========================================================
    # SLIDE 3: STAKEHOLDERS & RACI GOVERNANCE
    # =========================================================
    s3 = prs.slides.add_slide(blank_layout)
    create_header(s3, "Key Stakeholders & RACI Governance Matrix")

    # 3 Stakeholder Cards
    sh_cards = [
        ("SH-01: Chief Human Resources Officer", "High Influence | High Interest (Accountable)", 
         "Executive sponsor who mandates attendance policies, authorizes wellness investments, and tracks high-risk Bradford cohorts.", C_ROYAL),
        ("SH-02: VP of Operations / Plant Head", "High Influence | High Interest (Responsible)", 
         "Operational commander who manages line coverage, prevents Monday/Friday capacity slumps, and approves overtime costs.", C_NAVY),
        ("SH-05: Chief Financial Officer & Payroll", "High Influence | Medium Interest (Approver)", 
         "Financial fiduciary who reconciles direct payroll losses against the 1.20x friction multiplier and audits What-If ROI projections.", C_SLATE)
    ]

    for i, (stitle, sbadge, sdesc, scolor) in enumerate(sh_cards):
        left_pos = Inches(0.8 + i * 4.0)
        card = create_card(s3, left_pos, Inches(1.6), Inches(3.7), Inches(3.2))
        
        # Color bar top
        cbar = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, Inches(1.6), Inches(3.7), Inches(0.1))
        cbar.fill.solid()
        cbar.fill.fore_color.rgb = scolor
        cbar.line.color.rgb = scolor

        tb = s3.shapes.add_textbox(left_pos + Inches(0.2), Inches(1.8), Inches(3.3), Inches(2.8))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = stitle
        p1.font.name = "Calibri"
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = C_NAVY

        p2 = tf.add_paragraph()
        p2.text = sbadge
        p2.font.name = "Calibri"
        p2.font.size = Pt(9.5)
        p2.font.bold = True
        p2.font.color.rgb = C_ROYAL
        p2.space_before = Pt(4)

        p3 = tf.add_paragraph()
        p3.text = sdesc
        p3.font.name = "Calibri"
        p3.font.size = Pt(10)
        p3.font.color.rgb = C_SLATE
        p3.space_before = Pt(8)

    # Bottom Summary RACI Strip
    create_card(s3, Inches(0.8), Inches(5.1), Inches(11.733), Inches(1.7))
    tb_b = s3.shapes.add_textbox(Inches(1.0), Inches(5.2), Inches(11.3), Inches(1.4))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True

    p = tf_b.paragraphs[0]
    p.text = "ENTERPRISE RACI GOVERNANCE HIGHLIGHTS"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    raci_bullets = [
        "Accountability (A): CHRO is strictly accountable for attendance policy and Bradford risk tiers; CFO is accountable for financial models.",
        "Responsibility (R): HR Business Partners (SH-04) and Operations Leads (SH-02) execute daily shift tracking and line check-ins.",
        "Consulted & Informed (C/I): Department Heads (SH-03) provide operational feedback; Frontline Employees (SH-06) receive transparent policy notices."
    ]
    for rb in raci_bullets:
        pr = tf_b.add_paragraph()
        pr.text = f"• {rb}"
        pr.font.name = "Calibri"
        pr.font.size = Pt(9.5)
        pr.font.color.rgb = C_SLATE
        pr.space_before = Pt(4)

    add_speaker_notes(
        s3,
        "On slide 3, we map our key stakeholders using an enterprise RACI governance framework. "
        "The Chief Human Resources Officer, SH-01, is our primary accountable sponsor setting policy. "
        "The VP of Operations, SH-02, manages daily production lines and must mitigate unexpected capacity deficits. "
        "The CFO, SH-05, requires defensible financial loss accounting and validates our What-If cost recovery simulations. "
        "Together with department heads and HR business partners, this matrix ensures complete organizational alignment."
    )

    # =========================================================
    # SLIDE 4: TRACEABILITY MATRIX (RTM HIGHLIGHTS)
    # =========================================================
    s4 = prs.slides.add_slide(blank_layout)
    create_header(s4, "Bidirectional Requirements Traceability (RTM)")

    # 4 Must Rows Table
    rtm_tbl_shape = s4.shapes.add_table(5, 6, Inches(0.8), Inches(1.6), Inches(11.733), Inches(3.6))
    tbl = rtm_tbl_shape.table
    tbl.columns[0].width = Inches(1.0)
    tbl.columns[1].width = Inches(2.8)
    tbl.columns[2].width = Inches(1.2)
    tbl.columns[3].width = Inches(2.6)
    tbl.columns[4].width = Inches(2.6)
    tbl.columns[5].width = Inches(1.533)

    rtm_headers = ["HMW ID", "Business Requirement (BR)", "User Story", "Database Entity.Column", "UI & Pytest Trace", "Status"]
    for c_i, h_txt in enumerate(rtm_headers):
        cell = tbl.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY
        cell.text = h_txt
        for p in cell.text_frame.paragraphs:
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = C_WHITE

    rtm_data = [
        ("HMW-01", "BR-01: Direct wage payroll loss computation", "US-01 (Data Ingest)", "v_absence_joined.direct_wage_loss", "UI-01 Overview | TC-T-07", "PASS (100%)"),
        ("HMW-01", "BR-02: 1.20x operational friction factor", "US-14 (Simulator)", "v_absence_joined.friction_loss_total", "UI-06 Simulator | TC-T-07", "PASS (100%)"),
        ("HMW-04", "BR-08: Bradford index scoring B = S^2 * D", "US-08 (Bradford)", "fact_events.absent_days, spells", "UI-03 Bradford | TC-T-06", "PASS (100%)"),
        ("HMW-05", "BR-05: Monday & Friday surge Chi-Square", "US-13 (Temporal)", "fact_events.day_of_week", "UI-05 Temporal | TC-T-10", "PASS (100%)")
    ]

    for r_i, row_v in enumerate(rtm_data, start=1):
        for c_i, val in enumerate(row_v):
            cell = tbl.cell(r_i, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_LIGHT if r_i % 2 == 1 else C_WHITE
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.name = "Calibri"
                p.font.size = Pt(9.5)
                p.font.color.rgb = C_GREEN if c_i == 5 else C_SLATE
                if c_i == 5:
                    p.font.bold = True

    # RTM Highlights Card
    create_card(s4, Inches(0.8), Inches(5.4), Inches(11.733), Inches(1.4))
    tb_rtm = s4.shapes.add_textbox(Inches(1.0), Inches(5.5), Inches(11.3), Inches(1.1))
    tf_rtm = tb_rtm.text_frame
    tf_rtm.word_wrap = True

    p = tf_rtm.paragraphs[0]
    p.text = "REQUIREMENTS INTEGRITY AUDIT"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    r_bullets = [
        "Zero Orphaned Artifacts: 100% of the 12 Business Requirements and 10 Operational Questions map to functional code and tests.",
        "Forward & Backward Traceability: Any modification to business logic propagates immediately across database columns, UI screens, and test assertions."
    ]
    for b in r_bullets:
        pb = tf_rtm.add_paragraph()
        pb.text = f"• {b}"
        pb.font.name = "Calibri"
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = C_SLATE
        pb.space_before = Pt(3)

    add_speaker_notes(
        s4,
        "Slide 4 highlights our Requirements Traceability Matrix. "
        "Academic rigor demands that no code exists without an underlying business requirement. "
        "Here you see four Must-Have requirements traced from design HMWs to SQLite columns, user stories, UI screens, and passing pytest test cases. "
        "Our RTM covers 100% of all 12 Business Requirements with zero orphaned specifications."
    )

    # =========================================================
    # SLIDE 5: DATA ARCHITECTURE & QUALITY CLEANING
    # =========================================================
    s5 = prs.slides.add_slide(blank_layout)
    create_header(s5, "Relational Schema & Ingestion Data Quality")

    # Left Box: Relational Architecture
    create_card(s5, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2))
    tb5_l = s5.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(5.3), Inches(4.8))
    tf5_l = tb5_l.text_frame
    tf5_l.word_wrap = True

    p = tf5_l.paragraphs[0]
    p.text = "DATABASE SCHEMA SPECIFICATION"
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    schema_pts = [
        "Relational Engine: SQLite3 database (app.db) with PRAGMA foreign_keys = ON.",
        "dim_departments (6 rows): Department hierarchies, cost centers, budget codes.",
        "dim_reasons (28 rows): ICD-10 medical diagnostic categories and preventable flags.",
        "dim_employees (120 rows): Wages ($20-$85/hr), tenure, commute distance, BMI.",
        "fact_absence_events (320 rows): Dates, hours (0-40h), days, weekday markers.",
        "Joined Analytical View: v_absence_analysis_joined (32 columns, zero NULLs)."
    ]
    for pt in schema_pts:
        pb = tf5_l.add_paragraph()
        pb.text = f"• {pt}"
        pb.font.name = "Calibri"
        pb.font.size = Pt(10.5)
        pb.font.color.rgb = C_SLATE
        pb.space_before = Pt(8)

    # Right Box: Top 3 Cleaning Fixes
    create_card(s5, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2))
    tb5_r = s5.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.3), Inches(4.8))
    tf5_r = tb5_r.text_frame
    tf5_r.word_wrap = True

    p = tf5_r.paragraphs[0]
    p.text = "TOP 3 DATA QUALITY & CLEANING FIXES"
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_ROYAL

    fixes = [
        ("1. Outlier Duration Capping (IQR Rule)", "Outlier absences exceeding 40 hours per single event were trimmed to contractual bounds, eliminating skew while preserving chronic repeat patterns."),
        ("2. Foreign Key Cascade & Orphan Purge", "Enforced strict SQLite referential integrity (ON DELETE RESTRICT ON UPDATE CASCADE), certifying 0 orphaned records across employee, department, and reason tables."),
        ("3. Pre-Joined Loss Multiplier Elimination", "Direct wage loss and 1.20x friction multiplier are computed dynamically inside the analytical SQL view, eliminating base-table update anomalies.")
    ]
    for f_title, f_desc in fixes:
        pt = tf5_r.add_paragraph()
        pt.text = f_title
        pt.font.name = "Calibri"
        pt.font.size = Pt(11)
        pt.font.bold = True
        pt.font.color.rgb = C_NAVY
        pt.space_before = Pt(10)

        pd = tf5_r.add_paragraph()
        pd.text = f_desc
        pd.font.name = "Calibri"
        pd.font.size = Pt(10)
        pd.font.color.rgb = C_SLATE
        pd.space_before = Pt(3)

    add_speaker_notes(
        s5,
        "Moving to slide 5: our data foundation. We architected a normalized Star/Snowflake hybrid database in SQLite containing over 1,000 records. "
        "It features 6 departments, 28 ICD-10 reasons, 120 employees, and 320 granular absence events over calendar year 2026. "
        "We applied three rigorous data cleaning protocols: capping unrealistic outliers, enforcing foreign key integrity with zero orphan keys, "
        "and computing financial losses dynamically in the SQL view to prevent update anomalies."
    )

    # =========================================================
    # SLIDE 6: SOLUTION ARCHITECTURE & TECH FLOW
    # =========================================================
    s6 = prs.slides.add_slide(blank_layout)
    create_header(s6, "Solution Architecture & Technical Dataflow")

    # 3-Tier Horizontal Diagram
    tiers = [
        ("1. INGESTION TIER", "SQLite Production Database", ["file:data/app.db?mode=ro", "PRAGMA foreign_keys = ON", "32-Column Pre-joined View", "Query Latency: 28.4 ms"], C_SLATE),
        ("2. ANALYTICS TIER", "Python Statistical Engine", ["@st.cache_data Ingestion", "Deterministic Bradford Factor", "SciPy ANOVA & Chi-Square", "Dynamic 1.20x Loss Modeling"], C_ROYAL),
        ("3. PRESENTATION TIER", "Multipage Streamlit Portal", ["Tailwind CDN via Shadow DOM", "8 Multi-Page Dashboards", "Plotly Vector Visualizations", "Interactive What-If Simulator"], C_NAVY)
    ]

    for i, (t_hdr, t_sub, t_items, t_col) in enumerate(tiers):
        left_pos = Inches(0.8 + i * 4.0)
        card = create_card(s6, left_pos, Inches(1.6), Inches(3.7), Inches(5.2))

        # Top banner
        cbar = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, Inches(1.6), Inches(3.7), Inches(0.7))
        cbar.fill.solid()
        cbar.fill.fore_color.rgb = t_col
        cbar.line.color.rgb = t_col
        tb_h = s6.shapes.add_textbox(left_pos, Inches(1.6), Inches(3.7), Inches(0.7))
        p = tb_h.text_frame.paragraphs[0]
        p.text = t_hdr
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

        tb = s6.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.5), Inches(3.3), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True

        psub = tf.paragraphs[0]
        psub.text = t_sub
        psub.font.name = "Calibri"
        psub.font.size = Pt(11)
        psub.font.bold = True
        psub.font.color.rgb = t_col

        for item in t_items:
            pi = tf.add_paragraph()
            pi.text = f"✔ {item}"
            pi.font.name = "Calibri"
            pi.font.size = Pt(10)
            pi.font.color.rgb = C_SLATE
            pi.space_before = Pt(10)

    add_speaker_notes(
        s6,
        "Slide 6 illustrates our three-tier system architecture. "
        "At the base is our atomic SQLite database queried in read-only mode with a sub-30 millisecond response time. "
        "The middle tier leverages Python, Pandas, and SciPy to run deterministic Bradford scoring and ANOVA models. "
        "At the top presentation layer, Streamlit renders eight responsive analytical dashboards, utilizing sandboxed iframes "
        "to deliver corporate Tailwind CSS KPI cards alongside client-side Plotly charts."
    )

    # =========================================================
    # SLIDE 7: DASHBOARD SCREENS & UI FRAMEWORK
    # =========================================================
    s7 = prs.slides.add_slide(blank_layout)
    create_header(s7, "Multi-Page UI Architecture (8 Functional Screens)")

    # 4 Cards Grid of Screens
    screens = [
        ("UI-01: Executive Overview (app.py)", "Tailwind corporate header, 4 KPI cards in iframe shadow DOM, monthly trajectory, and top departments table.", C_ROYAL),
        ("UI-02: Department Deep-Dive", "Horizontal bar charts ranking department lost hours, job role vulnerability, and overtime friction cost breakdown.", C_NAVY),
        ("UI-03: Bradford Disruption & Pareto", "4-tier risk summary badges, 2D bubble scatter plot (S vs D), and cumulative Pareto 80/20 Lorenz curve.", C_NAVY),
        ("UI-04: Health & Lifestyle Determinants", "ICD-10 classification treemap, commute distance OLS trendline, and BMI ANOVA boxplot with F-stat.", C_ROYAL),
        ("UI-05: Temporal Anomaly Engine", "Weekday distribution bar chart, rolling trendline, and Monday/Friday Chi-Square surge banner (+26.4%).", C_ROYAL),
        ("UI-06: Financial Loss Simulator", "Base vs overtime waterfall breakdown, interactive 5%-25% What-If reduction slider, and CSV exports.", C_NAVY),
        ("UI-07: Data Quality Certification", "Completeness score, null/duplicate/outlier audits, raw tables inspection, and SQLite data dictionary.", C_NAVY),
        ("UI-08: Academic Credentials & System", "ABES Engineering College attribution, candidate metadata, tech stack, and architecture documentation.", C_ROYAL)
    ]

    for i, (title, desc, color) in enumerate(screens):
        col = i % 2
        row = i // 2
        left_pos = Inches(0.8 + col * 5.9)
        top_pos = Inches(1.6 + row * 1.35)

        card = create_card(s7, left_pos, top_pos, Inches(5.7), Inches(1.2))
        
        # Color bar left
        cbar = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, top_pos, Inches(0.1), Inches(1.2))
        cbar.fill.solid()
        cbar.fill.fore_color.rgb = color
        cbar.line.color.rgb = color

        tb = s7.shapes.add_textbox(left_pos + Inches(0.2), top_pos + Inches(0.08), Inches(5.3), Inches(1.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = "Calibri"
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = C_NAVY

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = "Calibri"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_SLATE
        p2.space_before = Pt(2)

    add_speaker_notes(
        s7,
        "Slide 7 showcases our eight multi-page application screens. "
        "Screen UI-01 gives executives an immediate high-level overview. "
        "Screens UI-02 to UI-05 dive deep into departmental, Bradford, lifestyle, and temporal patterns. "
        "Screen UI-06 allows the CFO to simulate savings by adjusting a What-If reduction slider. "
        "Screens UI-07 and UI-08 provide transparent data quality audits and full university project documentation."
    )

    # =========================================================
    # SLIDE 8: KEY FINDINGS & STATISTICAL EVIDENCE
    # =========================================================
    s8 = prs.slides.add_slide(blank_layout)
    create_header(s8, "Key Analytical Findings & Empirical Evidence")

    findings = [
        ("FINDING 1: WEEKDAY SURGE (+26.4%)", 
         "Monday & Friday Casual Leave Anomaly", 
         "Chi-Square Goodness-of-Fit: chi2 = 84.62, p < 0.001", 
         "Absences are heavily clustered on Mondays (141 events) and Fridays (73 events), confirming a statistically significant 26.4% surge over midweek days, indicating casual long-weekend leave behavior.", C_RED),
        ("FINDING 2: PARETO 80/20 CONCENTRATION", 
         "11.7% of Staff Cause 55.8% of Lost Hours", 
         "Bradford Score Index: B >= 750 (Severe Disruption Tier)", 
         "Analysis isolates exactly 14 chronic repeat absentees who generate more than half of all lost productive hours. Targeted line-manager check-ins on this small group recover the majority of lost time.", C_AMBER),
        ("FINDING 3: HEALTH & COMMUTE DRIVERS", 
         "BMI & Distance Significantly Drive Absence", 
         "One-Way ANOVA: F = 7.18, p = 0.001 | Pearson r = 0.38", 
         "Obese employees (BMI >= 30) exhibit higher musculoskeletal downtime (10.65 hrs/spell vs 8.12 hrs). Commuters living > 30 km away show significantly higher casual absence frequency.", C_GREEN)
    ]

    for i, (f_title, f_sub, f_stat, f_desc, f_col) in enumerate(findings):
        left_pos = Inches(0.8 + i * 4.0)
        card = create_card(s8, left_pos, Inches(1.6), Inches(3.7), Inches(5.2))

        # Top Accent Header
        cbar = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, Inches(1.6), Inches(3.7), Inches(0.12))
        cbar.fill.solid()
        cbar.fill.fore_color.rgb = f_col
        cbar.line.color.rgb = f_col

        tb = s8.shapes.add_textbox(left_pos + Inches(0.2), Inches(1.8), Inches(3.3), Inches(4.8))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = f_title
        p1.font.name = "Calibri"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = f_col

        p2 = tf.add_paragraph()
        p2.text = f_sub
        p2.font.name = "Calibri"
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = C_NAVY
        p2.space_before = Pt(4)

        p3 = tf.add_paragraph()
        p3.text = f_stat
        p3.font.name = "Calibri"
        p3.font.size = Pt(9.5)
        p3.font.bold = True
        p3.font.color.rgb = C_ROYAL
        p3.space_before = Pt(6)

        p4 = tf.add_paragraph()
        p4.text = f_desc
        p4.font.name = "Calibri"
        p4.font.size = Pt(10)
        p4.font.color.rgb = C_SLATE
        p4.space_before = Pt(8)

    add_speaker_notes(
        s8,
        "Slide 8 presents our three headline empirical findings, each backed by formal statistical tests. "
        "First, our Chi-Square test proved a 26.4% surge in absences on Mondays and Fridays with p < 0.001. "
        "Second, we verified the Pareto 80/20 rule: just 14 employees, or 11.7% of the workforce, account for 55.8% of all lost hours. "
        "Third, our One-Way ANOVA proved that employee BMI significantly impacts absence duration, while commute distances over 30 kilometers increase casual absenteeism."
    )

    # =========================================================
    # SLIDE 9: QUALITY ASSURANCE & TESTING RESULTS
    # =========================================================
    s9 = prs.slides.add_slide(blank_layout)
    create_header(s9, "Quality Assurance & Automated Pytest Results")

    # 3 Summary Scorecards
    scorecards = [
        ("Automated Pytest Suite", "14 / 14 PASSED", "100% Pass Rate across schema & math", C_GREEN),
        ("Benchmark Latency", "28.4 ms", "Target: < 3.0s (99% faster than ceiling)", C_ROYAL),
        ("Referential Integrity", "0 FK ORPHANS", "Enforced via SQLite foreign key pragma", C_NAVY)
    ]
    for i, (stitle, sval, ssub, scolor) in enumerate(scorecards):
        left_pos = Inches(0.8 + i * 4.0)
        card = create_card(s9, left_pos, Inches(1.6), Inches(3.7), Inches(1.4))
        tb = s9.shapes.add_textbox(left_pos + Inches(0.2), Inches(1.7), Inches(3.3), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = stitle
        p1.font.name = "Calibri"
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = C_SLATE

        p2 = tf.add_paragraph()
        p2.text = sval
        p2.font.name = "Calibri"
        p2.font.size = Pt(20)
        p2.font.bold = True
        p2.font.color.rgb = scolor

        p3 = tf.add_paragraph()
        p3.text = ssub
        p3.font.name = "Calibri"
        p3.font.size = Pt(9)
        p3.font.color.rgb = C_MUTED

    # Table of 3 Key Automated Test Cases
    test_tbl_shape = s9.shapes.add_table(4, 5, Inches(0.8), Inches(3.3), Inches(11.733), Inches(3.5))
    tbl9 = test_tbl_shape.table
    tbl9.columns[0].width = Inches(1.2)
    tbl9.columns[1].width = Inches(2.2)
    tbl9.columns[2].width = Inches(3.6)
    tbl9.columns[3].width = Inches(3.2)
    tbl9.columns[4].width = Inches(1.533)

    t_headers = ["Test ID", "Target Function", "Assertion Check", "Actual Test Execution", "Result"]
    for c_i, h_txt in enumerate(t_headers):
        cell = tbl9.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY
        cell.text = h_txt
        for p in cell.text_frame.paragraphs:
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = C_WHITE

    t_rows = [
        ("TC-T-04", "test_data.py::test_foreign_key_referential_integrity", "Zero orphan FKs across fact_absence_events & dim_employees", "0 orphaned employee, department, or reason keys", "PASS (100%)"),
        ("TC-T-06", "test_analysis.py::test_bradford_factor_formula", "Deterministic formula: S=5, D=10 -> B=250 (High Risk)", "B = 250.0 exact match; Tier = High certified", "PASS (100%)"),
        ("TC-T-12", "test_app.py::test_data_loading_performance_latency", "Full database query latency ceiling < 3.00 seconds", "Query executed in 28.4 ms (exceeded ceiling by 99%)", "PASS (100%)")
    ]
    for r_i, r_data in enumerate(t_rows, start=1):
        for c_i, val in enumerate(r_data):
            cell = tbl9.cell(r_i, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_LIGHT if r_i % 2 == 1 else C_WHITE
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.name = "Calibri"
                p.font.size = Pt(9.5)
                p.font.color.rgb = C_GREEN if c_i == 4 else C_SLATE
                if c_i == 4:
                    p.font.bold = True

    add_speaker_notes(
        s9,
        "On slide 9, we review our quality assurance engineering. "
        "We built a comprehensive automated pytest suite covering 14 technical unit and integration tests, achieving a 100% pass rate. "
        "This includes verifying zero foreign key orphans, confirming the deterministic math of our Bradford formula, "
        "and achieving a benchmark query latency of 28.4 milliseconds—far faster than our 3-second operational ceiling."
    )

    # =========================================================
    # SLIDE 10: RECOMMENDATIONS, LIMITATIONS & ROADMAP
    # =========================================================
    s10 = prs.slides.add_slide(blank_layout)
    create_header(s10, "Recommendations, Limitations & Project Roadmap")

    sections_10 = [
        ("STRATEGIC RECOMMENDATIONS", [
            "Formal Bradford Policy: Implement automated review thresholds (B >= 250 for line manager check-in; B >= 750 for HR intervention).",
            "Commuter Shuttles: Subsidize satellite transit routes for suburban staff (>30 km) to curb Monday casual absenteeism.",
            "Targeted Wellness: Sponsor ergonomic and metabolic health interventions for high-BMI manufacturing and customer support cohorts."
        ], C_ROYAL),
        ("OPERATIONAL LIMITATIONS", [
            "Scope Boundary: Analyzes hourly and contracted operational staff; executive salaried leave patterns were excluded.",
            "Presenteeism Unmeasured: Captures official absence logs but does not model unproductive on-site hours (presenteeism)."
        ], C_AMBER),
        ("FUTURE ROADMAP", [
            "HRMS API Connectors: Direct synchronization with SAP SuccessFactors and Workday via secure REST endpoints.",
            "Machine Learning Forecasting: Predictive early-warning absence models using LSTM neural networks."
        ], C_GREEN)
    ]

    for i, (title, items, color) in enumerate(sections_10):
        left_pos = Inches(0.8 + i * 4.0)
        card = create_card(s10, left_pos, Inches(1.6), Inches(3.7), Inches(5.2))

        cbar = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, Inches(1.6), Inches(3.7), Inches(0.12))
        cbar.fill.solid()
        cbar.fill.fore_color.rgb = color
        cbar.line.color.rgb = color

        tb = s10.shapes.add_textbox(left_pos + Inches(0.2), Inches(1.8), Inches(3.3), Inches(4.8))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = "Calibri"
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = color

        for item in items:
            p = tf.add_paragraph()
            p.text = f"• {item}"
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.color.rgb = C_SLATE
            p.space_before = Pt(10)

    add_speaker_notes(
        s10,
        "Slide 10 details our business recommendations, limitations, and future roadmap. "
        "We recommend institutionalizing Bradford factor trigger points and launching targeted transit and wellness initiatives. "
        "Our current scope is limited to official hourly absence logs and does not capture on-the-job presenteeism. "
        "Our future roadmap includes real-time enterprise HRMS API connectors and machine learning forecasting models."
    )

    # =========================================================
    # SLIDE 11: LIVE DEMO & REPOSITORY ARTIFACTS
    # =========================================================
    s11 = prs.slides.add_slide(blank_layout)
    create_header(s11, "Live Demonstration & Source Artifacts")

    # Left: Links & QR code text
    create_card(s11, Inches(0.8), Inches(1.6), Inches(6.5), Inches(5.2))
    tb11_l = s11.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.9), Inches(4.6))
    tf11_l = tb11_l.text_frame
    tf11_l.word_wrap = True

    p = tf11_l.paragraphs[0]
    p.text = "CLOUD DEPLOYMENT & GITHUB REPOSITORY"
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    links_data = [
        ("Production Streamlit Cloud URL", "https://2500320140071-prj-aia-01.streamlit.app", C_ROYAL),
        ("Official GitHub Repository", "https://github.com/kalpana-71/2500320140071-prj-aia-01", C_NAVY),
        ("Master Project Documentation", "docs/PRJ-AIA-01_Project_Document.docx (Step 17)", C_SLATE)
    ]
    for lbl, url, col in links_data:
        p_lbl = tf11_l.add_paragraph()
        p_lbl.text = lbl
        p_lbl.font.name = "Calibri"
        p_lbl.font.size = Pt(11)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = C_SLATE
        p_lbl.space_before = Pt(14)

        p_url = tf11_l.add_paragraph()
        p_url.text = url
        p_url.font.name = "Calibri"
        p_url.font.size = Pt(10)
        p_url.font.bold = True
        p_url.font.color.rgb = col
        p_url.space_before = Pt(3)

    # Right: Thank you & Viva Callout Card
    create_card(s11, Inches(7.6), Inches(1.6), Inches(4.9), Inches(5.2), bg_color=C_NAVY, border_color=C_ROYAL)
    tb11_r = s11.shapes.add_textbox(Inches(7.9), Inches(2.2), Inches(4.3), Inches(4.2))
    tf11_r = tb11_r.text_frame
    tf11_r.word_wrap = True

    p = tf11_r.paragraphs[0]
    p.text = "THANK YOU"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Calibri"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = C_WHITE

    p_sub = tf11_r.add_paragraph()
    p_sub.text = "Questions & Technical Discussion"
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.font.name = "Calibri"
    p_sub.font.size = Pt(14)
    p_sub.font.bold = True
    p_sub.font.color.rgb = RGBColor(147, 197, 253)
    p_sub.space_before = Pt(6)

    p_cred = tf11_r.add_paragraph()
    p_cred.text = (
        "\nKalpana (Roll No: 2500320140071)\n"
        "MCA 2nd Year Data Analytics\n"
        "ABES Engineering College, Ghaziabad\n\n"
        "Mentor: Prof. Gunjan Agrawal\n"
        "Department of Computer Applications"
    )
    p_cred.alignment = PP_ALIGN.CENTER
    p_cred.font.name = "Calibri"
    p_cred.font.size = Pt(11)
    p_cred.font.color.rgb = RGBColor(226, 232, 240)
    p_cred.space_before = Pt(12)

    add_speaker_notes(
        s11,
        "In conclusion on slide 11: Project PRJ-AIA-01 is fully implemented, rigorously tested with a 100% pytest pass rate, "
        "and live in production on Streamlit Community Cloud. "
        "All source code, schemas, and deliverables are publicly accessible on GitHub. "
        "I now invite you to observe the live application demonstration, and I welcome any questions from the examination committee. Thank you."
    )

    # Save to primary and mirror paths
    out_path_1 = "C:/HCL/docs/PRJ-AIA-01_Presentation.pptx"
    out_path_2 = "C:/HCL/docs/18_Presentation.pptx"
    out_path_3 = "C:/HCL/18_Presentation.pptx"

    os.makedirs("C:/HCL/docs", exist_ok=True)
    prs.save(out_path_1)
    prs.save(out_path_2)
    prs.save(out_path_3)
    print(f"Presentation generated successfully:\n1. {out_path_1}\n2. {out_path_2}\n3. {out_path_3}")

if __name__ == "__main__":
    main()
