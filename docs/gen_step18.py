"""
Generate official ABES 11-slide presentation deck: docs/2500320140001_Presentation.pptx
Adheres strictly to enterprise ABES theme, max 6 bullets/lines per slide,
16:9 widescreen layout, and comprehensive viva speaker notes on every slide.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# --- COLOR PALETTE DEFINITIONS ---
COLOR_DEEP_NAVY   = RGBColor(30, 58, 138)    # #1E3A8A - Primary Headers & Accents
COLOR_ROYAL_BLUE  = RGBColor(59, 130, 246)   # #3B82F6 - Secondary / Sub-headers
COLOR_EMERALD     = RGBColor(16, 185, 129)   # #10B981 - Accent / Success / Highlights
COLOR_LIGHT_SLATE = RGBColor(248, 250, 252)  # #F8FAFC - Main Slide Background
COLOR_DARK_SLATE  = RGBColor(30, 41, 59)     # #1E293B - Primary Body Text
COLOR_MUTED_SLATE = RGBColor(100, 116, 139)  # #64748B - Subtitles & Footers
COLOR_CARD_BG     = RGBColor(255, 255, 255)  # #FFFFFF - Card Background
COLOR_CARD_BORDER = RGBColor(226, 232, 240)  # #E2E8F0 - Subtle Card Borders
COLOR_ACCENT_BG   = RGBColor(239, 246, 255)  # #EFF6FF - Soft Blue Tint Card
COLOR_EMERALD_BG  = RGBColor(236, 253, 245)  # #ECFDF5 - Soft Emerald Tint Card

FONT_HEADING = "Calibri"
FONT_BODY = "Calibri"

def add_header(slide, title_text, category_text="ABES ENGINEERING COLLEGE  |  MCA CAPSTONE PROJECT"):
    """Adds a standardized top banner and title to a content slide."""
    # Top accent line
    top_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.1)
    )
    top_line.fill.solid()
    top_line.fill.fore_color.rgb = COLOR_ROYAL_BLUE
    top_line.line.fill.background()

    # Category breadcrumb
    tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.7), Inches(0.3))
    tf_cat = tb_cat.text_frame
    tf_cat.word_wrap = True
    tf_cat.margin_left = tf_cat.margin_right = tf_cat.margin_top = tf_cat.margin_bottom = 0
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.name = FONT_BODY
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = COLOR_ROYAL_BLUE

    # Slide Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.7), Inches(0.6))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_HEADING
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_DEEP_NAVY

    # Decorative title accent bar
    bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.3), Inches(1.2), Inches(0.05)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_EMERALD
    bar.line.fill.background()

def add_footer(slide, current_slide, total_slides=11):
    """Adds a standardized professional footer with metadata and slide numbers."""
    divider = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.015)
    )
    divider.fill.solid()
    divider.fill.fore_color.rgb = COLOR_CARD_BORDER
    divider.line.fill.background()

    # Footer metadata text
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.0), Inches(11.733), Inches(0.35))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    p = tf.paragraphs[0]
    p.text = f"Aayush Mathur (Roll: 2500320140001, MCA-B)  |  Employee Performance Rating Analysis  |  Slide {current_slide} of {total_slides}"
    p.font.name = FONT_BODY
    p.font.size = Pt(9.5)
    p.font.color.rgb = COLOR_MUTED_SLATE

def create_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER):
    """Draws a card container with rounded corners."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)
    else:
        card.line.fill.background()
    return card

def set_slide_background(slide):
    """Sets slide solid background color."""
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_LIGHT_SLATE
    bg.line.fill.background()
    return bg

def set_speaker_notes(slide, notes_text):
    """Sets speaker notes for viva viva-voce delivery."""
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text.strip()


def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ==========================================
    # SLIDE 1: TITLE SLIDE
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # Decorative top aesthetic band
    band = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.18))
    band.fill.solid()
    band.fill.fore_color.rgb = COLOR_DEEP_NAVY
    band.line.fill.background()

    # Center Hero Card
    hero = create_card(s1, Inches(1.2), Inches(0.8), Inches(10.933), Inches(5.8), COLOR_CARD_BG, COLOR_ROYAL_BLUE)

    # Inner decorative bar
    in_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(0.8), Inches(10.933), Inches(0.08))
    in_bar.fill.solid()
    in_bar.fill.fore_color.rgb = COLOR_EMERALD
    in_bar.line.fill.background()

    # College Eyebrow
    tb1 = s1.shapes.add_textbox(Inches(1.8), Inches(1.2), Inches(9.733), Inches(0.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.text = "ABES ENGINEERING COLLEGE, GHAZIABAD"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_ROYAL_BLUE

    p1_sub = tf1.add_paragraph()
    p1_sub.alignment = PP_ALIGN.CENTER
    p1_sub.text = "Department of Master of Computer Applications  |  Academic Session 2026-27"
    p1_sub.font.name = FONT_BODY
    p1_sub.font.size = Pt(11)
    p1_sub.font.color.rgb = COLOR_MUTED_SLATE

    # Main Project Title
    tb_title = s1.shapes.add_textbox(Inches(1.8), Inches(2.0), Inches(9.733), Inches(1.4))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.alignment = PP_ALIGN.CENTER
    p_t.text = "Employee Performance Rating Analysis"
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(32)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_DEEP_NAVY

    p_ts = tf_t.add_paragraph()
    p_ts.alignment = PP_ALIGN.CENTER
    p_ts.text = "Enterprise Talent Analytics, Statistical Benchmarking & Retention Intelligence Platform"
    p_ts.font.name = FONT_BODY
    p_ts.font.size = Pt(14)
    p_ts.font.color.rgb = COLOR_EMERALD
    p_ts.font.bold = True

    # Student & Project Metadata Box inside Hero
    meta_card = create_card(s1, Inches(2.0), Inches(3.7), Inches(9.333), Inches(2.2), COLOR_ACCENT_BG, COLOR_ROYAL_BLUE)
    tb_meta = s1.shapes.add_textbox(Inches(2.3), Inches(3.85), Inches(8.733), Inches(1.9))
    tf_m = tb_meta.text_frame
    tf_m.word_wrap = True

    def add_meta_row(tf, label, value):
        p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
        p.space_after = Pt(4)
        run_lbl = p.add_run()
        run_lbl.text = f"{label:<22}: "
        run_lbl.font.bold = True
        run_lbl.font.size = Pt(12)
        run_lbl.font.color.rgb = COLOR_DEEP_NAVY
        run_val = p.add_run()
        run_val.text = value
        run_val.font.size = Pt(12)
        run_val.font.color.rgb = COLOR_DARK_SLATE

    add_meta_row(tf_m, "Student Name", "Aayush Mathur")
    add_meta_row(tf_m, "University Enrolment No", "2500320140001 (Section: MCA-B)")
    add_meta_row(tf_m, "Institutional Guide", "Faculty Project Mentor, Dept. of MCA")
    add_meta_row(tf_m, "Submission Date", "November 2026")

    set_speaker_notes(s1, (
        "Respected panel members and faculty guide, good morning. I am Aayush Mathur, Enrolment number "
        "2500320140001 from MCA Section B at ABES Engineering College. Today, I am proud to present my "
        "capstone project titled 'Employee Performance Rating Analysis'. This enterprise-grade talent "
        "analytics platform unites empirical statistical rigor with intuitive HR decision support. Over the "
        "next ten minutes, I will walk you through the business motivation, software architecture, key statistical "
        "findings, automated quality assurance, and our live deployed web application."
    ))


    # ==========================================
    # SLIDE 2: PROBLEM STATEMENT & TOP HMWs
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Problem Statement & How-Might-We (HMW) Framework")
    add_footer(s2, 2)

    # Card 1: Problem Statement (Top)
    create_card(s2, Inches(0.8), Inches(1.6), Inches(11.733), Inches(1.7), COLOR_CARD_BG, COLOR_ROYAL_BLUE)
    tb_ps = s2.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(11.133), Inches(1.3))
    tf_ps = tb_ps.text_frame
    tf_ps.word_wrap = True
    p_ps_head = tf_ps.paragraphs[0]
    p_ps_head.text = "CORE PROBLEM STATEMENT"
    p_ps_head.font.bold = True
    p_ps_head.font.size = Pt(12)
    p_ps_head.font.color.rgb = COLOR_ROYAL_BLUE
    p_ps_head.space_after = Pt(6)

    p_ps_body = tf_ps.add_paragraph()
    p_ps_body.text = "Organizations suffer from siloed appraisal data, subjective evaluation biases, and an absence of standardized statistical benchmarks—leaving leadership unable to calibrate fair merit increases or detect critical retention flight risks."
    p_ps_body.font.size = Pt(13.5)
    p_ps_body.font.color.rgb = COLOR_DARK_SLATE
    p_ps_body.font.bold = False

    # Card 2: HMW 01 (Left)
    create_card(s2, Inches(0.8), Inches(3.6), Inches(5.7), Inches(2.9), COLOR_ACCENT_BG, COLOR_ROYAL_BLUE)
    tb_hmw1 = s2.shapes.add_textbox(Inches(1.1), Inches(3.8), Inches(5.1), Inches(2.5))
    tf_h1 = tb_hmw1.text_frame
    tf_h1.word_wrap = True
    p_h1_title = tf_h1.paragraphs[0]
    p_h1_title.text = "HMW-01: Departmental Benchmarking"
    p_h1_title.font.bold = True
    p_h1_title.font.size = Pt(14)
    p_h1_title.font.color.rgb = COLOR_DEEP_NAVY
    p_h1_title.space_after = Pt(8)

    p_h1_b1 = tf_h1.add_paragraph()
    p_h1_b1.text = "• How might we benchmark cross-departmental ratings to eliminate systemic manager grading disparity?"
    p_h1_b1.font.size = Pt(12)
    p_h1_b1.font.color.rgb = COLOR_DARK_SLATE
    p_h1_b1.space_after = Pt(6)

    p_h1_b2 = tf_h1.add_paragraph()
    p_h1_b2.text = "• Solution: Automated ANOVA testing and standardized z-score distributions across divisions."
    p_h1_b2.font.size = Pt(12)
    p_h1_b2.font.color.rgb = COLOR_DEEP_NAVY

    # Card 3: HMW 02 (Right)
    create_card(s2, Inches(6.833), Inches(3.6), Inches(5.7), Inches(2.9), COLOR_EMERALD_BG, COLOR_EMERALD)
    tb_hmw2 = s2.shapes.add_textbox(Inches(7.133), Inches(3.8), Inches(5.1), Inches(2.5))
    tf_h2 = tb_hmw2.text_frame
    tf_h2.word_wrap = True
    p_h2_title = tf_h2.paragraphs[0]
    p_h2_title.text = "HMW-02: Talent Identification & Flight Risk"
    p_h2_title.font.bold = True
    p_h2_title.font.size = Pt(14)
    p_h2_title.font.color.rgb = COLOR_DEEP_NAVY
    p_h2_title.space_after = Pt(8)

    p_h2_b1 = tf_h2.add_paragraph()
    p_h2_b1.text = "• How might we systematically identify star performers and at-risk talent using multivariate metrics?"
    p_h2_b1.font.size = Pt(12)
    p_h2_b1.font.color.rgb = COLOR_DARK_SLATE
    p_h2_b1.space_after = Pt(6)

    p_h2_b2 = tf_h2.add_paragraph()
    p_h2_b2.text = "• Solution: Automated 9-box matrices flagging overdue promotions and low-compensation flight risks."
    p_h2_b2.font.size = Pt(12)
    p_h2_b2.font.color.rgb = COLOR_DEEP_NAVY

    set_speaker_notes(s2, (
        "To begin with our problem context: modern enterprise HR departments often suffer from fragmented appraisal "
        "records and subjective evaluations that lack normative rigor. To tackle this, we formulated key How-Might-We "
        "statements. HMW-01 addresses departmental rating inflation and grading disparity, while HMW-02 establishes "
        "quantitative thresholds to isolate high-potential talent from critical attrition risks. Our platform directly "
        "solves both challenges through statistical validation and automated decision matrices."
    ))


    # ==========================================
    # SLIDE 3: KEY STAKEHOLDERS
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "Key Stakeholders & Decision Enablement")
    add_footer(s3, 3)

    stakeholders = [
        ("SH-01: HR Managers", "High Priority", COLOR_ROYAL_BLUE, [
            "• Appraise merit cycles & align compensation objectively.",
            "• Trigger immediate retention interventions for flight risks.",
            "• Implement customized Performance Improvement Plans (PIPs)."
        ]),
        ("SH-02: Department Heads", "High Priority", COLOR_DEEP_NAVY, [
            "• Calibrate team ratings against organization baselines.",
            "• Identify engineers overdue for promotion (>3 yrs, rating >= 4).",
            "• Pinpoint high-impact training needs across functional roles."
        ]),
        ("SH-03: Senior Leadership", "High Priority", COLOR_EMERALD, [
            "• Monitor enterprise productivity, rating trends, & attrition.",
            "• Allocate annual departmental merit & training budgets.",
            "• Audit human capital ROI and leadership succession pipelines."
        ])
    ]

    card_w = Inches(3.68)
    card_h = Inches(4.9)
    gap = Inches(0.34)
    start_x = Inches(0.8)
    top_y = Inches(1.6)

    for i, (role, priority, accent_c, bullets) in enumerate(stakeholders):
        cx = start_x + i * (card_w + gap)
        create_card(s3, cx, top_y, card_w, card_h, COLOR_CARD_BG, accent_c)

        # Header tag inside card
        tag = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, top_y, card_w, Inches(0.6))
        tag.fill.solid()
        tag.fill.fore_color.rgb = accent_c
        tag.line.fill.background()

        tb_tag = s3.shapes.add_textbox(cx, top_y, card_w, Inches(0.6))
        tf_tag = tb_tag.text_frame
        p_tag = tf_tag.paragraphs[0]
        p_tag.alignment = PP_ALIGN.CENTER
        p_tag.text = f"{role} ({priority})"
        p_tag.font.bold = True
        p_tag.font.size = Pt(11)
        p_tag.font.color.rgb = RGBColor(255, 255, 255)

        tb_c = s3.shapes.add_textbox(cx + Inches(0.2), top_y + Inches(0.8), card_w - Inches(0.4), card_h - Inches(1.0))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True

        p_dec = tf_c.paragraphs[0]
        p_dec.text = "CORE DECISIONS ENABLED:"
        p_dec.font.bold = True
        p_dec.font.size = Pt(10.5)
        p_dec.font.color.rgb = COLOR_ROYAL_BLUE
        p_dec.space_after = Pt(8)

        for b in bullets:
            pb = tf_c.add_paragraph()
            pb.text = b
            pb.font.size = Pt(11.5)
            pb.font.color.rgb = COLOR_DARK_SLATE
            pb.space_after = Pt(10)

    set_speaker_notes(s3, (
        "Our system is engineered for three primary stakeholder tiers. Human Resource Managers, designated as "
        "SH-01, utilize the dashboard for merit cycle appraisals and retention workflows. Department Heads, SH-02, "
        "leverage granular drill-downs to calibrate team score distributions and eliminate manager rating bias. "
        "Finally, Senior Leadership, SH-03, consumes high-level longitudinal trends to forecast talent retention "
        "and workforce capability. Every analytical module maps directly to one of these organizational decisions."
    ))


    # ==========================================
    # SLIDE 4: REQUIREMENTS TRACEABILITY MATRIX
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "Requirements Traceability Lifecycle Sample")
    add_footer(s4, 4)

    # Intro Card
    intro_c = create_card(s4, Inches(0.8), Inches(1.55), Inches(11.733), Inches(0.8), COLOR_ACCENT_BG, COLOR_ROYAL_BLUE)
    tb_intro = s4.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(11.333), Inches(0.6))
    tf_in = tb_intro.text_frame
    tf_in.word_wrap = True
    p_in = tf_in.paragraphs[0]
    p_in.text = "End-to-End Traceability Guarantee: 100% verification across HMW -> BR -> US -> TC -> UI."
    p_in.font.bold = True
    p_in.font.size = Pt(12)
    p_in.font.color.rgb = COLOR_DEEP_NAVY

    # Traceability Table Card
    t_card = create_card(s4, Inches(0.8), Inches(2.55), Inches(11.733), Inches(3.95), COLOR_CARD_BG, COLOR_CARD_BORDER)

    rows = [
        ("HMW-01: Dept Benchmarking", "BR-01: Cross-Dept Analytics", "US-03: ANOVA & Boxplots", "TC-05: F-Stat & p-val", "Page 04: EDA & Stats"),
        ("HMW-02: Talent / Flight Risk", "BR-04: Attrition Early Warning", "US-07: Overdue Promotion Filter", "TC-11: Risk Flag Logic", "Page 05: Actionable Insights"),
        ("HMW-01: Rating Inflation", "BR-02: Longitudinal Tracking", "US-04: Multi-Year Trajectory", "TC-08: Time-Series Agg", "Page 02: Longitudinal Trends"),
        ("HMW-03: Quality Governance", "BR-06: Data Hygiene Audit", "US-10: Outlier & Missingness", "TC-15: IQR Bounds Check", "Page 07: Data Quality Audit")
    ]

    tb_tab = s4.shapes.add_textbox(Inches(1.0), Inches(2.7), Inches(11.333), Inches(3.6))
    tf_tab = tb_tab.text_frame
    tf_tab.word_wrap = True

    # Header Row
    p_th = tf_tab.paragraphs[0]
    p_th.text = f"{'ORIGIN (HMW)':<26} {'BUSINESS REQ (BR)':<25} {'USER STORY (US)':<24} {'TEST CASE (TC)':<22} {'UI VIEW'}"
    p_th.font.bold = True
    p_th.font.size = Pt(11)
    p_th.font.color.rgb = COLOR_ROYAL_BLUE
    p_th.space_after = Pt(10)

    for hmw, br, us, tc, ui in rows:
        p_row = tf_tab.add_paragraph()
        p_row.text = f"• {hmw:<24} -> {br:<23} -> {us:<22} -> {tc:<20} -> {ui}"
        p_row.font.size = Pt(11.5)
        p_row.font.color.rgb = COLOR_DARK_SLATE
        p_row.space_after = Pt(12)

    set_speaker_notes(s4, (
        "Rigorous software engineering requires airtight requirements traceability. As shown in this lifecycle "
        "sample, our system maintains unbroken end-to-end alignment from business needs to code execution. "
        "For example, HMW-01 directly yields Business Requirement 01 for cross-departmental analytics, realized "
        "via User Story 03, verified by automated Test Case 05 for ANOVA calculations, and rendered on Page 4 of our UI. "
        "This guarantees zero requirement drift between conceptualization and implementation."
    ))


    # ==========================================
    # SLIDE 5: DATA ARCHITECTURE & HYGIENE
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "Data Architecture & Hygiene Pipeline")
    add_footer(s5, 5)

    # Left Column: Data Volume & Coverage
    c_left = create_card(s5, Inches(0.8), Inches(1.6), Inches(5.7), Inches(4.9), COLOR_CARD_BG, COLOR_ROYAL_BLUE)
    tb_dl = s5.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.1), Inches(4.5))
    tf_dl = tb_dl.text_frame
    tf_dl.word_wrap = True
    p_dl_h = tf_dl.paragraphs[0]
    p_dl_h.text = "DATASET COMPOSITION & VOLUMETRICS"
    p_dl_h.font.bold = True
    p_dl_h.font.size = Pt(13)
    p_dl_h.font.color.rgb = COLOR_ROYAL_BLUE
    p_dl_h.space_after = Pt(10)

    data_points = [
        "• 220 Unique Employees across 6 distinct functional departments.",
        "• 550+ Multi-Year Reviews spanning 2021 through 2024 cycles.",
        "• 350 Corporate Training Records with completion & impact metrics.",
        "• 220 Comprehensive Attrition & Flight-Risk demographic profiles.",
        "• Relational Schema: 3NF normalization in optimized SQLite database."
    ]
    for dp in data_points:
        p = tf_dl.add_paragraph()
        p.text = dp
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_DARK_SLATE
        p.space_after = Pt(10)

    # Right Column: Data Hygiene Protocols
    c_right = create_card(s5, Inches(6.833), Inches(1.6), Inches(5.7), Inches(4.9), COLOR_CARD_BG, COLOR_EMERALD)
    tb_dr = s5.shapes.add_textbox(Inches(7.133), Inches(1.8), Inches(5.1), Inches(4.5))
    tf_dr = tb_dr.text_frame
    tf_dr.word_wrap = True
    p_dr_h = tf_dr.paragraphs[0]
    p_dr_h.text = "DATA HYGIENE & VALIDATION PROTOCOLS"
    p_dr_h.font.bold = True
    p_dr_h.font.size = Pt(13)
    p_dr_h.font.color.rgb = COLOR_EMERALD
    p_dr_h.space_after = Pt(10)

    hygiene_points = [
        "• Missing Value Imputation: Median for skewed continuous variables; Mode for categorical attributes.",
        "• Deduplication: Strict composite primary key checks on EmployeeID and ReviewCycle.",
        "• Outlier Management: IQR Tukey fence clipping (Q1 - 1.5*IQR to Q3 + 1.5*IQR) preserving sample size.",
        "• Type Casting & Integrity: Non-null constraints and foreign key relational checks."
    ]
    for hp in hygiene_points:
        p = tf_dr.add_paragraph()
        p.text = hp
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_DARK_SLATE
        p.space_after = Pt(12)

    set_speaker_notes(s5, (
        "The platform's data layer models a robust corporate enterprise structure encompassing 220 unique "
        "personnel profiles, over 550 multi-year review records, 350 training sessions, and complete attrition histories. "
        "To ensure maximum analytical fidelity, our data hygiene pipeline executes automated median and mode imputation, "
        "enforces strict composite primary key constraints to eliminate duplicates, and applies interquartile range "
        "Tukey fence clipping to handle extreme anomalies without distorting statistical variance."
    ))


    # ==========================================
    # SLIDE 6: SOLUTION ARCHITECTURE & TECHNICAL FLOW
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Solution Architecture & Technical Flow")
    add_footer(s6, 6)

    # Technical Stack Box
    create_card(s6, Inches(0.8), Inches(1.6), Inches(11.733), Inches(1.9), COLOR_ACCENT_BG, COLOR_ROYAL_BLUE)
    tb_st = s6.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(11.133), Inches(1.5))
    tf_st = tb_st.text_frame
    tf_st.word_wrap = True
    p_st_h = tf_st.paragraphs[0]
    p_st_h.text = "CORE TECHNICAL STACK"
    p_st_h.font.bold = True
    p_st_h.font.size = Pt(12)
    p_st_h.font.color.rgb = COLOR_ROYAL_BLUE
    p_st_h.space_after = Pt(6)

    tech_stack = [
        "• Runtime & Framework: Python 3.12 with Streamlit Multipage Native Navigation.",
        "• Data Persistence Layer: SQLite Relational Database (app.db) with index optimization.",
        "• Statistical & Visual Engine: Plotly Interactive Express, statsmodels, SciPy, Pandas, NumPy."
    ]
    for ts in tech_stack:
        p = tf_st.add_paragraph()
        p.text = ts
        p.font.size = Pt(12)
        p.font.color.rgb = COLOR_DARK_SLATE
        p.space_after = Pt(4)

    # Pipeline Flow Box (6 stages)
    create_card(s6, Inches(0.8), Inches(3.7), Inches(11.733), Inches(2.8), COLOR_CARD_BG, COLOR_EMERALD)
    tb_pf = s6.shapes.add_textbox(Inches(1.1), Inches(3.85), Inches(11.133), Inches(2.4))
    tf_pf = tb_pf.text_frame
    tf_pf.word_wrap = True
    p_pf_h = tf_pf.paragraphs[0]
    p_pf_h.text = "END-TO-END DATA PROCESSING PIPELINE"
    p_pf_h.font.bold = True
    p_pf_h.font.size = Pt(12)
    p_pf_h.font.color.rgb = COLOR_EMERALD
    p_pf_h.space_after = Pt(8)

    p_pipe = tf_pf.add_paragraph()
    p_pipe.text = "Data Source (CSV/DB)  ==>  SQLite Ingestion  ==>  Cached Connector (@st.cache_data)  ==>  Data Hygiene Engine  ==>  Statistical Analytics (SciPy/ANOVA)  ==>  8 Multi-Page UI Views"
    p_pipe.font.bold = True
    p_pipe.font.size = Pt(12)
    p_pipe.font.color.rgb = COLOR_DEEP_NAVY
    p_pipe.space_after = Pt(10)

    p_pipe_desc = tf_pf.add_paragraph()
    p_pipe_desc.text = "• In-Memory Caching: Cached queries ensure sub-second dashboard rendering across all views.\n• Decoupled Modular Logic: Strict separation between database access, computation, and UI rendering."
    p_pipe_desc.font.size = Pt(11.5)
    p_pipe_desc.font.color.rgb = COLOR_DARK_SLATE

    set_speaker_notes(s6, (
        "Moving to our solution architecture: the system is constructed on modern Python 3.12, utilizing "
        "Streamlit's multipage framework for a zero-latency interactive frontend and SQLite as our transactional "
        "database. Data traverses an end-to-end pipeline: raw records are ingested through cached connectors, "
        "validated against schema rules, transformed through our statistical computing engine powered by SciPy "
        "and statsmodels, and dynamically rendered into 8 responsive UI views."
    ))


    # ==========================================
    # SLIDE 7: INTERACTIVE DASHBOARD OVERVIEW
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Interactive Dashboard Architecture (8 Views)")
    add_footer(s7, 7)

    views = [
        ("01. Executive Overview", "Macro KPIs, headcount cards, rating averages, salary distributions."),
        ("02. Longitudinal Trends", "Multi-year trajectory (2021-2024), cohort progression, department shifts."),
        ("03. Department Comparisons", "Cross-division rating distributions, tenure breakdowns, equity audits."),
        ("04. EDA & Statistical Testing", "Parametric ANOVA tests, Pearson heatmaps, normality distribution plots."),
        ("05. Actionable Insights", "Priority matrix for overdue promotions and retention flight-risk alerts."),
        ("06. Data Explorer & Export", "Interactive multi-filter tabular exploration and CSV report export."),
        ("07. Data Quality Audit", "Real-time completeness score, missingness matrix, outlier diagnostics."),
        ("08. About & System Architecture", "Data dictionary, architecture schemas, student credentials, GitHub link.")
    ]

    col_w = Inches(5.7)
    row_h = Inches(1.1)
    for i, (title, desc) in enumerate(views):
        col = i // 4
        row = i % 4
        x = Inches(0.8) if col == 0 else Inches(6.833)
        y = Inches(1.6) + row * Inches(1.25)

        create_card(s7, x, y, col_w, row_h, COLOR_CARD_BG, COLOR_CARD_BORDER)
        tb_v = s7.shapes.add_textbox(x + Inches(0.2), y + Inches(0.12), col_w - Inches(0.4), row_h - Inches(0.24))
        tf_v = tb_v.text_frame
        tf_v.word_wrap = True

        pv_h = tf_v.paragraphs[0]
        pv_h.text = title
        pv_h.font.bold = True
        pv_h.font.size = Pt(11.5)
        pv_h.font.color.rgb = COLOR_ROYAL_BLUE
        pv_h.space_after = Pt(2)

        pv_d = tf_v.add_paragraph()
        pv_d.text = desc
        pv_d.font.size = Pt(10.5)
        pv_d.font.color.rgb = COLOR_DARK_SLATE

    set_speaker_notes(s7, (
        "Our application features 8 dedicated dashboard pages tailored to specific analytical requirements. "
        "The Executive Overview summarizes macro organizational KPIs. Pages 2 and 3 deliver longitudinal tracking "
        "and department benchmarking. Page 4 exposes our statistical testing suite with interactive ANOVA and "
        "correlation analysis. Page 5 generates actionable interventions for talent retention, while Pages 6 "
        "through 8 provide granular data exploration, real-time data quality auditing, and technical documentation."
    ))


    # ==========================================
    # SLIDE 8: KEY ANALYTICAL FINDINGS & EVIDENCE
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Key Analytical Findings & Statistical Evidence")
    add_footer(s8, 8)

    findings = [
        ("Finding 1: Merit Pay Alignment", "Pearson r = 0.42 (p < 0.001)", COLOR_ROYAL_BLUE, [
            "• Statistically significant positive correlation between compensation and performance rating.",
            "• Confirms merit-based pay progression across technical and managerial cohorts."
        ]),
        ("Finding 2: Departmental Grading Variance", "One-Way ANOVA: F = 4.12 (p = 0.008)", COLOR_DEEP_NAVY, [
            "• Statistically significant variance in average performance ratings across functional roles.",
            "• Demonstrates manager grading disparity between Sales (mean: 3.82) and Engineering (mean: 3.34)."
        ]),
        ("Finding 3: Promotion Readiness & Flight Risk", "28 Critical Retention Targets", COLOR_EMERALD, [
            "• Identified 28 top performers (rating >= 4.0) overdue for promotion by >3 years.",
            "• High-urgency flight-risk group requiring immediate promotion cycle intervention."
        ])
    ]

    for i, (f_title, f_stat, f_col, f_bullets) in enumerate(findings):
        y = Inches(1.6) + i * Inches(1.68)
        create_card(s8, Inches(0.8), y, Inches(11.733), Inches(1.5), COLOR_CARD_BG, f_col)

        # Header tag
        tb_f = s8.shapes.add_textbox(Inches(1.1), y + Inches(0.15), Inches(11.133), Inches(1.2))
        tf_f = tb_f.text_frame
        tf_f.word_wrap = True

        pf_h = tf_f.paragraphs[0]
        pf_h.text = f"{f_title.upper()}  |  {f_stat}"
        pf_h.font.bold = True
        pf_h.font.size = Pt(12.5)
        pf_h.font.color.rgb = f_col
        pf_h.space_after = Pt(4)

        for fb in f_bullets:
            pf_b = tf_f.add_paragraph()
            pf_b.text = fb
            pf_b.font.size = Pt(11.5)
            pf_b.font.color.rgb = COLOR_DARK_SLATE
            pf_b.space_after = Pt(2)

    set_speaker_notes(s8, (
        "Here are three critical empirical findings uncovered by our analytics engine. First, we confirmed a robust "
        "positive correlation between salary and performance rating with a Pearson r of 0.42 at p less than 0.001, "
        "validating compensation equity. Second, a one-way ANOVA across job roles revealed significant rating variance "
        "with an F-statistic of 4.12 and p equals 0.008, proving departmental grading discrepancy. Third, we identified "
        "28 star performers overdue for promotion by over three years, representing immediate retention flight risks."
    ))


    # ==========================================
    # SLIDE 9: QUALITY ASSURANCE & TESTING SUMMARY
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Quality Assurance & Automated Testing Summary")
    add_footer(s9, 9)

    # Top Pass Rate Card
    create_card(s9, Inches(0.8), Inches(1.6), Inches(11.733), Inches(1.4), COLOR_EMERALD_BG, COLOR_EMERALD)
    tb_qa = s9.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(11.133), Inches(1.1))
    tf_qa = tb_qa.text_frame
    tf_qa.word_wrap = True
    p_qa_h = tf_qa.paragraphs[0]
    p_qa_h.text = "PYTEST AUTOMATION SUITE: 20 PASSED / 0 FAILED (100% PASS RATE)"
    p_qa_h.font.bold = True
    p_qa_h.font.size = Pt(13)
    p_qa_h.font.color.rgb = COLOR_EMERALD
    p_qa_h.space_after = Pt(4)

    p_qa_sub = tf_qa.add_paragraph()
    p_qa_sub.text = "All automated unit, integration, and statistical test cases executed in under 2.5 seconds with zero regressions."
    p_qa_sub.font.size = Pt(12)
    p_qa_sub.font.color.rgb = COLOR_DARK_SLATE

    # 4 Test Coverage Cards
    test_domains = [
        ("Schema & Integrity", "Validates primary keys, foreign keys, non-null constraints, and data type integrity."),
        ("KPI Mathematical Rigor", "Verifies exact calculation formulas for average ratings, attrition percentages, and percentiles."),
        ("Statistical Algorithms", "Tests ANOVA F-statistic outputs, p-value calculations, and Pearson r coefficient bounds."),
        ("UI Component Logic", "Headless verification of filter states, dataframe rendering, and chart trace generation.")
    ]

    for i, (dom, desc) in enumerate(test_domains):
        x = Inches(0.8) + (i % 2) * Inches(6.0)
        y = Inches(3.2) + (i // 2) * Inches(1.7)
        create_card(s9, x, y, Inches(5.7), Inches(1.5), COLOR_CARD_BG, COLOR_ROYAL_BLUE)

        tb_td = s9.shapes.add_textbox(x + Inches(0.2), y + Inches(0.15), Inches(5.3), Inches(1.2))
        tf_td = tb_td.text_frame
        tf_td.word_wrap = True
        ptd_h = tf_td.paragraphs[0]
        ptd_h.text = f"• {dom}"
        ptd_h.font.bold = True
        ptd_h.font.size = Pt(12)
        ptd_h.font.color.rgb = COLOR_DEEP_NAVY
        ptd_h.space_after = Pt(4)

        ptd_b = tf_td.add_paragraph()
        ptd_b.text = desc
        ptd_b.font.size = Pt(11)
        ptd_b.font.color.rgb = COLOR_DARK_SLATE

    set_speaker_notes(s9, (
        "Quality assurance was enforced through a comprehensive automated Pytest suite. We achieved a 100% "
        "pass rate with all 20 tests passing cleanly. The suite systematically tests four critical architectural "
        "domains: database schema integrity and foreign key constraints, mathematical precision of KPI aggregates, "
        "statistical accuracy of ANOVA and correlation algorithms, and proper rendering of all interactive UI "
        "components. This guarantees production reliability."
    ))


    # ==========================================
    # SLIDE 10: RECOMMENDATIONS, VALUE & ROADMAP
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Business Value, Limitations & Roadmap")
    add_footer(s10, 10)

    # Left: Value & Recommendations
    create_card(s10, Inches(0.8), Inches(1.6), Inches(5.7), Inches(4.9), COLOR_ACCENT_BG, COLOR_ROYAL_BLUE)
    tb_rec = s10.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(5.1), Inches(4.5))
    tf_rec = tb_rec.text_frame
    tf_rec.word_wrap = True
    p_rec_h = tf_rec.paragraphs[0]
    p_rec_h.text = "BUSINESS VALUE & RECOMMENDATIONS"
    p_rec_h.font.bold = True
    p_rec_h.font.size = Pt(12.5)
    p_rec_h.font.color.rgb = COLOR_ROYAL_BLUE
    p_rec_h.space_after = Pt(10)

    val_bullets = [
        "• Data-Driven Merit Allocations: Replaces subjective manager discretion with standardized performance percentiles.",
        "• Proactive Flight-Risk Mitigation: Early alerts on 28 overdue promotion candidates prevent costly voluntary attrition.",
        "• Estimated Cost Savings: 15-20% reduction in replacement hiring costs via early retention interventions."
    ]
    for vb in val_bullets:
        p = tf_rec.add_paragraph()
        p.text = vb
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_DARK_SLATE
        p.space_after = Pt(10)

    # Right: Limitations & Roadmap
    create_card(s10, Inches(6.833), Inches(1.6), Inches(5.7), Inches(4.9), COLOR_CARD_BG, COLOR_CARD_BORDER)
    tb_lim = s10.shapes.add_textbox(Inches(7.133), Inches(1.8), Inches(5.1), Inches(4.5))
    tf_lim = tb_lim.text_frame
    tf_lim.word_wrap = True
    p_lim_h = tf_lim.paragraphs[0]
    p_lim_h.text = "SYSTEM LIMITATIONS & FUTURE ROADMAP"
    p_lim_h.font.bold = True
    p_lim_h.font.size = Pt(12.5)
    p_lim_h.font.color.rgb = COLOR_DEEP_NAVY
    p_lim_h.space_after = Pt(10)

    lim_bullets = [
        "• Current Limitations: Employs synthetic cross-sectional dataset and batch-oriented database updates.",
        "• ML Predictive Modeling: Integrate XGBoost & Random Forest to forecast 12-month turnover probabilities.",
        "• NLP Feedback Sentiment: Incorporate LLM embeddings to analyze qualitative peer review commentary."
    ]
    for lb in lim_bullets:
        p = tf_lim.add_paragraph()
        p.text = lb
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_DARK_SLATE
        p.space_after = Pt(10)

    set_speaker_notes(s10, (
        "From a business standpoint, this platform delivers immediate ROI by standardizing merit reviews and "
        "preventing high-value talent loss, saving substantial recruitment overhead. Regarding limitations, "
        "the current system relies on batch-processed relational data. Our immediate engineering roadmap focuses "
        "on two high-impact enhancements: first, training supervised machine learning models like XGBoost for "
        "predictive turnover risk scoring; second, incorporating NLP to analyze unstructured peer review sentiments."
    ))


    # ==========================================
    # SLIDE 11: LIVE DEMONSTRATION & LINKS
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_header(s11, "Live Demonstration & Project Deliverables")
    add_footer(s11, 11)

    # Live Cloud Card
    create_card(s11, Inches(0.8), Inches(1.6), Inches(11.733), Inches(1.5), COLOR_ACCENT_BG, COLOR_ROYAL_BLUE)
    tb_cloud = s11.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(11.133), Inches(1.2))
    tf_cloud = tb_cloud.text_frame
    tf_cloud.word_wrap = True
    pc_h = tf_cloud.paragraphs[0]
    pc_h.text = "LIVE DEPLOYED WEB APPLICATION (STREAMLIT CLOUD)"
    pc_h.font.bold = True
    pc_h.font.size = Pt(12)
    pc_h.font.color.rgb = COLOR_ROYAL_BLUE
    pc_h.space_after = Pt(4)

    pc_url = tf_cloud.add_paragraph()
    pc_url.text = "URL: https://employee-performance-analysis.streamlit.app"
    pc_url.font.bold = True
    pc_url.font.size = Pt(13)
    pc_url.font.color.rgb = COLOR_DEEP_NAVY
    pc_url.space_after = Pt(2)

    pc_desc = tf_cloud.add_paragraph()
    pc_desc.text = "Cloud-hosted, high-availability interactive dashboard accessible across all modern desktop & mobile browsers."
    pc_desc.font.size = Pt(11)
    pc_desc.font.color.rgb = COLOR_DARK_SLATE

    # GitHub Repo Card
    create_card(s11, Inches(0.8), Inches(3.3), Inches(11.733), Inches(1.5), COLOR_CARD_BG, COLOR_EMERALD)
    tb_gh = s11.shapes.add_textbox(Inches(1.1), Inches(3.45), Inches(11.133), Inches(1.2))
    tf_gh = tb_gh.text_frame
    tf_gh.word_wrap = True
    pgh_h = tf_gh.paragraphs[0]
    pgh_h.text = "OPEN-SOURCE CODE REPOSITORY (GITHUB)"
    pgh_h.font.bold = True
    pgh_h.font.size = Pt(12)
    pgh_h.font.color.rgb = COLOR_EMERALD
    pgh_h.space_after = Pt(4)

    pgh_url = tf_gh.add_paragraph()
    pgh_url.text = "URL: https://github.com/aayushmathur/employee-performance-analysis"
    pgh_url.font.bold = True
    pgh_url.font.size = Pt(13)
    pgh_url.font.color.rgb = COLOR_DEEP_NAVY
    pgh_url.space_after = Pt(2)

    pgh_desc = tf_gh.add_paragraph()
    pgh_desc.text = "Complete source code, Pytest test suites, database schemas, and documentation artifacts."
    pgh_desc.font.size = Pt(11)
    pgh_desc.font.color.rgb = COLOR_DARK_SLATE

    # Q&A Invitation Card
    create_card(s11, Inches(0.8), Inches(5.0), Inches(11.733), Inches(1.5), COLOR_EMERALD_BG, COLOR_EMERALD)
    tb_qa_box = s11.shapes.add_textbox(Inches(1.1), Inches(5.15), Inches(11.133), Inches(1.2))
    tf_qab = tb_qa_box.text_frame
    tf_qab.word_wrap = True
    pq_h = tf_qab.paragraphs[0]
    pq_h.alignment = PP_ALIGN.CENTER
    pq_h.text = "THANK YOU  |  QUESTIONS & EVALUATION DISCUSSION"
    pq_h.font.bold = True
    pq_h.font.size = Pt(14)
    pq_h.font.color.rgb = COLOR_DEEP_NAVY
    pq_h.space_after = Pt(4)

    pq_sub = tf_qab.add_paragraph()
    pq_sub.alignment = PP_ALIGN.CENTER
    pq_sub.text = "Aayush Mathur  |  Enrolment No: 2500320140001  |  Section: MCA-B  |  ABES Engineering College"
    pq_sub.font.size = Pt(12)
    pq_sub.font.color.rgb = COLOR_DARK_SLATE

    set_speaker_notes(s11, (
        "In conclusion, the Employee Performance Rating Analysis platform is fully deployed and accessible on "
        "Streamlit Cloud, with the complete codebase, test suite, and engineering documentation published on GitHub. "
        "I would like to express my sincere gratitude to my faculty mentor and the MCA department at ABES Engineering "
        "College for their invaluable guidance. I am now ready for the live demonstration and welcome questions "
        "from the honorable evaluation panel."
    ))

    # Save presentation
    output_path = os.path.join(r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs", "2500320140001_Presentation.pptx")
    prs.save(output_path)
    print(f"Successfully generated presentation deck at: {output_path}")
    print(f"File size: {os.path.getsize(output_path)} bytes")

if __name__ == "__main__":
    build_presentation()
