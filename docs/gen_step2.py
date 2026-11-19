"""Generate 02_HMW_Statements.docx for the Employee Performance Rating Analysis project."""

import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

OUT_PATH = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\02_HMW_Statements.docx"

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(2.8)
    section.right_margin  = Cm(2.8)

# ── Helper utilities ──────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    """Set table cell background colour."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)

def header_row(table, labels, bg="1F4E79"):
    """Style the first row as a coloured header."""
    hdr = table.rows[0]
    for i, label in enumerate(labels):
        cell = hdr.cells[i]
        cell.text = ""
        run  = cell.paragraphs[0].add_run(label)
        run.bold      = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_bg(cell, bg)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

def add_heading1(doc, text):
    p = doc.add_heading(text, level=1)
    p.runs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    return p

def add_heading2(doc, text):
    p = doc.add_heading(text, level=2)
    p.runs[0].font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    return p

def normal(doc, text, bold=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    return p

# ══════════════════════════════════════════════════════════════════════════════
# TITLE PAGE
# ══════════════════════════════════════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()

title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run("EMPLOYEE PERFORMANCE RATING ANALYSIS")
run.bold = True
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

doc.add_paragraph()

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub_p.add_run("Step 2 — How Might We (HMW) Statements")
run.bold = True
run.font.size = Pt(16)
run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

doc.add_paragraph()
doc.add_paragraph()

# Student details block
details = [
    ("Student Name",     "Aayush Mathur"),
    ("Enrolment No.",    "2500320140001"),
    ("Section",          "MCA-B"),
    ("Institution",      "ABES Engineering College"),
    ("Programme",        "Master of Computer Applications (MCA)"),
    ("Subject",          "Data Visualisation / Business Intelligence Project"),
    ("Date",             "19 November 2026"),
]
for label, value in details:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run(f"{label}: ")
    r1.bold = True
    r1.font.size = Pt(12)
    r2 = p.add_run(value)
    r2.font.size = Pt(12)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — INTRODUCTION
# ══════════════════════════════════════════════════════════════════════════════
add_heading1(doc, "1. Introduction")
normal(doc,
    "How Might We (HMW) statements convert identified stakeholder pain points into focused "
    "design and analytical questions. Each statement reframes a specific problem as an "
    "opportunity — narrow enough to be actionable, yet open enough to invite multiple "
    "solutions. This document derives HMW statements directly from the operational questions "
    "and stakeholder map established in Step 1 (Project Understanding).")

normal(doc,
    "The HMW statements in this document guide the design of the Employee Performance "
    "Rating Analysis dashboard: every chart, filter, and KPI must address at least one "
    "statement below.")

doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — STAKEHOLDER PAIN POINTS
# ══════════════════════════════════════════════════════════════════════════════
add_heading1(doc, "2. Stakeholder Pain Points")
normal(doc,
    "The following pain points were identified through stakeholder analysis. Each pain point "
    "represents a gap between what the stakeholder currently experiences and what they need "
    "to make effective, data-driven decisions about employee performance.")

doc.add_paragraph()

pain_points = {
    "SH-01 — HR Manager": [
        ("PP-01-A", "Cannot easily identify which departments are underperforming relative to the company average without manually compiling and cross-referencing multiple spreadsheets."),
        ("PP-01-B", "The manual appraisal rating process introduces unconscious bias because there is no standardised benchmark or comparative view to validate individual manager scores."),
        ("PP-01-C", "No consolidated view of promotion readiness — HR must track tenure, training, and rating data separately and combine them manually, leading to missed promotion cycles."),
    ],
    "SH-02 — Department Head": [
        ("PP-02-A", "Cannot benchmark their own team's performance against the company average or peer departments without requesting a custom report from HR."),
        ("PP-02-B", "Hard to identify star performers and outliers objectively, as performance ratings are communicated in aggregate form without role-level or seniority-level breakdown."),
        ("PP-02-C", "No data to justify salary adjustment recommendations — cannot link an employee's performance trajectory to compensation changes in a defensible, data-backed way."),
    ],
    "SH-03 — Senior Management": [
        ("PP-03-A", "No bird's-eye view of organisational performance health — currently receives quarterly PDF reports that are outdated by the time decisions are made."),
        ("PP-03-B", "Cannot predict attrition risk early because low-rating patterns in specific departments are not surfaced proactively; attrition is noticed only after resignations occur."),
        ("PP-03-C", "Hard to verify whether appraisal policies are applied fairly and consistently across departments, roles, and demographic groups, exposing the organisation to compliance risk."),
    ],
    "SH-04 — Employee": [
        ("PP-04-A", "No transparency on how performance ratings are calculated — employees receive a final score with no explanation of which dimensions (attendance, KPIs, training) were weighted and how."),
        ("PP-04-B", "Cannot see career progression benchmarks for their role — unclear what rating, tenure, or training milestones trigger a promotion or salary band change."),
        ("PP-04-C", "Unclear what specific factors (e.g. number of training hours, years of experience) most strongly predict high ratings among peers in the same job role."),
    ],
    "SH-05 — BI Analyst": [
        ("PP-05-A", "Raw HR data exports require extensive cleaning (missing values, inconsistent department codes, duplicate employee records) before any reliable analysis can begin."),
        ("PP-05-B", "No standardised KPI definitions across the organisation — 'high performer' means different things in different departments, making cross-department comparisons unreliable."),
        ("PP-05-C", "Performance trends over time are hard to spot without dedicated time-series views; current static pivot tables make it difficult to identify inflection points or seasonal patterns."),
    ],
    "SH-06 — IT Admin": [
        ("PP-06-A", "No automated data refresh pipeline — analysts must manually export, transform, and re-upload HR data every reporting cycle, creating version-control and freshness issues."),
        ("PP-06-B", "Data quality issues (null rating scores, out-of-range age values, mismatched employee IDs) are not flagged automatically; errors propagate silently into dashboards and distort KPIs."),
    ],
}

for sh, points in pain_points.items():
    add_heading2(doc, sh)
    tbl = doc.add_table(rows=1 + len(points), cols=2)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_row(tbl, ["Pain Point ID", "Description"])
    set_col_width(tbl, 0, 2.5)
    set_col_width(tbl, 1, 12.5)
    for i, (pid, desc) in enumerate(points):
        row = tbl.rows[i + 1]
        row.cells[0].text = pid
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[1].text = desc
        if i % 2 == 0:
            set_cell_bg(row.cells[0], "DEEAF1")
            set_cell_bg(row.cells[1], "DEEAF1")
    doc.add_paragraph()

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — HMW STATEMENTS TABLE
# ══════════════════════════════════════════════════════════════════════════════
add_heading1(doc, "3. How Might We (HMW) Statements")
normal(doc,
    "Each HMW statement below maps directly to one or more stakeholders, their pain points, "
    "and the operational questions (OQ-01 to OQ-10) defined in Step 1. Statements marked "
    "PRIORITY address the highest-impact, highest-frequency problems and must be solved in "
    "the first release of the dashboard.")

doc.add_paragraph()

PRIORITY_IDS = {"HMW-01", "HMW-02", "HMW-05", "HMW-06", "HMW-07"}

hmw_rows = [
    ("HMW-01",
     "How might we help HR managers instantly compare department-level average performance ratings against the company-wide benchmark, with drill-down to individual employees?",
     "SH-01", "PP-01-A", "OQ-01"),
    ("HMW-02",
     "How might we surface the top 10% and bottom 10% performers across the organisation in a single ranked view, to support objective, evidence-based appraisal decisions?",
     "SH-01", "PP-01-B", "OQ-02"),
    ("HMW-03",
     "How might we visualise the correlation between salary bands and performance ratings for each department, so department heads can build a data-backed case for pay adjustments?",
     "SH-02", "PP-02-C", "OQ-03"),
    ("HMW-04",
     "How might we identify and display which manager's team consistently outperforms all other teams quarter over quarter, enabling recognition and best-practice sharing?",
     "SH-02", "PP-02-A", "OQ-04"),
    ("HMW-05",
     "How might we detect and flag statistically significant demographic patterns in performance ratings — by gender and age group — to proactively surface potential appraisal bias?",
     "SH-03", "PP-03-C", "OQ-05"),
    ("HMW-06",
     "How might we build a promotion-readiness scorecard that automatically ranks employees by a weighted combination of tenure, latest performance rating, and training hours completed?",
     "SH-04", "PP-04-B, PP-01-C", "OQ-10"),
    ("HMW-07",
     "How might we give senior management a real-time attrition-risk heatmap by department, based on sustained low-rating trends, before the risk crystallises into actual resignations?",
     "SH-03", "PP-03-B", "OQ-07"),
    ("HMW-08",
     "How might we show each employee the top factors that most influence high performance ratings among peers in the same job role, so they can focus their development efforts?",
     "SH-04", "PP-04-A, PP-04-C", "OQ-09"),
    ("HMW-09",
     "How might we produce a 3-year rolling performance trend line chart that automatically highlights inflection points (significant rises or drops) with contextual annotations?",
     "SH-05", "PP-05-C", "OQ-08"),
    ("HMW-10",
     "How might we detect, quantify, and report data-quality issues in the HR dataset — missing ratings, duplicate IDs, out-of-range values — before they distort any analytical output?",
     "SH-06", "PP-06-B", "OQ-01 to OQ-10"),
    ("HMW-11",
     "How might we let department heads slice and filter performance data interactively by job role and seniority level to uncover hidden performance patterns within their teams?",
     "SH-02", "PP-02-B", "OQ-06"),
    ("HMW-12",
     "How might we help senior management track, year on year, whether each appraisal cycle has resulted in a measurable improvement in average performance ratings organisation-wide?",
     "SH-03", "PP-03-A", "OQ-08"),
]

tbl = doc.add_table(rows=1 + len(hmw_rows), cols=5)
tbl.style = "Table Grid"
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(tbl, ["HMW ID", "How might we …?", "Stakeholder", "Pain Point", "Linked OQ"])
set_col_width(tbl, 0, 1.5)
set_col_width(tbl, 1, 8.5)
set_col_width(tbl, 2, 1.5)
set_col_width(tbl, 3, 2.0)
set_col_width(tbl, 4, 2.0)

for i, (hmw_id, statement, sh, pp, oq) in enumerate(hmw_rows):
    row = tbl.rows[i + 1]
    is_priority = hmw_id in PRIORITY_IDS
    bg = "FFF2CC" if is_priority else ("EBF3FB" if i % 2 == 0 else "FFFFFF")
    for ci in range(5):
        set_cell_bg(row.cells[ci], bg)

    # HMW ID cell
    p0 = row.cells[0].paragraphs[0]
    r0 = p0.add_run(hmw_id)
    r0.bold = True
    r0.font.size = Pt(9)
    if is_priority:
        p0.add_run("\n")
        rp = p0.add_run("★ PRIORITY")
        rp.bold = True
        rp.font.size = Pt(8)
        rp.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    row.cells[1].text = statement
    row.cells[1].paragraphs[0].runs[0].font.size = Pt(9)
    row.cells[2].text = sh
    row.cells[2].paragraphs[0].runs[0].font.size = Pt(9)
    row.cells[3].text = pp
    row.cells[3].paragraphs[0].runs[0].font.size = Pt(9)
    row.cells[4].text = oq
    row.cells[4].paragraphs[0].runs[0].font.size = Pt(9)

doc.add_paragraph()
doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — PRIORITY STATEMENTS
# ══════════════════════════════════════════════════════════════════════════════
add_heading1(doc, "4. Priority HMW Statements — Justification")
normal(doc,
    "The following five statements are designated PRIORITY because they address the most "
    "frequent stakeholder needs, have the broadest organisational impact, and are directly "
    "answerable from the IBM HR Analytics dataset available to this project.")

doc.add_paragraph()

priority_reasons = [
    ("HMW-01",
     "Department-vs-benchmark comparison",
     "This is the single most-requested HR metric. Without it, every other analysis lacks "
     "a reference point. It directly feeds the opening screen of the dashboard and is used "
     "daily by HR managers during appraisal season."),
    ("HMW-02",
     "Top/bottom 10% performer ranking",
     "Identifying outlier performers is the core output of any performance review system. "
     "Surfacing this automatically reduces the risk of recency bias and subjective gut-feel "
     "decisions that currently dominate the appraisal process."),
    ("HMW-05",
     "Demographic pattern detection for bias",
     "Legal and compliance risk: if ratings systematically skew by gender or age, the "
     "organisation may violate equal-opportunities policies. Early detection through the "
     "dashboard reduces liability and builds employee trust."),
    ("HMW-06",
     "Promotion-readiness scorecard",
     "Missed or delayed promotions are a leading driver of high-performer attrition. "
     "Automating the scorecard eliminates the manual effort HR currently spends cross-referencing "
     "three separate data sources and prevents eligible employees from being overlooked."),
    ("HMW-07",
     "Attrition-risk heatmap",
     "Reactive attrition management is expensive — replacing an employee costs 50–200 % of "
     "their annual salary. A proactive, real-time risk heatmap gives senior management the "
     "lead time to intervene (targeted retention bonuses, manager coaching) before "
     "resignations occur."),
]

tbl2 = doc.add_table(rows=1 + len(priority_reasons), cols=4)
tbl2.style = "Table Grid"
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(tbl2, ["HMW ID", "Theme", "Why PRIORITY?", "Linked OQ"])
set_col_width(tbl2, 0, 1.5)
set_col_width(tbl2, 1, 3.5)
set_col_width(tbl2, 2, 9.0)
set_col_width(tbl2, 3, 1.5)

priority_oqs = {
    "HMW-01": "OQ-01",
    "HMW-02": "OQ-02",
    "HMW-05": "OQ-05",
    "HMW-06": "OQ-10",
    "HMW-07": "OQ-07",
}

for i, (hmw_id, theme, reason) in enumerate(priority_reasons):
    row = tbl2.rows[i + 1]
    bg = "FFF2CC"
    for ci in range(4):
        set_cell_bg(row.cells[ci], bg)
    id_run = row.cells[0].paragraphs[0].add_run(hmw_id)
    id_run.bold = True; id_run.font.size = Pt(9)
    row.cells[1].text = theme
    row.cells[1].paragraphs[0].runs[0].bold = True
    row.cells[1].paragraphs[0].runs[0].font.size = Pt(9)
    row.cells[2].text = reason
    row.cells[2].paragraphs[0].runs[0].font.size = Pt(9)
    row.cells[3].text = priority_oqs[hmw_id]
    row.cells[3].paragraphs[0].runs[0].font.size = Pt(9)

doc.add_paragraph()
doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — OQ COVERAGE CHECK
# ══════════════════════════════════════════════════════════════════════════════
add_heading1(doc, "5. OQ Coverage Check")
normal(doc,
    "The table below verifies that every Operational Question (OQ-01 to OQ-10) defined in "
    "Step 1 is addressed by at least one HMW statement. This ensures no analytical "
    "requirement has been left without a design question.")

doc.add_paragraph()

coverage = [
    ("OQ-01", "Average performance rating per department",                                     "HMW-01, HMW-10", "✔"),
    ("OQ-02", "Consistently high/low performers over time",                                    "HMW-02, HMW-10", "✔"),
    ("OQ-03", "Salary vs performance correlation",                                             "HMW-03, HMW-10", "✔"),
    ("OQ-04", "Manager's team with highest average performance",                               "HMW-04, HMW-10", "✔"),
    ("OQ-05", "Rating distribution by gender/age group",                                       "HMW-05, HMW-10", "✔"),
    ("OQ-06", "Significant performance differences across job roles",                          "HMW-11, HMW-10", "✔"),
    ("OQ-07", "Departments with highest attrition risk (low ratings)",                         "HMW-07, HMW-10", "✔"),
    ("OQ-08", "Overall company performance trend over 3 years",                                "HMW-09, HMW-12, HMW-10", "✔"),
    ("OQ-09", "Factors (experience, education, training) predicting high performance",         "HMW-08, HMW-10", "✔"),
    ("OQ-10", "Employees overdue for promotion (tenure + rating)",                             "HMW-06, HMW-10", "✔"),
]

tbl3 = doc.add_table(rows=1 + len(coverage), cols=4)
tbl3.style = "Table Grid"
tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
header_row(tbl3, ["OQ ID", "Operational Question", "Covered by HMW(s)", "Covered?"])
set_col_width(tbl3, 0, 1.3)
set_col_width(tbl3, 1, 8.5)
set_col_width(tbl3, 2, 4.0)
set_col_width(tbl3, 3, 1.5)

for i, (oq_id, question, covered_by, status) in enumerate(coverage):
    row = tbl3.rows[i + 1]
    bg = "E2EFDA" if i % 2 == 0 else "FFFFFF"
    for ci in range(4):
        set_cell_bg(row.cells[ci], bg)
    id_run = row.cells[0].paragraphs[0].add_run(oq_id)
    id_run.bold = True; id_run.font.size = Pt(9)
    row.cells[1].text = question
    row.cells[1].paragraphs[0].runs[0].font.size = Pt(9)
    row.cells[2].text = covered_by
    row.cells[2].paragraphs[0].runs[0].font.size = Pt(9)
    p_status = row.cells[3].paragraphs[0]
    r_status = p_status.add_run(status)
    r_status.bold = True
    r_status.font.size = Pt(11)
    r_status.font.color.rgb = RGBColor(0x37, 0x5A, 0x23)
    p_status.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

normal(doc,
    "✔  All 10 operational questions are covered. No analytical gap remains. "
    "HMW-10 (data quality) is listed as a secondary cover for every OQ because "
    "data-quality failure would invalidate the answer to any question.", size=10)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading1(doc, "6. Summary and Next Steps")
normal(doc,
    "This document has translated stakeholder pain points into 12 specific, actionable "
    "How Might We statements. Each statement is traceable back to a named stakeholder, "
    "a concrete pain point, and one or more operational questions from Step 1.")

doc.add_paragraph()
normal(doc, "Key outputs of this step:", bold=True)
bullets = [
    "6 stakeholders mapped to 16 distinct pain points across HR, management, analytics, and IT domains.",
    "12 HMW statements covering all 10 operational questions.",
    "5 statements marked PRIORITY, with explicit justification linking each to business risk or decision frequency.",
    "Full OQ coverage verified — no question is left without a design prompt.",
]
for b in bullets:
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(b).font.size = Pt(11)

doc.add_paragraph()
normal(doc, "Next Steps (Step 3):", bold=True)
next_steps = [
    "Step 3 — Data Profiling: Profile the IBM HR Analytics dataset, document column-level statistics, missing values, and data types to validate that OQ answers are achievable.",
    "Step 4 — Dashboard Wireframing: Translate HMW statements into low-fidelity screen wireframes, one wireframe per PRIORITY HMW.",
    "Step 5 — KPI Definition: Define measurable KPIs for each HMW, specifying formula, unit, data source column, and refresh frequency.",
]
for ns in next_steps:
    p = doc.add_paragraph(style="List Number")
    p.add_run(ns).font.size = Pt(11)

# ── Save ──────────────────────────────────────────────────────────────────────
doc.save(OUT_PATH)
print(f"Saved: {OUT_PATH}")
print(f"Size : {os.path.getsize(OUT_PATH):,} bytes")
assert os.path.exists(OUT_PATH), "File not found after save!"
print("Verification: OK")
