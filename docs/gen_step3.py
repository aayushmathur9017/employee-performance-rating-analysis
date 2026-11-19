"""
Generate Step 3 - Needs Analysis document for Employee Performance Rating Analysis project.
"""
import os
from datetime import date
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT_PATH = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\03_Needs_Analysis.docx"

# ── helpers ──────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def bold_cell(cell, text, font_size=10, center=False, color=None):
    para = cell.paragraphs[0]
    para.clear()
    run = para.add_run(text)
    run.bold = True
    run.font.size = Pt(font_size)
    if color:
        run.font.color.rgb = RGBColor(*bytes.fromhex(color))
    if center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

def fill_cell(cell, text, font_size=9.5, center=False):
    para = cell.paragraphs[0]
    para.clear()
    run = para.add_run(text)
    run.font.size = Pt(font_size)
    if center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

def add_header_row(table, headers, bg="1F3864", fg="FFFFFF"):
    row = table.rows[0]
    for i, h in enumerate(headers):
        cell = row.cells[i]
        set_cell_bg(cell, bg)
        bold_cell(cell, h, font_size=10, center=True, color=fg)

def add_data_row(table, values, alt=False):
    row = table.add_row()
    bg = "EBF0FA" if alt else "FFFFFF"
    for i, v in enumerate(values):
        cell = row.cells[i]
        set_cell_bg(cell, bg)
        fill_cell(cell, v)

def set_col_widths(table, widths_in):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Inches(widths_in[i])

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.style.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    return h

# ── document ─────────────────────────────────────────────────────────────────

doc = Document()

# --- page margins ---
for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1.1)
    section.right_margin  = Inches(1.1)

# ════════════════════════════════════════════════════════════════════════════
# TITLE PAGE
# ════════════════════════════════════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("EMPLOYEE PERFORMANCE RATING ANALYSIS")
run.bold = True
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

doc.add_paragraph()

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run("Step 3 – Needs Analysis")
r.bold = True
r.font.size = Pt(16)
r.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

doc.add_paragraph()
doc.add_paragraph()

details = [
    ("Student Name",    "Aayush Mathur"),
    ("Enrolment No.",   "2500320140001"),
    ("Section",         "MCA-B"),
    ("Institution",     "ABES Engineering College"),
    ("Subject",         "Data Analysis / Business Intelligence Project"),
    ("Date",            str(date.today())),
    ("Document",        "03_Needs_Analysis.docx"),
]

tbl = doc.add_table(rows=len(details), cols=2)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(details):
    r0, r1 = tbl.rows[i].cells
    set_cell_bg(r0, "1F3864")
    bold_cell(r0, k, color="FFFFFF")
    set_cell_bg(r1, "EBF0FA" if i % 2 else "FFFFFF")
    fill_cell(r1, v)
set_col_widths(tbl, [2.2, 3.8])

doc.add_page_break()

# ════════════════════════════════════════════════════════════════════════════
# 1. INTRODUCTION
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "1. Introduction", 1)

intro = (
    "Needs Analysis is the bridge between understanding a problem and specifying what a solution "
    "must do. In the context of the Employee Performance Rating Analysis project, this step "
    "translates the twelve 'How Might We' (HMW) problem statements—identified during the Design "
    "Thinking empathise and define phases—into concrete, testable Business Requirements (BRs). "
    "Each requirement describes a capability the system must provide, the stakeholder it serves, "
    "its MoSCoW priority, and a clear acceptance criterion by which its delivery can be verified. "
    "The document also maps every HMW to at least one BR (traceability) and records any data needed, "
    "ensuring no stakeholder need is left unaddressed before solution design begins."
)
p = doc.add_paragraph(intro)
p.style.font.size = Pt(10.5)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# 2. BUSINESS REQUIREMENTS TABLE
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "2. Business Requirements", 1)

doc.add_paragraph(
    "The table below lists 14 business requirements derived from the HMW statements. "
    "Requirements are written as 'The business needs to <capability> so that <benefit>.' "
    "Priority follows the MoSCoW model: Must (critical), Should (important), Could (nice-to-have), Won't (deferred)."
).style.font.size = Pt(10)

doc.add_paragraph()

br_data = [
    # BR ID | Requirement | Stakeholder | Source HMW | Priority | Measurement
    ("BR-01",
     "Display department-wise average performance ratings on a central dashboard so that HR managers can benchmark teams against the company-wide average.",
     "SH-01", "HMW-01", "Must",
     "KPI card per department shows mean rating ± std dev; dashboard loads within 3 s."),
    ("BR-02",
     "Rank all employees by performance rating and highlight the top 10 % and bottom 10 % so that appraisal committees can make objective promotion and improvement decisions.",
     "SH-01", "HMW-02", "Must",
     "Sortable table includes a 'Percentile' column; top/bottom rows visually distinguished."),
    ("BR-03",
     "Compute and display the Pearson correlation coefficient between monthly salary and performance rating per department so that pay equity issues are made visible.",
     "SH-02", "HMW-03", "Should",
     "Scatter plot with best-fit line and r-value label is present for each department."),
    ("BR-04",
     "Aggregate performance ratings by reporting manager and rank managers by their team's average score so that leadership effectiveness is measurable.",
     "SH-02", "HMW-04", "Must",
     "Horizontal bar chart ranks ≥5 managers; average and employee-count labels shown."),
    ("BR-05",
     "Break down rating distributions by gender and age group and apply a statistical test so that potential demographic bias in ratings can be detected and flagged.",
     "SH-03", "HMW-05", "Must",
     "Box plots by gender and by age bracket shown; independent-samples t-test p-value displayed."),
    ("BR-06",
     "Calculate a weighted promotion-readiness score (rating 50 %, tenure 30 %, training hours 20 %) for every employee so that promotion decisions are data-driven and consistent.",
     "SH-04", "HMW-06", "Must",
     "Score 0–100 displayed per employee; employees scoring > 70 are automatically flagged 'Promotion Ready'."),
    ("BR-07",
     "Compute an attrition-risk level for each department based on average low rating and absenteeism rate so that senior management can take proactive retention actions.",
     "SH-03", "HMW-07", "Must",
     "Each department is classified as High / Medium / Low risk; heatmap colour-coded accordingly."),
    ("BR-08",
     "Display a 3-year monthly or quarterly performance trend line chart for the entire company with year-over-year percentage change so that management can track progress over time.",
     "SH-05", "HMW-09 / HMW-12", "Must",
     "Line chart with trend line visible; YoY % change annotation present; inflection points flagged."),
    ("BR-09",
     "Run a regression or feature-importance model to identify which employee attributes (education, experience, training) correlate most with high performance so that HR can design targeted training programmes.",
     "SH-05", "HMW-08", "Should",
     "Feature-importance chart or regression-coefficient table shows top 5 predictors with effect sizes."),
    ("BR-10",
     "Provide interactive filters (department, job role, gender, year, manager) that apply simultaneously to all report screens so that any stakeholder can slice the data to their specific context.",
     "SH-01, SH-02", "HMW-11", "Must",
     "Every chart and table updates within 2 s whenever a filter value is changed."),
    ("BR-11",
     "Generate an automated data-quality report that flags missing values, outliers, and duplicate records before analysis runs so that all insights are based on clean, reliable data.",
     "SH-06", "HMW-10", "Must",
     "Report shows count and percentage of missing values per column, outlier count, and % clean rows."),
    ("BR-12",
     "Allow any user to download the currently filtered dataset as a CSV or Excel file so that analysts and HR staff can perform offline review and audit.",
     "SH-05", "HMW-10", "Could",
     "Download button present on the Data page; exported file matches on-screen filtered rows exactly."),
    ("BR-13",
     "Display a statistical comparison of performance ratings across all job roles using ANOVA or Kruskal-Wallis tests so that role-based performance disparities are validated beyond visual inspection.",
     "SH-02", "HMW-11", "Should",
     "Test p-value and per-group mean table displayed; roles with significant differences highlighted."),
    ("BR-14",
     "Flag employees who have not been promoted in more than 3 years and hold a performance rating ≥ 4.0 so that HR can identify and act on overdue promotions promptly.",
     "SH-01, SH-04", "HMW-06", "Must",
     "Filtered list shows employee name, years since last promotion, and current rating; exportable."),
]

headers_br = ["BR ID", "Requirement", "Stakeholder", "Source HMW", "Priority", "Acceptance Criterion / Measurement"]
col_widths_br = [0.55, 2.8, 0.75, 0.9, 0.65, 2.0]

tbl_br = doc.add_table(rows=1, cols=6)
tbl_br.style = "Table Grid"
tbl_br.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_br, headers_br)

priority_colors = {"Must": "1F3864", "Should": "2E74B5", "Could": "5B9BD5", "Won't": "808080"}

for idx, row in enumerate(br_data):
    tr = tbl_br.add_row()
    bg = "EBF0FA" if idx % 2 else "FFFFFF"
    for ci, val in enumerate(row):
        cell = tr.cells[ci]
        set_cell_bg(cell, bg)
        if ci == 4:  # priority
            p_clr = priority_colors.get(val, "000000")
            para = cell.paragraphs[0]
            para.clear()
            run = para.add_run(val)
            run.bold = True
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(*bytes.fromhex(p_clr))
        else:
            fill_cell(cell, val)

set_col_widths(tbl_br, col_widths_br)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# 3. DATA REQUIREMENTS
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "3. Data Requirements", 1)

doc.add_paragraph(
    "The table below specifies the data columns required by each business requirement and the source table in the IBM HR Analytics dataset."
).style.font.size = Pt(10)
doc.add_paragraph()

data_req = [
    ("BR-01", "Department, PerformanceRating, EmployeeID", "employee_data"),
    ("BR-02", "EmployeeID, EmployeeName, Department, PerformanceRating, YearsAtCompany", "employee_data"),
    ("BR-03", "Department, MonthlyIncome, PerformanceRating", "employee_data"),
    ("BR-04", "ManagerID, ManagerName, EmployeeID, PerformanceRating", "employee_data"),
    ("BR-05", "Gender, Age, PerformanceRating, Department", "employee_data"),
    ("BR-06", "EmployeeID, PerformanceRating, YearsAtCompany, TrainingTimesLastYear", "employee_data"),
    ("BR-07", "Department, PerformanceRating, Attrition, AbsenceRate (derived)", "employee_data, attrition_data"),
    ("BR-08", "PerformanceRating, ReviewDate (Year, Month/Quarter)", "performance_history"),
    ("BR-09", "PerformanceRating, Education, TotalWorkingYears, TrainingTimesLastYear, JobLevel, BusinessTravel", "employee_data"),
    ("BR-10", "All columns: Department, JobRole, Gender, ReviewYear, ManagerName", "employee_data, performance_history"),
    ("BR-11", "All columns (data-quality scan)", "employee_data, performance_history"),
    ("BR-12", "All filtered columns (user-selected)", "employee_data, performance_history"),
    ("BR-13", "JobRole, PerformanceRating", "employee_data"),
    ("BR-14", "EmployeeID, EmployeeName, YearsSinceLastPromotion, PerformanceRating", "employee_data"),
]

headers_dr = ["BR ID", "Data Columns Needed", "Source Table"]
col_widths_dr = [0.7, 4.3, 2.0]

tbl_dr = doc.add_table(rows=1, cols=3)
tbl_dr.style = "Table Grid"
tbl_dr.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_dr, headers_dr)

for idx, row in enumerate(data_req):
    tr = tbl_dr.add_row()
    bg = "EBF0FA" if idx % 2 else "FFFFFF"
    for ci, val in enumerate(row):
        cell = tr.cells[ci]
        set_cell_bg(cell, bg)
        fill_cell(cell, val)

set_col_widths(tbl_dr, col_widths_dr)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# 4. GAPS ANALYSIS
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "4. Gaps Analysis", 1)

doc.add_paragraph(
    "A gap exists when a HMW statement has no corresponding BR, or when a BR cannot be traced back to a HMW. "
    "Both indicate either an unmet stakeholder need or unnecessary scope. "
    "The analysis below confirms full coverage."
).style.font.size = Pt(10)
doc.add_paragraph()

add_heading(doc, "4.1 HMWs Without a Business Requirement", 2)

hmw_gap_data = [
    ("HMW-01", "Compare dept ratings vs benchmark", "BR-01", "Covered – KPI card per dept"),
    ("HMW-02", "Surface top/bottom 10 % performers", "BR-02", "Covered – ranked table with percentile"),
    ("HMW-03", "Salary vs performance correlation", "BR-03", "Covered – scatter plot + r-value"),
    ("HMW-04", "Manager whose team outperforms", "BR-04", "Covered – manager ranking bar chart"),
    ("HMW-05", "Demographic patterns / bias detection", "BR-05", "Covered – box plots + t-test"),
    ("HMW-06", "Promotion-readiness scorecard", "BR-06, BR-14", "Covered – weighted score + overdue flag"),
    ("HMW-07", "Attrition-risk heatmap", "BR-07", "Covered – dept risk level heatmap"),
    ("HMW-08", "Factors influencing ratings", "BR-09", "Covered – feature-importance model"),
    ("HMW-09", "3-year trend chart with inflection points", "BR-08", "Covered – trend line chart"),
    ("HMW-10", "Data-quality issues detection", "BR-11, BR-12", "Covered – quality report + download"),
    ("HMW-11", "Filter by role / seniority", "BR-10, BR-13", "Covered – interactive filters + ANOVA"),
    ("HMW-12", "Track whether appraisal cycle improved YoY", "BR-08", "Covered – YoY % change on trend chart"),
]

headers_gap = ["HMW ID", "HMW Summary", "Mapped BR(s)", "Status"]
col_widths_gap = [0.75, 2.5, 1.5, 2.5]

tbl_gap = doc.add_table(rows=1, cols=4)
tbl_gap.style = "Table Grid"
tbl_gap.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_gap, headers_gap)

for idx, row in enumerate(hmw_gap_data):
    tr = tbl_gap.add_row()
    bg = "EBF0FA" if idx % 2 else "FFFFFF"
    for ci, val in enumerate(row):
        cell = tr.cells[ci]
        set_cell_bg(cell, bg)
        fill_cell(cell, val)

set_col_widths(tbl_gap, col_widths_gap)
doc.add_paragraph()

add_heading(doc, "4.2 Business Requirements Without a HMW", 2)
p = doc.add_paragraph(
    "All 14 BRs are traceable to at least one HMW statement. "
    "No orphan requirements exist. "
    "BR-08 is mapped to both HMW-09 and HMW-12 because a single trend-line chart satisfies both the "
    "'3-year trend with inflection points' and the 'YoY appraisal improvement tracking' needs, "
    "which is an intentional consolidation that reduces implementation overhead."
)
p.style.font.size = Pt(10)
doc.add_paragraph()

add_heading(doc, "4.3 Identified Gaps and Resolutions", 2)

gaps_data = [
    ("GAP-01",
     "HMW-12 (track appraisal cycle YoY improvement) partially overlaps HMW-09 (trend chart).",
     "Merged into BR-08. The YoY % change annotation on the trend chart satisfies both HMWs. No separate BR required.",
     "Resolved"),
    ("GAP-02",
     "HMW-10 (data quality) is broad and applies to all BRs.",
     "Addressed by BR-11 (data-quality report) and BR-12 (download for offline audit). "
     "Quality checks are run as a prerequisite to all analysis BRs.",
     "Resolved"),
    ("GAP-03",
     "No explicit requirement for role-based access control (who sees what).",
     "Out of scope for this academic project. Noted for future work. All stakeholders access the full dataset.",
     "Deferred – Won't"),
    ("GAP-04",
     "No real-time data feed (live HRIS integration).",
     "Dataset is static (IBM HR Analytics CSV). Real-time integration is out of scope.",
     "Deferred – Won't"),
]

headers_gaps = ["Gap ID", "Gap Description", "Resolution", "Status"]
col_widths_gaps_tbl = [0.65, 2.3, 3.4, 0.9]

tbl_gaps = doc.add_table(rows=1, cols=4)
tbl_gaps.style = "Table Grid"
tbl_gaps.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_gaps, headers_gaps)

for idx, row in enumerate(gaps_data):
    tr = tbl_gaps.add_row()
    bg = "EBF0FA" if idx % 2 else "FFFFFF"
    for ci, val in enumerate(row):
        cell = tr.cells[ci]
        set_cell_bg(cell, bg)
        fill_cell(cell, val)

set_col_widths(tbl_gaps, col_widths_gaps_tbl)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# 5. TRACEABILITY SUMMARY
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "5. Traceability Summary – HMW → Business Requirement", 1)

doc.add_paragraph(
    "The matrix below shows which Business Requirement(s) address each HMW statement, "
    "confirming complete coverage. A '✓' indicates the HMW is fully satisfied by the mapped BR; "
    "a '(partial)' indicates co-ownership with another BR."
).style.font.size = Pt(10)
doc.add_paragraph()

trace_data = [
    ("HMW-01", "Compare dept ratings vs company benchmark",    "BR-01",        "Full"),
    ("HMW-02", "Surface top/bottom 10 % performers",           "BR-02",        "Full"),
    ("HMW-03", "Salary vs performance correlation per dept",    "BR-03",        "Full"),
    ("HMW-04", "Manager whose team consistently outperforms",   "BR-04",        "Full"),
    ("HMW-05", "Detect demographic patterns / flag bias",       "BR-05",        "Full"),
    ("HMW-06", "Promotion-readiness scorecard",                 "BR-06, BR-14", "Full"),
    ("HMW-07", "Attrition-risk heatmap by dept",               "BR-07",        "Full"),
    ("HMW-08", "Factors influencing ratings in role",          "BR-09",        "Full"),
    ("HMW-09", "3-year trend chart with inflection points",    "BR-08",        "Full"),
    ("HMW-10", "Detect data-quality issues",                   "BR-11, BR-12", "Full"),
    ("HMW-11", "Filter by role / seniority",                   "BR-10, BR-13", "Full"),
    ("HMW-12", "Track appraisal-cycle YoY improvement",        "BR-08",        "Partial – shared with HMW-09"),
]

headers_tr = ["HMW ID", "HMW Statement", "Linked BR(s)", "Coverage"]
col_widths_tr = [0.75, 3.0, 1.5, 2.0]

tbl_tr = doc.add_table(rows=1, cols=4)
tbl_tr.style = "Table Grid"
tbl_tr.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_tr, headers_tr)

for idx, row in enumerate(trace_data):
    tr2 = tbl_tr.add_row()
    bg = "EBF0FA" if idx % 2 else "FFFFFF"
    for ci, val in enumerate(row):
        cell = tr2.cells[ci]
        set_cell_bg(cell, bg)
        fill_cell(cell, val)

set_col_widths(tbl_tr, col_widths_tr)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# 6. MoSCoW SUMMARY
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "6. MoSCoW Priority Summary", 1)

moscow_data = [
    ("Must",   "9", "BR-01, BR-02, BR-04, BR-05, BR-06, BR-07, BR-08, BR-10, BR-11, BR-14", "64 %"),
    ("Should", "3", "BR-03, BR-09, BR-13",                                                   "21 %"),
    ("Could",  "1", "BR-12",                                                                  "7 %"),
    ("Won't",  "1", "Real-time HRIS integration, Role-based access control",                  "7 %"),
]

headers_ms = ["Priority", "Count", "Requirements", "% of Total"]
col_widths_ms = [0.9, 0.7, 4.8, 0.85]

tbl_ms = doc.add_table(rows=1, cols=4)
tbl_ms.style = "Table Grid"
tbl_ms.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_ms, headers_ms)

for idx, row in enumerate(moscow_data):
    tr = tbl_ms.add_row()
    bg = "EBF0FA" if idx % 2 else "FFFFFF"
    for ci, val in enumerate(row):
        cell = tr.cells[ci]
        set_cell_bg(cell, bg)
        fill_cell(cell, val)

set_col_widths(tbl_ms, col_widths_ms)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════════════════════
os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
doc.save(OUTPUT_PATH)

exists = os.path.exists(OUTPUT_PATH)
size   = os.path.getsize(OUTPUT_PATH)
print(f"File created: {exists}")
print(f"Path: {OUTPUT_PATH}")
print(f"Size: {size:,} bytes")
