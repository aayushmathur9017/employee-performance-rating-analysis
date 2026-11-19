"""
Generate 05_BRD.docx — Business Requirements Document
Employee Performance Rating Analysis
Student: Aayush Mathur | Enrolment: 2500320140001 | Section: MCA-B | ABES Engineering College
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
from datetime import date

OUTPUT_PATH = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\05_BRD.docx"

# ── helpers ──────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_bold(cell, bold=True):
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = bold

def style_header_row(table, hex_color='1F3864'):
    for cell in table.rows[0].cells:
        set_cell_bg(cell, hex_color)
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    return h

def add_body(doc, text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Cm(0)
    return p

def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header row
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
    style_header_row(table)
    # Data rows
    for ri, row_data in enumerate(rows):
        row_cells = table.rows[ri + 1].cells
        for ci, val in enumerate(row_data):
            row_cells[ci].text = str(val)
            row_cells[ci].paragraphs[0].paragraph_format.space_before = Pt(2)
            row_cells[ci].paragraphs[0].paragraph_format.space_after = Pt(2)
        # Alternate shading
        if ri % 2 == 0:
            for cell in table.rows[ri + 1].cells:
                set_cell_bg(cell, 'DCE6F1')
    # Column widths
    if col_widths:
        for row in table.rows:
            for ci, width in enumerate(col_widths):
                row.cells[ci].width = Inches(width)
    doc.add_paragraph()
    return table

def page_break(doc):
    doc.add_page_break()

# ── document ─────────────────────────────────────────────────────────────────

doc = Document()

# ── Page margins ──
section = doc.sections[0]
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin   = Cm(3.0)
section.right_margin  = Cm(2.5)

# ── Default styles ──
style = doc.styles['Normal']
style.font.name  = 'Calibri'
style.font.size  = Pt(11)

# ══════════════════════════════════════════════════════════════════════════════
# TITLE PAGE
# ══════════════════════════════════════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run('BUSINESS REQUIREMENTS DOCUMENT')
run.bold = True
run.font.size = Pt(22)
run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

doc.add_paragraph()
sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run2 = sub_p.add_run('Employee Performance Rating Analysis')
run2.bold = True
run2.font.size = Pt(16)
run2.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

doc.add_paragraph()
doc.add_paragraph()

meta_lines = [
    ('Project Name', 'Employee Performance Rating Analysis'),
    ('Document Type', 'Business Requirements Document (BRD)'),
    ('Version', '1.0'),
    ('Date', '19 November 2026'),
    ('Status', 'Draft'),
    ('Prepared By', 'Aayush Mathur'),
    ('Enrolment No.', '2500320140001'),
    ('Section', 'MCA-B'),
    ('Institution', 'ABES Engineering College'),
]
meta_table = doc.add_table(rows=len(meta_lines), cols=2)
meta_table.style = 'Table Grid'
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
for ri, (k, v) in enumerate(meta_lines):
    cells = meta_table.rows[ri].cells
    cells[0].text = k
    cells[1].text = v
    cells[0].paragraphs[0].runs[0].bold = True
    if ri % 2 == 0:
        set_cell_bg(cells[0], 'DCE6F1')
        set_cell_bg(cells[1], 'DCE6F1')
    for cell in cells:
        cell.width = Inches(2.8)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# TABLE OF CONTENTS (static)
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, 'Table of Contents', 1)
toc_items = [
    '1. Document Control',
    '2. Executive Summary',
    '3. Business Problem and Opportunity',
    '4. Objectives and Success Metrics (KPIs)',
    '5. Stakeholders — RACI Matrix',
    '6. Scope',
    '7. Business Requirements (BR-01 to BR-14)',
    '8. Business Rules',
    '9. Data Requirements',
    '10. Assumptions, Constraints, and Dependencies',
    '11. Risks and Mitigation',
    '12. Glossary',
    '13. Sign-off',
]
for item in toc_items:
    p = doc.add_paragraph(item, style='List Number')
    p.paragraph_format.space_after = Pt(4)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — DOCUMENT CONTROL
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '1. Document Control', 1)

add_heading(doc, '1.1 Document Information', 2)
doc_info = [
    ['Field', 'Details'],
    ['Document Title', 'Business Requirements Document — Employee Performance Rating Analysis'],
    ['Document ID', 'BRD-EPRA-2026-001'],
    ['Version', '1.0'],
    ['Author', 'Aayush Mathur'],
    ['Date Created', '19 November 2026'],
    ['Last Modified', '19 November 2026'],
    ['Status', 'Draft'],
    ['Reviewers', 'Faculty Guide'],
    ['Approvers', 'Faculty Guide'],
]
t = doc.add_table(rows=len(doc_info), cols=2)
t.style = 'Table Grid'
t.alignment = WD_TABLE_ALIGNMENT.CENTER
for ri, row_data in enumerate(doc_info):
    cells = t.rows[ri].cells
    cells[0].text = row_data[0]
    cells[1].text = row_data[1]
    if ri == 0:
        set_cell_bg(cells[0], '1F3864')
        set_cell_bg(cells[1], '1F3864')
        for cell in cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    else:
        cells[0].paragraphs[0].runs[0].bold = True
        if ri % 2 == 0:
            set_cell_bg(cells[0], 'DCE6F1')
            set_cell_bg(cells[1], 'DCE6F1')
doc.add_paragraph()

add_heading(doc, '1.2 Change History', 2)
add_table(doc,
    headers=['Version', 'Date', 'Author', 'Description of Change'],
    rows=[
        ['1.0', '19 Nov 2026', 'Aayush Mathur', 'Initial draft — all sections populated for college project submission.'],
    ],
    col_widths=[0.8, 1.2, 1.5, 3.5]
)

add_heading(doc, '1.3 Distribution List', 2)
add_table(doc,
    headers=['Name / Role', 'Purpose', 'Date Sent'],
    rows=[
        ['Faculty Guide', 'Review and approval', '19 Nov 2026'],
        ['Aayush Mathur (Author)', 'Retained copy', '19 Nov 2026'],
        ['Project Supervisor', 'Information', '19 Nov 2026'],
    ],
    col_widths=[2.0, 2.5, 2.0]
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '2. Executive Summary', 1)

add_body(doc,
    'The Employee Performance Rating Analysis project is an academic data analytics initiative '
    'undertaken as part of the Master of Computer Applications (MCA) programme at ABES Engineering '
    'College, Ghaziabad. The project aims to design, develop, and deploy a fully interactive, '
    'multi-page analytics dashboard that enables human-resource professionals, department heads, and '
    'senior management to derive actionable insights from structured employee performance data. '
    'Leveraging a synthetic dataset of 150–300 employee records spanning a minimum of three performance '
    'years, the project integrates data engineering, descriptive statistics, inferential analysis, and '
    'modern web-based visualisation into a single cohesive platform.'
)

add_body(doc,
    'The core business problem being addressed is the prevalence of manual, subjective, and siloed '
    'employee appraisal processes in many mid-sized organisations. Without a unified analytical view, '
    'HR managers struggle to benchmark performance across departments, identify at-risk employees, detect '
    'salary inequities, or make fair and timely promotion decisions. Decisions that should be data-driven '
    'are instead made on intuition and incomplete spreadsheets, leading to employee dissatisfaction, '
    'talent attrition, and suboptimal organisational outcomes. This project demonstrates how applied data '
    'science can transform these ad-hoc processes into evidence-based workflows.'
)

add_body(doc,
    'The primary deliverable is a Streamlit multi-page web application backed by a SQLite relational '
    'database. The application provides eight interactive dashboard pages covering departmental '
    'benchmarking, individual performance ranking, salary–rating correlation analysis, manager '
    'effectiveness scoring, demographic bias detection, promotion readiness scoring, attrition risk '
    'assessment, and three-year trend visualisation. The application is version-controlled via GitHub '
    'and deployed to a public cloud platform, making it accessible to any authorised stakeholder with '
    'a web browser. All source code, data-generation scripts, and analytical artefacts are documented '
    'and published as part of the project deliverables.'
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — BUSINESS PROBLEM AND OPPORTUNITY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '3. Business Problem and Opportunity', 1)

add_heading(doc, '3.1 Current State', 2)
add_body(doc,
    'In a typical mid-sized organisation, employee performance management relies heavily on annual '
    'paper-based or spreadsheet-driven appraisal cycles. The following pain-points characterise the '
    'current state:'
)
current_state_rows = [
    ('CS-01', 'No Unified View', 'Performance data is scattered across multiple spreadsheets maintained independently by department heads. There is no single source of truth.'),
    ('CS-02', 'Subjectivity Bias', 'Ratings assigned by individual managers are rarely calibrated against organisational or departmental norms, introducing unconscious bias.'),
    ('CS-03', 'No Trend Analysis', 'Historical performance is not tracked year-over-year, making it impossible to identify improvement or decline patterns.'),
    ('CS-04', 'Delayed Promotion Decisions', 'Employees eligible for promotion based on tenure and performance are overlooked due to the absence of automated flagging mechanisms.'),
    ('CS-05', 'Invisible Attrition Risk', 'HR teams are unable to predict which departments or individuals are at high risk of leaving the organisation.'),
    ('CS-06', 'Salary Inequity Blind Spot', 'The relationship between salary bands and performance ratings is never formally analysed, leaving potential pay inequities undetected.'),
]
add_table(doc,
    headers=['ID', 'Pain Point', 'Description'],
    rows=current_state_rows,
    col_widths=[0.6, 1.5, 4.9]
)

add_heading(doc, '3.2 Desired State', 2)
add_body(doc,
    'Post-implementation, the organisation (or academic simulation thereof) should achieve the '
    'following desired outcomes:'
)
desired_rows = [
    ('DS-01', 'Centralised Dashboard', 'All performance metrics available in one interactive web application accessible via any browser.'),
    ('DS-02', 'Objective Benchmarking', 'Department-level average ratings computed automatically and displayed with industry-standard KPI cards.'),
    ('DS-03', 'Fair Promotion Pipeline', 'A transparent, formula-driven promotion-readiness score eliminates subjective promotion decisions.'),
    ('DS-04', 'Proactive Attrition Management', 'Risk scores highlight departments and individuals requiring immediate HR intervention.'),
    ('DS-05', 'Gender and Age Equity', 'Statistical tests automatically surface any significant rating disparities across demographic groups.'),
    ('DS-06', 'Manager Accountability', 'Manager rankings by team average rating encourage equitable and constructive leadership.'),
]
add_table(doc,
    headers=['ID', 'Outcome', 'Description'],
    rows=desired_rows,
    col_widths=[0.6, 1.8, 4.6]
)

add_heading(doc, '3.3 Business Opportunity', 2)
add_body(doc,
    'The convergence of open-source data science libraries (Pandas, SciPy, Plotly), no-cost web '
    'frameworks (Streamlit), and free cloud hosting platforms (Streamlit Community Cloud) means that '
    'a production-grade analytics dashboard can be built and deployed at virtually zero infrastructure '
    'cost. For this academic project, a synthetic dataset of 150–300 records will serve as a '
    'representative proxy for real organisational data. The analytical techniques applied — descriptive '
    'statistics, Pearson correlation, ANOVA, Kruskal-Wallis tests, and linear regression — are directly '
    'transferable to real-world deployments. The project thus serves as both a learning exercise and a '
    'reference architecture for practitioners.'
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — OBJECTIVES AND SUCCESS METRICS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '4. Objectives and Success Metrics (KPIs)', 1)

add_body(doc,
    'The table below defines the measurable objectives of the project, the Key Performance Indicators '
    '(KPIs) used to gauge success, the target value or threshold for each KPI, and the measurement '
    'method applied within the dashboard.'
)

obj_rows = [
    ('OBJ-01', 'Enable dept-level performance benchmarking', 'Average rating displayed per department', 'Dashboard loads in < 3 seconds; all depts visible', 'KPI card on Department Overview page'),
    ('OBJ-02', 'Identify top and bottom performers within each department', 'Percentage of employees assigned a performance rank', '100% coverage — all employees ranked', 'Interactive ranking table with top/bottom 10% flags'),
    ('OBJ-03', 'Detect and quantify salary–rating correlation', 'Pearson correlation coefficient (r) between salary and rating', 'Value of r is computed and displayed with p-value', 'Scatter plot with regression line on Salary Analysis page'),
    ('OBJ-04', 'Assess and communicate attrition risk per department', 'Risk level (High / Medium / Low) assigned to each department', 'All departments assigned a risk level', 'Colour-coded heatmap on Attrition Risk page'),
    ('OBJ-05', 'Promote data-driven, transparent promotion decisions', 'Promotion-readiness score (0–100) computed for every employee', 'Score displayed for all employees; overdue flags visible', 'Promotion Scorecard page with formula transparency'),
    ('OBJ-06', 'Ensure high data quality throughout the pipeline', 'Percentage of rows passing all data-quality checks', '≥ 95% clean rows in the dataset', 'Automated Data Quality Report with row-level flags'),
    ('OBJ-07', 'Detect potential demographic bias in performance ratings', 'Statistical significance (p-value) of gender/age rating difference', 'p-value computed and interpreted on the dashboard', 'Gender & Age Bias Analysis page with t-test results'),
    ('OBJ-08', 'Visualise three-year performance trend', 'Year-over-year rating trend for each department', 'Trend chart available for at least 3 years of data', 'Line chart on Trend Analysis page'),
]
add_table(doc,
    headers=['Obj. ID', 'Objective', 'KPI', 'Target', 'Measurement Method'],
    rows=obj_rows,
    col_widths=[0.7, 1.7, 1.5, 1.5, 1.6]
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — STAKEHOLDERS (RACI)
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '5. Stakeholders — RACI Matrix', 1)

add_heading(doc, '5.1 Stakeholder Profiles', 2)
sh_rows = [
    ('SH-01', 'HR Manager', 'Oversees the appraisal process, compensation, and HR policy. Primary consumer of the dashboard.'),
    ('SH-02', 'Department Head', 'Reviews team performance, manager rankings, and department benchmark data.'),
    ('SH-03', 'Senior Management', 'Uses high-level KPIs, attrition risk summaries, and trend charts for strategic decision-making.'),
    ('SH-04', 'Employee', 'Indirect stakeholder whose performance records form the core dataset; benefits from transparent promotion scoring.'),
    ('SH-05', 'BI Analyst', 'Responsible for designing analytical models, writing Python scripts, and maintaining the dashboard.'),
    ('SH-06', 'IT Administrator', 'Manages database provisioning, cloud deployment, access control, and backup.'),
]
add_table(doc,
    headers=['ID', 'Role', 'Responsibilities'],
    rows=sh_rows,
    col_widths=[0.7, 1.5, 4.8]
)

add_heading(doc, '5.2 RACI Matrix', 2)
add_body(doc, 'Legend: R = Responsible | A = Accountable | C = Consulted | I = Informed')
raci_rows = [
    ('Define Business Requirements', 'A', 'C', 'C', 'I', 'R', 'I'),
    ('Provide Source Data / Synthetic Data Design', 'C', 'C', 'I', 'I', 'R', 'A'),
    ('Design Database Schema', 'I', 'I', 'I', 'I', 'R', 'A'),
    ('Develop Dashboard Pages', 'C', 'I', 'I', 'I', 'R', 'I'),
    ('Perform Statistical Analysis', 'C', 'C', 'I', 'I', 'R', 'I'),
    ('Review Dashboard Outputs', 'A', 'R', 'R', 'I', 'C', 'I'),
    ('Approve KPIs and Business Rules', 'A', 'C', 'C', 'I', 'R', 'I'),
    ('Use Analytics for Decisions', 'R', 'R', 'R', 'I', 'C', 'I'),
    ('Maintain System and Database', 'I', 'I', 'I', 'I', 'C', 'R'),
    ('Conduct Data Quality Audit', 'C', 'I', 'I', 'I', 'R', 'A'),
    ('Deploy to Cloud Platform', 'I', 'I', 'I', 'I', 'C', 'R'),
    ('Document and Archive Deliverables', 'I', 'I', 'I', 'I', 'R', 'A'),
]
add_table(doc,
    headers=['Activity', 'SH-01\nHR Mgr', 'SH-02\nDept Head', 'SH-03\nSr Mgmt', 'SH-04\nEmployee', 'SH-05\nBI Analyst', 'SH-06\nIT Admin'],
    rows=raci_rows,
    col_widths=[2.4, 0.9, 0.9, 0.9, 0.9, 0.9, 0.9]
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — SCOPE
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '6. Scope', 1)

add_heading(doc, '6.1 In-Scope', 2)
in_scope = [
    ('SC-01', 'Synthetic employee dataset (150–300 records) generated using Python Faker / NumPy with realistic distributions across departments, roles, tenures, genders, and age groups.'),
    ('SC-02', 'SQLite relational database with five normalised tables: Employee_Master, Performance_Records, Salary_Data, Attrition_Data, and Training_Data.'),
    ('SC-03', 'Eight-page Streamlit multi-page application with a shared sidebar for global filters.'),
    ('SC-04', 'Statistical analyses: descriptive stats, Pearson correlation, t-tests, one-way ANOVA, Kruskal-Wallis H-test, and multiple linear regression.'),
    ('SC-05', 'Automated data-quality checks with a downloadable report flagging nulls, outliers, and type mismatches.'),
    ('SC-06', 'GitHub version control with a complete, documented repository including README, requirements.txt, and project report.'),
    ('SC-07', 'Cloud deployment to Streamlit Community Cloud for public accessibility.'),
    ('SC-08', 'CSV and Excel export functionality for all major data views.'),
    ('SC-09', 'Promotion-readiness scoring and overdue-promotion flagging algorithm.'),
    ('SC-10', 'Attrition-risk scoring at department level using business-rule logic.'),
]
add_table(doc,
    headers=['ID', 'Item'],
    rows=in_scope,
    col_widths=[0.7, 6.3]
)

add_heading(doc, '6.2 Out-of-Scope', 2)
out_scope = [
    ('OS-01', 'Real employee personally identifiable information (PII) — the project uses exclusively synthetic/anonymised data.'),
    ('OS-02', 'Real-time integration with payroll, HRMS, or ERP systems (e.g., SAP, Workday).'),
    ('OS-03', 'Mobile application or native iOS/Android development.'),
    ('OS-04', 'Machine-learning model deployment in a production environment (e.g., churn prediction API).'),
    ('OS-05', 'LDAP, Active Directory, or OAuth-based authentication systems.'),
    ('OS-06', 'Multi-language (i18n) or accessibility (WCAG) compliance beyond basic readability.'),
    ('OS-07', 'Real-time data ingestion pipelines (Kafka, Spark Streaming).'),
]
add_table(doc,
    headers=['ID', 'Exclusion'],
    rows=out_scope,
    col_widths=[0.7, 6.3]
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — BUSINESS REQUIREMENTS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '7. Business Requirements', 1)

add_body(doc,
    'This section details each business requirement identified through stakeholder analysis and design-'
    'thinking exercises. Requirements are prioritised using the MoSCoW framework: Must Have (M), '
    'Should Have (S), Could Have (C), and Won\'t Have (W) for this release.'
)

brs = [
    {
        'id': 'BR-01',
        'title': 'Department-wise Average Rating Dashboard',
        'priority': 'Must Have',
        'hmw': 'HMW-01, HMW-04',
        'oq': 'OQ-01',
        'description': (
            'The system shall compute and display the average employee performance rating for each '
            'department in an interactive dashboard page. The page shall include at minimum: a KPI '
            'card per department showing the average rating (1–5 scale), a horizontal bar chart '
            'ranking all departments by their average rating, and a data table with supporting '
            'statistics (mean, median, standard deviation, count of employees).'
        ),
        'rationale': (
            'Department-level benchmarking is the most fundamental HR analytics requirement. Without '
            'a consolidated departmental view, HR managers and department heads cannot identify '
            'performance gaps or celebrate high-performing teams. This requirement directly addresses '
            'OQ-01 and the pain point CS-01 (No Unified View).'
        ),
        'acceptance': [
            'All departments present in the dataset are represented on the dashboard.',
            'Average rating is accurate to two decimal places.',
            'Dashboard page loads in under 3 seconds with 300 records.',
            'Global filters (year, department) update the chart dynamically.',
            'KPI cards change colour based on rating band (e.g., green ≥ 4.0, amber 3.0–3.9, red < 3.0).',
        ],
    },
    {
        'id': 'BR-02',
        'title': 'Employee Ranking — Top and Bottom 10%',
        'priority': 'Must Have',
        'hmw': 'HMW-02',
        'oq': 'OQ-02',
        'description': (
            'The system shall rank all employees within their respective departments based on their '
            'average performance rating and flag those in the top 10% (high performers) and bottom '
            '10% (low performers). The ranking page shall display an interactive sortable table with '
            'employee ID, name, department, role, average rating, rank, and performer flag. '
            'Colour-coded badges shall visually distinguish high, average, and low performers.'
        ),
        'rationale': (
            'Identifying outlier performers — both high and low — is critical for targeted HR '
            'interventions. High performers should be recognised and retained; low performers may '
            'require coaching or performance improvement plans. Without a data-driven ranking, '
            'managers rely on memory and recency bias.'
        ),
        'acceptance': [
            'Every employee in the dataset has a rank and a performer flag.',
            'Top 10% badge is displayed in green; bottom 10% badge in red.',
            'Rankings are recalculated when the year filter is changed.',
            'The table is exportable to CSV.',
            'Percentile cut-offs are based on within-department calculation, not overall.',
        ],
    },
    {
        'id': 'BR-03',
        'title': 'Salary–Rating Correlation Analysis',
        'priority': 'Should Have',
        'hmw': 'HMW-05',
        'oq': 'OQ-03',
        'description': (
            'The system shall compute and visualise the Pearson correlation coefficient between '
            'employee salary and their average performance rating. The analysis page shall include '
            'a scatter plot with a linear regression line, the computed r value, the p-value for '
            'statistical significance, and an interpretive statement (e.g., "Moderate positive '
            'correlation: r = 0.42, p < 0.05").'
        ),
        'rationale': (
            'Organisations need to validate that compensation is aligned with performance. A weak '
            'or negative correlation indicates potential pay inequity or ineffective salary review '
            'processes. This analysis supports evidence-based compensation benchmarking.'
        ),
        'acceptance': [
            'Scatter plot renders for all employees with valid salary and rating records.',
            'Pearson r and p-value are computed using SciPy stats module.',
            'Interpretive label (weak/moderate/strong, positive/negative) is shown.',
            'Filter by department narrows the scatter plot to that department.',
            'Colour encoding distinguishes salary bands on the scatter plot.',
        ],
    },
    {
        'id': 'BR-04',
        'title': 'Manager Ranking by Team Average Rating',
        'priority': 'Must Have',
        'hmw': 'HMW-06',
        'oq': 'OQ-04',
        'description': (
            'The system shall compute the average performance rating of all direct reports for each '
            'manager and rank managers accordingly. The manager ranking page shall display a sortable '
            'table with manager ID, manager name, department, number of direct reports, team average '
            'rating, and overall rank. A bar chart shall visualise the ranking across all managers.'
        ),
        'rationale': (
            'Manager effectiveness has a direct impact on team performance. By surfacing which managers '
            'lead the highest-performing teams, senior management can identify leadership best practices '
            'and support managers of lower-performing teams with targeted coaching or resources.'
        ),
        'acceptance': [
            'All managers with at least two direct reports are included in the ranking.',
            'Team average rating is accurate to two decimal places.',
            'Bar chart is sorted in descending order by team average.',
            'Clicking a manager row opens a drill-down view of that manager\'s team.',
            'The ranking updates when the performance year filter is applied.',
        ],
    },
    {
        'id': 'BR-05',
        'title': 'Rating Distribution by Gender and Age Group',
        'priority': 'Must Have',
        'hmw': 'HMW-07',
        'oq': 'OQ-05',
        'description': (
            'The system shall display the distribution of performance ratings across demographic '
            'groups — specifically gender (Male, Female, Non-binary) and age group (<30, 30–40, '
            '41–50, >50). Visualisations shall include grouped box plots, violin plots, or histogram '
            'overlays. A two-sample t-test (for gender) and one-way ANOVA (for age groups) shall be '
            'performed and results displayed with p-values and interpretations.'
        ),
        'rationale': (
            'Detecting demographic bias in performance ratings is both an ethical obligation and a '
            'legal compliance requirement in many jurisdictions. If ratings differ significantly '
            'across gender or age groups without a valid performance-based reason, the appraisal '
            'process may be biased.'
        ),
        'acceptance': [
            'Box plots or violin plots are rendered for each demographic group.',
            'T-test result (gender) and ANOVA result (age groups) are displayed with p-values.',
            'A visual flag or alert is shown when p-value < 0.05 (statistically significant bias).',
            'Demographic breakdown is filterable by department and year.',
            'All charts are labelled and include a legend.',
        ],
    },
    {
        'id': 'BR-06',
        'title': 'Promotion-Readiness Score (0–100)',
        'priority': 'Must Have',
        'hmw': 'HMW-09, HMW-10',
        'oq': 'OQ-10',
        'description': (
            'The system shall compute a Promotion-Readiness Score (PRS) for every employee on a '
            '0–100 scale using the formula: PRS = (0.5 × rating_normalised) + '
            '(0.3 × tenure_normalised) + (0.2 × training_score_normalised), where each component '
            'is normalised to [0, 1] before scaling to 100. The scorecard page shall display the '
            'score for each employee with a progress bar or gauge chart and highlight employees '
            'who have never been promoted and have a PRS ≥ 70.'
        ),
        'rationale': (
            'Without a transparent, quantitative promotion criterion, promotion decisions are '
            'perceived as subjective and unfair, leading to disengagement and attrition. A '
            'formula-driven score provides employees and managers with a clear, shared understanding '
            'of what it takes to be promoted.'
        ),
        'acceptance': [
            'PRS is computed for every employee with complete rating, tenure, and training data.',
            'The formula and weighting are displayed on the page for transparency.',
            'Employees with PRS ≥ 70 and never_promoted = True are highlighted.',
            'The scorecard is filterable by department and sortable by PRS descending.',
            'Score values are bounded between 0 and 100 inclusive.',
        ],
    },
    {
        'id': 'BR-07',
        'title': 'Attrition-Risk Score per Department',
        'priority': 'Must Have',
        'hmw': 'HMW-08',
        'oq': 'OQ-07',
        'description': (
            'The system shall assign an attrition-risk level (High, Medium, Low) to each department '
            'based on the business rule: High if avg_rating < 3.0 AND absenteeism > 10%; Medium if '
            'avg_rating < 3.5; Low otherwise. The attrition risk page shall display a colour-coded '
            'heatmap of departments by risk level, a risk summary table, and trend data showing '
            'whether risk has increased or decreased over the available performance years.'
        ),
        'rationale': (
            'Employee attrition is costly — studies estimate replacement costs at 50–200% of an '
            'employee\'s annual salary. Early identification of high-risk departments enables '
            'proactive interventions such as engagement programmes, manager coaching, or compensation '
            'reviews before talent is lost.'
        ),
        'acceptance': [
            'All departments are assigned a risk level (H/M/L) using the defined business rule.',
            'High-risk departments are displayed in red, medium in amber, low in green.',
            'Risk level updates dynamically when the year filter is changed.',
            'Risk trend (improving / stable / worsening) is shown for each department.',
            'An export button allows downloading the risk summary table.',
        ],
    },
    {
        'id': 'BR-08',
        'title': '3-Year Performance Trend Chart',
        'priority': 'Must Have',
        'hmw': 'HMW-11',
        'oq': 'OQ-08',
        'description': (
            'The system shall display an interactive line chart showing the year-over-year average '
            'performance rating for each department across a minimum of three performance years. '
            'The chart shall support multi-line overlays (one line per department), hover tooltips '
            'with exact values, and optional smoothing. A secondary chart shall show the overall '
            'organisation-level rating trend.'
        ),
        'rationale': (
            'Single-year snapshots are insufficient for strategic workforce planning. A multi-year '
            'trend reveals whether a department is on a trajectory of improvement or decline, '
            'enabling timely structural or managerial interventions.'
        ),
        'acceptance': [
            'Line chart displays data for at least 3 years from the synthetic dataset.',
            'Each department has a distinct line colour and legend entry.',
            'Hovering over a data point shows the year, department, and exact average rating.',
            'A department filter allows users to show/hide individual department lines.',
            'Overall organisational trend is displayed as a dashed reference line.',
        ],
    },
    {
        'id': 'BR-09',
        'title': 'Regression Analysis for Performance Predictors',
        'priority': 'Should Have',
        'hmw': 'HMW-12',
        'oq': 'OQ-09',
        'description': (
            'The system shall perform a multiple linear regression analysis to identify the '
            'significant predictors of employee performance rating. Independent variables shall '
            'include tenure, age, training score, absenteeism rate, and salary band (encoded). '
            'The results page shall display coefficient values, R² score, p-values per predictor, '
            'and a feature-importance bar chart.'
        ),
        'rationale': (
            'Understanding which factors predict high performance allows HR to design targeted '
            'training programmes, adjust hiring criteria, and allocate development budgets more '
            'effectively. Regression analysis transforms correlation data into actionable insights.'
        ),
        'acceptance': [
            'Regression is performed using scikit-learn or statsmodels library.',
            'Model coefficients, R², and p-values are displayed in a formatted table.',
            'A feature-importance bar chart ranks predictors by absolute coefficient value.',
            'Model assumptions (linearity, homoscedasticity) are briefly noted on the page.',
            'Results update when the year or department filter is changed.',
        ],
    },
    {
        'id': 'BR-10',
        'title': 'Interactive Global Filters',
        'priority': 'Must Have',
        'hmw': 'HMW-01, HMW-02',
        'oq': 'OQ-01, OQ-02',
        'description': (
            'The system shall provide a persistent sidebar with global interactive filters that '
            'apply across all dashboard pages. Filters shall include: Performance Year (single or '
            'multi-select), Department (multi-select with Select All option), Gender (multi-select), '
            'Role / Job Title (multi-select), and Rating Range (slider, 1–5).'
        ),
        'rationale': (
            'Without cross-cutting filters, users must navigate to individual charts and apply '
            'separate controls, creating a fragmented experience. A persistent sidebar with global '
            'filters dramatically improves usability and analytical efficiency.'
        ),
        'acceptance': [
            'Sidebar filters are present on every page of the application.',
            'Changing a filter value re-renders all charts on the current page within 2 seconds.',
            'A "Reset All Filters" button restores all filters to their default values.',
            'Filter state is preserved when navigating between pages in the same session.',
            'The filter sidebar is collapsible to maximise chart viewing area.',
        ],
    },
    {
        'id': 'BR-11',
        'title': 'Data Quality Report',
        'priority': 'Must Have',
        'hmw': 'HMW-03',
        'oq': 'OQ-01',
        'description': (
            'The system shall automatically generate and display a data quality report covering '
            'the loaded dataset. The report shall include: total row count, count and percentage of '
            'missing values per column, count of duplicate rows, out-of-range rating values, '
            'invalid date formats, and an overall data-quality score (% clean rows).'
        ),
        'rationale': (
            'Data quality is a prerequisite for analytical credibility. Without a formal quality '
            'assessment, analysts cannot trust that the insights they derive are based on clean, '
            'complete, and consistent data. The report also satisfies academic requirements for '
            'demonstrating data-engineering best practices.'
        ),
        'acceptance': [
            'Report runs automatically on application startup and is available on a dedicated page.',
            'Each column has a missing-value count and percentage displayed.',
            'Rows failing any quality check are flagged and listed in a "Bad Rows" table.',
            'An overall quality score (0–100%) is displayed prominently.',
            'Report can be downloaded as a CSV file.',
        ],
    },
    {
        'id': 'BR-12',
        'title': 'CSV and Excel Data Export',
        'priority': 'Could Have',
        'hmw': 'HMW-04',
        'oq': 'OQ-01',
        'description': (
            'The system shall provide export buttons on all major data-table views, enabling users '
            'to download the displayed (filtered) data as a CSV file or as a formatted Excel '
            '(.xlsx) workbook. The Excel export shall include column headers, frozen first row, '
            'and auto-fitted column widths.'
        ),
        'rationale': (
            'Many HR managers and department heads require offline access to data for further '
            'analysis in Excel or for inclusion in management reports. Export functionality '
            'significantly increases the practical utility of the dashboard beyond the browser.'
        ),
        'acceptance': [
            'Each data table page has a "Download CSV" and "Download Excel" button.',
            'Downloaded file reflects currently applied filters.',
            'File is named descriptively (e.g., dept_avg_ratings_2024.csv).',
            'Excel file opens without errors in Microsoft Excel 2016 and later.',
            'Export works for files with up to 300 rows.',
        ],
    },
    {
        'id': 'BR-13',
        'title': 'ANOVA / Kruskal-Wallis Test Across Roles',
        'priority': 'Should Have',
        'hmw': 'HMW-07',
        'oq': 'OQ-06',
        'description': (
            'The system shall perform one-way ANOVA or Kruskal-Wallis H-test (chosen based on '
            'normality, assessed via Shapiro-Wilk test) to determine whether performance ratings '
            'differ significantly across job roles. Results shall include the test statistic, '
            'p-value, degrees of freedom, and a post-hoc pairwise comparison table (Tukey HSD '
            'for ANOVA or Dunn\'s test for Kruskal-Wallis) if the omnibus test is significant.'
        ),
        'rationale': (
            'Role-based performance differences may indicate structural factors such as differing '
            'workload, training availability, or rating calibration practices by role. Statistical '
            'testing moves the conversation from anecdote to evidence.'
        ),
        'acceptance': [
            'Shapiro-Wilk test is applied to determine normality and select appropriate test.',
            'ANOVA or Kruskal-Wallis result is displayed with test name, statistic, and p-value.',
            'If p < 0.05, post-hoc pairwise comparisons are computed and displayed.',
            'Results are filterable by department.',
            'All statistical outputs include a plain-language interpretation.',
        ],
    },
    {
        'id': 'BR-14',
        'title': 'Flag Overdue Promotions',
        'priority': 'Must Have',
        'hmw': 'HMW-09, HMW-10',
        'oq': 'OQ-10',
        'description': (
            'The system shall flag employees who are overdue for promotion based on the rule: '
            'tenure > 3 years AND average rating ≥ 4.0 AND never_promoted = True. Flagged employees '
            'shall appear in a dedicated "Overdue Promotions" table on the Promotion Scorecard page, '
            'with their employee ID, name, department, role, tenure, average rating, and PRS. '
            'An alert banner shall appear on the HR Manager view if any overdue employees exist.'
        ),
        'rationale': (
            'Failing to promote high-performing, long-tenured employees is one of the leading '
            'causes of voluntary attrition. Automated flagging ensures that no eligible employee '
            'is overlooked due to administrative oversight or manager inaction.'
        ),
        'acceptance': [
            'Employees meeting all three criteria are listed in the Overdue Promotions table.',
            'The flag logic correctly applies AND conditions (all three must be true).',
            'The table shows tenure in years (rounded to one decimal) and avg rating.',
            'A count badge ("X employees overdue") is visible on the sidebar navigation.',
            'The list updates dynamically when filters are applied.',
        ],
    },
]

for br in brs:
    add_heading(doc, f'7.{br["id"].split("-")[1].lstrip("0") if br["id"].split("-")[1].lstrip("0") else "0"}. {br["id"]}: {br["title"]}', 2)
    
    # Details table
    details = [
        ['Requirement ID', br['id']],
        ['Priority (MoSCoW)', br['priority']],
        ['Source OQ', br['oq']],
        ['Source HMW', br['hmw']],
    ]
    t = doc.add_table(rows=len(details), cols=2)
    t.style = 'Table Grid'
    for ri, row_data in enumerate(details):
        cells = t.rows[ri].cells
        cells[0].text = row_data[0]
        cells[1].text = row_data[1]
        cells[0].paragraphs[0].runs[0].bold = True
        set_cell_bg(cells[0], 'DCE6F1')
        cells[0].width = Inches(1.8)
        cells[1].width = Inches(5.2)
    doc.add_paragraph()

    # Description
    p = doc.add_paragraph()
    r = p.add_run('Description: ')
    r.bold = True
    p.add_run(br['description'])

    # Rationale
    p2 = doc.add_paragraph()
    r2 = p2.add_run('Rationale: ')
    r2.bold = True
    p2.add_run(br['rationale'])

    # Acceptance criteria
    p3 = doc.add_paragraph()
    r3 = p3.add_run('Acceptance Criteria:')
    r3.bold = True
    for ac in br['acceptance']:
        doc.add_paragraph(ac, style='List Bullet')

    doc.add_paragraph()

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — BUSINESS RULES
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '8. Business Rules', 1)

add_body(doc,
    'Business rules are formal, non-negotiable constraints and calculation definitions that govern '
    'how data is interpreted, computed, and acted upon within the system. All analytical modules '
    'must implement these rules exactly as specified.'
)

rules = [
    ('BR-R01', 'Rating Scale', 'Rating Scale Constraint',
     'All performance ratings stored in the database must be integers in the inclusive range [1, 5]. '
     'No decimal ratings are permitted. Any imported row with a rating outside this range must be '
     'flagged in the Data Quality Report and excluded from analytical calculations.',
     'Rating ∈ {1, 2, 3, 4, 5}',
     'High'),
    ('BR-R02', 'Promotion Score Formula', 'Promotion Score Calculation',
     'The Promotion-Readiness Score (PRS) for each employee is calculated as:\n'
     'PRS = (0.5 × rating_normalised + 0.3 × tenure_normalised + 0.2 × training_score_normalised) × 100\n'
     'Where each component is Min-Max normalised to [0, 1] within the entire employee population '
     'before the weighted sum is computed. The resulting PRS is bounded to [0, 100].',
     'PRS = round((0.5×R_norm + 0.3×T_norm + 0.2×TR_norm) × 100, 2)',
     'High'),
    ('BR-R03', 'Attrition Risk Classification', 'Attrition Risk Scoring',
     'Each department is classified by attrition risk as follows:\n'
     '• High: dept_avg_rating < 3.0 AND dept_avg_absenteeism > 10%\n'
     '• Medium: dept_avg_rating < 3.5 (and not High)\n'
     '• Low: all other departments\n'
     'The classification is computed separately for each performance year.',
     'IF avg_rating < 3.0 AND absenteeism > 10% → High; ELIF avg_rating < 3.5 → Medium; ELSE → Low',
     'High'),
    ('BR-R04', 'Top / Bottom Performer Definition', 'Performance Ranking Threshold',
     'An employee is classified as a Top Performer if their average rating for the selected year '
     'falls within the top 10 percentile of their department\'s rating distribution. Similarly, '
     'an employee is a Bottom Performer if their rating falls within the bottom 10 percentile. '
     'Percentile calculation is strictly within-department, not across the entire organisation.',
     'Top: rating ≥ P90 within dept; Bottom: rating ≤ P10 within dept',
     'High'),
    ('BR-R05', 'Demographic Bias Detection Threshold', 'Statistical Bias Flag',
     'A statistically significant demographic bias in performance ratings is flagged when the '
     'p-value of the two-sample independent t-test (for gender comparisons) or one-way ANOVA '
     '(for age group comparisons) is less than 0.05. Flagged results are displayed with a '
     'red warning badge and an interpretive statement.',
     'Flag if p-value < 0.05 (α = 0.05 significance level)',
     'Medium'),
    ('BR-R06', 'Overdue Promotion Flag', 'Promotion Oversight Detection',
     'An employee is flagged as "Overdue for Promotion" if ALL of the following conditions are '
     'simultaneously true:\n'
     '1. tenure_years > 3 (more than 3 years since the join date)\n'
     '2. avg_rating ≥ 4.0 (average performance rating across all available years)\n'
     '3. never_promoted = True (the employee has not been promoted since joining)\n'
     'The flag is a boolean field computed at query time and refreshes with each data load.',
     'tenure > 3 AND avg_rating ≥ 4.0 AND never_promoted = True',
     'High'),
]

for rule in rules:
    add_heading(doc, f'{rule[0]}: {rule[1]}', 2)
    t = doc.add_table(rows=4, cols=2)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    pairs = [
        ('Rule Name', rule[2]),
        ('Rule Statement', rule[3]),
        ('Formula / Condition', rule[4]),
        ('Priority', rule[5]),
    ]
    for ri, (k, v) in enumerate(pairs):
        cells = t.rows[ri].cells
        cells[0].text = k
        cells[1].text = v
        cells[0].paragraphs[0].runs[0].bold = True
        set_cell_bg(cells[0], 'DCE6F1')
        cells[0].width = Inches(1.8)
        cells[1].width = Inches(5.2)
    doc.add_paragraph()

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — DATA REQUIREMENTS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '9. Data Requirements', 1)

add_body(doc,
    'The system relies on five interrelated data entities stored in a SQLite database. All data for '
    'this academic project is synthetically generated using Python libraries (Faker, NumPy, Pandas). '
    'The table below provides a comprehensive specification of each data element.'
)

data_rows = [
    # Employee Master
    ('Employee Master', 'emp_id', 'Unique employee identifier', 'Synthetic Generator', 'VARCHAR(10), PK', 'One-time (at data generation)'),
    ('Employee Master', 'emp_name', 'Full name of the employee', 'Faker library', 'VARCHAR(100)', 'One-time'),
    ('Employee Master', 'department', 'Department name (e.g., Sales, IT, HR)', 'Defined list (8 depts)', 'VARCHAR(50)', 'One-time'),
    ('Employee Master', 'role', 'Job title / role within department', 'Defined list per dept', 'VARCHAR(50)', 'One-time'),
    ('Employee Master', 'gender', 'Gender identity (Male, Female, Non-binary)', 'Random categorical', 'VARCHAR(20)', 'One-time'),
    ('Employee Master', 'age', 'Age in years at time of data generation', 'NumPy normal dist (μ=35, σ=8)', 'INTEGER', 'One-time'),
    ('Employee Master', 'education', 'Highest education qualification', 'Defined levels list', 'VARCHAR(50)', 'One-time'),
    ('Employee Master', 'join_date', 'Date the employee joined the organisation', 'Random date 2015–2023', 'DATE', 'One-time'),
    # Performance Records
    ('Performance Records', 'perf_id', 'Unique performance record identifier', 'Auto-increment', 'INTEGER, PK', 'Annual'),
    ('Performance Records', 'emp_id', 'Foreign key to Employee Master', 'Linked from master', 'VARCHAR(10), FK', 'Annual'),
    ('Performance Records', 'year', 'Performance appraisal year', 'Range: 2022–2024', 'INTEGER', 'Annual'),
    ('Performance Records', 'quarter', 'Performance quarter (Q1–Q4)', 'Uniform distribution', 'INTEGER', 'Quarterly'),
    ('Performance Records', 'rating', 'Performance rating (1–5 integer)', 'NumPy randint, biased normal', 'INTEGER [1–5]', 'Annual / Quarterly'),
    ('Performance Records', 'manager_id', 'ID of the rating manager', 'Linked from manager pool', 'VARCHAR(10), FK', 'Annual'),
    # Salary Data
    ('Salary Data', 'emp_id', 'Foreign key to Employee Master', 'Linked from master', 'VARCHAR(10), FK', 'One-time'),
    ('Salary Data', 'salary', 'Annual salary in Indian Rupees', 'Normal dist by salary band', 'NUMERIC(10,2)', 'One-time'),
    ('Salary Data', 'salary_band', 'Salary band label (L1–L5)', 'Defined levels', 'VARCHAR(10)', 'One-time'),
    ('Salary Data', 'last_increment_date', 'Date of last salary increment', 'Random date post join', 'DATE', 'One-time'),
    # Attrition Data
    ('Attrition Data', 'emp_id', 'Foreign key to Employee Master', 'Linked from master', 'VARCHAR(10), FK', 'One-time'),
    ('Attrition Data', 'attrition_flag', 'Whether the employee has left (0/1)', 'Random Bernoulli (p=0.15)', 'BOOLEAN', 'One-time'),
    ('Attrition Data', 'attrition_reason', 'Reason for leaving if applicable', 'Categorical list', 'VARCHAR(100)', 'One-time'),
    ('Attrition Data', 'absenteeism_pct', 'Percentage of working days absent', 'Beta distribution (α=2, β=18)', 'FLOAT [0–1]', 'Annual'),
    # Training Data
    ('Training Data', 'training_id', 'Unique training record identifier', 'Auto-increment', 'INTEGER, PK', 'Per-completion'),
    ('Training Data', 'emp_id', 'Foreign key to Employee Master', 'Linked from master', 'VARCHAR(10), FK', 'Per-completion'),
    ('Training Data', 'training_name', 'Name of training programme', 'Defined training catalogue', 'VARCHAR(100)', 'Per-completion'),
    ('Training Data', 'completion_date', 'Date training was completed', 'Random date', 'DATE', 'Per-completion'),
    ('Training Data', 'score', 'Training assessment score (0–100)', 'Normal dist (μ=72, σ=12)', 'FLOAT [0–100]', 'Per-completion'),
]

add_table(doc,
    headers=['Table', 'Column', 'Description', 'Source', 'Data Type', 'Refresh'],
    rows=data_rows,
    col_widths=[1.3, 1.3, 1.6, 1.4, 1.1, 0.8]
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10 — ASSUMPTIONS, CONSTRAINTS, DEPENDENCIES
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '10. Assumptions, Constraints, and Dependencies', 1)

add_heading(doc, '10.1 Assumptions', 2)
add_body(doc, 'The following assumptions are made in the design and development of this project:')
assumptions = [
    ('A-01', 'The synthetic dataset is a sufficiently representative proxy for a real mid-sized organisation\'s employee performance data for the purpose of this academic project.'),
    ('A-02', 'All employees in the dataset have at least one complete performance record per year. No employee has zero performance records.'),
    ('A-03', 'The five data tables (Employee Master, Performance Records, Salary Data, Attrition Data, Training Data) together provide all information required to satisfy all 14 business requirements.'),
    ('A-04', 'The end user has access to a modern web browser (Chrome, Firefox, Edge) and a stable internet connection for accessing the deployed Streamlit application.'),
    ('A-05', 'The Faculty Guide will evaluate the project based on the analytical depth, code quality, documentation completeness, and adherence to the stated business requirements.'),
    ('A-06', 'The SQLite database will not be used in a concurrent multi-user write scenario; it is read-only during dashboard operation.'),
]
add_table(doc,
    headers=['ID', 'Assumption'],
    rows=assumptions,
    col_widths=[0.7, 6.3]
)

add_heading(doc, '10.2 Constraints', 2)
add_body(doc, 'The following constraints limit the solution design:')
constraints = [
    ('C-01', 'The project must be completed within the academic semester timeline (approximately 12 weeks from kick-off to submission).'),
    ('C-02', 'The technology stack is constrained to Python 3.10+, Streamlit, Pandas, Plotly, SciPy, scikit-learn, and SQLite — no commercial BI tools or cloud databases.'),
    ('C-03', 'No real employee personally identifiable information (PII) may be used. All data must be synthetic or fully anonymised.'),
    ('C-04', 'The deployed application must run within the free tier of Streamlit Community Cloud (memory limit ~1 GB RAM).'),
    ('C-05', 'The dataset size is bounded to 150–300 employee records to ensure reasonable performance within the academic environment.'),
    ('C-06', 'The project is developed by a single student; no distributed version control branching strategy is required beyond basic main/dev branches.'),
]
add_table(doc,
    headers=['ID', 'Constraint'],
    rows=constraints,
    col_widths=[0.7, 6.3]
)

add_heading(doc, '10.3 Dependencies', 2)
add_body(doc, 'The following external dependencies must be in place for successful project delivery:')
dependencies = [
    ('D-01', 'Python 3.10 or later must be installed on the development machine.'),
    ('D-02', 'All required Python libraries (Streamlit, Pandas, Plotly, SciPy, scikit-learn, python-docx, openpyxl) must be installable via pip from PyPI.'),
    ('D-03', 'GitHub account and repository must be created and accessible for version control and Streamlit Community Cloud deployment.'),
    ('D-04', 'Streamlit Community Cloud account must be registered and linked to the GitHub repository.'),
    ('D-05', 'The Faculty Guide must provide timely review feedback within 3 business days of document submission to allow iterative improvements.'),
    ('D-06', 'The synthetic data generation script must be executed successfully before the dashboard can be run, as it populates the SQLite database.'),
]
add_table(doc,
    headers=['ID', 'Dependency'],
    rows=dependencies,
    col_widths=[0.7, 6.3]
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 11 — RISKS AND MITIGATION
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '11. Risks and Mitigation', 1)

add_body(doc,
    'The following risk register identifies potential threats to the successful delivery of this '
    'project, along with their assessed probability and impact, and the corresponding mitigation '
    'strategy. Risk levels are assessed as High (H), Medium (M), or Low (L).'
)

risks = [
    ('R-01', 'Data Quality Issues in Synthetic Dataset', 'M', 'H',
     'Implement an automated data-quality check script (BR-11) that runs at application startup and '
     'flags any rows that fail validation rules. Set target of ≥ 95% clean rows. Re-generate or '
     'repair synthetic data before submission.'),
    ('R-02', 'Scope Creep — Adding Unrequested Features', 'H', 'M',
     'Strictly adhere to the 14 Business Requirements defined in Section 7. Any additional feature '
     'ideas are logged in a backlog and explicitly marked as out-of-scope for this version. The BRD '
     'serves as the change-control baseline.'),
    ('R-03', 'Bias in Synthetic Data Generation', 'M', 'M',
     'Use seed-controlled random number generation to ensure reproducibility. Validate the distribution '
     'of ratings, genders, and departments against realistic benchmarks. Document generation parameters '
     'in the codebase.'),
    ('R-04', 'Academic Deadline Risk', 'M', 'H',
     'Develop a 12-week project plan with weekly milestones. Prioritise Must-Have requirements (BR-01, '
     '02, 04–08, 10–11, 14) for the core sprint and treat Should/Could-Have items as stretch goals. '
     'Buffer the final week for testing and documentation.'),
    ('R-05', 'Streamlit Cloud Deployment Failure', 'L', 'H',
     'Test deployment to Streamlit Community Cloud at least two weeks before the submission deadline. '
     'Maintain a local fallback (run the app locally with `streamlit run app.py`) and document '
     'instructions in the README.'),
    ('R-06', 'Statistical Test Misinterpretation', 'M', 'H',
     'Include plain-language interpretive statements alongside all statistical outputs on the dashboard. '
     'Cite the test name, assumptions, and p-value threshold (α = 0.05) explicitly. Have the Faculty '
     'Guide review the statistical analysis section before final submission.'),
    ('R-07', 'Performance Degradation with 300 Records', 'L', 'M',
     'Profile Streamlit app load times during development. Use Streamlit\'s @st.cache_data decorator '
     'for all heavy data-loading and computation functions. Target sub-3-second page load for all pages.'),
    ('R-08', 'Library Version Incompatibility', 'M', 'M',
     'Pin all library versions in requirements.txt. Test the full application against the pinned '
     'versions on a fresh virtual environment before deployment. Document the Python version and OS '
     'used for development.'),
]

add_table(doc,
    headers=['Risk ID', 'Risk Description', 'Prob.', 'Impact', 'Mitigation Strategy'],
    rows=risks,
    col_widths=[0.7, 1.8, 0.6, 0.7, 3.2]
)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 12 — GLOSSARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '12. Glossary', 1)

add_body(doc, 'The following terms are used throughout this document and the broader project deliverables:')

glossary = [
    ('Rating (Performance Rating)',
     'A numerical score (integer, 1–5) assigned by a manager to an employee at the end of an '
     'appraisal period, reflecting the employee\'s overall job performance relative to expectations. '
     '1 = Unsatisfactory, 2 = Below Expectations, 3 = Meets Expectations, 4 = Exceeds Expectations, '
     '5 = Outstanding.'),
    ('Attrition',
     'The voluntary or involuntary departure of employees from an organisation. High attrition rates '
     'indicate workforce instability and are costly due to recruitment, onboarding, and lost '
     'productivity expenses.'),
    ('Promotion-Readiness Score (PRS)',
     'A composite, formula-driven score (0–100) that quantifies how ready an individual employee is '
     'for a promotion, based on their performance rating, tenure, and training score. Defined in BR-R02.'),
    ('Percentile',
     'A statistical measure that indicates the value below which a given percentage of observations '
     'in a dataset falls. For example, the 90th percentile rating means 90% of employees in the same '
     'department have a rating at or below that value.'),
    ('KPI (Key Performance Indicator)',
     'A quantifiable metric used to evaluate the success of an objective. In this project, KPIs include '
     'average department rating, percentage of employees ranked, and percentage of clean data rows.'),
    ('ANOVA (Analysis of Variance)',
     'A parametric statistical test used to determine whether there are statistically significant '
     'differences between the means of three or more independent groups. In this project, one-way '
     'ANOVA is applied to compare performance ratings across age groups or job roles.'),
    ('Pearson Correlation Coefficient (r)',
     'A measure of the linear relationship between two continuous variables, ranging from -1 '
     '(perfect negative correlation) to +1 (perfect positive correlation). Values near 0 indicate '
     'no linear relationship. In this project, r measures the association between salary and rating.'),
    ('Department Average (Dept Avg)',
     'The arithmetic mean of performance ratings for all employees within a specific department for '
     'a given appraisal period. Used as the primary benchmarking metric in BR-01.'),
    ('Tenure',
     'The length of time (measured in years) that an employee has been continuously employed by the '
     'organisation, calculated as (current_date - join_date). Tenure is a key component of the '
     'Promotion-Readiness Score formula.'),
    ('Bias (Demographic Bias)',
     'A systematic, statistically significant difference in performance ratings across demographic '
     'groups (e.g., gender, age) that cannot be explained by legitimate performance differences. '
     'Detected using t-tests and ANOVA at α = 0.05 significance level.'),
    ('Benchmark',
     'A standard or reference point against which performance can be compared. In this project, '
     'the department average rating serves as an internal benchmark for individual employee performance.'),
    ('MoSCoW Prioritisation',
     'A requirements prioritisation framework using four categories: Must Have (critical for the '
     'current release), Should Have (important but not critical), Could Have (desirable if time '
     'permits), and Won\'t Have (explicitly excluded from this release).'),
    ('Kruskal-Wallis H-Test',
     'A non-parametric alternative to one-way ANOVA, used when the assumption of normality cannot '
     'be met. It tests whether samples originate from the same distribution. Applied in BR-13 when '
     'Shapiro-Wilk test indicates non-normality of rating distributions.'),
]

for i, (term, definition) in enumerate(glossary):
    p = doc.add_paragraph()
    r = p.add_run(f'{term}: ')
    r.bold = True
    p.add_run(definition)
    p.paragraph_format.space_after = Pt(6)

page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 13 — SIGN-OFF
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '13. Sign-off', 1)

add_body(doc,
    'The undersigned acknowledge that they have reviewed and approved this Business Requirements '
    'Document. By signing below, they confirm that the document accurately represents the agreed '
    'scope, requirements, and constraints for the Employee Performance Rating Analysis project.'
)

signoff_rows = [
    ('Student / Author', 'Aayush Mathur (Enrolment: 2500320140001)', '___________________', '___________'),
    ('Faculty Guide', 'To be filled by Faculty', '___________________', '___________'),
    ('Reviewer', 'To be filled by Reviewer', '___________________', '___________'),
    ('Project Supervisor', 'To be filled by Supervisor', '___________________', '___________'),
]
add_table(doc,
    headers=['Role', 'Name', 'Signature', 'Date'],
    rows=signoff_rows,
    col_widths=[1.5, 2.5, 1.8, 1.2]
)

doc.add_paragraph()
add_body(doc, '— End of Document —')
p = doc.paragraphs[-1]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ── Save ──────────────────────────────────────────────────────────────────────
os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
doc.save(OUTPUT_PATH)
print(f"[OK] Saved: {OUTPUT_PATH}")
print(f"[OK] File size: {os.path.getsize(OUTPUT_PATH):,} bytes")
