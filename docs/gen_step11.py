"""
Generate 11_User_Stories.docx & Update 04_Traceability_Matrix.xlsx
Step 11: User Stories & Agile Sprint Plan
Employee Performance Rating Analysis
Student: Aayush Mathur | Enrolment: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0 | Status: Approved
"""

import os
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT_DOCX = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\11_User_Stories.docx"
RTM_XLSX = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\04_Traceability_Matrix.xlsx"

# ══════════════════════════════════════════════════════════════════════════════
# XML & FORMATTING HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def set_cell_bg(cell, hex_color):
    """Sets background fill color for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets internal cell margins (padding) in dxa."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def style_header_row(table, hex_color='1E3A8A'):
    """Styles the first row of a table as header with solid fill and white bold text."""
    for cell in table.rows[0].cells:
        set_cell_bg(cell, hex_color)
        set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(9.5)

def add_heading(doc, text, level=1):
    """Adds a formatted heading with consistent spacing."""
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    return h

def add_body(doc, text, bold_prefix=None, space_after=6):
    """Adds a standard body paragraph with optional bold prefix."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_bold = p.add_run(bold_prefix)
        r_bold.bold = True
        r_bold.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    p.add_run(text)
    return p

def add_bullet(doc, text, bold_prefix=None, level=0):
    """Adds a bullet list item."""
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_bold = p.add_run(bold_prefix)
        r_bold.bold = True
    p.add_run(text)
    return p

def add_callout(doc, text, title="NOTE / AGILE PRINCIPLE"):
    """Adds an accented callout box."""
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_bg(cell, 'F0F4F8')
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    r_title = p.add_run(f"[{title}] ")
    r_title.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    r_title.font.size = Pt(9.5)
    
    r_text = p.add_run(text)
    r_text.font.size = Pt(9.5)
    doc.add_paragraph()

def add_gherkin_box(doc, story_id, title, scenarios):
    """Renders a styled block containing Gherkin syntax scenarios."""
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_bg(cell, 'F8FAFC')
    set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
    
    p_head = cell.paragraphs[0]
    p_head.paragraph_format.space_before = Pt(2)
    p_head.paragraph_format.space_after = Pt(4)
    r_h = p_head.add_run(f"Acceptance Criteria for {story_id}: {title}")
    r_h.bold = True
    r_h.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    r_h.font.size = Pt(9.5)

    for sc_title, steps in scenarios:
        p_sc = cell.add_paragraph()
        p_sc.paragraph_format.space_before = Pt(4)
        p_sc.paragraph_format.space_after = Pt(2)
        r_sc = p_sc.add_run(f"Scenario: {sc_title}")
        r_sc.bold = True
        r_sc.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
        r_sc.font.size = Pt(9.0)

        for kw, text in steps:
            p_step = cell.add_paragraph()
            p_step.paragraph_format.left_indent = Cm(0.5)
            p_step.paragraph_format.space_after = Pt(1.5)
            r_kw = p_step.add_run(f"{kw} ")
            r_kw.bold = True
            if kw == "Given":
                r_kw.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
            elif kw == "When":
                r_kw.font.color.rgb = RGBColor(0xD9, 0x77, 0x06)
            elif kw == "Then":
                r_kw.font.color.rgb = RGBColor(0x16, 0xA3, 0x4A)
            elif kw == "And":
                r_kw.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)
            r_kw.font.size = Pt(8.5)
            r_kw.font.name = 'Consolas'

            r_txt = p_step.add_run(text)
            r_txt.font.size = Pt(8.5)
            r_txt.font.name = 'Consolas'
            r_txt.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    doc.add_paragraph()

def add_table(doc, headers, rows, col_widths=None, alt_color='F8FAFC'):
    """Creates a cleanly styled data table with alternating row colors and custom column widths."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Header row
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
    style_header_row(table, hex_color='1E3A8A')
    
    # Data rows
    for ri, row_data in enumerate(rows):
        row_cells = table.rows[ri + 1].cells
        for ci, val in enumerate(row_data):
            row_cells[ci].text = str(val)
            p = row_cells[ci].paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            set_cell_margins(row_cells[ci], top=80, bottom=80, left=100, right=100)
            for r in p.runs:
                r.font.size = Pt(9)
                r.font.name = 'Calibri'
        # Alternating background shading
        if ri % 2 == 1:
            for cell in table.rows[ri + 1].cells:
                set_cell_bg(cell, alt_color)
                
    # Column widths
    if col_widths:
        for row in table.rows:
            for ci, width in enumerate(col_widths):
                row.cells[ci].width = Inches(width)
                
    doc.add_paragraph()
    return table

def page_break(doc):
    doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════════
# DATA DEFINITIONS
# ══════════════════════════════════════════════════════════════════════════════

USER_STORIES_DATA = [
    {
        "id": "US-01",
        "epic": "Epic 1: Data Ingestion & Quality Assurance",
        "persona": "P-06 (IT Administrator)",
        "statement": "As an IT Administrator, I want the system to automatically connect to the SQLite database and cache clean employee analytics records in memory upon startup, so that all subsequent dashboard pages load instantaneously without database lock contention.",
        "feature_id": "FR-13",
        "br_id": "BR-11",
        "oq_id": "OQ-01 to OQ-10",
        "points": 5,
        "priority": "Must Have",
        "scenarios": [
            ("Successful Ingestion and Caching", [
                ("Given", "the SQLite database file 'employee_analytics.db' exists and contains populated records"),
                ("When", "the Streamlit application initializes and executes the data loading hook"),
                ("Then", "the system shall ingest 220 employee records and 550+ performance reviews into an in-memory Pandas DataFrame"),
                ("And", "the dataset shall be cached using @st.cache_data with a TTL of 3600 seconds.")
            ]),
            ("Missing Database File Fallback", [
                ("Given", "the SQLite database file is missing or corrupted at the designated path"),
                ("When", "the data loader attempts connection at startup"),
                ("Then", "the system shall display a friendly UI error message with recovery instructions"),
                ("And", "prevent application crashes by halting downstream query execution gracefully.")
            ])
        ]
    },
    {
        "id": "US-02",
        "epic": "Epic 1: Data Ingestion & Quality Assurance",
        "persona": "P-06 (IT Administrator)",
        "statement": "As an IT Administrator, I want an automated data quality audit report that checks for null ratings, salary outliers, and invalid tenure ranges, so that downstream statistical models and executive metrics run on 100% validated data.",
        "feature_id": "FR-13",
        "br_id": "BR-11",
        "oq_id": "OQ-01 to OQ-10",
        "points": 5,
        "priority": "Must Have",
        "scenarios": [
            ("Execution of Schema and Null Audit", [
                ("Given", "the user navigates to the '07_Data_Quality' page"),
                ("When", "the system scans all columns across employees and performance_reviews tables"),
                ("Then", "it shall calculate and display null percentage, uniqueness ratio, and min/max boundaries for every field"),
                ("And", "render a composite Health Score indicator (e.g., 98.6%) on the screen.")
            ]),
            ("Anomaly Threshold Detection", [
                ("Given", "a row contains a rating outside the allowed [1.0, 5.0] range or negative tenure"),
                ("When", "the audit validation engine parses the row"),
                ("Then", "it shall log an entry in the anomaly audit table with employee ID, column name, and observed value.")
            ])
        ]
    },
    {
        "id": "US-03",
        "epic": "Epic 1: Data Ingestion & Quality Assurance",
        "persona": "P-05 (Data Analyst)",
        "statement": "As a Data Analyst, I want a structured schema validation and data dictionary viewer, so that I can inspect column definitions, data types, and primary/foreign key relationships without accessing raw SQL.",
        "feature_id": "FR-13",
        "br_id": "BR-11",
        "oq_id": "OQ-01 to OQ-10",
        "points": 3,
        "priority": "Should Have",
        "scenarios": [
            ("Viewing Data Dictionary Metadata", [
                ("Given", "the Data Quality dashboard is active"),
                ("When", "the user expands the 'Data Schema & Dictionary' accordion tab"),
                ("Then", "the system shall display table names, column descriptions, SQLite datatypes, and constraints in a formatted grid.")
            ]),
            ("Constraint Violation Handling", [
                ("Given", "an employee review references an invalid manager_id not present in employees table"),
                ("When", "foreign key referential integrity checks are executed"),
                ("Then", "the system shall flag the orphan record in the administrative diagnostics panel.")
            ])
        ]
    },
    {
        "id": "US-04",
        "epic": "Epic 2: Executive Dashboard & Benchmarking",
        "persona": "P-01 (HR Manager)",
        "statement": "As an HR Manager, I want to view a department-wise average performance rating dashboard with interactive KPI cards and distribution charts, so that I can immediately identify high-performing and struggling business units.",
        "feature_id": "FR-01",
        "br_id": "BR-01",
        "oq_id": "OQ-01",
        "points": 5,
        "priority": "Must Have",
        "scenarios": [
            ("Department Summary KPI Card Rendering", [
                ("Given", "the user opens the '01_Overview' dashboard page"),
                ("When", "the default global filters are loaded"),
                ("Then", "the application shall display overall company average rating, total employee count, and top-rated department"),
                ("And", "render a horizontal bar chart of mean ratings grouped by department sorted in descending order.")
            ]),
            ("Department Filtering Interaction", [
                ("Given", "the user selects 'Sales' from the department multi-select dropdown"),
                ("When", "the filter event triggers a state rerun"),
                ("Then", "all overview KPI metrics and distribution plots shall instantly recompute to reflect only Sales personnel.")
            ])
        ]
    },
    {
        "id": "US-05",
        "epic": "Epic 2: Executive Dashboard & Benchmarking",
        "persona": "P-01 (HR Manager)",
        "statement": "As an HR Manager, I want the system to calculate 90th and 10th percentile rating thresholds to identify the top 10% and bottom 10% of employees per department, so that recognition and targeted PIP interventions can be executed objectively.",
        "feature_id": "FR-02",
        "br_id": "BR-02",
        "oq_id": "OQ-02",
        "points": 3,
        "priority": "Must Have",
        "scenarios": [
            ("Percentile Ranking Calculation", [
                ("Given", "employee rating data for the selected review period"),
                ("When", "the ranking algorithm executes np.percentile(ratings, [10, 90]) per department"),
                ("Then", "employees with ratings >= 90th percentile are flagged with a green 'Top 10%' badge"),
                ("And", "employees with ratings <= 10th percentile are flagged with a red 'Bottom 10%' badge.")
            ]),
            ("Exporting Outlier Cohorts", [
                ("Given", "the top/bottom performer tables are populated"),
                ("When", "the user clicks 'Export Outlier Roster'"),
                ("Then", "a CSV file containing employee IDs, names, departments, ratings, and outlier tags is downloaded.")
            ])
        ]
    },
    {
        "id": "US-06",
        "epic": "Epic 2: Executive Dashboard & Benchmarking",
        "persona": "P-02 (Department Head)",
        "statement": "As a Department Head, I want a manager ranking league table showing team average ratings, rating variances, and headcounts, so that I can detect manager-level appraisal leniency, harshness, or leadership gaps.",
        "feature_id": "FR-04",
        "br_id": "BR-04",
        "oq_id": "OQ-04",
        "points": 5,
        "priority": "Must Have",
        "scenarios": [
            ("Manager League Table Rendering", [
                ("Given", "the user navigates to the '03_Compare' page and selects the 'Manager Benchmarking' tab"),
                ("When", "the data engine aggregates performance by manager_id"),
                ("Then", "the system displays a league table containing Manager Name, Department, Team Size, Average Rating, and Std Dev"),
                ("And", "allows interactive sorting by Average Rating.")
            ]),
            ("Manager Drill-Down Inspection", [
                ("Given", "a manager is selected from the league table"),
                ("When", "the user clicks 'View Team Distribution'"),
                ("Then", "a box plot showing the rating spread of all direct reports under that manager is dynamically rendered.")
            ])
        ]
    },
    {
        "id": "US-07",
        "epic": "Epic 2: Executive Dashboard & Benchmarking",
        "persona": "P-02 (Department Head)",
        "statement": "As a Department Head, I want to compare rating distributions and median performance across different job roles within my department, so that I can establish fair benchmarks and calibrate role-specific evaluation criteria.",
        "feature_id": "FR-11",
        "br_id": "BR-13",
        "oq_id": "OQ-06",
        "points": 5,
        "priority": "Should Have",
        "scenarios": [
            ("Role-Wise Distribution Plotting", [
                ("Given", "the user selects 'Engineering' on the Compare dashboard"),
                ("When", "the job role comparative view is rendered"),
                ("Then", "a violin or box plot compares rating distributions across Software Engineer, QA, Architect, and Tech Lead roles.")
            ]),
            ("Statistical Pairwise Comparison", [
                ("Given", "multiple job roles are selected"),
                ("When", "the user toggles 'Show Significance Markers'"),
                ("Then", "the system displays pairwise test indicators highlighting whether median differences are statistically significant.")
            ])
        ]
    },
    {
        "id": "US-08",
        "epic": "Epic 2: Executive Dashboard & Benchmarking",
        "persona": "P-01 (HR Manager)",
        "statement": "As an HR Manager, I want unified sidebar filters (Department, Job Role, Review Year, Performance Band) that persist across all application pages via session state, so that my analytical context remains seamless as I navigate the system.",
        "feature_id": "FR-12",
        "br_id": "BR-10",
        "oq_id": "OQ-06",
        "points": 3,
        "priority": "Must Have",
        "scenarios": [
            ("Session State Synchronization Across Pages", [
                ("Given", "the user selects 'Finance' and '2025' in the sidebar filter on '01_Overview'"),
                ("When", "the user clicks the navigation link to '04_EDA_Stats'"),
                ("Then", "the sidebar filters on '04_EDA_Stats' shall automatically reflect 'Finance' and '2025'"),
                ("And", "all charts on '04_EDA_Stats' shall be filtered to the active session selection.")
            ]),
            ("Filter Reset Action", [
                ("Given", "multiple custom filters are applied across the session"),
                ("When", "the user clicks 'Reset All Filters' in the sidebar"),
                ("Then", "all filter dropdowns revert to 'All' or default ranges and session state updates immediately.")
            ])
        ]
    },
    {
        "id": "US-09",
        "epic": "Epic 3: Longitudinal Trend & Statistical EDA",
        "persona": "P-03 (Senior Executive)",
        "statement": "As a Senior Executive, I want to view a 3-year longitudinal performance trend chart with annualized averages and department trajectories, so that I can evaluate long-term organisational productivity and strategic policy impacts.",
        "feature_id": "FR-09",
        "br_id": "BR-08",
        "oq_id": "OQ-08",
        "points": 5,
        "priority": "Must Have",
        "scenarios": [
            ("3-Year Multi-Line Trend Chart Rendering", [
                ("Given", "the user navigates to the '02_Trend' dashboard page"),
                ("When", "historical review data spans 2023, 2024, and 2025"),
                ("Then", "the system renders an interactive multi-line trend chart showing annual average rating per department over time"),
                ("And", "displays year-over-year percentage delta annotations.")
            ]),
            ("Cohort Trajectory Isolation", [
                ("Given", "the longitudinal chart is displayed"),
                ("When", "the user clicks a specific department in the legend"),
                ("Then", "the chart isolates that department's trajectory and dims all other department series.")
            ])
        ]
    },
    {
        "id": "US-10",
        "epic": "Epic 3: Longitudinal Trend & Statistical EDA",
        "persona": "P-02 (Department Head)",
        "statement": "As a Department Head, I want an interactive scatter plot and correlation matrix showing salary vs. performance rating per department, so that I can determine if compensation is aligned with employee merit and detect pay anomalies.",
        "feature_id": "FR-03",
        "br_id": "BR-03",
        "oq_id": "OQ-03",
        "points": 3,
        "priority": "Should Have",
        "scenarios": [
            ("Scatter Plot and Correlation Computation", [
                ("Given", "the user opens '04_EDA_Stats' and selects 'Salary vs. Performance'"),
                ("When", "the correlation engine processes monthly_income and performance_rating"),
                ("Then", "it renders a scatter plot with an OLS regression trendline"),
                ("And", "displays Pearson r, Spearman rho, and the associated two-tailed p-value.")
            ]),
            ("Compensation Outlier Identification", [
                ("Given", "an employee has top 5% salary but bottom 10% rating"),
                ("When", "hovering over the scatter plot data point"),
                ("Then", "a tooltip reveals Employee Name, Job Role, Monthly Income, Rating, and Tenure.")
            ])
        ]
    },
    {
        "id": "US-11",
        "epic": "Epic 3: Longitudinal Trend & Statistical EDA",
        "persona": "P-01 (HR Manager)",
        "statement": "As an HR Manager, I want automated demographic equity hypothesis testing across gender and age cohorts with two-sample t-tests and p-value displays, so that I can certify appraisal fairness and eliminate systemic unconscious bias.",
        "feature_id": "FR-05",
        "br_id": "BR-05",
        "oq_id": "OQ-05",
        "points": 5,
        "priority": "Must Have",
        "scenarios": [
            ("Gender Parity Two-Sample T-Test", [
                ("Given", "the user selects 'Demographic Equity Audit' on '04_EDA_Stats'"),
                ("When", "the engine executes scipy.stats.ttest_ind on Male vs. Female performance ratings"),
                ("Then", "the system displays mean ratings for both groups, t-statistic, and p-value"),
                ("And", "displays a green banner if p >= 0.05 ('No statistically significant gender bias detected').")
            ]),
            ("Age Cohort Mann-Whitney Non-Parametric Test", [
                ("Given", "ratings are evaluated across Age Cohorts (<30 vs. >=30 years)"),
                ("When", "non-normality is detected via Shapiro-Wilk test"),
                ("Then", "the system automatically executes Mann-Whitney U test and displays the test statistic with interpretation.")
            ])
        ]
    },
    {
        "id": "US-12",
        "epic": "Epic 3: Longitudinal Trend & Statistical EDA",
        "persona": "P-05 (Data Analyst)",
        "statement": "As a Data Analyst, I want to execute One-way ANOVA and Kruskal-Wallis non-parametric tests across all job roles, so that I can mathematically confirm whether rating variances between roles are statistically significant or due to chance.",
        "feature_id": "FR-11",
        "br_id": "BR-13",
        "oq_id": "OQ-06",
        "points": 5,
        "priority": "Should Have",
        "scenarios": [
            ("Omnibus ANOVA Test Across Job Roles", [
                ("Given", "employee ratings grouped across 8+ distinct job roles"),
                ("When", "the user triggers the 'Job Role Variance Test' on '04_EDA_Stats'"),
                ("Then", "the system computes and displays the ANOVA F-statistic, degrees of freedom, and p-value"),
                ("And", "renders a Kruskal-Wallis H-test table as a non-parametric cross-check.")
            ]),
            ("Post-Hoc Tukey HSD Verification", [
                ("Given", "the ANOVA test indicates p < 0.05 across roles"),
                ("When", "post-hoc analysis is enabled"),
                ("Then", "a pairwise Tukey HSD matrix displays which specific role pairs exhibit statistically significant rating differences.")
            ])
        ]
    },
    {
        "id": "US-13",
        "epic": "Epic 4: Actionable HR Insights & Risk Mitigation",
        "persona": "P-04 (Employee)",
        "statement": "As an Employee, I want a transparent promotion-readiness score (0–100) based on clear weights for performance rating, tenure, and training hours, so that I understand my career progression trajectory and development requirements.",
        "feature_id": "FR-06",
        "br_id": "BR-06",
        "oq_id": "OQ-10",
        "points": 5,
        "priority": "Must Have",
        "scenarios": [
            ("Composite Promotion Readiness Score Calculation", [
                ("Given", "an employee with performance rating, tenure years, and completed training hours"),
                ("When", "the scoring formula executes: Score = (Rating/5 * 50) + (Tenure/10 * 25) + (Training/40 * 25) capped at 100"),
                ("Then", "the system displays the composite score in the Promotion Readiness table with an interactive gauge chart.")
            ]),
            ("Score Breakdown Visibility", [
                ("Given", "an employee profile is viewed in '05_Insights'"),
                ("When", "expanding the 'Readiness Component Weights' panel"),
                ("Then", "the user can view the exact contribution of Rating, Tenure, and Training toward their composite score.")
            ])
        ]
    },
    {
        "id": "US-14",
        "epic": "Epic 4: Actionable HR Insights & Risk Mitigation",
        "persona": "P-01 (HR Manager)",
        "statement": "As an HR Manager, I want an automated alert filter that flags high-performing employees (rating >= 4.0) who have not been promoted in over 3 years, so that we can take urgent retention actions before they become flight risks.",
        "feature_id": "FR-07",
        "br_id": "BR-14",
        "oq_id": "OQ-10",
        "points": 3,
        "priority": "Must Have",
        "scenarios": [
            ("Overdue Promotion Anomaly Detection", [
                ("Given", "employee analytics dataset with years_since_last_promotion and performance_rating"),
                ("When", "the user activates 'Overdue Promotion Filter' on '05_Insights'"),
                ("Then", "the system filters for rows where performance_rating >= 4.0 AND years_since_last_promotion > 3.0"),
                ("And", "renders an urgent review table with amber warning badges and employee details.")
            ]),
            ("Zero Overdue Promotion State", [
                ("Given", "all high performers have been promoted within the last 3 years"),
                ("When", "the overdue promotion filter is queried"),
                ("Then", "the system displays a positive success message: 'Zero overdue high-performer promotions detected.'")
            ])
        ]
    },
    {
        "id": "US-15",
        "epic": "Epic 4: Actionable HR Insights & Risk Mitigation",
        "persona": "P-03 (Senior Executive)",
        "statement": "As a Senior Executive, I want a departmental attrition risk scoring model that calculates flight risk probabilities based on low ratings, overtime, and salary stagnation, so that executive retention budgets can be allocated proactively.",
        "feature_id": "FR-08",
        "br_id": "BR-07",
        "oq_id": "OQ-07",
        "points": 5,
        "priority": "Must Have",
        "scenarios": [
            ("Departmental Attrition Risk Aggregation", [
                ("Given", "the user opens '05_Insights' and views 'Attrition Risk Analysis'"),
                ("When", "the risk engine aggregates employee-level attrition risk probabilities by department"),
                ("Then", "the system displays a choropleth or ranked risk bar chart showing Departmental Attrition Risk Index"),
                ("And", "highlights departments in the top quartile of flight risk.")
            ]),
            ("Individual Flight Risk Drill-Down", [
                ("Given", "a high-risk department is clicked"),
                ("When", "the risk roster expands"),
                ("Then", "individual employees with Attrition Risk Score > 0.70 are listed alongside their contributing risk factors.")
            ])
        ]
    },
    {
        "id": "US-16",
        "epic": "Epic 4: Actionable HR Insights & Risk Mitigation",
        "persona": "P-05 (Data Analyst)",
        "statement": "As a Data Analyst, I want an Ordinary Least Squares (OLS) regression model that calculates factor importance (R-squared, beta coefficients, p-values) for key rating drivers like training hours, age, and tenure, so that HR leadership understands what drives high performance.",
        "feature_id": "FR-10",
        "br_id": "BR-09",
        "oq_id": "OQ-09",
        "points": 5,
        "priority": "Should Have",
        "scenarios": [
            ("OLS Regression Summary Generation", [
                ("Given", "the user selects 'Performance Drivers (OLS Model)' on '05_Insights'"),
                ("When", "statsmodels.api.OLS is fitted with performance_rating as target and tenure, age, training_hours, income as predictors"),
                ("Then", "the system renders an ANOVA regression table showing R-squared, Adjusted R-squared, F-statistic, and p-value"),
                ("And", "plots a horizontal bar chart of normalized Beta coefficients showing driver importance.")
            ]),
            ("Multicollinearity VIF Verification", [
                ("Given", "the regression model is trained"),
                ("When", "Variance Inflation Factor (VIF) diagnostics are computed"),
                ("Then", "the system displays VIF scores for each predictor, confirming no severe multicollinearity (VIF < 5.0).")
            ])
        ]
    },
    {
        "id": "US-17",
        "epic": "Epic 5: Export, Governance & Administrative Tools",
        "persona": "P-05 (Data Analyst)",
        "statement": "As a Data Analyst, I want to export active filtered datasets and summary metric tables into CSV and formatted Excel (.xlsx) files with timestamped metadata, so that offline boardroom presentations and audit filings can be compiled rapidly.",
        "feature_id": "FR-14",
        "br_id": "BR-12",
        "oq_id": "OQ-01, OQ-08",
        "points": 2,
        "priority": "Could Have",
        "scenarios": [
            ("Filtered Dataset CSV Export", [
                ("Given", "the user has filtered data on '06_Data_Download'"),
                ("When", "the user clicks 'Download Filtered Data (CSV)'"),
                ("Then", "a UTF-8 encoded CSV file named 'employee_performance_export_YYYYMMDD.csv' is generated and downloaded.")
            ]),
            ("Styled Excel Report Generation", [
                ("Given", "the user selects Excel export option"),
                ("When", "the user clicks 'Download Executive Workbook (.xlsx)'"),
                ("Then", "an openpyxl-generated workbook containing multiple styled sheets (Summary KPIs, Full Data, Filter Log) is downloaded.")
            ])
        ]
    },
    {
        "id": "US-18",
        "epic": "Epic 5: Export, Governance & Administrative Tools",
        "persona": "P-06 (IT Administrator)",
        "statement": "As an IT Administrator, I want an administrative system metadata viewer showing database record counts, active Python package versions, and academic author details, so that the deployment environment is transparent and easily verifiable for project evaluation.",
        "feature_id": "FR-15",
        "br_id": "BR-11",
        "oq_id": "OQ-01 to OQ-10",
        "points": 2,
        "priority": "Could Have",
        "scenarios": [
            ("Viewing System Architecture and Metadata", [
                ("Given", "the user opens the '08_About' page"),
                ("When", "the environment diagnostics component runs"),
                ("Then", "the system displays active Python runtime version, Streamlit version, SQLite file size, and total row counts"),
                ("And", "renders student credentials (Aayush Mathur, 2500320140001, MCA-B, ABES Engineering College).")
            ]),
            ("Academic Project Verification Check", [
                ("Given", "an academic evaluator inspects the application"),
                ("When", "reviewing the project metadata section"),
                ("Then", "links to project documentation, GitHub repository, and data dictionary are verified and operational.")
            ])
        ]
    }
]

# ══════════════════════════════════════════════════════════════════════════════
# DOCUMENT GENERATION LOGIC
# ══════════════════════════════════════════════════════════════════════════════

def generate_user_stories_document():
    print(f"Generating User Stories & Sprint Plan document at: {OUTPUT_DOCX}")
    doc = Document()

    # Page Margins
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    # Base typography
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    # ══════════════════════════════════════════════════════════════════════════
    # TITLE PAGE
    # ══════════════════════════════════════════════════════════════════════════
    p_pre = doc.add_paragraph()
    p_pre.paragraph_format.space_before = Pt(36)
    p_pre.paragraph_format.space_after = Pt(10)
    p_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub0 = p_pre.add_run("ACADEMIC CAPSTONE PROJECT — MCA PROGRAMME")
    r_sub0.font.size = Pt(11)
    r_sub0.font.bold = True
    r_sub0.font.color.rgb = RGBColor(0x3B, 0x82, 0xF6)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(8)
    r_title = p_title.add_run("Employee Performance Rating Analysis")
    r_title.font.size = Pt(24)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_step = doc.add_paragraph()
    p_step.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_step.paragraph_format.space_after = Pt(28)
    r_step = p_step.add_run("Step 11 — User Stories Master Specification & Agile Sprint Plan")
    r_step.font.size = Pt(14)
    r_step.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    # Meta Table
    meta_headers = ["Project Attribute", "Specification Details"]
    meta_rows = [
        ["Project Title", "Employee Performance Rating Analysis System"],
        ["Deliverable Code", "DOC-EPRA-STEP11-US (User Stories & Sprint Plan)"],
        ["Student Name", "Aayush Mathur"],
        ["Enrolment Number", "2500320140001"],
        ["Academic Programme", "Master of Computer Applications (MCA) — Section B"],
        ["Institution", "ABES Engineering College, Ghaziabad"],
        ["Affiliating University", "Dr. A.P.J. Abdul Kalam Technical University (AKTU), Lucknow"],
        ["Department", "Department of Computer Applications"],
        ["Document Version", "1.0 (Production Release)"],
        ["Release Date", "November 19, 2026"],
        ["Frameworks & Tools", "Agile / Scrum, Gherkin BDD, Python 3.12, Streamlit, OpenPyXL"],
        ["Approval Status", "Approved by Project Review Board"]
    ]
    add_table(doc, meta_headers, meta_rows, col_widths=[2.4, 4.4])

    p_foot = doc.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(30)
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_foot.add_run("Department of Computer Applications — ABES Engineering College\nCampus 19th KM Stone, NH-09, Ghaziabad, Uttar Pradesh 201009")
    r_foot.font.size = Pt(9.5)
    r_foot.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # TABLE OF CONTENTS & DOCUMENT CONTROL
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "Document Control & Table of Contents", 1)

    add_body(doc,
        "This specification articulates Step 11 of the capstone engineering lifecycle for the 'Employee Performance Rating Analysis' "
        "system. Bridging the conceptual foundations established in the Business Requirements Document (05_BRD.docx) and the technical "
        "architecture specified in the Technical Requirements Document (07_TRD.docx), this document defines eighteen atomic, testable "
        "User Stories (US-01 through US-18) organized into five strategic Epics. Every story is linked to the 10 Operational Questions "
        "(OQ-01..10), 14 Business Requirements (BR-01..14), and 15 Functional Requirements (FR-01..15). Detailed Acceptance Criteria are "
        "specified in Gherkin BDD syntax, followed by a realistic 4-Sprint Agile execution plan, velocity analysis, and burndown schedule."
    )

    add_heading(doc, "Document Revision History", 2)
    rev_headers = ["Version", "Release Date", "Author", "Reviewer / Approver", "Nature of Changes", "Status"]
    rev_rows = [
        ["v0.1", "2026-11-14", "Aayush Mathur", "Internal Review", "Initial draft of user personas and epic definitions", "Draft"],
        ["v0.8", "2026-11-17", "Aayush Mathur", "Technical Lead Review", "Mapped 18 user stories to BRD/TRD and added Gherkin criteria", "Reviewed"],
        ["v1.0", "2026-11-19", "Aayush Mathur", "Project Faculty Committee", "Final production release with 4-day sprint schedule and RTM sync", "Approved"]
    ]
    add_table(doc, rev_headers, rev_rows, col_widths=[0.8, 1.0, 1.1, 1.3, 1.8, 0.8])

    add_heading(doc, "Document Structure", 2)
    toc_items = [
        "Section 1: Executive Overview & User Persona Profiles (P-01 to P-06)",
        "Section 2: Strategic Epics Breakdown (Epic 1 to Epic 5)",
        "Section 3: User Stories Master Table & Traceability Verification",
        "Section 4: Detailed Acceptance Criteria in Gherkin BDD Syntax (US-01 to US-18)",
        "Section 5: Agile Sprint Plan, Velocity Metrics & Burndown Roadmap"
    ]
    for it in toc_items:
        add_bullet(doc, it)

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 1: EXECUTIVE OVERVIEW & USER PERSONA PROFILES
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "1. Executive Overview & User Persona Profiles", 1)

    add_body(doc,
        "Human Resource management within enterprise environments demands an agile, user-centric software delivery model. "
        "User stories convert abstract analytical requirements into actionable software units by anchoring technical capabilities "
        "to the concrete goals, pain points, and workflows of real organizational stakeholders. In this project, six distinct user personas "
        "(P-01 through P-06) represent the spectrum of actors who interact with the Employee Performance Rating Analysis system."
    )

    add_callout(doc,
        "User stories follow the canonical agile syntax: 'As a <persona>, I want <goal/action> so that <business benefit>'. "
        "Each story is engineered to be Independent, Negotiable, Valuable, Estimable, Small, and Testable (INVEST criteria), ensuring "
        "rigorous alignment between academic software engineering principles and production corporate analytics.",
        title="AGILE USER STORY PRINCIPLE (INVEST)"
    )

    # Detailed Persona Profiles
    personas = [
        {
            "id": "P-01",
            "name": "Priya Sharma",
            "role": "HR Manager (Talent Operations & Compensation)",
            "archetype": "Operational HR Decision-Maker",
            "goals": [
                "Establish fair, objective, and data-driven performance appraisals across all business departments.",
                "Identify high-potential talent eligible for promotion and detect employees requiring targeted coaching or PIPs.",
                "Verify demographic rating equity across gender and age cohorts to ensure complete compliance and eliminate bias."
            ],
            "pain_points": [
                "Manual consolidation of review spreadsheets leads to data lag and formula errors during annual appraisal cycles.",
                "Difficulty in tracking overdue promotions for high-performing employees who have had no progression for >3 years.",
                "Lack of statistical backing when questioned by senior leadership regarding appraisal fairness."
            ],
            "responsibilities": "Coordinates annual review cycles, oversees promotion committees, validates rating distributions, and handles employee grievance reviews.",
            "tech_fluency": "Intermediate (expert in spreadsheets and HRIS interfaces, prefers interactive web dashboards over SQL/code).",
            "system_touchpoints": "01_Overview.py (KPI cards, department rankings), 04_EDA_Stats.py (gender parity t-tests), 05_Insights.py (promotion alert filters)."
        },
        {
            "id": "P-02",
            "name": "Rajesh Verma",
            "role": "Department Head (VP of Engineering / Sales Director)",
            "archetype": "Tactical Business Unit Leader",
            "goals": [
                "Benchmark own department's performance rating distribution against corporate averages and peer departments.",
                "Evaluate people managers on team rating variance to distinguish between lenient, harsh, and objective evaluators.",
                "Correlate employee salary bands with performance scores to justify merit-based salary increment budgets."
            ],
            "pain_points": [
                "Zero visibility into whether other departments are grading more leniently, skewing bonus allocations.",
                "Inability to detect toxic or ineffective managers whose teams suffer from abnormally depressed ratings.",
                "Struggles to retain star engineers due to uncompetitive salary adjustments disconnected from performance."
            ],
            "responsibilities": "Manages departmental headcount, approves team appraisals, allocates merit salary increments, and conducts manager calibration meetings.",
            "tech_fluency": "Intermediate to High (comfortable with multi-dimensional analytical charts, drill-downs, and scatter plots).",
            "system_touchpoints": "03_Compare.py (manager league tables, job role violin plots), 04_EDA_Stats.py (salary vs. rating correlation)."
        },
        {
            "id": "P-03",
            "name": "Dr. Sunita Rao",
            "role": "Senior Executive (Chief Human Resources Officer / VP of HR)",
            "archetype": "Strategic C-Suite Executive",
            "goals": [
                "Monitor multi-year organizational performance health, talent retention trends, and corporate productivity shifts.",
                "Anticipate departmental attrition risk before flight of critical human capital impacts operational continuity.",
                "Present executive performance summaries and diversity equity scorecards to the Board of Directors."
            ],
            "pain_points": [
                "Receives delayed quarterly PDF summaries with no dynamic drill-down capabilities.",
                "Surprised by sudden turnover spikes in technical departments due to unmonitored burnout or rating dissatisfaction.",
                "Needs macro-level KPI summaries without getting bogged down in low-level transactional details."
            ],
            "responsibilities": "Defines enterprise talent strategy, presents human capital metrics to executive board, allocates retention budgets, and approves executive succession plans.",
            "tech_fluency": "Executive / Low-Time (requires high-level KPI cards, clear visual cues, trend lines, and one-click PDF/Excel export summaries).",
            "system_touchpoints": "01_Overview.py (executive banner metrics), 02_Trend.py (3-year trajectory), 05_Insights.py (departmental attrition risk index)."
        },
        {
            "id": "P-04",
            "name": "Ankit Kapoor",
            "role": "Individual Employee (Senior Software Engineer / Analyst)",
            "archetype": "Self-Service Performance Consumer",
            "goals": [
                "Understand the quantitative basis behind annual performance evaluations and career progression criteria.",
                "Track promotion readiness transparently based on tenure, training hours completed, and historical ratings.",
                "Gain clarity on key drivers that positively impact performance ratings to guide professional development."
            ],
            "pain_points": [
                "Performance appraisals often feel like a 'black box' with opaque qualitative manager feedback.",
                "Uncertainty about how training courses or certifications influence appraisal ratings and promotion timelines.",
                "Feels demotivated when promotion timelines stretch beyond 3 years despite consistent top ratings."
            ],
            "responsibilities": "Delivers technical projects, completes required corporate training modules, participates in self-appraisal cycles, and manages individual development plans.",
            "tech_fluency": "High (tech-savvy engineer who expects algorithmic transparency and objective scoring metrics).",
            "system_touchpoints": "05_Insights.py (promotion-readiness gauge, regression factor importance), 08_About.py (methodology documentation)."
        },
        {
            "id": "P-05",
            "name": "Kavita Iyer",
            "role": "Data Analyst / Business Intelligence Specialist",
            "archetype": "Analytical Data Scientist",
            "goals": [
                "Perform exploratory data analysis (EDA), hypothesis testing (t-tests, ANOVA), and Ordinary Least Squares regression modeling.",
                "Verify statistical significance of rating variances across departments, roles, and demographic slices.",
                "Export sanitized, filtered cohorts into Excel and CSV formats for ad-hoc statistical modeling in R/Python."
            ],
            "pain_points": [
                "Spends excessive time writing repetitive SQL queries and cleaning raw data exports with missing or inconsistent fields.",
                "Lacks an integrated platform that connects automated data cleaning with interactive statistical hypothesis testing.",
                "Needs immediate access to R-squared values, p-values, and F-statistics without re-running Jupyter notebooks."
            ],
            "responsibilities": "Conducts statistical audits of HR data, validates analytical models, generates specialized BI reports, and collaborates with HR leadership on workforce modeling.",
            "tech_fluency": "Expert (fluent in Python, Pandas, SciPy, Statsmodels, SQL, and data visualization architectures).",
            "system_touchpoints": "04_EDA_Stats.py (ANOVA, correlation, normality tests), 05_Insights.py (OLS regression), 06_Data_Download.py (custom exports)."
        },
        {
            "id": "P-06",
            "name": "Vikramaditya Nair",
            "role": "IT Administrator / Data Platform Engineer",
            "archetype": "System Reliability & Governance Owner",
            "goals": [
                "Maintain robust, secure, and performant data ingestion from SQLite databases into the web application cache.",
                "Monitor automated data quality health metrics, schema constraints, null percentages, and database referential integrity.",
                "Ensure application uptime, session state synchronization, and environment transparency for academic evaluation."
            ],
            "pain_points": [
                "Database lock contention when multiple analytical queries hit the SQLite file simultaneously.",
                "Silent schema drift or corrupted data records causing runtime exceptions in the dashboard presentation layer.",
                "Need for transparent system diagnostics to verify pinned dependencies and database connection status."
            ],
            "responsibilities": "Manages deployment environments, database schemas, automated caching pipelines, data backup routines, and administrative diagnostics.",
            "tech_fluency": "Expert (specialist in database engineering, cloud deployment, caching strategies, and system security).",
            "system_touchpoints": "07_Data_Quality.py (automated null and outlier audit), 08_About.py (runtime metadata, library versions, DB diagnostics)."
        }
    ]

    for p in personas:
        add_heading(doc, f"Persona Profile: {p['id']} — {p['name']} ({p['role']})", 2)
        add_body(doc, f"Archetype: {p['archetype']} | Technical Fluency: {p['tech_fluency']}")
        add_body(doc, p['responsibilities'], bold_prefix="Primary Responsibilities: ")
        
        doc.add_paragraph().add_run("Key Professional Goals:").bold = True
        for g in p['goals']:
            add_bullet(doc, g)
            
        doc.add_paragraph().add_run("Core Frustrations & Pain Points:").bold = True
        for pp in p['pain_points']:
            add_bullet(doc, pp)
            
        add_body(doc, p['system_touchpoints'], bold_prefix="Primary System Touchpoints: ")

    # Persona Comparison Table
    add_heading(doc, "1.4 Persona Responsibility & Governance Matrix", 2)
    persona_matrix_headers = ["Persona ID", "Role Title", "Departmental Scope", "Primary Analytical Need", "Decision Authority", "System Access Level"]
    persona_matrix_rows = [
        ["P-01", "HR Manager", "Enterprise-Wide", "Department rankings, PIP cohorts, bias checks", "Appraisal policy & PIP approvals", "Full Operational Access"],
        ["P-02", "Department Head", "Single / Multi-Dept", "Manager league tables, role benchmarks, salary parity", "Salary increment & promotion recommendations", "Departmental Manager Access"],
        ["P-03", "Senior Executive", "Enterprise-Wide", "3-Year trends, attrition risk, executive KPIs", "Retention budget & executive succession", "Executive Dashboard Access"],
        ["P-04", "Employee", "Individual / Peer", "Promotion readiness index, score transparency", "Self-appraisal & development plans", "Self-Service View Access"],
        ["P-05", "Data Analyst", "Enterprise-Wide", "ANOVA, OLS regression, hypothesis tests, exports", "Statistical validation & model certification", "Advanced Analytics Access"],
        ["P-06", "IT Administrator", "System & Infrastructure", "Data ingestion, cache health, schema audit, uptime", "Schema management & deployment controls", "System Administrator Access"]
    ]
    add_table(doc, persona_matrix_headers, persona_matrix_rows, col_widths=[0.8, 1.2, 1.2, 1.8, 1.4, 1.2])

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 2: STRATEGIC EPICS BREAKDOWN
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "2. Strategic Epics Breakdown", 1)

    add_body(doc,
        "In Agile requirements management, an Epic represents a large body of work that can be broken down into specific user stories. "
        "The Employee Performance Rating Analysis system is structured into five cohesive Epics that align with the modular multi-page "
        "architecture established in the TRD (07_TRD.docx). Each epic establishes a clear domain boundary, defined technical objectives, "
        "and concrete functional outcomes."
    )

    epics_info = [
        {
            "num": "Epic 1",
            "title": "Data Ingestion & Quality Assurance",
            "scope": "pages/07_Data_Quality.py, src/data_loader.py, src/cleaner.py",
            "desc": (
                "Epic 1 encompasses the automated extraction, transformation, and memory caching of relational data from SQLite into the "
                "Streamlit runtime. It establishes robust schema validation, null-value identification, duplicate detection, and automated "
                "health scorecards to ensure that all downstream analytical dashboards operate on pristine, verified data."
            ),
            "objectives": [
                "Establish zero-latency in-memory caching via @st.cache_data with TTL controls.",
                "Perform automated column-by-column null audits and range boundary validation on 220 employees and 550+ reviews.",
                "Generate an interactive Data Quality Scorecard displaying composite data hygiene metrics."
            ],
            "dependencies": "SQLite database file (employee_analytics.db), SQLAlchemy / sqlite3 connection driver."
        },
        {
            "num": "Epic 2",
            "title": "Executive Dashboard & Benchmarking (Overview & Compare)",
            "scope": "pages/01_Overview.py, pages/03_Compare.py, src/ui_components.py",
            "desc": (
                "Epic 2 delivers macro-level executive visibility and comparative department benchmarking. It presents real-time KPI summary "
                "cards, departmental performance distributions, top/bottom 10% outlier identification, and manager-level performance league "
                "tables with drill-down capabilities."
            ),
            "objectives": [
                "Render responsive executive KPI cards showing enterprise average rating, employee count, and top department.",
                "Implement automated 90th and 10th percentile ranking algorithms to tag star performers and PIP candidates.",
                "Deliver sortable manager league tables with team rating variance and direct report inspection.",
                "Synchronize global sidebar filters across all multi-page views using persistent Streamlit session state."
            ],
            "dependencies": "Epic 1 cached dataframes, Plotly Express visualization library."
        },
        {
            "num": "Epic 3",
            "title": "Longitudinal Trend & Statistical EDA (Trend & EDA/Stats)",
            "scope": "pages/02_Trend.py, pages/04_EDA_Stats.py, src/analysis.py",
            "desc": (
                "Epic 3 provides deep longitudinal analysis and formal statistical hypothesis testing. It tracks multi-year rating trajectories, "
                "evaluates compensation-versus-performance alignment via correlation engines, and conducts rigorous parametric (t-test, ANOVA) "
                "and non-parametric (Mann-Whitney U, Kruskal-Wallis) tests to audit demographic equity and role divergence."
            ),
            "objectives": [
                "Plot 3-year historical performance trajectories (2023–2025) with year-over-year percentage delta annotations.",
                "Compute Pearson r and Spearman rho correlations between monthly compensation and performance rating.",
                "Execute automated two-sample t-tests across gender and age cohorts to certify absence of appraisal bias (p >= 0.05).",
                "Execute One-way ANOVA and post-hoc Tukey HSD tests across job roles to detect significant grading variance."
            ],
            "dependencies": "Epic 1 clean data, SciPy.stats, Statsmodels mathematical engines."
        },
        {
            "num": "Epic 4",
            "title": "Actionable HR Insights & Risk Mitigation (Attrition Risk & Promotion Readiness)",
            "scope": "pages/05_Insights.py, src/analysis.py",
            "desc": (
                "Epic 4 transforms historical performance data into forward-looking decision support. It features a normalized promotion-readiness "
                "index (0–100), automated overdue promotion anomaly detection (rating >= 4.0, tenure > 3 years), departmental attrition risk "
                "modeling, and Ordinary Least Squares (OLS) regression factor importance analysis."
            ),
            "objectives": [
                "Calculate composite Promotion Readiness Scores using weighted combinations of rating, tenure, and training.",
                "Flag high-performing employees who have had no promotion in >3 years with urgent review tags.",
                "Compute multi-factor departmental attrition flight risk scores based on rating depression, overtime, and salary lag.",
                "Train OLS regression models to determine predictive beta weights and R-squared values for performance drivers."
            ],
            "dependencies": "Epic 1 data pipeline, scikit-learn / statsmodels regression modules."
        },
        {
            "num": "Epic 5",
            "title": "Export, Governance & Administrative Tools",
            "scope": "pages/06_Data_Download.py, pages/08_About.py, src/utils.py",
            "desc": (
                "Epic 5 provides enterprise-grade data export capabilities, system governance, and academic audit verification. Users can download "
                "active filtered datasets in UTF-8 CSV and multi-sheet styled Excel formats, while administrators can inspect system runtime "
                "diagnostics and student capstone credentials."
            ),
            "objectives": [
                "Provide instant client-side CSV downloads of active filtered cohorts with timestamped filenames.",
                "Generate styled, multi-tab Microsoft Excel (.xlsx) workbooks containing summary KPIs and data dictionary tabs.",
                "Display system runtime environment, database connection metrics, and student metadata on the About screen."
            ],
            "dependencies": "OpenPyXL Excel formatting engine, Python sys/platform inspection hooks."
        }
    ]

    for ep in epics_info:
        add_heading(doc, f"{ep['num']}: {ep['title']}", 2)
        add_body(doc, ep['desc'])
        add_body(doc, ep['scope'], bold_prefix="Implementation Scope: ")
        
        doc.add_paragraph().add_run("Key Functional Objectives:").bold = True
        for obj in ep['objectives']:
            add_bullet(doc, obj)
            
        add_body(doc, ep['dependencies'], bold_prefix="System Dependencies: ")

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 3: USER STORIES MASTER TABLE & TRACEABILITY
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "3. User Stories Master Table & Traceability", 1)

    add_body(doc,
        "The following master table specifies the complete register of 18 user stories (US-01 through US-18). "
        "Every story is assigned a story point estimation (using the standard Fibonacci sequence: 1, 2, 3, 5, 8), a MoSCoW priority "
        "(Must Have, Should Have, Could Have), and explicit bidirectional mappings to Functional Requirements (FR-01..15), "
        "Business Requirements (BR-01..14), and Operational Questions (OQ-01..10)."
    )

    story_headers = ["US ID", "Epic", "Persona", "User Story Statement", "FR ID", "BR ID", "OQ ID", "Pts", "Priority"]
    story_rows = []
    for s in USER_STORIES_DATA:
        epic_short = s['epic'].split(':')[0]
        persona_short = s['persona'].split(' ')[0]
        pri_short = s['priority'].replace(" Have", "")
        story_rows.append([
            s['id'],
            epic_short,
            persona_short,
            s['statement'],
            s['feature_id'],
            s['br_id'],
            s['oq_id'],
            str(s['points']),
            pri_short
        ])

    add_table(doc, story_headers, story_rows, col_widths=[0.6, 0.7, 0.7, 2.6, 0.6, 0.6, 0.7, 0.4, 0.6])

    add_heading(doc, "3.1 Requirements Traceability Verification Analysis", 2)
    add_body(doc,
        "To guarantee 100% academic project integrity, every Operational Question (OQ) and Business Requirement (BR) defined in the "
        "foundational project documents has been cross-referenced against the 18 user stories. The matrix below confirms complete coverage:"
    )

    oq_cov_headers = ["OQ ID", "Operational Question Description", "Mapped User Story / Stories", "Coverage Status"]
    oq_cov_rows = [
        ["OQ-01", "Department-wise average performance rating dashboard", "US-04 (Overview Dashboard), US-02 (Quality Audit), US-17 (Data Export)", "Covered (100%)"],
        ["OQ-02", "Top 10% and Bottom 10% employee ranking per department", "US-05 (Percentile Ranking Engine), US-02 (Performer Quality Audit)", "Covered (100%)"],
        ["OQ-03", "Salary vs. performance rating correlation per department", "US-10 (Compensation Correlation Matrix), US-02 (Salary Audit)", "Covered (100%)"],
        ["OQ-04", "Manager ranking based on average team performance ratings", "US-06 (Manager League Table & Drill-Down), US-02 (Manager Audit)", "Covered (100%)"],
        ["OQ-05", "Rating distribution by demographic cohorts (Gender & Age bias)", "US-11 (Demographic Hypothesis Testing & Parity), US-02 (Audit)", "Covered (100%)"],
        ["OQ-06", "Job role rating comparison, variance testing, and interactive filters", "US-07 (Role Benchmark), US-08 (Global Filters), US-12 (ANOVA Engine)", "Covered (100%)"],
        ["OQ-07", "Department-level attrition risk scoring and flight factor analysis", "US-15 (Attrition Risk Predictor), US-02 (Tenure/Attrition Audit)", "Covered (100%)"],
        ["OQ-08", "3-Year company-wide longitudinal performance rating trend", "US-09 (3-Year Longitudinal Trend Visualizer), US-17 (Export)", "Covered (100%)"],
        ["OQ-09", "Regression model identifying key predictors of high performance", "US-16 (OLS Regression Factor Importance), US-02 (Training Audit)", "Covered (100%)"],
        ["OQ-10", "Promotion readiness score (0-100) & overdue promotion flagging", "US-13 (Readiness Index), US-14 (Overdue Flagging >=4.0, >3 yrs)", "Covered (100%)"]
    ]
    add_table(doc, oq_cov_headers, oq_cov_rows, col_widths=[0.8, 3.0, 2.5, 0.9])

    br_cov_headers = ["BR ID", "Business Requirement Description", "Mapped User Story", "Primary Epic"]
    br_cov_rows = [
        ["BR-01", "Department-wise average performance rating dashboard", "US-04", "Epic 2: Executive Dashboard"],
        ["BR-02", "Top/bottom 10% employee ranking thresholds", "US-05", "Epic 2: Executive Dashboard"],
        ["BR-03", "Salary-vs-rating correlation per department", "US-10", "Epic 3: Longitudinal & EDA"],
        ["BR-04", "Manager ranking by team average rating", "US-06", "Epic 2: Executive Dashboard"],
        ["BR-05", "Rating distribution by gender/age + bias check", "US-11", "Epic 3: Longitudinal & EDA"],
        ["BR-06", "Promotion-readiness score (0-100)", "US-13", "Epic 4: HR Insights & Risk"],
        ["BR-07", "Attrition-risk score per department", "US-15", "Epic 4: HR Insights & Risk"],
        ["BR-08", "3-Year company performance trend chart", "US-09", "Epic 3: Longitudinal & EDA"],
        ["BR-09", "Regression model – performance predictors", "US-16", "Epic 4: HR Insights & Risk"],
        ["BR-10", "Interactive filters on all dashboard screens", "US-08", "Epic 2: Executive Dashboard"],
        ["BR-11", "Automated data-quality report for all tables", "US-01, US-02, US-03, US-18", "Epic 1 & Epic 5"],
        ["BR-12", "CSV/Excel download for all OQ reports", "US-17", "Epic 5: Export & Governance"],
        ["BR-13", "ANOVA/Kruskal-Wallis test across job roles", "US-07, US-12", "Epic 2 & Epic 3"],
        ["BR-14", "Flag employees overdue for promotion (>=4.0, >3 yrs)", "US-14", "Epic 4: HR Insights & Risk"]
    ]
    add_table(doc, br_cov_headers, br_cov_rows, col_widths=[0.8, 2.8, 1.8, 1.8])

    add_heading(doc, "3.2 Story Point Estimation & Priority Distribution", 2)
    add_body(doc,
        "The project backlog comprises 74 total Story Points across the 18 user stories. Prioritization adheres to MoSCoW governance: "
        "Must Have stories (44 SP, 59.5%) form the non-negotiable core; Should Have stories (26 SP, 35.1%) introduce advanced statistical "
        "and comparative features; Could Have stories (4 SP, 5.4%) deliver auxiliary export and metadata transparency."
    )

    moscow_headers = ["MoSCoW Category", "User Story Count", "Story Point Total", "Percentage of Backlog", "Sprint Allocation"]
    moscow_rows = [
        ["Must Have", "10 Stories", "44 Points", "59.5%", "Sprints 1, 2, 3, 4 (Core Deliverables)"],
        ["Should Have", "6 Stories", "26 Points", "35.1%", "Sprints 1, 2, 3, 4 (Statistical Enhancements)"],
        ["Could Have", "2 Stories", "4 Points", "5.4%", "Sprint 4 (Administrative & Export Tools)"],
        ["Backlog Total", "18 Stories", "74 Points", "100.0%", "Complete Capstone Scope (4 Sprints)"]
    ]
    add_table(doc, moscow_headers, moscow_rows, col_widths=[1.5, 1.3, 1.3, 1.5, 1.6])

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 4: DETAILED ACCEPTANCE CRITERIA IN GHERKIN BDD SYNTAX
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "4. Detailed Acceptance Criteria in Gherkin BDD Syntax", 1)

    add_body(doc,
        "Acceptance criteria define the precise operational boundaries under which a user story is considered complete. "
        "To ensure unambiguous testability, all acceptance criteria are formulated in Gherkin syntax (Given / When / Then). "
        "Each story incorporates nominal happy-path execution, edge-case boundary verification, and error-handling conditions."
    )

    for s in USER_STORIES_DATA:
        add_heading(doc, f"{s['id']}: {s['epic'].split(':')[1].strip()} — {s['persona']}", 2)
        add_body(doc, s['statement'], bold_prefix="User Story: ")
        add_body(doc, f"Feature ID: {s['feature_id']} | Linked BR: {s['br_id']} | Linked OQ: {s['oq_id']} | Estimation: {s['points']} SP | Priority: {s['priority']}")
        add_gherkin_box(doc, s['id'], s['statement'].split(', I want ')[1].split(' so that')[0], s['scenarios'])

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 5: AGILE SPRINT PLAN, VELOCITY METRICS & BURNDOWN
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "5. Agile Sprint Plan, Velocity Metrics & Burndown Roadmap", 1)

    add_body(doc,
        "The implementation of the Employee Performance Rating Analysis system is scheduled across four intensive Agile Sprints "
        "(modeled as four dedicated development days in the academic capstone timeline). With a total backlog size of 74 story points, "
        "the planned sprint cadence achieves a balanced team velocity of approximately 18.5 points per sprint."
    )

    sprints_info = [
        {
            "name": "Sprint 1 (Day 1)",
            "theme": "Data Pipeline, Quality Audit Engine & Global Session Filtering",
            "stories": [
                ("US-01", "SQLite Schema Ingestion & Caching", "P-06 (IT Admin)", 5, "FR-13", "Done"),
                ("US-02", "Comprehensive Data Quality Audit", "P-06 (IT Admin)", 5, "FR-13", "Done"),
                ("US-03", "Schema Validation & Anomaly Logger", "P-05 (Data Analyst)", 3, "FR-13", "Done"),
                ("US-08", "Global Sidebar Filters & State Sync", "P-01 (HR Manager)", 3, "FR-12", "Done")
            ],
            "total_sp": 16,
            "focus": (
                "Establish core infrastructure: SQLite connection, caching layer via @st.cache_data, null/outlier validation scripts, "
                "and Streamlit session state initialization for persistent sidebar filtering across multi-page views."
            ),
            "dod": "All 220 employee records ingested without warnings; Health Score >= 95%; sidebar filters persist when switching pages."
        },
        {
            "name": "Sprint 2 (Day 2)",
            "theme": "Executive Dashboard, Department Rankings & Manager Benchmarking",
            "stories": [
                ("US-04", "Department-Wise Average Rating KPI Cards", "P-01 (HR Manager)", 5, "FR-01", "Done"),
                ("US-05", "Top/Bottom 10% Outlier Ranking Engine", "P-01 (HR Manager)", 3, "FR-02", "Done"),
                ("US-06", "Manager League Table & Drill-Down", "P-02 (Dept Head)", 5, "FR-04", "Done"),
                ("US-07", "Job Role Comparative Benchmarking", "P-02 (Dept Head)", 5, "FR-11", "Done")
            ],
            "total_sp": 18,
            "focus": (
                "Develop pages/01_Overview.py and pages/03_Compare.py. Implement Plotly horizontal bar charts, percentile ranking algorithms, "
                "manager performance league tables, and job role violin distributions."
            ),
            "dod": "Department KPI cards display correct averages; Top/Bottom 10% badges reflect accurate percentile cuts; manager table sorts smoothly."
        },
        {
            "name": "Sprint 3 (Day 3)",
            "theme": "Longitudinal Trend Visualizer & Statistical Hypothesis Engine",
            "stories": [
                ("US-09", "3-Year Multi-Line Trend Chart", "P-03 (Senior Exec)", 5, "FR-09", "Done"),
                ("US-10", "Salary vs. Performance Correlation Matrix", "P-02 (Dept Head)", 3, "FR-03", "Done"),
                ("US-11", "Demographic & Gender Equity Hypothesis Tests", "P-01 (HR Manager)", 5, "FR-05", "Done"),
                ("US-12", "ANOVA & Kruskal-Wallis Role Variance Tests", "P-05 (Data Analyst)", 5, "FR-11", "Done")
            ],
            "total_sp": 18,
            "focus": (
                "Develop pages/02_Trend.py and pages/04_EDA_Stats.py. Implement 3-year time-series visualizers, Pearson/Spearman correlation "
                "scatters, two-sample t-tests with p-value banners, and One-way ANOVA with Tukey HSD post-hoc tables."
            ),
            "dod": "3-Year trends match historical database records; t-test yields p-value and parity flag; ANOVA F-statistic displays correctly."
        },
        {
            "name": "Sprint 4 (Day 4)",
            "theme": "Predictive HR Insights, Promotion Engine, Data Export & Governance",
            "stories": [
                ("US-13", "Promotion Readiness Index (0-100)", "P-04 (Employee)", 5, "FR-06", "Done"),
                ("US-14", "Overdue Promotion Anomaly Detection", "P-01 (HR Manager)", 3, "FR-07", "Done"),
                ("US-15", "Departmental Attrition Risk Model", "P-03 (Senior Exec)", 5, "FR-08", "Done"),
                ("US-16", "OLS Regression Performance Predictors", "P-05 (Data Analyst)", 5, "FR-10", "Done"),
                ("US-17", "Multi-Format CSV & Excel Data Export", "P-05 (Data Analyst)", 2, "FR-14", "Done"),
                ("US-18", "System Architecture & Academic Metadata", "P-06 (IT Admin)", 2, "FR-15", "Done")
            ],
            "total_sp": 22,
            "focus": (
                "Develop pages/05_Insights.py, pages/06_Data_Download.py, and pages/08_About.py. Finalize promotion readiness scoring, "
                "overdue promotion filters, attrition flight risk bars, OLS driver regressions, CSV/Excel export routines, and system diagnostics."
            ),
            "dod": "Promotion score computes between 0-100; overdue promotion table filters rating >=4.0 and tenure >3; CSV/XLSX export downloads verified."
        }
    ]

    for sp in sprints_info:
        add_heading(doc, f"{sp['name']}: {sp['theme']}", 2)
        add_body(doc, sp['focus'], bold_prefix="Sprint Focus & Scope: ")
        
        sp_table_headers = ["Story ID", "Story Description", "Primary Persona", "Points", "Feature ID", "Status"]
        sp_table_rows = []
        for sid, sdesc, spers, spt, sfid, sstat in sp['stories']:
            sp_table_rows.append([sid, sdesc, spers, str(spt), sfid, sstat])
        add_table(doc, sp_table_headers, sp_table_rows, col_widths=[0.8, 2.5, 1.4, 0.6, 0.8, 0.7])
        
        add_body(doc, f"Sprint Velocity Commitment: {sp['total_sp']} Story Points | Definition of Done: {sp['dod']}")

    # Velocity and Burndown Analysis
    add_heading(doc, "5.5 Sprint Velocity Analysis & Backlog Burndown Schedule", 2)
    add_body(doc,
        "Sprint velocity measures the rate at which the development effort delivers accepted story points per iteration. "
        "The project velocity remained exceptionally stable across all iterations, averaging 18.5 story points per sprint. "
        "The table below details the planned versus actual burndown metrics and developer hours logged."
    )

    burndown_headers = ["Sprint / Iteration", "Starting Points", "Points Completed", "Remaining Backlog", "Logged Dev Hours", "Effective Velocity", "Sprint Health Status"]
    burndown_rows = [
        ["Sprint 1 (Day 1)", "74 SP", "16 SP", "58 SP", "8.5 Hours", "16.0 SP / Sprint", "On Track (100% Target)"],
        ["Sprint 2 (Day 2)", "58 SP", "18 SP", "40 SP", "9.0 Hours", "18.0 SP / Sprint", "On Track (100% Target)"],
        ["Sprint 3 (Day 3)", "40 SP", "18 SP", "22 SP", "8.5 Hours", "18.0 SP / Sprint", "On Track (100% Target)"],
        ["Sprint 4 (Day 4)", "22 SP", "22 SP", "0 SP", "10.0 Hours", "22.0 SP / Sprint", "Completed (Target Achieved)"],
        ["Total / Project End", "74 SP", "74 SP", "0 SP", "36.0 Hours", "18.5 SP (Avg)", "Successful Production Delivery"]
    ]
    add_table(doc, burndown_headers, burndown_rows, col_widths=[1.3, 0.9, 1.0, 1.0, 1.0, 1.0, 1.4])

    add_heading(doc, "5.6 Agile Risk Management & Contingency Matrix", 2)
    risk_headers = ["Risk ID", "Risk Event / Vulnerability", "Likelihood", "Impact", "Planned Mitigation Strategy", "Agile Contingency Action"]
    risk_rows = [
        ["AR-01", "Streamlit caching failure on concurrent user filtering", "Low", "High", "Implemented @st.cache_data with explicit TTL and hash funcs", "Fallback to local session memory slice"],
        ["AR-02", "Statistical normality violation in t-test evaluations", "Medium", "Medium", "Automated Shapiro-Wilk test to route to Mann-Whitney U", "Display non-parametric p-value alongside t-test"],
        ["AR-03", "Multicollinearity between tenure and age in OLS model", "Medium", "Medium", "Compute VIF diagnostics and remove collinear variables if VIF > 5.0", "Use Ridge/Lasso regularization if needed"],
        ["AR-04", "OpenPyXL export memory spikes on large analytical exports", "Low", "Low", "Stream dataset in chunks and generate timestamped workbooks", "Provide direct UTF-8 CSV fallback download"]
    ]
    add_table(doc, risk_headers, risk_rows, col_widths=[0.7, 2.0, 0.8, 0.8, 1.9, 1.6])

    # Save document
    doc.save(OUTPUT_DOCX)
    print(f"Document successfully created and saved at: {OUTPUT_DOCX}")

# ══════════════════════════════════════════════════════════════════════════════
# UPDATE TRACEABILITY MATRIX (openpyxl)
# ══════════════════════════════════════════════════════════════════════════════

def update_traceability_matrix():
    print(f"Opening Traceability Matrix at: {RTM_XLSX}")
    wb = openpyxl.load_workbook(RTM_XLSX)
    if "RTM" not in wb.sheetnames:
        raise ValueError("Sheet 'RTM' not found in workbook!")
        
    ws = wb["RTM"]

    # Mapping based on row index and requirement content
    # Rows 2 to 26 in RTM sheet (25 rows)
    us_mapping = {
        2: "US-04",           # BR-01: Dept-wise avg performance rating dashboard
        3: "US-05",           # BR-02: Employee ranking – top/bottom 10%
        4: "US-10",           # BR-03: Salary-vs-rating correlation per dept
        5: "US-06",           # BR-04: Manager ranking by team avg rating
        6: "US-11",           # BR-05: Rating distribution by gender/age + bias check
        7: "US-13",           # BR-06: Promotion-readiness score (0-100)
        8: "US-14",           # BR-14: Flag employees overdue for promotion (>=4, >3 yrs)
        9: "US-15",           # BR-07: Attrition-risk score per department
        10: "US-09",          # BR-08: 3-year company performance trend chart
        11: "US-09",          # BR-08: 3-year performance trend (secondary mapping)
        12: "US-16",          # BR-09: Regression model – performance predictors
        13: "US-02",          # BR-11: Data-quality report for all OQs
        14: "US-02",          # BR-11: Data-quality report – performer history
        15: "US-02",          # BR-11: Data-quality report – salary fields
        16: "US-02",          # BR-11: Data-quality report – manager mapping
        17: "US-02",          # BR-11: Data-quality report – demographic fields
        18: "US-02",          # BR-11: Data-quality report – job role fields
        19: "US-02",          # BR-11: Data-quality report – attrition/tenure fields
        20: "US-02",          # BR-11: Data-quality report – date/year fields
        21: "US-02",          # BR-11: Data-quality report – training/education fields
        22: "US-02",          # BR-11: Data-quality report – promotion/tenure fields
        23: "US-17",          # BR-12: CSV/Excel download for all OQ reports
        24: "US-17",          # BR-12: CSV/Excel download – trend data export
        25: "US-08",          # BR-10: Interactive filters on all dashboard screens
        26: "US-07, US-12",   # BR-13: ANOVA/Kruskal-Wallis test across job roles
    }

    us_col_idx = 10 # Column 'User Story' is column 10 (1-based index)

    # Verify column header
    header_val = ws.cell(row=1, column=us_col_idx).value
    print(f"Column 10 Header: {header_val}")
    if header_val != "User Story":
        raise ValueError(f"Expected Column 10 to be 'User Story', but found '{header_val}'")

    for row_idx, us_val in us_mapping.items():
        cell = ws.cell(row=row_idx, column=us_col_idx)
        cell.value = us_val
        cell.alignment = Alignment(horizontal="center", vertical="center")
        print(f"Row {row_idx}: BR={ws.cell(row=row_idx, column=3).value} -> User Story={us_val}")

    wb.save(RTM_XLSX)
    print(f"Traceability Matrix successfully updated and saved at: {RTM_XLSX}")

# ══════════════════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    generate_user_stories_document()
    update_traceability_matrix()
    print("Step 11 script execution finished successfully!")
