import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_PATH = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\08_Database_Design.xlsx"
RTM_PATH = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\04_Traceability_Matrix.xlsx"

wb = openpyxl.Workbook()
wb.remove(wb.active)  # Remove default blank sheet

# Colors
COLOR_NAVY       = "1E3A8A"   # Primary Header Deep Blue
COLOR_STEEL      = "4472C4"   # Secondary Header Steel Blue
COLOR_ROYAL      = "2563EB"   # Accent Blue
COLOR_ALT        = "F8FAFC"   # Alternating Row Light Slate
COLOR_BORDER     = "D1D5DB"   # Grid Border
COLOR_WHITE      = "FFFFFF"
COLOR_TEXT_MAIN  = "111827"
COLOR_CODE_BG    = "F1F5F9"

# Fonts
FONT_TITLE       = Font(name="Calibri", size=15, bold=True, color=COLOR_NAVY)
FONT_SUBTITLE    = Font(name="Calibri", size=11, italic=True, color="4B5563")
FONT_SECTION     = Font(name="Calibri", size=12, bold=True, color=COLOR_NAVY)
FONT_HEADER      = Font(name="Calibri", size=11, bold=True, color=COLOR_WHITE)
FONT_SUBHEADER   = Font(name="Calibri", size=10, bold=True, color=COLOR_WHITE)
FONT_CELL        = Font(name="Calibri", size=10, bold=False, color=COLOR_TEXT_MAIN)
FONT_CELL_BOLD   = Font(name="Calibri", size=10, bold=True, color=COLOR_TEXT_MAIN)
FONT_CODE        = Font(name="Consolas", size=9, bold=False, color="0F172A")
FONT_CODE_BLOCK  = Font(name="Consolas", size=9, bold=False, color="1E293B")

# Fills
FILL_NAVY        = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
FILL_STEEL       = PatternFill(start_color=COLOR_STEEL, end_color=COLOR_STEEL, fill_type="solid")
FILL_ROYAL       = PatternFill(start_color=COLOR_ROYAL, end_color=COLOR_ROYAL, fill_type="solid")
FILL_ALT         = PatternFill(start_color=COLOR_ALT, end_color=COLOR_ALT, fill_type="solid")
FILL_CODE_BG     = PatternFill(start_color=COLOR_CODE_BG, end_color=COLOR_CODE_BG, fill_type="solid")

# Alignments
ALIGN_LEFT       = Alignment(horizontal="left", vertical="center")
ALIGN_LEFT_WRAP  = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER     = Alignment(horizontal="center", vertical="center")
ALIGN_CENTER_WRAP= Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT      = Alignment(horizontal="right", vertical="center")
ALIGN_SQL_BLOCK  = Alignment(horizontal="left", vertical="top", wrap_text=True)

# Borders
def thin_border():
    s = Side(style="thin", color=COLOR_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def style_header_row(ws, row_idx, num_cols, fill=FILL_NAVY):
    ws.row_dimensions[row_idx].height = 28
    for c in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = fill
        cell.font = FONT_HEADER
        cell.alignment = ALIGN_CENTER_WRAP
        cell.border = thin_border()

def apply_table_styles(ws, start_row, num_cols, align_map=None, code_cols=None):
    for r in range(start_row, ws.max_row + 1):
        ws.row_dimensions[r].height = 22
        is_alt = (r % 2 == 1)
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border()
            if is_alt and not cell.fill.start_color.index:
                cell.fill = FILL_ALT
            
            # Font
            if code_cols and c in code_cols:
                cell.font = FONT_CODE
            else:
                cell.font = FONT_CELL
            
            # Alignment
            if align_map and c in align_map:
                cell.alignment = align_map[c]
            else:
                cell.alignment = ALIGN_LEFT_WRAP

# ==============================================================================
# 1. SHEET: Overview
# ==============================================================================
ws_ov = wb.create_sheet("Overview")
ws_ov.sheet_properties.tabColor = COLOR_NAVY
ws_ov.views.sheetView[0].showGridLines = True

# Title Block
ws_ov.merge_cells("A1:F1")
cell_title = ws_ov["A1"]
cell_title.value = "Employee Performance Rating Analysis — Database Architecture & Specification"
cell_title.font = FONT_TITLE
cell_title.alignment = Alignment(horizontal="left", vertical="center")
ws_ov.row_dimensions[1].height = 32

ws_ov.merge_cells("A2:F2")
cell_sub = ws_ov["A2"]
cell_sub.value = "Academic Capstone Project Specification | Student: Aayush Mathur (Enrolment No: 2500320140001, MCA-B, ABES Engineering College)"
cell_sub.font = FONT_SUBTITLE
cell_sub.alignment = Alignment(horizontal="left", vertical="center")
ws_ov.row_dimensions[2].height = 20

# Section 1: Project Metadata
ws_ov.cell(row=4, column=1, value="1. Project & Document Metadata").font = FONT_SECTION
meta_headers = ["Attribute", "Value", "Technical Specification Details"]
ws_ov.row_dimensions[5].height = 24
for c_idx, h in enumerate(meta_headers, start=1):
    cell = ws_ov.cell(row=5, column=c_idx, value=h)
    cell.fill = FILL_STEEL
    cell.font = FONT_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = thin_border()

meta_rows = [
    ("Project Title", "Employee Performance Rating Analysis", "Comprehensive data-driven evaluation of organizational appraisal trends"),
    ("Document Identifier", "DOC-08-DB-SPEC", "Definitive Relational Database Design, Schema & Data Dictionary Specification"),
    ("Student Author", "Aayush Mathur", "Enrolment No: 2500320140001, Section: MCA-B, Batch: 2025-2027"),
    ("Institution", "ABES Engineering College", "Department of Master of Computer Applications (MCA), Ghaziabad, UP"),
    ("Current Version", "Version 1.0 (Production Release)", "Fully approved database architecture aligned with Steps 1-7"),
    ("Publication Date", "2026-11-19", "Formal capstone delivery milestone"),
    ("Database Engine", "SQLite 3 (ACID-Compliant Serverless)", "Embedded relational storage at data/app.db (zero cloud DB latency)"),
    ("Connection Protocol", "Python 3.12 sqlite3 / URI mode=ro", "Thread-safe, read-only analytics execution with Streamlit @st.cache_data"),
    ("Data Normalization", "Third Normal Form (3NF) Standard", "Decomposed entity structure eliminating update anomalies & data redundancy"),
    ("Primary Analytical View", "v_employee_analytics", "Flattened 38-column joined view providing single-pane OQ query execution"),
]

for r_idx, (k, v, desc) in enumerate(meta_rows, start=6):
    ws_ov.row_dimensions[r_idx].height = 20
    ws_ov.cell(row=r_idx, column=1, value=k).font = FONT_CELL_BOLD
    ws_ov.cell(row=r_idx, column=2, value=v).font = FONT_CELL
    ws_ov.cell(row=r_idx, column=3, value=desc).font = FONT_CELL
    for c in range(1, 4):
        ws_ov.cell(row=r_idx, column=c).border = thin_border()
        if r_idx % 2 == 1:
            ws_ov.cell(row=r_idx, column=c).fill = FILL_ALT

# Section 2: Purpose & Architecture Principles
start_r = 17
ws_ov.cell(row=start_r, column=1, value="2. Database Purpose & Architectural Tenets").font = FONT_SECTION
ws_ov.row_dimensions[start_r].height = 24

tenets = [
    ("Zero Setup Dependency", "The SQLite database is embedded locally at data/app.db, allowing the application to run instantaneously on any developer workstation or cloud environment (Streamlit Community Cloud) without external DB provisioning."),
    ("Strict 3NF Integrity", "Entity relations are cleanly decomposed into 5 relational tables (departments, employees, performance_reviews, training_records, attrition_records) with enforced Foreign Key constraints and cascading rules."),
    ("Multi-Year Longitudinal Grain", "The performance_reviews table captures 3 years of historical appraisal cycles (2024-2026) per employee, enabling historical CAGR trend analysis and regression modeling."),
    ("Sub-Second Query Caching", "Data queries are decoupled from runtime UI rendering via Streamlit's @st.cache_data(ttl=3600), delivering instantaneous page transitions (< 5ms) after initial load."),
    ("High Data Hygiene Standards", "Comprehensive CHECK constraints enforce numeric bounds on ratings (1.0-5.0), ages (18-75), tenure (>=0), and categorical domain sets (gender, education, cycles)."),
]
ws_ov.row_dimensions[start_r+1].height = 24
for c_idx, h in enumerate(["Architecture Principle", "Implementation Strategy & Value"], start=1):
    cell = ws_ov.cell(row=start_r+1, column=c_idx, value=h)
    cell.fill = FILL_STEEL
    cell.font = FONT_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = thin_border()

for idx, (p, s) in enumerate(tenets, start=start_r+2):
    ws_ov.row_dimensions[idx].height = 28
    ws_ov.cell(row=idx, column=1, value=p).font = FONT_CELL_BOLD
    ws_ov.cell(row=idx, column=2, value=s).font = FONT_CELL
    for c in range(1, 3):
        ws_ov.cell(row=idx, column=c).border = thin_border()
        ws_ov.cell(row=idx, column=c).alignment = ALIGN_LEFT_WRAP
        if idx % 2 == 1:
            ws_ov.cell(row=idx, column=c).fill = FILL_ALT

# Section 3: Database Naming Conventions
start_r = 25
ws_ov.cell(row=start_r, column=1, value="3. Database Naming Conventions & Standardizations").font = FONT_SECTION
ws_ov.row_dimensions[start_r].height = 24

conv_headers = ["Object Class", "Standard Convention", "Pattern / Syntax", "Concrete Project Example", "Engineering Rationale"]
ws_ov.row_dimensions[start_r+1].height = 24
for c_idx, h in enumerate(conv_headers, start=1):
    cell = ws_ov.cell(row=start_r+1, column=c_idx, value=h)
    cell.fill = FILL_STEEL
    cell.font = FONT_HEADER
    cell.alignment = ALIGN_CENTER
    cell.border = thin_border()

conv_rows = [
    ("Table Entities", "Lowercase plural snake_case", "tbl_names", "departments, employees, performance_reviews", "Standard SQL convention representing relational sets of entities"),
    ("Primary Keys", "Singular entity name + '_id'", "entity_id", "dept_id, emp_id, review_id, training_id", "Instantly disambiguates key attributes during complex relational joins"),
    ("Foreign Keys", "Exact referenced PK column name", "fk_col_id", "employees.dept_id -> departments.dept_id", "Eliminates ambiguity in natural joining and automatic schema introspection"),
    ("Analytical Views", "Prefix 'v_' + descriptive name", "v_entity_role", "v_employee_analytics", "Explicitly separates materialized or virtual query layers from raw tables"),
    ("Performance Indexes", "Prefix 'idx_' + table + '_' + col", "idx_table_col", "idx_reviews_emp, idx_reviews_year", "Consistent index discovery and execution query planner clarity"),
    ("Boolean / Flags", "Descriptive suffix '_flag' or score", "is_attrited / flag", "attrition_flag (0 or 1), overdue_promotion_flag", "Clear binary semantics avoiding ambiguous null/true interpretations"),
    ("Date Columns", "Descriptive suffix '_date'", "event_date", "hire_date, completion_date, attrition_date", "Enforces ISO-8601 string storage ('YYYY-MM-DD') for deterministic sorting"),
]

for idx, (oc, sc, syn, ex, rat) in enumerate(conv_rows, start=start_r+2):
    ws_ov.row_dimensions[idx].height = 22
    ws_ov.cell(row=idx, column=1, value=oc).font = FONT_CELL_BOLD
    ws_ov.cell(row=idx, column=2, value=sc).font = FONT_CELL
    ws_ov.cell(row=idx, column=3, value=syn).font = FONT_CODE
    ws_ov.cell(row=idx, column=4, value=ex).font = FONT_CODE
    ws_ov.cell(row=idx, column=5, value=rat).font = FONT_CELL
    for c in range(1, 6):
        ws_ov.cell(row=idx, column=c).border = thin_border()
        if idx % 2 == 1:
            ws_ov.cell(row=idx, column=c).fill = FILL_ALT
        if c in (1, 2, 5):
            ws_ov.cell(row=idx, column=c).alignment = ALIGN_LEFT_WRAP
        else:
            ws_ov.cell(row=idx, column=c).alignment = ALIGN_CENTER_WRAP

ws_ov.column_dimensions["A"].width = 25
ws_ov.column_dimensions["B"].width = 38
ws_ov.column_dimensions["C"].width = 30
ws_ov.column_dimensions["D"].width = 40
ws_ov.column_dimensions["E"].width = 45
ws_ov.freeze_panes = "A4"

# ==============================================================================
# 2. SHEET: Tables_Summary
# ==============================================================================
ws_ts = wb.create_sheet("Tables_Summary")
ws_ts.sheet_properties.tabColor = COLOR_STEEL
ws_ts.views.sheetView[0].showGridLines = True

tbl_summary_headers = ["Table Name", "Grain", "Primary Key", "Foreign Keys", "Columns Count", "Planned Rows", "Purpose"]
style_header_row(ws_ts, 1, len(tbl_summary_headers), fill=FILL_NAVY)

tbl_summary_data = [
    (
        "departments",
        "One row per functional organizational department unit",
        "dept_id (INTEGER)",
        "None (Root organizational dimension)",
        5,
        "6 - 8 departments",
        "Master registry of business departments, official unit names, managerial cost center codes, department heads, and operational facility campus locations."
    ),
    (
        "employees",
        "One row per distinct individual employee master profile",
        "emp_id (TEXT)",
        "dept_id -> departments(dept_id), manager_id -> employees(emp_id)",
        13,
        "150 - 300 employees (Planned: 250)",
        "Central enterprise workforce master recording core demographic attributes (gender, age), compensation, job roles, academic education levels, hire dates, manager hierarchies, and tenure."
    ),
    (
        "performance_reviews",
        "One row per employee appraisal review cycle / evaluation year",
        "review_id (INTEGER)",
        "emp_id -> employees(emp_id)",
        11,
        "450 - 900 records (Planned: 750, 3 cycles/emp)",
        "Longitudinal appraisal evaluation repository storing multi-year performance ratings (1.00-5.00), goal completion percentages, completed projects, peer review feedback, and calculated promotion readiness indexes."
    ),
    (
        "training_records",
        "One row per individual professional training course completion per employee",
        "training_id (INTEGER)",
        "emp_id -> employees(emp_id)",
        8,
        "300 - 600 records (Planned: 500 courses)",
        "Continuous professional development ledger capturing technical, domain, and leadership training programs, completed workshop hours, evaluation test scores, and formal professional certifications."
    ),
    (
        "attrition_records",
        "One row per employee retention, sentiment, and turnover risk profile",
        "attrition_id (INTEGER)",
        "emp_id -> employees(emp_id) [Unique 1:1]",
        8,
        "150 - 300 records (Planned: 250 profiles)",
        "Workforce stability and retention intelligence ledger tracking binary attrition states, formal separation exit dates and motivations, predictive attrition risk scores, work-life balance, and job satisfaction ratings."
    ),
]

for r_idx, row in enumerate(tbl_summary_data, start=2):
    for c_idx, val in enumerate(row, start=1):
        ws_ts.cell(row=r_idx, column=c_idx, value=val)

apply_table_styles(
    ws_ts, start_row=2, num_cols=len(tbl_summary_headers),
    align_map={
        1: ALIGN_LEFT_WRAP,
        2: ALIGN_LEFT_WRAP,
        3: ALIGN_CENTER_WRAP,
        4: ALIGN_LEFT_WRAP,
        5: ALIGN_CENTER,
        6: ALIGN_CENTER_WRAP,
        7: ALIGN_LEFT_WRAP
    },
    code_cols=[1, 3, 4]
)

ws_ts.column_dimensions["A"].width = 22
ws_ts.column_dimensions["B"].width = 34
ws_ts.column_dimensions["C"].width = 20
ws_ts.column_dimensions["D"].width = 38
ws_ts.column_dimensions["E"].width = 16
ws_ts.column_dimensions["F"].width = 24
ws_ts.column_dimensions["G"].width = 50
ws_ts.freeze_panes = "A2"

# ==============================================================================
# 3. SHEET: Data_Dictionary
# ==============================================================================
ws_dd = wb.create_sheet("Data_Dictionary")
ws_dd.sheet_properties.tabColor = COLOR_NAVY
ws_dd.views.sheetView[0].showGridLines = True

dd_headers = [
    "Table Name", "Column Name", "Data Type", "Key", "Nullable",
    "Allowed Values / Range", "Description", "OQ / BR Supported"
]
style_header_row(ws_dd, 1, len(dd_headers), fill=FILL_NAVY)

# Comprehensive Data Dictionary Rows (45 attributes across all 5 tables)
dd_rows = [
    # Table 1: departments (5 attributes)
    ("departments", "dept_id", "INTEGER", "PK", "No", "1 .. 999 (AUTOINCREMENT)", "Surrogate primary key uniquely identifying each functional organizational department.", "BR-01, OQ-01"),
    ("departments", "dept_name", "TEXT", "Unique", "No", "'Engineering', 'Sales', 'Marketing', 'Finance', 'Human Resources', 'Operations'", "Standardized corporate department name used for organizational slicing and aggregation.", "BR-01, BR-03, BR-07, OQ-01, OQ-03, OQ-07"),
    ("departments", "cost_center", "TEXT", "None", "No", "Format: CC-[DEPT]-[0-9]{3} (e.g., 'CC-ENG-101')", "Internal accounting and financial cost center allocation code assigned to the unit.", "BR-11, OQ-01"),
    ("departments", "department_head", "TEXT", "None", "No", "Text string (e.g. 'Dr. Rajesh Sharma')", "Executive leader / Director officially responsible for the operational unit.", "BR-04, OQ-04"),
    ("departments", "location", "TEXT", "None", "No", "Physical facility (e.g. 'Building 1 - North Wing')", "Corporate facility campus or floor location identifier where the department resides.", "BR-10, OQ-06"),

    # Table 2: employees (13 attributes)
    ("employees", "emp_id", "TEXT", "PK", "No", "Format: EMP-[0-9]{4} (EMP-0001 to EMP-0250)", "Unique immutable enterprise identification number assigned to each employee.", "BR-02, BR-11, OQ-02"),
    ("employees", "first_name", "TEXT", "None", "No", "Alphabetic text (2-50 characters)", "Official legal first name of the employee.", "BR-02, BR-11, OQ-02"),
    ("employees", "last_name", "TEXT", "None", "No", "Alphabetic text (2-50 characters)", "Official legal surname / family name of the employee.", "BR-02, BR-11, OQ-02"),
    ("employees", "gender", "TEXT", "None", "No", "'Male', 'Female', 'Non-Binary', 'Other'", "Employee self-identified demographic gender category used for equity auditing.", "BR-05, BR-11, OQ-05"),
    ("employees", "age", "INTEGER", "None", "No", "18 to 75 (years)", "Chronological age of the employee in completed years as of reporting date.", "BR-05, BR-09, BR-11, OQ-05, OQ-09"),
    ("employees", "dept_id", "INTEGER", "FK", "No", "References departments(dept_id)", "Foreign key linking the employee to their primary organizational home department.", "BR-01, BR-03, BR-07, OQ-01, OQ-03, OQ-07"),
    ("employees", "job_role", "TEXT", "None", "No", "'Software Engineer', 'Data Scientist', 'Sales Exec', 'HR Specialist', 'Financial Analyst', 'Operations Lead'", "Specific functional role title and professional career stream classification.", "BR-06, BR-10, BR-13, OQ-06"),
    ("employees", "education_level", "TEXT", "None", "No", "'High School', 'Bachelor', 'Master', 'Doctorate'", "Highest formalized academic educational degree or credential obtained.", "BR-09, BR-11, OQ-09"),
    ("employees", "monthly_income", "REAL", "None", "No", ">= 15000.0 (Range: 18,000 to 280,000 INR/USD)", "Gross monthly base compensation in currency units, excluding bonuses.", "BR-03, BR-09, BR-11, OQ-03, OQ-09"),
    ("employees", "hire_date", "TEXT", "None", "No", "ISO-8601 Date: 'YYYY-MM-DD' (2012-01-01 to 2026-01-01)", "Official enterprise onboarding date when continuous employment commenced.", "BR-08, BR-11, OQ-08"),
    ("employees", "manager_id", "TEXT", "FK", "Yes", "References employees(emp_id) (or NULL for executive head)", "Self-referencing foreign key linking employee to their immediate line manager.", "BR-04, BR-11, OQ-04"),
    ("employees", "marital_status", "TEXT", "None", "Yes", "'Single', 'Married', 'Divorced'", "Demographic marital status category collected for benefits and diversity reporting.", "BR-05, BR-11, OQ-05"),
    ("employees", "tenure_years", "REAL", "None", "No", ">= 0.0 (Range: 0.2 to 14.5 years)", "Calculated accumulated continuous length of corporate service in fractional years.", "BR-07, BR-09, BR-11, BR-14, OQ-07, OQ-09, OQ-10"),

    # Table 3: performance_reviews (11 attributes)
    ("performance_reviews", "review_id", "INTEGER", "PK", "No", "1 .. 999999 (AUTOINCREMENT)", "Surrogate primary key uniquely identifying an annual performance appraisal event.", "BR-08, BR-11, OQ-08"),
    ("performance_reviews", "emp_id", "TEXT", "FK", "No", "References employees(emp_id)", "Foreign key linking the performance review record to the specific evaluated employee.", "BR-01, BR-02, BR-08, OQ-01, OQ-02, OQ-08"),
    ("performance_reviews", "review_year", "INTEGER", "None", "No", "2020 to 2030 (Specifically: 2024, 2025, 2026)", "Calendar appraisal year covering the twelve-month performance cycle.", "BR-08, BR-11, OQ-08"),
    ("performance_reviews", "review_cycle", "TEXT", "None", "No", "'Annual', 'Mid-Year', 'Q1', 'Q2', 'Q3', 'Q4'", "Formal corporate cadence frequency at which performance appraisal was administered.", "BR-08, BR-11, OQ-08"),
    ("performance_reviews", "performance_rating", "REAL", "None", "No", "1.00 to 5.00 (Continuous two-decimal scale)", "Final normalized composite performance evaluation rating score assigned to the employee.", "BR-01..05, BR-08..09, BR-13..14, OQ-01..10"),
    ("performance_reviews", "goals_completed_pct", "REAL", "None", "No", "0.0 to 100.0 (Percentage)", "Percentage of quarterly and annual strategic key performance objectives achieved.", "BR-06, BR-09, OQ-09, OQ-10"),
    ("performance_reviews", "projects_completed", "INTEGER", "None", "No", ">= 0 (Typical range: 1 to 18 projects)", "Total count of successfully delivered client deliverables or internal projects.", "BR-06, BR-09, OQ-09, OQ-10"),
    ("performance_reviews", "manager_feedback_score", "REAL", "None", "Yes", "1.00 to 5.00", "Quantitative leadership evaluation rating provided by the direct line supervisor.", "BR-04, BR-11, OQ-04"),
    ("performance_reviews", "peer_review_score", "REAL", "None", "Yes", "1.00 to 5.00", "Consolidated 360-degree peer feedback assessment score across teammates.", "BR-02, BR-11, OQ-02"),
    ("performance_reviews", "years_since_last_promotion", "REAL", "None", "No", ">= 0.0 (Range: 0.0 to 9.5 years)", "Elapsed continuous service tenure in years since the employee's last formal promotion.", "BR-06, BR-11, BR-14, OQ-10"),
    ("performance_reviews", "promotion_readiness_score", "REAL", "None", "No", "0.0 to 100.0 (Composite index)", "Algorithmic readiness index calculated from rating history, tenure velocity, and goal delivery.", "BR-06, BR-14, OQ-10"),

    # Table 4: training_records (8 attributes)
    ("training_records", "training_id", "INTEGER", "PK", "No", "1 .. 999999 (AUTOINCREMENT)", "Surrogate primary key uniquely identifying a completed training course participation.", "BR-09, BR-11, OQ-09"),
    ("training_records", "emp_id", "TEXT", "FK", "No", "References employees(emp_id)", "Foreign key mapping the professional training course to the enrolled employee.", "BR-09, BR-11, OQ-09"),
    ("training_records", "training_course_name", "TEXT", "None", "No", "Text (e.g., 'Cloud Infrastructure Architecture', 'Python for AI')", "Title of the professional curriculum or training program attended by the employee.", "BR-09, BR-11, OQ-09"),
    ("training_records", "training_category", "TEXT", "None", "No", "'Technical', 'Leadership', 'Compliance', 'Soft Skills', 'Domain Specialization'", "Curriculum subject classification grouping the nature of competency development.", "BR-09, BR-11, OQ-09"),
    ("training_records", "training_hours", "REAL", "None", "No", ">= 0.0 (Range: 4.0 to 120.0 hours)", "Total classroom, lab, or webinar instructional hours completed for the course.", "BR-09, BR-11, OQ-09"),
    ("training_records", "completion_date", "TEXT", "None", "No", "ISO-8601 Date: 'YYYY-MM-DD'", "Calendar date on which all examination and coursework requirements were finalized.", "BR-08, BR-09, OQ-08, OQ-09"),
    ("training_records", "training_score", "REAL", "None", "Yes", "0.0 to 100.0 (Percentage)", "Standardized post-course assessment exam score achieved by the employee.", "BR-09, BR-11, OQ-09"),
    ("training_records", "certification_earned", "INTEGER", "None", "No", "0 (No Credential) or 1 (Certified)", "Binary flag indicating whether a recognized industry certification was attained.", "BR-09, BR-11, OQ-09"),

    # Table 5: attrition_records (8 attributes)
    ("attrition_records", "attrition_id", "INTEGER", "PK", "No", "1 .. 999999 (AUTOINCREMENT)", "Surrogate primary key uniquely identifying an employee retention profile record.", "BR-07, BR-11, OQ-07"),
    ("attrition_records", "emp_id", "TEXT", "FK/Unique", "No", "References employees(emp_id) [Unique 1:1]", "Foreign key establishing a strict 1:1 relationship with the employee master record.", "BR-07, BR-11, OQ-07"),
    ("attrition_records", "attrition_flag", "INTEGER", "None", "No", "0 (Active Employee) or 1 (Departed Employee)", "Binary status indicator confirming whether employee is active or has separated.", "BR-07, BR-11, OQ-07"),
    ("attrition_records", "attrition_date", "TEXT", "None", "Yes", "ISO-8601 Date: 'YYYY-MM-DD' or NULL (Active)", "Formal date of corporate resignation, departure, or employment termination.", "BR-07, BR-08, OQ-07, OQ-08"),
    ("attrition_records", "attrition_reason", "TEXT", "None", "Yes", "'Better Compensation', 'Career Growth', 'Commute/Location', 'Work-Life Balance', 'Personal/Family', 'Active'", "Primary validated motivation stated during exit interview proceedings.", "BR-07, BR-11, OQ-07"),
    ("attrition_records", "attrition_risk_score", "REAL", "None", "No", "0.00 to 1.00 (Probability scale)", "Statistical / Machine Learning predicted likelihood of voluntary resignation.", "BR-07, BR-11, OQ-07"),
    ("attrition_records", "work_life_balance_score", "INTEGER", "None", "No", "1 (Poor) to 5 (Excellent)", "Employee annual engagement survey rating assessing work-life equilibrium.", "BR-07, BR-11, OQ-07"),
    ("attrition_records", "job_satisfaction_score", "INTEGER", "None", "No", "1 (Very Dissatisfied) to 5 (Highly Satisfied)", "Self-reported employee sentiment and operational job satisfaction rating.", "BR-07, BR-11, OQ-07"),
]

for r_idx, row in enumerate(dd_rows, start=2):
    for c_idx, val in enumerate(row, start=1):
        ws_dd.cell(row=r_idx, column=c_idx, value=val)

apply_table_styles(
    ws_dd, start_row=2, num_cols=len(dd_headers),
    align_map={
        1: ALIGN_LEFT_WRAP,
        2: ALIGN_LEFT_WRAP,
        3: ALIGN_CENTER,
        4: ALIGN_CENTER,
        5: ALIGN_CENTER,
        6: ALIGN_LEFT_WRAP,
        7: ALIGN_LEFT_WRAP,
        8: ALIGN_CENTER_WRAP
    },
    code_cols=[1, 2, 3, 4, 6]
)

ws_dd.column_dimensions["A"].width = 22
ws_dd.column_dimensions["B"].width = 25
ws_dd.column_dimensions["C"].width = 14
ws_dd.column_dimensions["D"].width = 12
ws_dd.column_dimensions["E"].width = 12
ws_dd.column_dimensions["F"].width = 32
ws_dd.column_dimensions["G"].width = 50
ws_dd.column_dimensions["H"].width = 25
ws_dd.freeze_panes = "A2"

# ==============================================================================
# 4. SHEET: Relationships
# ==============================================================================
ws_rel = wb.create_sheet("Relationships")
ws_rel.sheet_properties.tabColor = COLOR_STEEL
ws_rel.views.sheetView[0].showGridLines = True

rel_headers = [
    "Parent Table", "Child Table", "Key / Foreign Key",
    "Cardinality [1:M, M:1]", "On Update / Delete Rule", "Business Meaning"
]
style_header_row(ws_rel, 1, len(rel_headers), fill=FILL_NAVY)

rel_data = [
    (
        "departments",
        "employees",
        "departments.dept_id = employees.dept_id",
        "1:M (One department contains many employees)",
        "ON UPDATE CASCADE / ON DELETE RESTRICT",
        "Enforces organizational unit allocation. Every employee must belong to a valid department. Departments containing active personnel cannot be dropped, safeguarding organizational history."
    ),
    (
        "employees",
        "employees",
        "employees.emp_id = employees.manager_id",
        "1:M (One manager supervises many employees)",
        "ON UPDATE CASCADE / ON DELETE SET NULL",
        "Establishes internal operational reporting hierarchies. Line managers are themselves active employees in the system. If a manager departs, subordinate records safely default to NULL until reassigned."
    ),
    (
        "employees",
        "performance_reviews",
        "employees.emp_id = performance_reviews.emp_id",
        "1:M (One employee has multiple annual review records)",
        "ON UPDATE CASCADE / ON DELETE CASCADE",
        "Stores longitudinal appraisal histories across multiple consecutive calendar years (2024-2026). If an employee master record is purged, associated appraisal ratings cascade to maintain integrity."
    ),
    (
        "employees",
        "training_records",
        "employees.emp_id = training_records.emp_id",
        "1:M (One employee completes multiple training courses)",
        "ON UPDATE CASCADE / ON DELETE CASCADE",
        "Tracks employee professional development, skill certifications, and workshop hours over time. Deleting an employee cascades to clear their associated training records."
    ),
    (
        "employees",
        "attrition_records",
        "employees.emp_id = attrition_records.emp_id",
        "1:1 (One employee has exactly one retention profile)",
        "ON UPDATE CASCADE / ON DELETE CASCADE",
        "Maintains predictive flight risk indicators, employee satisfaction survey scores, and separation reasons. The unique foreign key constraint guarantees exactly one sentiment record per worker."
    ),
]

for r_idx, row in enumerate(rel_data, start=2):
    for c_idx, val in enumerate(row, start=1):
        ws_rel.cell(row=r_idx, column=c_idx, value=val)

apply_table_styles(
    ws_rel, start_row=2, num_cols=len(rel_headers),
    align_map={
        1: ALIGN_LEFT_WRAP,
        2: ALIGN_LEFT_WRAP,
        3: ALIGN_LEFT_WRAP,
        4: ALIGN_CENTER_WRAP,
        5: ALIGN_CENTER_WRAP,
        6: ALIGN_LEFT_WRAP
    },
    code_cols=[1, 2, 3, 5]
)

ws_rel.column_dimensions["A"].width = 20
ws_rel.column_dimensions["B"].width = 22
ws_rel.column_dimensions["C"].width = 38
ws_rel.column_dimensions["D"].width = 28
ws_rel.column_dimensions["E"].width = 30
ws_rel.column_dimensions["F"].width = 55
ws_rel.freeze_panes = "A2"

# ==============================================================================
# 5. SHEET: KPIs_Measures
# ==============================================================================
ws_kpi = wb.create_sheet("KPIs_Measures")
ws_kpi.sheet_properties.tabColor = COLOR_NAVY
ws_kpi.views.sheetView[0].showGridLines = True

kpi_headers = [
    "Measure / KPI ID", "Measure Name", "Formula / SQL Expression",
    "Source Tables", "Supported OQ", "Business Purpose"
]
style_header_row(ws_kpi, 1, len(kpi_headers), fill=FILL_NAVY)

kpi_data = [
    (
        "KPI-01",
        "Department Average Performance Rating",
        "SELECT d.dept_name, ROUND(AVG(pr.performance_rating), 2) AS avg_dept_rating FROM performance_reviews pr JOIN employees e ON pr.emp_id = e.emp_id JOIN departments d ON e.dept_id = d.dept_id GROUP BY d.dept_name;",
        "performance_reviews, employees, departments",
        "OQ-01, BR-01",
        "Calculates baseline talent performance per department to identify highest-performing business units and under-performing divisions requiring resource intervention."
    ),
    (
        "KPI-02",
        "Top Performer Headcount (Rating >= 4.5)",
        "SELECT COUNT(DISTINCT e.emp_id) AS top_performers_count FROM employees e JOIN performance_reviews pr ON e.emp_id = pr.emp_id WHERE pr.review_year = 2026 AND pr.performance_rating >= 4.5;",
        "performance_reviews, employees",
        "OQ-02, BR-02",
        "Isolates the upper decile / elite talent tier for succession planning, fast-track leadership programs, and merit bonus compensation allocation."
    ),
    (
        "KPI-03",
        "Under-Performer Headcount (Rating < 3.0)",
        "SELECT COUNT(DISTINCT e.emp_id) AS under_performers_count FROM employees e JOIN performance_reviews pr ON e.emp_id = pr.emp_id WHERE pr.review_year = 2026 AND pr.performance_rating < 3.0;",
        "performance_reviews, employees",
        "OQ-02, BR-02",
        "Identifies workers requiring structured Performance Improvement Plans (PIPs), upskilling re-assignments, or contractual review."
    ),
    (
        "KPI-04",
        "Salary-to-Rating Correlation Coefficient (r)",
        "Pearson r = [N*Σ(XY) - ΣX*ΣY] / SQRT([N*ΣX² - (ΣX)²] * [N*ΣY² - (ΣY)²]) where X = monthly_income and Y = performance_rating per department.",
        "employees, performance_reviews, departments",
        "OQ-03, BR-03",
        "Audits pay-for-performance meritocracy to verify whether higher appraisal ratings correspond to proportional financial compensation across departments."
    ),
    (
        "KPI-05",
        "Manager Supervised Team Average Rating",
        "SELECT e.manager_id, (m.first_name || ' ' || m.last_name) AS manager_name, ROUND(AVG(pr.performance_rating), 2) AS team_avg_rating, COUNT(e.emp_id) AS team_size FROM employees e JOIN performance_reviews pr ON e.emp_id = pr.emp_id JOIN employees m ON e.manager_id = m.emp_id GROUP BY e.manager_id HAVING team_size >= 3;",
        "employees (self-join), performance_reviews",
        "OQ-04, BR-04",
        "Ranks managerial supervisory performance, flagging leadership scoring anomalies such as excessive grading leniency, grading severity, or team delivery bottlenecks."
    ),
    (
        "KPI-06",
        "Gender & Demographic Rating Equity Disparity",
        "Rating Disparity = AVG(pr.performance_rating | Female) - AVG(pr.performance_rating | Male); Welchs t-test / Mann-Whitney U test p-value to evaluate statistical significance (alpha = 0.05).",
        "employees, performance_reviews",
        "OQ-05, BR-05",
        "Quantifies demographic fairness across gender and age tiers to guarantee unbiased performance appraisal distributions and fulfill corporate DEI governance."
    ),
    (
        "KPI-07",
        "Cross-Job Role Appraisal Variance (ANOVA F)",
        "F = Between-Group Mean Square / Within-Group Mean Square across job_role categories; evaluates whether rating distributions diverge significantly across specialized positions.",
        "employees, performance_reviews",
        "OQ-06, BR-13",
        "Validates standard deviation and appraisal consistency across diverse functional roles (e.g. Sales vs Software Engineers vs HR Specialists)."
    ),
    (
        "KPI-08",
        "Departmental Voluntary Attrition Rate (%)",
        "SELECT d.dept_name, ROUND((SUM(ar.attrition_flag) * 100.0) / COUNT(e.emp_id), 2) AS dept_attrition_pct FROM departments d JOIN employees e ON d.dept_id = e.dept_id JOIN attrition_records ar ON e.emp_id = ar.emp_id GROUP BY d.dept_name;",
        "attrition_records, employees, departments",
        "OQ-07, BR-07",
        "Tracks percentage turnover per operational department, establishing correlation with workload, work-life balance scores, and compensation bands."
    ),
    (
        "KPI-09",
        "3-Year Organizational Performance CAGR Trend",
        "Performance Growth % = [ (AVG(rating_2026) - AVG(rating_2024)) / AVG(rating_2024) ] * 100.0 across all continuous active corporate personnel.",
        "performance_reviews",
        "OQ-08, BR-08",
        "Measures longitudinal company-wide talent productivity trajectory across 3 continuous appraisal cycles to track executive organizational health."
    ),
    (
        "KPI-10",
        "Promotion-Readiness Index (Score 0-100)",
        "Readiness Score = (0.40 * (performance_rating / 5.0) * 100) + (0.30 * MIN(years_since_last_promotion / 4.0, 1.0) * 100) + (0.20 * goals_completed_pct) + (0.10 * MIN(total_training_hours / 40.0, 1.0) * 100);",
        "performance_reviews, training_records, employees",
        "OQ-10, BR-06",
        "Provides an objective, multi-criteria mathematical index scoring employee readiness for elevation to the next career band, eliminating promotional favoritism."
    ),
    (
        "KPI-11",
        "Overdue for Promotion Star Performer Headcount",
        "SELECT COUNT(DISTINCT e.emp_id) AS overdue_count FROM employees e JOIN performance_reviews pr ON e.emp_id = pr.emp_id WHERE pr.review_year = 2026 AND pr.performance_rating >= 4.0 AND pr.years_since_last_promotion > 3.0;",
        "employees, performance_reviews",
        "OQ-10, BR-14",
        "Flags high-performing personnel (rating >= 4.0) who have remained unpromoted for over 3 years, alerting leadership to imminent voluntary flight risk."
    ),
    (
        "KPI-12",
        "Multivariate Performance Predictor Model (OLS)",
        "performance_rating = β0 + β1*(age) + β2*(tenure_years) + β3*(monthly_income) + β4*(training_hours) + β5*(work_life_balance) + ε; reports R², F-statistic, and p-values.",
        "employees, performance_reviews, training_records, attrition_records",
        "OQ-09, BR-09",
        "Quantifies the predictive strength of employee tenure, professional upskilling, compensation, and satisfaction on annual appraisal outcomes."
    ),
]

for r_idx, row in enumerate(kpi_data, start=2):
    for c_idx, val in enumerate(row, start=1):
        ws_kpi.cell(row=r_idx, column=c_idx, value=val)

apply_table_styles(
    ws_kpi, start_row=2, num_cols=len(kpi_headers),
    align_map={
        1: ALIGN_CENTER,
        2: ALIGN_LEFT_WRAP,
        3: ALIGN_LEFT_WRAP,
        4: ALIGN_LEFT_WRAP,
        5: ALIGN_CENTER_WRAP,
        6: ALIGN_LEFT_WRAP
    },
    code_cols=[1, 3]
)

ws_kpi.column_dimensions["A"].width = 16
ws_kpi.column_dimensions["B"].width = 30
ws_kpi.column_dimensions["C"].width = 55
ws_kpi.column_dimensions["D"].width = 28
ws_kpi.column_dimensions["E"].width = 18
ws_kpi.column_dimensions["F"].width = 45
ws_kpi.freeze_panes = "A2"

# ==============================================================================
# 6. SHEET: SQLite_DDL
# ==============================================================================
ws_ddl = wb.create_sheet("SQLite_DDL")
ws_ddl.sheet_properties.tabColor = COLOR_STEEL
ws_ddl.views.sheetView[0].showGridLines = True

ddl_headers = [
    "Statement ID", "Object Name", "Object Type", "DDL SQL Statement", "Constraints & Technical Purpose"
]
style_header_row(ws_ddl, 1, len(ddl_headers), fill=FILL_NAVY)

ddl_data = [
    (
        "DDL-01",
        "PRAGMA Configuration",
        "System Setting",
        "PRAGMA foreign_keys = ON;\nPRAGMA journal_mode = WAL;\nPRAGMA synchronous = NORMAL;",
        "Enforces strict relational foreign key constraint validation and configures Write-Ahead Logging for high-concurrency read operations."
    ),
    (
        "DDL-02",
        "departments",
        "TABLE",
        "CREATE TABLE IF NOT EXISTS departments (\n"
        "    dept_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    dept_name TEXT NOT NULL UNIQUE,\n"
        "    cost_center TEXT NOT NULL,\n"
        "    department_head TEXT NOT NULL,\n"
        "    location TEXT NOT NULL\n"
        ");",
        "Primary entity table storing 6-8 corporate functional business units. UNIQUE constraint guarantees dept_name uniqueness."
    ),
    (
        "DDL-03",
        "employees",
        "TABLE",
        "CREATE TABLE IF NOT EXISTS employees (\n"
        "    emp_id TEXT PRIMARY KEY,\n"
        "    first_name TEXT NOT NULL,\n"
        "    last_name TEXT NOT NULL,\n"
        "    gender TEXT NOT NULL CHECK(gender IN ('Male', 'Female', 'Non-Binary', 'Other')),\n"
        "    age INTEGER NOT NULL CHECK(age >= 18 AND age <= 75),\n"
        "    dept_id INTEGER NOT NULL REFERENCES departments(dept_id) ON UPDATE CASCADE ON DELETE RESTRICT,\n"
        "    job_role TEXT NOT NULL,\n"
        "    education_level TEXT NOT NULL CHECK(education_level IN ('High School', 'Bachelor', 'Master', 'Doctorate')),\n"
        "    monthly_income REAL NOT NULL CHECK(monthly_income >= 15000.0),\n"
        "    hire_date TEXT NOT NULL,\n"
        "    manager_id TEXT REFERENCES employees(emp_id) ON UPDATE CASCADE ON DELETE SET NULL,\n"
        "    marital_status TEXT CHECK(marital_status IN ('Single', 'Married', 'Divorced')),\n"
        "    tenure_years REAL NOT NULL CHECK(tenure_years >= 0.0)\n"
        ");",
        "Central workforce master table with 13 columns. Enforces foreign key to departments and self-referencing hierarchy manager_id with CHECK constraints on age, income, and gender."
    ),
    (
        "DDL-04",
        "performance_reviews",
        "TABLE",
        "CREATE TABLE IF NOT EXISTS performance_reviews (\n"
        "    review_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    emp_id TEXT NOT NULL REFERENCES employees(emp_id) ON UPDATE CASCADE ON DELETE CASCADE,\n"
        "    review_year INTEGER NOT NULL CHECK(review_year BETWEEN 2020 AND 2030),\n"
        "    review_cycle TEXT NOT NULL CHECK(review_cycle IN ('Annual', 'Mid-Year', 'Q1', 'Q2', 'Q3', 'Q4')),\n"
        "    performance_rating REAL NOT NULL CHECK(performance_rating >= 1.0 AND performance_rating <= 5.0),\n"
        "    goals_completed_pct REAL NOT NULL CHECK(goals_completed_pct >= 0.0 AND goals_completed_pct <= 100.0),\n"
        "    projects_completed INTEGER NOT NULL CHECK(projects_completed >= 0),\n"
        "    manager_feedback_score REAL CHECK(manager_feedback_score >= 1.0 AND manager_feedback_score <= 5.0),\n"
        "    peer_review_score REAL CHECK(peer_review_score >= 1.0 AND peer_review_score <= 5.0),\n"
        "    years_since_last_promotion REAL NOT NULL CHECK(years_since_last_promotion >= 0.0),\n"
        "    promotion_readiness_score REAL NOT NULL CHECK(promotion_readiness_score >= 0.0 AND promotion_readiness_score <= 100.0)\n"
        ");",
        "Appraisal transaction ledger capturing multi-year rating cycles. Enforces continuous rating boundaries (1.0-5.0), project bounds, and cascade deletion upon employee purge."
    ),
    (
        "DDL-05",
        "training_records",
        "TABLE",
        "CREATE TABLE IF NOT EXISTS training_records (\n"
        "    training_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    emp_id TEXT NOT NULL REFERENCES employees(emp_id) ON UPDATE CASCADE ON DELETE CASCADE,\n"
        "    training_course_name TEXT NOT NULL,\n"
        "    training_category TEXT NOT NULL CHECK(training_category IN ('Technical', 'Leadership', 'Compliance', 'Soft Skills', 'Domain Specialization')),\n"
        "    training_hours REAL NOT NULL CHECK(training_hours >= 0.0),\n"
        "    completion_date TEXT NOT NULL,\n"
        "    training_score REAL CHECK(training_score >= 0.0 AND training_score <= 100.0),\n"
        "    certification_earned INTEGER NOT NULL CHECK(certification_earned IN (0, 1))\n"
        ");",
        "Professional learning repository recording instructional hours and certifications. Foreign key links to employees with ON DELETE CASCADE."
    ),
    (
        "DDL-06",
        "attrition_records",
        "TABLE",
        "CREATE TABLE IF NOT EXISTS attrition_records (\n"
        "    attrition_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    emp_id TEXT NOT NULL UNIQUE REFERENCES employees(emp_id) ON UPDATE CASCADE ON DELETE CASCADE,\n"
        "    attrition_flag INTEGER NOT NULL CHECK(attrition_flag IN (0, 1)),\n"
        "    attrition_date TEXT,\n"
        "    attrition_reason TEXT CHECK(attrition_reason IN ('Better Compensation', 'Career Growth', 'Commute/Location', 'Work-Life Balance', 'Personal/Family', 'Active')),\n"
        "    attrition_risk_score REAL NOT NULL CHECK(attrition_risk_score >= 0.0 AND attrition_risk_score <= 1.0),\n"
        "    work_life_balance_score INTEGER NOT NULL CHECK(work_life_balance_score BETWEEN 1 AND 5),\n"
        "    job_satisfaction_score INTEGER NOT NULL CHECK(job_satisfaction_score BETWEEN 1 AND 5)\n"
        ");",
        "Turnover and sentiment profile table. UNIQUE constraint on emp_id enforces 1:1 relation with employees table."
    ),
    (
        "DDL-07",
        "idx_employees_dept / manager",
        "INDEXES",
        "CREATE INDEX IF NOT EXISTS idx_employees_dept ON employees(dept_id);\n"
        "CREATE INDEX IF NOT EXISTS idx_employees_manager ON employees(manager_id);\n"
        "CREATE INDEX IF NOT EXISTS idx_reviews_emp ON performance_reviews(emp_id);\n"
        "CREATE INDEX IF NOT EXISTS idx_reviews_year ON performance_reviews(review_year);\n"
        "CREATE INDEX IF NOT EXISTS idx_training_emp ON training_records(emp_id);\n"
        "CREATE INDEX IF NOT EXISTS idx_attrition_emp ON attrition_records(emp_id);",
        "Performance B-Tree indexes created on all primary join foreign keys and temporal slice columns, ensuring O(log N) join efficiency."
    ),
    (
        "DDL-08",
        "v_employee_analytics",
        "VIEW",
        "CREATE VIEW IF NOT EXISTS v_employee_analytics AS\n"
        "SELECT \n"
        "    e.emp_id,\n"
        "    e.first_name,\n"
        "    e.last_name,\n"
        "    (e.first_name || ' ' || e.last_name) AS full_name,\n"
        "    e.gender,\n"
        "    e.age,\n"
        "    CASE \n"
        "        WHEN e.age < 30 THEN '<30'\n"
        "        WHEN e.age <= 45 THEN '30-45'\n"
        "        ELSE '>45'\n"
        "    END AS age_group,\n"
        "    e.dept_id,\n"
        "    d.dept_name,\n"
        "    d.cost_center,\n"
        "    d.location,\n"
        "    e.job_role,\n"
        "    e.education_level,\n"
        "    e.monthly_income,\n"
        "    CASE \n"
        "        WHEN e.monthly_income < 30000 THEN 'Low'\n"
        "        WHEN e.monthly_income <= 70000 THEN 'Mid'\n"
        "        ELSE 'High'\n"
        "    END AS salary_band,\n"
        "    e.hire_date,\n"
        "    e.tenure_years,\n"
        "    e.manager_id,\n"
        "    (m.first_name || ' ' || m.last_name) AS manager_name,\n"
        "    pr.review_id,\n"
        "    pr.review_year,\n"
        "    pr.review_cycle,\n"
        "    pr.performance_rating,\n"
        "    CASE \n"
        "        WHEN pr.performance_rating >= 4.5 THEN 'Top Performer'\n"
        "        WHEN pr.performance_rating < 3.0 THEN 'Low Performer'\n"
        "        ELSE 'Average Performer'\n"
        "    END AS rating_category,\n"
        "    pr.goals_completed_pct,\n"
        "    pr.projects_completed,\n"
        "    pr.years_since_last_promotion,\n"
        "    pr.promotion_readiness_score,\n"
        "    CASE \n"
        "        WHEN pr.performance_rating >= 4.0 AND pr.years_since_last_promotion > 3.0 THEN 1 \n"
        "        ELSE 0 \n"
        "    END AS overdue_promotion_flag,\n"
        "    COALESCE(t.total_training_hours, 0.0) AS total_training_hours,\n"
        "    COALESCE(t.courses_completed, 0) AS courses_completed,\n"
        "    COALESCE(t.avg_training_score, 0.0) AS avg_training_score,\n"
        "    COALESCE(ar.attrition_flag, 0) AS attrition_flag,\n"
        "    ar.attrition_date,\n"
        "    COALESCE(ar.attrition_reason, 'Active') AS attrition_reason,\n"
        "    COALESCE(ar.attrition_risk_score, 0.0) AS attrition_risk_score,\n"
        "    COALESCE(ar.work_life_balance_score, 3) AS work_life_balance_score,\n"
        "    COALESCE(ar.job_satisfaction_score, 3) AS job_satisfaction_score\n"
        "FROM employees e\n"
        "JOIN departments d ON e.dept_id = d.dept_id\n"
        "LEFT JOIN employees m ON e.manager_id = m.emp_id\n"
        "LEFT JOIN performance_reviews pr ON e.emp_id = pr.emp_id\n"
        "LEFT JOIN (\n"
        "    SELECT \n"
        "        emp_id,\n"
        "        ROUND(SUM(training_hours), 2) AS total_training_hours,\n"
        "        COUNT(training_id) AS courses_completed,\n"
        "        ROUND(AVG(training_score), 2) AS avg_training_score\n"
        "    FROM training_records\n"
        "    GROUP BY emp_id\n"
        ") t ON e.emp_id = t.emp_id\n"
        "LEFT JOIN attrition_records ar ON e.emp_id = ar.emp_id;",
        "Joined comprehensive analytical view containing exactly 38 columns uniting demographic, department, appraisal history, manager name, training sums, and retention indicators."
    ),
]

for r_idx, (sid, obj, otype, sql, notes) in enumerate(ddl_data, start=2):
    ws_ddl.cell(row=r_idx, column=1, value=sid)
    ws_ddl.cell(row=r_idx, column=2, value=obj)
    ws_ddl.cell(row=r_idx, column=3, value=otype)
    ws_ddl.cell(row=r_idx, column=4, value=sql)
    ws_ddl.cell(row=r_idx, column=5, value=notes)

for r in range(2, len(ddl_data) + 2):
    # Adjust row heights based on lines in SQL
    sql_text = ws_ddl.cell(row=r, column=4).value or ''
    line_count = len(sql_text.splitlines())
    ws_ddl.row_dimensions[r].height = max(35, min(line_count * 14 + 10, 400))
    is_alt = (r % 2 == 1)
    for c in range(1, 6):
        cell = ws_ddl.cell(row=r, column=c)
        cell.border = thin_border()
        if is_alt:
            cell.fill = FILL_ALT
        if c == 1:
            cell.font = FONT_CELL_BOLD
            cell.alignment = ALIGN_CENTER
        elif c in (2, 3):
            cell.font = FONT_CELL_BOLD
            cell.alignment = ALIGN_CENTER_WRAP
        elif c == 4:
            cell.font = FONT_CODE_BLOCK
            cell.alignment = ALIGN_SQL_BLOCK
            cell.fill = FILL_CODE_BG
        else:
            cell.font = FONT_CELL
            cell.alignment = ALIGN_LEFT_WRAP

ws_ddl.column_dimensions["A"].width = 16
ws_ddl.column_dimensions["B"].width = 24
ws_ddl.column_dimensions["C"].width = 16
ws_ddl.column_dimensions["D"].width = 75
ws_ddl.column_dimensions["E"].width = 45
ws_ddl.freeze_panes = "A2"

# Save 08_Database_Design.xlsx
wb.save(EXCEL_PATH)
print(f"Successfully generated: {EXCEL_PATH}")

# ==============================================================================
# UPDATE docs/04_Traceability_Matrix.xlsx
# ==============================================================================
wb_rtm = openpyxl.load_workbook(RTM_PATH)
ws_rtm = wb_rtm["RTM"]

# Mapping dictionary for the 25 rows
# Based on (OQ, BR ID, Requirement)
db_col_mappings = {
    # (OQ, BR_ID): DB Mapping
    ("OQ-01", "BR-01"): "departments.dept_name, performance_reviews.performance_rating, performance_reviews.review_year",
    ("OQ-02", "BR-02"): "employees.emp_id, employees.first_name, employees.last_name, performance_reviews.performance_rating",
    ("OQ-03", "BR-03"): "employees.monthly_income, performance_reviews.performance_rating, departments.dept_name",
    ("OQ-04", "BR-04"): "employees.manager_id, performance_reviews.performance_rating",
    ("OQ-05", "BR-05"): "employees.gender, employees.age, performance_reviews.performance_rating",
    ("OQ-10", "BR-06"): "performance_reviews.promotion_readiness_score, performance_reviews.performance_rating",
    ("OQ-10", "BR-14"): "performance_reviews.years_since_last_promotion, performance_reviews.performance_rating",
    ("OQ-07", "BR-07"): "attrition_records.attrition_flag, attrition_records.attrition_risk_score, departments.dept_name",
    ("OQ-08", "BR-08"): "performance_reviews.review_year, performance_reviews.performance_rating",
    ("OQ-09", "BR-09"): "employees.age, employees.tenure_years, employees.monthly_income, training_records.training_hours, performance_reviews.performance_rating",
    ("OQ-01", "BR-11"): "v_employee_analytics.dept_name, v_employee_analytics.performance_rating",
    ("OQ-02", "BR-11"): "v_employee_analytics.emp_id, v_employee_analytics.performance_rating",
    ("OQ-03", "BR-11"): "employees.monthly_income, v_employee_analytics.monthly_income",
    ("OQ-04", "BR-11"): "employees.manager_id, v_employee_analytics.manager_name",
    ("OQ-05", "BR-11"): "employees.gender, employees.age, v_employee_analytics.age_group",
    ("OQ-06", "BR-11"): "employees.job_role, departments.dept_name",
    ("OQ-07", "BR-11"): "attrition_records.attrition_flag, employees.tenure_years",
    ("OQ-08", "BR-11"): "performance_reviews.review_year, employees.hire_date",
    ("OQ-09", "BR-11"): "training_records.training_hours, employees.education_level",
    ("OQ-10", "BR-11"): "performance_reviews.years_since_last_promotion, employees.tenure_years",
    ("OQ-01", "BR-12"): "v_employee_analytics.*",
    ("OQ-08", "BR-12"): "performance_reviews.review_year, performance_reviews.performance_rating",
    ("OQ-06", "BR-10"): "departments.dept_name, employees.job_role, performance_reviews.review_year",
    ("OQ-06", "BR-13"): "employees.job_role, performance_reviews.performance_rating",
}

# In RTM, row 11 is secondary mapping for OQ-08 / HMW-12 / BR-08
updated_count = 0
for r in range(2, ws_rtm.max_row + 1):
    oq = ws_rtm.cell(r, 1).value
    hmw = ws_rtm.cell(r, 2).value
    br = ws_rtm.cell(r, 3).value
    req = ws_rtm.cell(r, 4).value
    
    key = (oq, br)
    mapping = None
    if key in db_col_mappings:
        mapping = db_col_mappings[key]
    
    # Specific row overrides if needed
    if r == 11 and oq == "OQ-08" and br == "BR-08":
        mapping = "performance_reviews.review_year, performance_reviews.performance_rating"

    if mapping:
        ws_rtm.cell(r, 9, value=mapping)
        updated_count += 1
        print(f"Row {r:2d} ({oq}, {br}): {mapping}")

# Column I width adjustment
ws_rtm.column_dimensions["I"].width = 42
wb_rtm.save(RTM_PATH)
print(f"Successfully updated {updated_count} rows in: {RTM_PATH}")
