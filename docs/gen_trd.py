"""
Generate 07_TRD.docx — Technical Requirements Document (TRD)
Employee Performance Rating Analysis
Student: Aayush Mathur | Enrolment: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0 | Status: Approved
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT_PATH = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\07_TRD.docx"

# ── XML & Formatting Helpers ──────────────────────────────────────────────────

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
    """Adds a formatted heading."""
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

def add_code_block(doc, text):
    """Adds a formatted monospaced code or configuration block."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.right_indent = Cm(0.8)
    r = p.add_run(text)
    r.font.name = 'Consolas'
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
    return p

def add_callout(doc, text, title="NOTE / ARCHITECTURE PRINCIPLE"):
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
# MAIN BUILDER FUNCTION
# ══════════════════════════════════════════════════════════════════════════════

def generate_trd():
    print("Initializing document...")
    doc = Document()

    # ── Page Setup ──
    section = doc.sections[0]
    section.top_margin    = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin   = Cm(2.54)
    section.right_margin  = Cm(2.54)

    # ── Default Styles ──
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)

    # ══════════════════════════════════════════════════════════════════════════
    # TITLE PAGE
    # ══════════════════════════════════════════════════════════════════════════
    p_pre = doc.add_paragraph()
    p_pre.paragraph_format.space_before = Pt(40)
    p_pre.paragraph_format.space_after = Pt(10)
    p_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub0 = p_pre.add_run("ACADEMIC CAPSTONE PROJECT — MCA PROGRAMME")
    r_sub0.font.size = Pt(11)
    r_sub0.font.bold = True
    r_sub0.font.color.rgb = RGBColor(0x3B, 0x82, 0xF6)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after = Pt(12)
    r_title = title_p.add_run("TECHNICAL REQUIREMENTS DOCUMENT\n(TRD)")
    r_title.bold = True
    r_title.font.size = Pt(24)
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_p.paragraph_format.space_after = Pt(36)
    r_sub = subtitle_p.add_run("System Architecture, Data Pipelines, Statistical Specifications & Engineering Blueprint\nProject: Employee Performance Rating Analysis")
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    # Metadata Card Table
    meta_headers = ["Attribute", "Technical Specification"]
    meta_rows = [
        ["Project Title", "Employee Performance Rating Analysis"],
        ["Document ID", "TRD-EPRA-2026-007"],
        ["Document Version", "v1.0 (Production Release Specification)"],
        ["Document Status", "Approved"],
        ["Student Author", "Aayush Mathur (Enrolment No: 2500320140001)"],
        ["Academic Cohort", "MCA-B, Final Year (Class of 2026)"],
        ["Department / Institution", "Department of Computer Applications, ABES Engineering College, Ghaziabad"],
        ["Target Technology Stack", "Python 3.12, Streamlit 1.38, SQLite 3, pandas 2.2, SciPy/Statsmodels, Plotly 5.24, Tailwind CSS"],
        ["Deployment Target", "Streamlit Community Cloud (Linux container) via GitHub Repository Sync"],
        ["Release Date", "19 November 2026"],
    ]
    add_table(doc, meta_headers, meta_rows, col_widths=[2.0, 4.5])

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # TABLE OF CONTENTS
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "Table of Contents", 1)
    toc_items = [
        "1. Document Control & Executive Summary",
        "2. Architecture Overview & High-Level Pipeline (with Mermaid Diagram)",
        "3. Directory Tree & System Folder Structure",
        "4. Functional Requirements Traceability (FR-01 to FR-15 Table)",
        "5. Data Layer Engineering & SQLite Caching Strategy",
        "6. Statistical Methods & Operational Question Mapping (OQ-01 to OQ-10)",
        "7. User Interface Design, Styling & Session State Synchronisation",
        "8. Performance Optimization, Error Handling & System Security",
        "9. Pinned Dependency Specifications (requirements.txt)",
        "10. Deployment Requirements & Cloud Infrastructure",
        "11. Technical Risks, Mitigations & Contingency Matrix",
    ]
    for item in toc_items:
        p_toc = doc.add_paragraph()
        p_toc.paragraph_format.space_after = Pt(4)
        r_toc = p_toc.add_run(item)
        r_toc.font.size = Pt(10.5)
        r_toc.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 1: DOCUMENT CONTROL
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "1. Document Control & Executive Summary", 1)
    
    add_body(doc,
        "This Technical Requirements Document (TRD) establishes the definitive architectural, algorithmic, data engineering, "
        "and presentation layer standards for the 'Employee Performance Rating Analysis' capstone project. Authored for student "
        "Aayush Mathur (Enrolment No: 2500320140001, MCA-B, ABES Engineering College), this specification bridges the business "
        "needs identified in the BRD (05_BRD.docx) and the User Stories (06_User_Stories.docx) with production-ready software "
        "engineering practices under Python 3.12 and the Streamlit multipage application framework."
    )

    add_heading(doc, "1.1 Revision History & Approval Log", 2)
    rev_headers = ["Version", "Release Date", "Primary Author", "Reviewer / Approver", "Nature of Changes", "Sign-Off Status"]
    rev_rows = [
        ["v0.1", "2026-11-12", "Aayush Mathur", "Internal Self-Review", "Initial technical architecture draft & data pipeline schema", "Draft"],
        ["v0.8", "2026-11-16", "Aayush Mathur", "Peer Code Review", "Added statistical hypothesis testing engine & Tailwind CSS integration", "Reviewed"],
        ["v1.0", "2026-11-19", "Aayush Mathur", "Faculty Guide / Committee", "Final production TRD approval; full alignment with BR-01..14", "Approved"],
    ]
    add_table(doc, rev_headers, rev_rows, col_widths=[0.8, 1.0, 1.2, 1.2, 1.8, 0.8])

    add_heading(doc, "1.2 Project Scope & Technical Objectives", 2)
    add_body(doc,
        "The project delivers a multi-faceted analytical platform engineered to evaluate organizational appraisal distributions, "
        "quantify demographic and tenure fairness, calculate promotion velocity, and predict retention/attrition vectors. "
        "To ensure uncompromising responsiveness and maintainability, the system enforces the following core architectural tenets:"
    )
    add_bullet(doc, "Strict Modular Separation: Isolation of ingestion, data transformation, statistical modeling, and UI rendering.", "Decoupled Architecture: ")
    add_bullet(doc, "Zero Cloud Database Dependency: Embedding SQLite at `data/app.db` ensures zero-setup reproducibility across local and cloud environments.", "Self-Contained Data Layer: ")
    add_bullet(doc, "Sub-2-Second Page Renders: Heavy aggregations and regression models are cached in memory using `@st.cache_data`.", "High-Performance Execution: ")
    add_bullet(doc, "Statistically Rigorous Analytics: Incorporating ANOVA, Kruskal-Wallis, Pearson/Spearman, OLS regression, and Mann-Whitney U tests.", "Mathematical Depth: ")
    add_bullet(doc, "Executive Aesthetic Standards: Combining Streamlit native components with Tailwind CSS CDN utility classes.", "Polished Human Interface: ")

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 2: ARCHITECTURE OVERVIEW & HIGH-LEVEL PIPELINE
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "2. Architecture Overview & High-Level Pipeline", 1)
    
    add_body(doc,
        "The system follows a multi-tiered, reactive pipeline designed to run efficiently inside single-node or containerized "
        "Python runtimes. The architecture transitions raw tabular employee records through validation, clean normalization, "
        "rigorous statistical computation, and dynamic interactive rendering."
    )

    add_callout(doc,
        "Pipeline Flow: Raw Data Source (CSV/Parquet) -> Ingestion & SQLite Engine (`data/app.db`) -> "
        "Data Loader (`src/data_loader.py` with `@st.cache_data`) -> Data Cleaning & Type Enforcement (`src/cleaner.py`) -> "
        "Statistical & Analytical Engine (`src/analysis.py`) -> UI Layout / Multipage Presentation Layer (`app.py` & `pages/*.py`).",
        title="END-TO-END DATA TRANSFORMATION PIPELINE"
    )

    add_heading(doc, "2.1 Mermaid Architecture Flowchart", 2)
    add_body(doc,
        "The following diagram formally specifies the flow of data, state synchronization, caching checkpoints, and component boundaries:"
    )

    mermaid_code = (
        "flowchart TD\n"
        "    subgraph DataStorage [Tier 1: Data Storage Layer]\n"
        "        RAW[Raw Dataset: IBM HR Analytics CSV / Records]\n"
        "        SQLITE[(SQLite Database: data/app.db)]\n"
        "        RAW -->|ETL Pre-seed Script| SQLITE\n"
        "    end\n\n"
        "    subgraph IngestionClean [Tier 2: Ingestion & Quality Layer]\n"
        "        LOADER[src/data_loader.py\n- sqlite3 connection\n- Parameterized SQL\n- @st.cache_data (TTL=3600s)\n- Schema Validation]\n"
        "        CLEANER[src/cleaner.py\n- Null value imputation\n- Type casting & Categorical encoding\n- Outlier capping via IQR\n- Derived metric generation]\n"
        "        SQLITE -->|Read Query| LOADER\n"
        "        LOADER -->|Raw DataFrame| CLEANER\n"
        "    end\n\n"
        "    subgraph AnalyticsEngine [Tier 3: Analytical & Statistical Engine]\n"
        "        ANALYSIS[src/analysis.py\n- Descriptive aggregates: mean, std, percentiles\n- Inferential tests: OLS, ANOVA, Kruskal-Wallis, t-test\n- ML scoring: Logistic Attrition Risk, Promotion Index\n- Correlation Matrices: Pearson & Spearman]\n"
        "        CLEANER -->|Cleaned DataFrame| ANALYSIS\n"
        "    end\n\n"
        "    subgraph PresentationLayer [Tier 4: Presentation & UI Layer]\n"
        "        ENTRY[app.py: Main Hub & Session State Sync]\n"
        "        P1[pages/01_Overview.py: Exec KPIs & Dept Benchmarks]\n"
        "        P2[pages/02_Trend.py: Multi-Year Longitudinal Trends]\n"
        "        P3[pages/03_Compare.py: Dept, Role & Manager ANOVA]\n"
        "        P4[pages/04_EDA_Stats.py: Demographic Bias & Correlations]\n"
        "        P5[pages/05_Insights.py: Promotion Index & Attrition Scoring]\n"
        "        P6[pages/06_Data_Download.py: Filtered CSV/Excel Export]\n"
        "        P7[pages/07_Data_Quality.py: Missing Values & Health Audit]\n"
        "        P8[pages/08_About.py: System Metadata & Student Profile]\n"
        "        STYLE[src/ui_components.py\n- Tailwind CSS CDN Injection\n- Custom Responsive Metric Cards\n- Unified Color Palettes]\n"
        "    end\n\n"
        "    ANALYSIS --> ENTRY\n"
        "    ENTRY --> P1\n"
        "    ENTRY --> P2\n"
        "    ENTRY --> P3\n"
        "    ENTRY --> P4\n"
        "    ENTRY --> P5\n"
        "    ENTRY --> P6\n"
        "    ENTRY --> P7\n"
        "    ENTRY --> P8\n"
        "    STYLE -.->|Visual Inject| ENTRY\n"
        "    STYLE -.->|Visual Inject| P1\n"
        "    STYLE -.->|Visual Inject| P3\n"
        "    STYLE -.->|Visual Inject| P5\n"
    )
    add_code_block(doc, mermaid_code)

    add_heading(doc, "2.2 Layer Decomposition & Architectural Responsibilities", 2)
    add_body(doc,
        "1. Data Storage Layer: Utilizes a standalone, zero-dependency SQLite 3 database located at `data/app.db`. "
        "The schema isolates employee demographic attributes, annual review cycles, managerial mappings, and salary bands. "
        "Database access is strictly read-only during runtime dashboard operations to preserve data integrity."
    )
    add_body(doc,
        "2. Ingestion & Quality Layer: Handled through `src/data_loader.py` and `src/cleaner.py`. Caching decorators ensure "
        "disk reads happen exactly once per session or cache epoch. The cleaner component detects anomalous input, enforces "
        "type consistency (e.g., date formats, integer IDs, categorical strings), and handles missing data gracefully without crashing."
    )
    add_body(doc,
        "3. Analytical & Statistical Engine: Encapsulated within `src/analysis.py`. This module operates as a pure functional engine. "
        "It accepts cleaned pandas DataFrames and parameter tuples, returning mathematical summaries, test statistics (p-values, "
        "F-ratios, t-statistics, R-squared values), and risk scores. It has zero coupling to Streamlit UI elements, enabling automated unit testing."
    )
    add_body(doc,
        "4. Presentation Layer: Implemented as an 8-page Streamlit multipage application. Each page reads shared filter state "
        "from `st.session_state`, invokes `src/analysis.py` for calculations, and renders visuals using Plotly Express/Graph Objects "
        "alongside HTML/Tailwind KPI cards rendered via `st.markdown(..., unsafe_allow_html=True)` or `st.components.v1.html`."
    )

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 3: DIRECTORY TREE & SYSTEM FOLDER STRUCTURE
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "3. Directory Tree & System Folder Structure", 1)
    
    add_body(doc,
        "The application codebase strictly adheres to standard Python and Streamlit production packaging practices. "
        "Separation of concerns is maintained by keeping UI entry points, page views, computational modules, raw data assets, "
        "automated test suites, and documentation isolated in dedicated directories."
    )

    tree_str = (
        "Employee Performance/\n"
        "├── .streamlit/\n"
        "│   ├── config.toml                  # UI theme colors, font families, telemetry disable, server config\n"
        "│   └── secrets.toml.example         # Template for environment-specific secrets (if applicable)\n"
        "├── data/\n"
        "│   ├── app.db                       # Pre-seeded SQLite 3 database containing employee appraisal records\n"
        "│   ├── schema.sql                   # DDL table creation and indexing scripts for SQLite schema\n"
        "│   └── raw_employees.csv            # Original baseline dataset backup (IBM HR Analytics derived)\n"
        "├── docs/\n"
        "│   ├── 01_Project_Understanding.docx # Problem definition, operational questions, stakeholder matrix\n"
        "│   ├── 02_HMW_Statements.docx       # How Might We design-thinking problem decomposition\n"
        "│   ├── 03_Needs_Analysis.docx       # User personas, empathy maps, and feature hierarchy\n"
        "│   ├── 04_Traceability_Matrix.xlsx  # Requirements Traceability Matrix (RTM) linking OQ, HMW, BR, FR\n"
        "│   ├── 05_BRD.docx                  # Business Requirements Document (BR-01 through BR-14, RACI)\n"
        "│   ├── 06_User_Stories.docx         # Agile User Stories with Gherkin Given-When-Then acceptance criteria\n"
        "│   └── 07_TRD.docx                  # Technical Requirements Document (System Architecture & Blueprint)\n"
        "├── pages/\n"
        "│   ├── 01_Overview.py               # Executive KPI dashboard, rating distribution, top/bottom performers\n"
        "│   ├── 02_Trend.py                  # 3-year historical rating trends, rolling averages, seasonality\n"
        "│   ├── 03_Compare.py                # Department & Manager comparative benchmarks, ANOVA/Kruskal-Wallis\n"
        "│   ├── 04_EDA_Stats.py              # Exploratory stats, demographic bias hypothesis tests, correlations\n"
        "│   ├── 05_Insights.py               # Predictive models: Promotion readiness scoring & Attrition risk index\n"
        "│   ├── 06_Data_Download.py          # Interactive dataset explorer and multi-format export (CSV/Excel)\n"
        "│   ├── 07_Data_Quality.py           # Automated schema integrity check, missingness audit, anomaly report\n"
        "│   └── 08_About.py                  # Project metadata, student academic details, architecture documentation\n"
        "├── src/\n"
        "│   ├── __init__.py                  # Package initialization\n"
        "│   ├── data_loader.py               # SQLite connection management, parameterized queries, st.cache_data\n"
        "│   ├── cleaner.py                   # Data cleaning, null handling, IQR outlier capping, categorical encoding\n"
        "│   ├── analysis.py                  # Statistical engine: descriptive, inferential, regression, risk algorithms\n"
        "│   └── ui_components.py             # Reusable UI widgets, Tailwind CSS KPI card generators, Plotly themes\n"
        "├── tests/\n"
        "│   ├── __init__.py                  # Test package initialization\n"
        "│   ├── test_data_loader.py          # Unit tests verifying SQLite ingestion, connection failures, cache validity\n"
        "│   ├── test_cleaner.py              # Unit tests for null handling, type casting, and outlier thresholds\n"
        "│   ├── test_analysis.py             # Mathematical verification of ANOVA, t-tests, Pearson r, and risk scores\n"
        "│   └── test_app_ui.py               # Streamlit AppTest integration tests verifying page renders and filters\n"
        "├── .gitignore                       # Standard Python ignore rules (__pycache__, .pytest_cache, virtualenvs)\n"
        "├── app.py                           # Application main entry point, navigation bar, global session state init\n"
        "├── README.md                        # Comprehensive setup guide, architecture summary, and local execution steps\n"
        "└── requirements.txt                 # Strictly pinned Python dependency versions for reproducible deployment\n"
    )
    add_code_block(doc, tree_str)

    add_heading(doc, "3.1 Key File Responsibilities", 2)
    file_spec_headers = ["File / Path", "Core Function", "Primary Dependencies", "Output Artifact"]
    file_spec_rows = [
        ["app.py", "Global application bootstrap, sidebar navigation, filter initialization", "streamlit", "Shared session state & landing view"],
        ["src/data_loader.py", "SQLite querying, schema validation, exception trapping", "sqlite3, pandas, streamlit", "Cached pd.DataFrame"],
        ["src/cleaner.py", "Data sanitization, type casting, outlier bounding", "pandas, numpy", "Clean validated DataFrame"],
        ["src/analysis.py", "Pure mathematical, inferential, and predictive calculations", "scipy, statsmodels, numpy", "Statistical dictionaries & models"],
        ["src/ui_components.py", "HTML/Tailwind CDN injection, custom card generation", "streamlit, streamlit.components.v1", "Styled HTML blocks & charts"],
        ["tests/test_analysis.py", "Automated validation of statistical outputs and algorithms", "pytest, numpy, scipy", "Test pass/fail assertion report"],
        [".streamlit/config.toml", "Streamlit environment configuration and theme definition", "TOML parser", "Global layout and color bindings"],
    ]
    add_table(doc, file_spec_headers, file_spec_rows, col_widths=[1.5, 2.2, 1.4, 1.6])

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 4: FUNCTIONAL REQUIREMENTS TRACEABILITY
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "4. Functional Requirements Traceability (FR-01 to FR-15)", 1)
    
    add_body(doc,
        "The functional requirements (FR) represent the explicit technical capabilities that the system must deliver. "
        "Each requirement maps directly to the Business Requirements (BR-01 through BR-14) established in the BRD, "
        "linking them to the specific application module, primary source code file, and execution priority."
    )

    fr_headers = ["FR ID", "Functional Requirement Description", "Linked BR / US", "Responsible Module / File", "Priority"]
    fr_rows = [
        [
            "FR-01",
            "Department Performance Summary & KPI Dashboard: Compute aggregate mean, median, standard deviation, and count of ratings grouped by department. Render interactive KPI cards and comparative bar charts.",
            "BR-01\nUS-01",
            "pages/01_Overview.py\nsrc/analysis.py",
            "Must Have"
        ],
        [
            "FR-02",
            "Top & Bottom 10% Performer Identification: Calculate 90th and 10th percentile rating thresholds per department. Generate flagged ranking tables with color-coded status badges for high and low performers.",
            "BR-02\nUS-02",
            "pages/01_Overview.py\nsrc/analysis.py",
            "Must Have"
        ],
        [
            "FR-03",
            "Salary vs. Performance Correlation Engine: Compute Pearson and Spearman rank correlation coefficients between annual compensation and performance rating. Render scatter plots with trend lines and p-values.",
            "BR-03\nUS-03",
            "pages/04_EDA_Stats.py\nsrc/analysis.py",
            "Should Have"
        ],
        [
            "FR-04",
            "Manager Team Performance Benchmark: Group employees by Manager ID, calculate team rating mean and variance, and rank managers. Display sortable league tables with drill-down capability into team rosters.",
            "BR-04\nUS-04",
            "pages/03_Compare.py\nsrc/analysis.py",
            "Must Have"
        ],
        [
            "FR-05",
            "Demographic & Gender Bias Hypothesis Testing: Execute two-sample independent t-tests and non-parametric Mann-Whitney U tests across gender and age cohorts to evaluate rating parity and statistical significance (alpha=0.05).",
            "BR-05\nUS-05",
            "pages/04_EDA_Stats.py\nsrc/analysis.py",
            "Must Have"
        ],
        [
            "FR-06",
            "Promotion Readiness Scoring Algorithm: Compute a normalized 0–100 composite index combining current rating, historical tenure, training hours completed, and projects delivered using weighted linear models.",
            "BR-06\nUS-06",
            "pages/05_Insights.py\nsrc/analysis.py",
            "Must Have"
        ],
        [
            "FR-07",
            "Overdue Promotion Anomaly Detection: Filter and flag employees who maintain a performance rating >= 4.0 and have had no promotion for > 3.0 years. Render alert badges and priority review tables.",
            "BR-14\nUS-07",
            "pages/05_Insights.py\nsrc/analysis.py",
            "Must Have"
        ],
        [
            "FR-08",
            "Departmental Attrition Risk Predictor: Calculate attrition probabilities using multi-factor risk scoring based on low ratings (< 3.0), overtime frequency, low salary hike, and tenure. Group by department.",
            "BR-07\nUS-08",
            "pages/05_Insights.py\nsrc/analysis.py",
            "Must Have"
        ],
        [
            "FR-09",
            "Multi-Year Longitudinal Performance Trend: Ingest multi-period review history, compute annualized rolling averages, trend slopes, and year-over-year rating changes across the organization.",
            "BR-08\nUS-09",
            "pages/02_Trend.py\nsrc/analysis.py",
            "Must Have"
        ],
        [
            "FR-10",
            "Predictive Factor Importance & OLS Regression: Fit Ordinary Least Squares (OLS) regression models predicting performance ratings from tenure, education, training, and work-life balance. Output R2 and beta coefficients.",
            "BR-09\nUS-10",
            "pages/05_Insights.py\nsrc/analysis.py",
            "Should Have"
        ],
        [
            "FR-11",
            "Job Role Comparative & ANOVA Engine: Perform One-way ANOVA and Kruskal-Wallis omnibus tests across job roles to detect statistically significant rating divergence; provide post-hoc Tukey HSD test summaries.",
            "BR-13\nUS-11",
            "pages/03_Compare.py\nsrc/analysis.py",
            "Should Have"
        ],
        [
            "FR-12",
            "Global Interactive Filtering & Session State Sync: Provide unified sidebar widgets (Department, Role, Year, Performance Band) that persist filter selections across all multipage views via st.session_state.",
            "BR-10\nUS-12",
            "app.py\nsrc/ui_components.py",
            "Must Have"
        ],
        [
            "FR-13",
            "Automated Data Quality Audit & Schema Validation: Inspect SQLite tables upon load for null value percentages, negative values, type mismatches, and duplicate records; display health scores on dedicated audit screen.",
            "BR-11\nUS-13",
            "pages/07_Data_Quality.py\nsrc/cleaner.py",
            "Must Have"
        ],
        [
            "FR-14",
            "Multi-Format Data & Report Export Engine: Provide client-side export capabilities allowing users to download active filtered subsets into CSV or styled Microsoft Excel (.xlsx) formats with timestamped filenames.",
            "BR-12\nUS-14",
            "pages/06_Data_Download.py\nsrc/utils.py",
            "Could Have"
        ],
        [
            "FR-15",
            "System Architecture, Metadata & Profile Viewer: Display technical specifications, active library versions, database connection status, and academic student details on the About screen.",
            "BR-11\nUS-15",
            "pages/08_About.py\nsrc/data_loader.py",
            "Should Have"
        ],
    ]
    add_table(doc, fr_headers, fr_rows, col_widths=[0.7, 2.6, 0.9, 1.6, 0.9])

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 5: DATA LAYER ENGINEERING & CACHING STRATEGY
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "5. Data Layer Engineering & SQLite Caching Strategy", 1)
    
    add_body(doc,
        "The data layer is built to guarantee zero external database setup overhead while enforcing relational integrity "
        "and sub-second query execution. This section details the storage mechanism, query architecture, validation schema, "
        "and caching policy."
    )

    add_heading(doc, "5.1 Storage Engine & Connection Protocol", 2)
    add_body(doc,
        "Data is stored in a local SQLite 3 database located at `data/app.db`. SQLite provides a fully ACID-compliant, "
        "serverless relational database embedded directly inside the application repository. Connection strings use standard Python paths:"
    )
    add_code_block(doc,
        "import sqlite3\n"
        "from pathlib import Path\n\n"
        "DB_PATH = Path(__file__).resolve().parent.parent / 'data' / 'app.db'\n\n"
        "def get_connection():\n"
        "    if not DB_PATH.exists():\n"
        "        raise FileNotFoundError(f'Critical Error: Database file not located at {DB_PATH}')\n"
        "    # URI=True enables read-only query mode to guarantee thread safety and prevent accidental writes\n"
        "    conn = sqlite3.connect(f'file:{DB_PATH}?mode=ro', uri=True, check_same_thread=False)\n"
        "    return conn\n"
    )

    add_heading(doc, "5.2 Caching Strategy & Streamlit Lifecycle", 2)
    add_body(doc,
        "Streamlit reruns the entire Python script upon every user interaction (e.g., clicking a dropdown or slider). "
        "Without aggressive caching, repeated disk reads and SQL queries would degrade responsiveness. The system enforces "
        "in-memory caching via `@st.cache_data`:"
    )
    add_bullet(doc, "The primary ingestion function `load_employee_data()` is wrapped with `@st.cache_data(ttl=3600, show_spinner=False)`.", "Cache Decoration: ")
    add_bullet(doc, "Once loaded into memory, subsequent page switches and filter queries read from RAM in < 5 milliseconds.", "Zero Disk I/O: ")
    add_bullet(doc, "The cache key incorporates the database file modification timestamp (`st_mtime`). If `app.db` is updated, the cache invalidates automatically.", "Cache Invalidation: ")

    add_heading(doc, "5.3 Relational Schema Definition", 2)
    add_body(doc,
        "The relational schema inside `data/app.db` comprises three normalized entity tables linked by foreign keys:"
    )
    schema_code = (
        "-- Schema definition for Employee Performance Rating Analysis (data/schema.sql)\n\n"
        "CREATE TABLE IF NOT EXISTS departments (\n"
        "    dept_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    dept_name TEXT NOT NULL UNIQUE,\n"
        "    cost_center TEXT NOT NULL\n"
        ");\n\n"
        "CREATE TABLE IF NOT EXISTS employees (\n"
        "    emp_id TEXT PRIMARY KEY,\n"
        "    first_name TEXT NOT NULL,\n"
        "    last_name TEXT NOT NULL,\n"
        "    gender TEXT CHECK(gender IN ('Male', 'Female', 'Non-Binary', 'Other')),\n"
        "    age INTEGER CHECK(age >= 18 AND age <= 75),\n"
        "    dept_id INTEGER REFERENCES departments(dept_id),\n"
        "    job_role TEXT NOT NULL,\n"
        "    education_level TEXT,\n"
        "    monthly_income REAL CHECK(monthly_income > 0),\n"
        "    hire_date DATE NOT NULL,\n"
        "    manager_id TEXT,\n"
        "    tenure_years REAL CHECK(tenure_years >= 0)\n"
        ");\n\n"
        "CREATE TABLE IF NOT EXISTS performance_reviews (\n"
        "    review_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    emp_id TEXT REFERENCES employees(emp_id),\n"
        "    review_year INTEGER NOT NULL,\n"
        "    performance_rating REAL CHECK(performance_rating >= 1.0 AND performance_rating <= 5.0),\n"
        "    training_hours INTEGER DEFAULT 0,\n"
        "    projects_completed INTEGER DEFAULT 0,\n"
        "    years_since_last_promotion REAL DEFAULT 0.0,\n"
        "    attrition_flag INTEGER CHECK(attrition_flag IN (0, 1)),\n"
        "    work_life_balance_score INTEGER CHECK(work_life_balance_score BETWEEN 1 AND 5)\n"
        ");\n\n"
        "CREATE INDEX IF NOT EXISTS idx_reviews_emp ON performance_reviews(emp_id);\n"
        "CREATE INDEX IF NOT EXISTS idx_reviews_year ON performance_reviews(review_year);\n"
        "CREATE INDEX IF NOT EXISTS idx_emp_dept ON employees(dept_id);\n"
    )
    add_code_block(doc, schema_code)

    add_heading(doc, "5.4 Error Handling & Fallback Protocol", 2)
    add_body(doc,
        "To protect the user experience from missing or corrupt database states, `src/data_loader.py` wraps database access "
        "in defensive boundary handlers:"
    )
    add_bullet(doc, "If `app.db` is missing, the application halts gracefully, displays an informative red alert card via Streamlit, and generates a fallback mock DataFrame so interface exploration does not crash.", "Missing DB Handling: ")
    add_bullet(doc, "The loader verifies that mandatory columns (`emp_id`, `performance_rating`, `dept_name`) exist. Missing columns raise a custom `SchemaValidationError`.", "Schema Verification: ")
    add_bullet(doc, "SQLite locks and database disk errors are captured with `sqlite3.OperationalError` and logged to standard error.", "Concurrency Protection: ")

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 6: STATISTICAL METHODS & OPERATIONAL QUESTION MAPPING
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "6. Statistical Methods & Operational Question Mapping", 1)
    
    add_body(doc,
        "To ensure that the platform delivers academic and industrial grade analytical depth, each Operational Question (OQ) "
        "defined in Step 1 is paired with exact mathematical algorithms, Python implementations, hypothesis formulations, "
        "and acceptance thresholds."
    )

    stat_headers = ["OQ ID", "Operational Question", "Statistical / Analytical Technique", "Python Package & Function", "Mathematical Formulation / Logic", "Threshold / Output"]
    stat_rows = [
        [
            "OQ-01",
            "What is the average performance rating per department?",
            "Descriptive Summary Statistics (Mean, Median, Standard Deviation, Standard Error)",
            "pandas.DataFrame.groupby()\nnumpy.mean()\nscipy.stats.sem()",
            "mu_d = (1 / N_d) * sum(r_i)\nstd_d = sqrt( sum((r_i - mu_d)^2) / (N_d - 1) )",
            "Aggregated KPI table; Bar chart with +/- 1 SE error bars"
        ],
        [
            "OQ-02",
            "Which employees are consistently high/low performers over time?",
            "Percentile Ranking & Quantile Discretization",
            "pandas.DataFrame.quantile()\nscipy.stats.percentileofscore()",
            "High: Rating >= Q_0.90\nLow: Rating <= Q_0.10\nRank_i = 100 * (Rank(r_i) / N)",
            "Top 10% (Green badge)\nBottom 10% (Red badge)"
        ],
        [
            "OQ-03",
            "How does salary correlate with performance rating?",
            "Bivariate Correlation Analysis (Pearson parametric & Spearman rank non-parametric)",
            "scipy.stats.pearsonr()\nscipy.stats.spearmanr()",
            "r_p = Cov(X,Y) / (sigma_X * sigma_Y)\nrho_s = 1 - (6 * sum(d_i^2) / (n(n^2 - 1)))",
            "r and rho coefficients with two-tailed p-value (alpha=0.05)"
        ],
        [
            "OQ-04",
            "Which manager's team has the highest average performance?",
            "Group Aggregation, League Table Ranking & Variance Ratio Test",
            "pandas.DataFrame.groupby('manager_id')\nscipy.stats.bartlett()",
            "mu_m = (1 / n_m) * sum(r_mj)\ns^2_m = (1 / (n_m - 1)) * sum(r_mj - mu_m)^2",
            "Manager leaderboard sorted by mu_m; Team spread metric"
        ],
        [
            "OQ-05",
            "What is the distribution of ratings by gender/age group, and does bias exist?",
            "Two-Sample Independent Student's t-Test & Mann-Whitney U Test",
            "scipy.stats.ttest_ind()\nscipy.stats.mannwhitneyu()",
            "t = (x_bar_1 - x_bar_2) / sqrt(s_p^2 (1/n1 + 1/n2))\nU = R1 - (n1(n1+1)/2)",
            "Reject H0 (bias present) if p < 0.05; Cohen's d effect size"
        ],
        [
            "OQ-06",
            "Are there significant performance differences across job roles?",
            "One-Way ANOVA (Parametric) & Kruskal-Wallis H-Test (Non-parametric) + Tukey HSD",
            "scipy.stats.f_oneway()\nscipy.stats.kruskal()\nstatsmodels.stats.multicomp.pairwise_tukeyhsd()",
            "F = MS_between / MS_within\nH = (12 / (N(N+1))) * sum(R_j^2 / n_j) - 3(N+1)",
            "F-stat, p-value; Pairwise confidence intervals for role differences"
        ],
        [
            "OQ-07",
            "Which departments have the highest attrition risk (low ratings)?",
            "Multi-Criteria Risk Scoring Model & Logistic Probability Function",
            "statsmodels.discrete.discrete_model.Logit()\nnumpy.clip()",
            "Risk_Score = w1*(5 - Rating) + w2*(Years_No_Promo) + w3*(Overtime) + w4*(Low_Hike)",
            "Risk index normalized to 0–100; High risk flag if score >= 70"
        ],
        [
            "OQ-08",
            "How has overall company performance trended over the last 3 years?",
            "Longitudinal Trend Analysis, Simple Moving Averages & Year-over-Year Delta",
            "pandas.DataFrame.rolling()\nnumpy.polyfit()",
            "SMA_t = (1 / k) * sum(r_{t-i})\nDelta_YoY = ((mu_{t} - mu_{t-1}) / mu_{t-1}) * 100",
            "3-Year trend line with linear slope beta and YoY percentage"
        ],
        [
            "OQ-09",
            "What factors (experience, education, training) predict high performance?",
            "Multiple Ordinary Least Squares (OLS) Linear Regression",
            "statsmodels.api.OLS()\nsklearn.ensemble.RandomForestRegressor()",
            "Y_Rating = beta_0 + beta_1*Tenure + beta_2*Training + beta_3*Education + epsilon",
            "Model R2, Adjusted R2, beta weights, p-values (< 0.05) & Feature Ranking"
        ],
        [
            "OQ-10",
            "Which employees are overdue for a promotion based on tenure and rating?",
            "Multi-Attribute Promotion Readiness Index & Logical Anomaly Filter",
            "src.analysis.calc_promotion_readiness()\npandas.DataFrame.query()",
            "P_Index = 0.45*(Rating/5) + 0.25*(Tenure_norm) + 0.20*(Training_norm) + 0.10*(Proj_norm)\nFlag: Rating >= 4.0 & Years_No_Promo > 3.0",
            "Readiness Score [0, 100]; Flagged roster of overdue candidates"
        ],
    ]
    add_table(doc, stat_headers, stat_rows, col_widths=[0.6, 1.4, 1.3, 1.3, 1.3, 1.1])

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 7: UI DESIGN, STYLING & SESSION STATE SYNCHRONISATION
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "7. User Interface Design, Styling & Session State Synchronisation", 1)
    
    add_body(doc,
        "The interface marries Streamlit's rapid execution with modern enterprise design standards. "
        "By injecting the Tailwind CSS CDN directly into the document object model and maintaining persistent state "
        "variables in `st.session_state`, the application guarantees seamless visual cohesion and cross-page filter synchronization."
    )

    add_heading(doc, "7.1 Tailwind CSS Integration Pattern", 2)
    add_body(doc,
        "Because Streamlit renders in an isolated DOM tree, Tailwind CSS utility classes are loaded globally within `app.py` "
        "via a custom HTML inject block using `st.markdown(..., unsafe_allow_html=True)`. This allows high-fidelity cards, "
        "badges, and layout grids without maintaining complex external CSS sheets:"
    )
    tailwind_snippet = (
        "# Global styling injector in app.py and src/ui_components.py\n"
        "import streamlit as st\n\n"
        "def inject_tailwind():\n"
        "    tailwind_cdn = '''\n"
        "    <script src=\"https://cdn.tailwindcss.com\"></script>\n"
        "    <style>\n"
        "        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');\n"
        "        html, body, [class*=\"css\"] { font-family: 'Inter', sans-serif; }\n"
        "    </style>\n"
        "    '''\n"
        "    st.markdown(tailwind_cdn, unsafe_allow_html=True)\n\n"
        "def render_kpi_card(title, value, delta, color='blue'):\n"
        "    # Generates a Tailwind-styled responsive executive KPI widget\n"
        "    card_html = f'''\n"
        "    <div class=\"bg-white rounded-xl shadow-sm border border-slate-200 p-5 mb-4 hover:shadow-md transition-shadow\">\n"
        "        <div class=\"text-xs font-semibold uppercase tracking-wider text-slate-500 mb-1\">{title}</div>\n"
        "        <div class=\"text-2xl font-bold text-slate-800\">{value}</div>\n"
        "        <div class=\"text-xs font-medium mt-2 text-{color}-600\">{delta}</div>\n"
        "    </div>\n"
        "    '''\n"
        "    st.markdown(card_html, unsafe_allow_html=True)\n"
    )
    add_code_block(doc, tailwind_snippet)

    add_heading(doc, "7.2 Unified Color Palette & Visual Hierarchy", 2)
    add_body(doc,
        "To convey immediate analytical meaning and adhere to WCAG 2.1 AA accessibility guidelines, "
        "the application adheres to a strict semantic color palette:"
    )
    palette_headers = ["Semantic Role", "Hex Code", "Tailwind Class", "Usage in System"]
    palette_rows = [
        ["Primary Brand / Corporate", "#1E3A8A", "bg-blue-900 / text-blue-900", "Headers, page titles, navigation banners, primary chart series"],
        ["Secondary Accent", "#3B82F6", "bg-blue-500 / text-blue-500", "Interactive buttons, selected tabs, correlation highlights, regression lines"],
        ["Success / High Performance", "#10B981", "bg-emerald-500 / text-emerald-500", "Top 10% badges, promotion ready flags, high appraisal bands (>= 4.0)"],
        ["Danger / Attrition Alert", "#EF4444", "bg-red-500 / text-red-500", "Bottom 10% badges, overdue promotion alerts, high attrition risk, ratings (< 3.0)"],
        ["Warning / Moderate", "#F59E0B", "bg-amber-500 / text-amber-500", "Mid-tier performers (3.0–3.9), moderate attrition watchlists, notice alerts"],
        ["Neutral Background", "#F8FAFC", "bg-slate-50", "Application background canvas, alternating table rows, callout containers"],
        ["Border / Card Outlines", "#E2E8F0", "border-slate-200", "Card container borders, table cell dividers, chart gridlines"],
    ]
    add_table(doc, palette_headers, palette_rows, col_widths=[1.5, 1.1, 1.7, 2.7])

    add_heading(doc, "7.3 Cross-Page Filter Synchronization via st.session_state", 2)
    add_body(doc,
        "In a multipage Streamlit application, sidebar widgets placed on individual pages reset when navigating between views "
        "unless explicitly bound to `st.session_state`. The system initializes global filter keys in `app.py`:"
    )
    session_snippet = (
        "# Session state filter initialization in app.py\n"
        "if 'selected_department' not in st.session_state:\n"
        "    st.session_state.selected_department = 'All'\n"
        "if 'selected_year' not in st.session_state:\n"
        "    st.session_state.selected_year = 'All'\n"
        "if 'min_rating_threshold' not in st.session_state:\n"
        "    st.session_state.min_rating_threshold = 1.0\n\n"
        "# Filter synchronizer widget inside sidebar\n"
        "def render_sidebar_filters(df):\n"
        "    dept_list = ['All'] + sorted(df['dept_name'].dropna().unique().tolist())\n"
        "    st.sidebar.selectbox('Filter Department', dept_list, key='selected_department')\n"
        "    year_list = ['All'] + sorted(df['review_year'].dropna().unique().tolist())\n"
        "    st.sidebar.selectbox('Filter Year', year_list, key='selected_year')\n"
    )
    add_code_block(doc, session_snippet)

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 8: PERFORMANCE, ERROR HANDLING & SECURITY
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "8. Performance Optimization, Error Handling & System Security", 1)
    
    add_body(doc,
        "To ensure production resilience, the application establishes strict non-functional constraints governing latency, "
        "defensive fault isolation, data sanitization, and database protection."
    )

    add_heading(doc, "8.1 Performance Targets & Latency Budgets", 2)
    add_body(doc,
        "The system enforces the following performance service level indicators (SLIs):"
    )
    perf_headers = ["Metric / Operation", "Target Latency", "Measurement Method", "Enforcement Strategy"]
    perf_rows = [
        ["Initial Application Boot", "< 2.0 seconds", "Streamlit cold start benchmark", "Pre-warmed SQLite connection, deferred heavy module imports"],
        ["Page-to-Page Navigation", "< 0.5 seconds", "Browser navigation timing API", "@st.cache_data for all raw and cleaned DataFrames"],
        ["Complex Statistical Calculation", "< 0.8 seconds", "Python time.perf_counter() timer", "Vectorized NumPy operations; avoiding iterative Python loops"],
        ["Data Export Generation (CSV/Excel)", "< 1.5 seconds", "Streamlit download button click", "In-memory BytesIO buffers; avoiding disk write bottlenecks"],
        ["Memory Footprint", "< 250 MB RAM", "Linux container cgroup memory metrics", "Downcasting integer and float dtypes in pandas"],
    ]
    add_table(doc, perf_headers, perf_rows, col_widths=[1.8, 1.2, 1.8, 2.2])

    add_heading(doc, "8.2 Error Handling & Try-Except Boundaries", 2)
    add_body(doc,
        "Uncaught Python exceptions produce unsightly tracebacks that confuse non-technical users. The system isolates "
        "every critical pipeline phase behind structured try-except blocks:"
    )
    add_bullet(doc, "Database I/O operations trap `sqlite3.DatabaseError` and display friendly recovery instructions.", "Database Faults: ")
    add_bullet(doc, "Zero-division or insufficient sample sizes in ANOVA/t-tests trap `ZeroDivisionError` and `ValueError`, rendering an informational warning card rather than breaking the UI.", "Statistical Edge Cases: ")
    add_bullet(doc, "User file downloads trap buffer errors and notify the user if an export times out.", "Export Resiliency: ")

    add_heading(doc, "8.3 Security Architecture & Data Protection", 2)
    add_body(doc,
        "Although designed as an analytical reporting application, the architecture follows the principle of least privilege:"
    )
    add_bullet(doc, "All database queries use parameterized placeholders (`?`), completely eliminating SQL injection risks.", "Parameterized Queries: ")
    add_bullet(doc, "The SQLite connection is opened with `mode=ro` (read-only), preventing runtime write or DROP injection attacks.", "Read-Only Database: ")
    add_bullet(doc, "Zero hard-coded credentials, passwords, or personal API keys exist in the repository; configurations are strictly file-based.", "Zero Hard-coded Secrets: ")
    add_bullet(doc, "Input widgets sanitize text queries to prevent cross-site scripting (XSS) when rendered into HTML cards.", "XSS Prevention: ")

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 9: PINNED DEPENDENCY SPECIFICATIONS (requirements.txt)
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "9. Pinned Dependency Specifications (requirements.txt)", 1)
    
    add_body(doc,
        "To guarantee 100% deterministic, reproducible builds across local developer machines, continuous integration runners, "
        "and Streamlit Community Cloud Linux containers, all external Python packages are pinned to exact semantic versions. "
        "The following specification constitutes the official `requirements.txt`:"
    )

    req_headers = ["Package Name", "Pinned Version", "Ecosystem Category", "Architectural Role & Justification"]
    req_rows = [
        ["streamlit", "==1.38.0", "Web UI Framework", "Core reactive presentation layer, multipage routing, widget state binding"],
        ["pandas", "==2.2.2", "Data Manipulation", "Tabular DataFrame ingestion, grouping, filtering, cleaning, and time series handling"],
        ["numpy", "==1.26.4", "Numerical Computing", "High-performance vectorized mathematical calculations, matrix math, array masks"],
        ["scipy", "==1.14.0", "Scientific Computing", "Inferential hypothesis testing (t-test, Mann-Whitney U, ANOVA, Pearson/Spearman)"],
        ["statsmodels", "==0.14.2", "Statistical Modeling", "Ordinary Least Squares (OLS) regression, Tukey HSD post-hoc, ANOVA tables"],
        ["plotly", "==5.24.1", "Data Visualization", "Interactive charts (bar, scatter, box, violin, line) with tooltips and zoom"],
        ["openpyxl", "==3.1.5", "Spreadsheet I/O", "Excel spreadsheet read/write engine for structured workbook data exports"],
        ["pytest", "==8.3.2", "Test Framework", "Automated unit and regression test suite execution across models and cleaner"],
        ["python-docx", "==1.2.0", "Document Automation", "Automated technical document and report generation engine"],
    ]
    add_table(doc, req_headers, req_rows, col_widths=[1.3, 1.0, 1.5, 3.2])

    add_heading(doc, "9.1 Exact requirements.txt Code Block", 2)
    req_text = (
        "# Employee Performance Rating Analysis — Official Pinned Dependencies\n"
        "# Python 3.12 Runtime Compatibility\n"
        "streamlit==1.38.0\n"
        "pandas==2.2.2\n"
        "numpy==1.26.4\n"
        "scipy==1.14.0\n"
        "statsmodels==0.14.2\n"
        "plotly==5.24.1\n"
        "openpyxl==3.1.5\n"
        "pytest==8.3.2\n"
        "python-docx==1.2.0\n"
    )
    add_code_block(doc, req_text)

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 10: DEPLOYMENT REQUIREMENTS & CLOUD INFRASTRUCTURE
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "10. Deployment Requirements & Cloud Infrastructure", 1)
    
    add_body(doc,
        "The application is engineered for turnkey deployment on Streamlit Community Cloud directly linked to a public or "
        "private GitHub repository. This section outlines repository configuration, file path resolution, and cloud container specifications."
    )

    add_heading(doc, "10.1 Path Resolution Standards (Cross-Platform Portability)", 2)
    add_body(doc,
        "A common point of failure in student capstone projects is hard-coding local Windows paths (e.g., `C:\\Users\\...`). "
        "The application strictly enforces relative path resolution using Python's standard `pathlib.Path` library. "
        "This ensures immediate execution across Windows, macOS, and Linux container runtimes:"
    )
    path_code = (
        "# Mandatory path resolution standard across all src/ and pages/ modules\n"
        "from pathlib import Path\n\n"
        "BASE_DIR = Path(__file__).resolve().parent.parent\n"
        "DATA_DIR = BASE_DIR / 'data'\n"
        "DB_FILE  = DATA_DIR / 'app.db'\n"
        "DOCS_DIR = BASE_DIR / 'docs'\n"
    )
    add_code_block(doc, path_code)

    add_heading(doc, "10.2 GitHub Repository Structure & Deployment Workflow", 2)
    add_body(doc,
        "The continuous delivery pipeline to Streamlit Community Cloud operates as follows:"
    )
    add_bullet(doc, "The repository root contains `app.py`, `requirements.txt`, and `.streamlit/config.toml`.", "Repository Root: ")
    add_bullet(doc, "The SQLite database `data/app.db` is committed directly to git (file size is approximately 1.5MB, well below GitHub's 100MB limit).", "Database Packaging: ")
    add_bullet(doc, "Developers link their GitHub repository in the Streamlit Cloud console, selecting `app.py` as the main file path and Python 3.12 as the target runtime.", "Cloud Console Linkage: ")
    add_bullet(doc, "Every `git push origin main` triggers an automated container rebuild, dependency pip install, and rolling restart in < 60 seconds.", "Continuous Deployment: ")

    add_heading(doc, "10.3 Streamlit Cloud Environment Configuration", 2)
    add_body(doc,
        "The file `.streamlit/config.toml` dictates server settings, layout options, and visual theming:"
    )
    config_code = (
        "# .streamlit/config.toml\n"
        "[server]\n"
        "headless = true\n"
        "port = 8501\n"
        "enableCORS = false\n"
        "enableXsrfProtection = true\n\n"
        "[theme]\n"
        "primaryColor = \"#1E3A8A\"\n"
        "backgroundColor = \"#FFFFFF\"\n"
        "secondaryBackgroundColor = \"#F8FAFC\"\n"
        "textColor = \"#1F2937\"\n"
        "font = \"sans serif\"\n\n"
        "[browser]\n"
        "gatherUsageStats = false\n"
    )
    add_code_block(doc, config_code)

    page_break(doc)

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 11: TECHNICAL RISKS & MITIGATION
    # ══════════════════════════════════════════════════════════════════════════
    add_heading(doc, "11. Technical Risks & Mitigation Matrix", 1)
    
    add_body(doc,
        "A rigorous engineering plan anticipates potential technical failure modes across runtime dependencies, data corruption, "
        "computational bottlenecks, and deployment drift. The following matrix details technical risks, severity ratings, "
        "and proactive mitigation strategies."
    )

    risk_headers = ["Risk ID", "Technical Risk Description", "Impact", "Likelihood", "Mitigation Strategy & Engineering Safeguard"]
    risk_rows = [
        [
            "TR-01",
            "SQLite database file lock or corruption during simultaneous access in Streamlit server threads",
            "High",
            "Low",
            "Open SQLite connection with read-only URI flag ('?mode=ro') and 'check_same_thread=False'. Wrap queries in caching layer."
        ],
        [
            "TR-02",
            "Dependency version drift or breaking changes in downstream libraries (e.g., Streamlit, pandas, statsmodels)",
            "High",
            "Medium",
            "Strictly pin all package versions in requirements.txt (e.g., streamlit==1.38.0); lock build environment using fixed Python 3.12."
        ],
        [
            "TR-03",
            "Memory exhaustion (OOM) on free-tier Streamlit Community Cloud containers (1GB RAM ceiling)",
            "High",
            "Low",
            "Downcast numeric datatypes (float64 -> float32, int64 -> int32), avoid caching duplicate copies of data, enforce single-pass aggregations."
        ],
        [
            "TR-04",
            "Statistical test failure due to zero variance or insufficient sample sizes (e.g., N < 3 in small departments)",
            "Medium",
            "Medium",
            "Implement defensive sample size check (N >= 5) before running ANOVA or t-tests. Fallback gracefully with an informational notice card."
        ],
        [
            "TR-05",
            "Relative path resolution failure between Windows local dev and Linux cloud container",
            "High",
            "Low",
            "Enforce strict usage of 'pathlib.Path(__file__).resolve().parent' across all code; prohibit hard-coded backslashes or absolute Windows paths."
        ],
        [
            "TR-06",
            "Tailwind CSS CDN loading failure due to client-side ad-blockers or network latency",
            "Low",
            "Low",
            "Provide fallback inline CSS styling in markdown templates ensuring critical layout and font structures remain intact if CDN fails."
        ],
        [
            "TR-07",
            "Multi-page session state loss when user bookmarks or navigates directly to a subpage URL",
            "Medium",
            "Medium",
            "Implement idempotent session state initializers at the top of every individual page script, ensuring defaults are restored if missing."
        ],
    ]
    add_table(doc, risk_headers, risk_rows, col_widths=[0.8, 2.3, 0.8, 0.8, 2.3])

    add_body(doc,
        "Conclusion: This Technical Requirements Document establishes a complete, unambiguous, and mathematically sound "
        "engineering roadmap. By adhering strictly to the architecture, data schemas, statistical models, and directory boundaries "
        "defined herein, the implementation of 'Employee Performance Rating Analysis' will achieve exceptional academic and industrial quality."
    )

    # ── Save Document ──
    print(f"Saving TRD document to {OUTPUT_PATH}...")
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    doc.save(OUTPUT_PATH)
    print("Document successfully created and saved.")

if __name__ == "__main__":
    generate_trd()
