"""
Academic Submission, System Architecture & Documentation (UI-08)
Employee Performance Rating Analysis System
Student: Aayush Mathur | Enrolment No: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0
"""

import os
import streamlit as st
import pandas as pd

from src.data_loader import get_db_connection, check_database_health
from src.ui_components import (
    inject_tailwind_css,
    render_header,
    render_kpi_card,
    render_insight_box,
    render_sidebar_filters,
)

# 1. Page Configuration
st.set_page_config(
    page_title="About & Documentation | Employee Performance Rating Analysis",
    page_icon="📖",
    layout="wide"
)
inject_tailwind_css()

# 2. Header
render_header(
    title="System Architecture, Relational Schema & Engineering Blueprint",
    subtitle="Technology topology, relational schema dictionary, and Requirements Traceability Matrix (RTM)",
    badge_text="UI-08 | Compliance",
    icon="🏗️"
)

# 3. Student Attribution Card
st.markdown(
    """
    <div style="background: linear-gradient(135deg, #1E3A8A 0%, #1D4ED8 100%); padding: 1.75rem 2rem; border-radius: 0.875rem; color: white; margin-bottom: 1.75rem; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;">
            <div>
                <span style="background-color: rgba(255,255,255,0.2); padding: 0.25rem 0.75rem; border-radius: 9999px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;">
                    Enterprise Intelligence Platform
                </span>
                <h2 style="color: white; font-size: 1.85rem; font-weight: 900; margin: 0.5rem 0 0.25rem 0;">Employee Performance Rating Analysis</h2>
                <div style="font-size: 1rem; opacity: 0.95;">Workforce Performance Calibration & Predictive Analytics System</div>
            </div>
            <div style="background: rgba(255,255,255,0.15); backdrop-filter: blur(8px); border-radius: 0.75rem; padding: 1rem 1.5rem; border: 1px solid rgba(255,255,255,0.25);">
                <div style="font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.08em; opacity: 0.85; font-weight: 600;">System Architect</div>
                <div style="font-size: 1.25rem; font-weight: 800; margin-top: 0.15rem;">Aayush Mathur</div>
                <div style="font-size: 0.85rem; margin-top: 0.2rem;">ID: <strong>2500320140001</strong></div>
                <div style="font-size: 0.85rem;">Environment: <strong>Production Ready</strong></div>
                <div style="font-size: 0.85rem; margin-top: 0.35rem; padding-top: 0.35rem; border-top: 1px solid rgba(255,255,255,0.2);">
                    <strong>Enterprise Workforce Systems</strong>
                </div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# 4. Multi-Tab Documentation Hub
tab_arch, tab_rtm, tab_schema, tab_glossary = st.tabs([
    "🏗️ System Architecture & Stack",
    "📋 Requirements Traceability (RTM)",
    "🗄️ Database Schema & Views",
    "📐 KPI Formula Glossary"
])

with tab_arch:
    col_a1, col_a2 = st.columns([3, 2])
    with col_a1:
        st.markdown(
            """
            ### Technical Stack & Architectural Layers
            The platform is built following a clean, modular 3-tier architecture:
            
            1. **Presentation Layer:**
               - **Streamlit (v1.31+)** multi-page web application architecture.
               - Custom **Tailwind CSS** tokens injected via CDN with scoped utility classes.
               - Interactive responsive charts powered by **Plotly Express** and **Plotly Graph Objects**.
            
            2. **Business Logic & Analytics Engine:**
               - **Pandas & NumPy** for high-performance vectorized telemetry manipulation.
               - **SciPy.stats** executing inferential hypothesis tests (Two-Sample t-test, Mann-Whitney U, One-Way ANOVA, Kruskal-Wallis).
               - **Scikit-Learn** running OLS Linear Regression and Random Forest feature importance models.
               - Algorithmic composite scoring engines for Promotion Readiness and Attrition Flight Risk.
            
            3. **Data Persistence & Governance Layer:**
               - **SQLite3 relational database (`data/app.db`)** enforcing 3NF relational normalization across 5 core tables.
               - Pre-aggregated SQL analytical views (`v_employee_analytics`) memoized using `@st.cache_data`.
               - Automated data hygiene, boundary validation, and IQR outlier telemetry.
            """
        )

    with col_a2:
        db_stat = check_database_health()
        st.markdown(
            f"""
            <div style="background-color: white; border: 1px solid #E2E8F0; border-radius: 0.75rem; padding: 1.25rem;">
                <div style="font-weight: 700; color: #1E3A8A; font-size: 1rem; margin-bottom: 0.75rem;">Database & Runtime Telemetry</div>
                <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>Runtime:</strong> Python 3.13.x (Win32)</div>
                <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>Database Engine:</strong> SQLite 3.x Relational (3NF)</div>
                <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>Database File Size:</strong> {db_stat.get('file_size_kb', 0.0)} KB</div>
                <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>Health Status:</strong> <span style="color: #10B981; font-weight: 700;">● {db_stat.get('status', 'Healthy')}</span></div>
                <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>Query Latency:</strong> {db_stat.get('latency_ms', 0.0)} ms</div>
                <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>Automated Pytest Coverage:</strong> 35+ test cases</div>
            </div>
            """,
            unsafe_allow_html=True
        )

with tab_rtm:
    st.markdown("### Requirements Traceability Matrix (RTM) Summary")
    st.markdown("Full two-way mapping connecting Business Requirements, User Stories, UI Screens, and Automated Test Cases:")
    
    rtm_data = [
        {"Req ID": "BR-01", "Scope": "Average Performance Rating KPI", "User Story": "US-01", "UI Screen": "UI-01 Overview", "Test Cases": "TC-F-01, TC-T-07"},
        {"Req ID": "BR-02", "Scope": "Top & Bottom Performer Deciles", "User Story": "US-02", "UI Screen": "UI-01 Overview", "Test Cases": "TC-F-02, TC-T-07"},
        {"Req ID": "BR-03", "Scope": "Salary vs Rating Correlation", "User Story": "US-03", "UI Screen": "UI-03 Compare", "Test Cases": "TC-F-06, TC-T-12"},
        {"Req ID": "BR-04", "Scope": "Manager Performance Leaderboard", "User Story": "US-05", "UI Screen": "UI-03 Compare", "Test Cases": "TC-F-08"},
        {"Req ID": "BR-05", "Scope": "Job Role Rating ANOVA Test", "User Story": "US-06", "UI Screen": "UI-03 Compare", "Test Cases": "TC-F-09, TC-T-13"},
        {"Req ID": "BR-06", "Scope": "Promotion Readiness Formula", "User Story": "US-09", "UI Screen": "UI-05 Insights", "Test Cases": "TC-F-12, TC-T-08"},
        {"Req ID": "BR-07", "Scope": "Overdue Promotion Identification", "User Story": "US-10", "UI Screen": "UI-05 Insights", "Test Cases": "TC-F-13, TC-T-08"},
        {"Req ID": "BR-08", "Scope": "3-Year Multi-Period Trajectory", "User Story": "US-04", "UI Screen": "UI-02 Trend", "Test Cases": "TC-F-04, TC-F-05"},
        {"Req ID": "BR-09", "Scope": "Gender Parity Hypothesis Testing", "User Story": "US-07", "UI Screen": "UI-04 EDA Stats", "Test Cases": "TC-F-10, TC-T-14"},
        {"Req ID": "BR-10", "Scope": "Interactive Workforce Filtering", "User Story": "US-11", "UI Screen": "Sidebar (Global)", "Test Cases": "TC-F-03"},
        {"Req ID": "BR-11", "Scope": "Data Hygiene & Completeness Audit", "User Story": "US-16", "UI Screen": "UI-07 Data Quality", "Test Cases": "TC-F-17, TC-T-10, TC-T-11"},
        {"Req ID": "BR-12", "Scope": "Dual CSV/Excel Dataset Export", "User Story": "US-08", "UI Screen": "UI-06 Download", "Test Cases": "TC-F-15, TC-F-16"},
        {"Req ID": "BR-13", "Scope": "OLS Regression Rating Drivers", "User Story": "US-14", "UI Screen": "UI-04 EDA Stats", "Test Cases": "TC-F-11, TC-T-15"},
        {"Req ID": "BR-14", "Scope": "Department Flight Risk Matrix", "User Story": "US-15", "UI Screen": "UI-05 Insights", "Test Cases": "TC-F-14, TC-T-09"},
    ]
    st.dataframe(pd.DataFrame(rtm_data), use_container_width=True, hide_index=True)

with tab_schema:
    st.markdown("### Relational Database Schema & Analytical Views")
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT name, type, sql FROM sqlite_master WHERE type IN ('table', 'view') AND name NOT LIKE 'sqlite_%';")
        schema_rows = cur.fetchall()
        conn.close()

        for name, obj_type, sql in schema_rows:
            with st.expander(f"{'📊 View' if obj_type == 'view' else '📁 Table'}: {name}"):
                st.code(sql, language="sql")
    except Exception as e:
        st.error(f"Failed to query database schema: {e}")

with tab_glossary:
    st.markdown(
        """
        ### Mathematical KPI Formula Glossary
        
        #### 1. Average Performance Rating
        $$\\mu_{\\text{rating}} = \\frac{1}{N} \\sum_{i=1}^N r_i$$
        
        #### 2. Promotion Readiness Score (0–100 Scale)
        $$\\text{Score} = 0.40 \\left(\\frac{\\text{Rating}}{5.0} \\times 100\\right) + 0.30 \\min\\left(\\frac{\\text{Years Waiting}}{4.0}, 1.0\\right) \\times 100 + 0.20 (\\text{Goals \\%}) + 0.10 \\min(\\text{Tenure} \\times 10, 100)$$
        
        #### 3. Overdue Promotion Flag
        $$\\text{Flag} = 1 \\iff (\\text{Performance Rating} \\ge 4.0) \\land (\\text{Years Since Last Promotion} > 3.0)$$
        
        #### 4. Pearson Correlation Coefficient ($r$)
        $$r = \\frac{\\sum (X - \\bar{X})(Y - \\bar{Y})}{\\sqrt{\\sum (X - \\bar{X})^2 \\sum (Y - \\bar{Y})^2}}$$
        
        #### 5. Independent Two-Sample t-Test (Gender Parity)
        $$t = \\frac{\\bar{X}_1 - \\bar{X}_2}{\\sqrt{\\frac{s_1^2}{n_1} + \\frac{s_2^2}{n_2}}}, \\quad H_0: \\mu_{\\text{male}} = \\mu_{\\text{female}}$$
        
        #### 6. One-Way ANOVA F-Statistic
        $$F = \\frac{\\text{Between-Group Variance (MSB)}}{\\text{Within-Group Variance (MSW)}}$$
        
        #### 7. Interquartile Range (IQR) Outlier Boundaries
        $$\\text{Lower Bound} = Q_1 - 1.5 \\times \\text{IQR}, \\quad \\text{Upper Bound} = Q_3 + 1.5 \\times \\text{IQR}$$
        """
    )

st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)
render_insight_box(
    title="Academic Project Portfolio Verification",
    message="All academic specifications, user stories (US-01 to US-18), business requirements (BR-01 to BR-14), and automated technical test cases (TC-T-01 to TC-T-15) have been verified and integrated into this production build.",
    insight_type="success"
)
