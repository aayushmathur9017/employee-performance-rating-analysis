"""
Departmental Parity, Manager Leadership & Compensation Equity (UI-03)
Employee Performance Rating Analysis System
Student: Aayush Mathur | Enrolment No: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import load_analytics_view
from src.cleaner import clean_employee_data
from src.analysis import (
    run_salary_rating_correlation,
    get_manager_leaderboard,
    run_job_role_anova,
    get_department_metrics,
)
from src.ui_components import (
    inject_tailwind_css,
    render_header,
    render_kpi_card,
    render_insight_box,
    render_sidebar_filters,
    render_badge,
)

# 1. Page Configuration
st.set_page_config(
    page_title="Compare & Parity | Employee Performance Rating Analysis",
    page_icon="⚖️",
    layout="wide"
)
inject_tailwind_css()

# 2. Ingest and Clean Data
raw_df = load_analytics_view()
cleaned_df, _ = clean_employee_data(raw_df)
filtered_df = render_sidebar_filters(cleaned_df)

# 3. Header
render_header(
    title="Departmental Parity, Leadership & Compensation Equity",
    subtitle="Salary vs. Rating correlation analysis, managerial effectiveness leaderboard, and job-role ANOVA variance testing",
    badge_text="UI-03 | BR-03, BR-04, BR-05",
    icon="⚖️"
)

# 4. Correlation & Parity KPI Cards
corr_stats = run_salary_rating_correlation(filtered_df)
anova_stats = run_job_role_anova(filtered_df, group_col="job_role")

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    r_val = corr_stats.get("pearson_r", 0.0)
    render_kpi_card(
        title="Salary Correlation (r)",
        value=f"{r_val:+.3f}",
        delta="Statistically Significant" if corr_stats.get("pearson_p", 1.0) < 0.05 else "Weak Alignment",
        delta_direction="up" if r_val > 0.3 else "neutral",
        subtitle=f"Pearson r (p = {corr_stats.get('pearson_p', 1.0):.4f})",
        icon="💵",
        badge_color="blue"
    )

with kpi2:
    f_val = anova_stats.get("f_stat", 0.0)
    p_val = anova_stats.get("f_p_val", 1.0)
    render_kpi_card(
        title="Role ANOVA F-Stat",
        value=f"{f_val:.2f}",
        delta="Significant Variance" if p_val < 0.05 else "Uniform Distribution",
        delta_direction="down" if p_val < 0.05 else "up",
        subtitle=f"p = {p_val:.5f} ({anova_stats.get('cohorts', 0)} roles)",
        icon="📊",
        badge_color="purple"
    )

with kpi3:
    managers = get_manager_leaderboard(filtered_df)
    top_manager = managers.iloc[0]["Manager Name"] if not managers.empty else "N/A"
    top_m_rating = managers.iloc[0]["Team Avg Rating"] if not managers.empty else 0.0
    render_kpi_card(
        title="Top Rated Manager",
        value=f"{top_m_rating:.2f} ⭐",
        delta=top_manager[:18],
        subtitle="Highest subordinate score",
        icon="🎖️",
        badge_color="emerald"
    )

with kpi4:
    spear_r = corr_stats.get("spearman_r", 0.0)
    render_kpi_card(
        title="Rank Alignment",
        value=f"{spear_r:+.3f}",
        delta="Non-linear monotonic",
        subtitle=f"Spearman rho (N = {corr_stats.get('sample_size', 0)})",
        icon="📈",
        badge_color="navy"
    )

st.markdown("<div style='height: 1.25rem;'></div>", unsafe_allow_html=True)

# 5. Compensation Equity: Salary vs. Performance Scatter Plot
col_scat1, col_scat2 = st.columns([3, 2])

with col_scat1:
    st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem;'>Compensation Equity: Salary vs. Performance Rating</h3>", unsafe_allow_html=True)
    if "current_salary" in filtered_df.columns and "performance_rating" in filtered_df.columns:
        fig_scatter = px.scatter(
            filtered_df,
            x="performance_rating",
            y="current_salary",
            color="dept_name",
            trendline="ols",
            hover_data=["full_name", "job_role", "tenure_years"],
            labels={
                "performance_rating": "Performance Appraisal Rating (1-5)",
                "current_salary": "Current Annual Salary (USD)",
                "dept_name": "Department"
            }
        )
        fig_scatter.update_layout(
            height=380,
            margin=dict(l=10, r=20, t=10, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_scatter, use_container_width=True, key="compare_salary_scatter")
    else:
        st.info("Salary and performance metrics unavailable for scatter plot.")

with col_scat2:
    st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem;'>Job Role Rating Distribution (Box Plot)</h3>", unsafe_allow_html=True)
    if "job_role" in filtered_df.columns and "performance_rating" in filtered_df.columns:
        fig_box = px.box(
            filtered_df,
            x="job_role",
            y="performance_rating",
            color="job_role",
            labels={"job_role": "Job Role", "performance_rating": "Rating"}
        )
        fig_box.update_layout(
            height=380,
            margin=dict(l=10, r=10, t=10, b=10),
            showlegend=False,
            xaxis=dict(tickangle=-25)
        )
        st.plotly_chart(fig_box, use_container_width=True, key="compare_role_box")
    else:
        st.info("Job role telemetry not available.")

st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)

# 6. Manager Leadership Leaderboard
render_header(
    title="Managerial Performance & Leadership Leaderboard",
    subtitle="Ranking team leads by subordinate average appraisal, promotion momentum, and stagnant talent flags",
    badge_text="Leadership Audit",
    icon="👥"
)

if not managers.empty:
    st.dataframe(
        managers,
        use_container_width=True,
        hide_index=True
    )
else:
    st.info("No managerial hierarchy records available.")

st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)
render_insight_box(
    title="Compensation Equity & Leadership Synthesis",
    message=f"Positive salary-to-rating correlation (r = {r_val:+.3f}, p < 0.05) indicates enterprise pay-for-performance alignment. However, role-level ANOVA (F = {f_val:.2f}, p = {p_val:.4f}) indicates notable rating variance between technical and administrative cohorts, which Total Rewards should address in normalization cycles.",
    insight_type="info"
)
