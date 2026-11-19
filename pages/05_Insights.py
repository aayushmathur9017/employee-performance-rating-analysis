"""
Actionable Insights, Promotion Readiness & Retention Risk (UI-05)
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
    identify_overdue_promotions,
    compute_department_attrition_risk,
    calculate_promotion_scores,
)
from src.ui_components import (
    inject_tailwind_css,
    render_header,
    render_kpi_card,
    render_insight_box,
    render_sidebar_filters,
    render_alert,
)

# 1. Page Configuration
st.set_page_config(
    page_title="Actionable Insights | Employee Performance Rating Analysis",
    page_icon="💡",
    layout="wide"
)
inject_tailwind_css()

# 2. Ingest and Clean Data
raw_df = load_analytics_view()
cleaned_df, _ = clean_employee_data(raw_df)
filtered_df = render_sidebar_filters(cleaned_df)

# Ensure promotion_readiness_score column exists
if "promotion_readiness_score" not in filtered_df.columns:
    filtered_df["promotion_readiness_score"] = calculate_promotion_scores(filtered_df)

# 3. Header
render_header(
    title="Prescriptive Talent Insights: Promotion Readiness & Flight Risk",
    subtitle="Algorithmic candidate scoring (0–100), overdue promotion alerts, and departmental attrition flight risk matrix",
    badge_text="UI-05 | BR-06, BR-07, BR-14",
    icon="💡"
)

# 4. Prescriptive Telemetry Cards
overdue_df = identify_overdue_promotions(filtered_df)
promo_candidates = filtered_df[filtered_df["promotion_readiness_score"] >= 80.0]
attr_risk_df = compute_department_attrition_risk(filtered_df)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    render_kpi_card(
        title="Promo Candidates",
        value=f"{len(promo_candidates)} Staff",
        delta="Score >= 80.0",
        delta_direction="up",
        subtitle="Eligible for advancement",
        icon="🚀",
        badge_color="emerald"
    )

with kpi2:
    overdue_count = len(overdue_df)
    render_kpi_card(
        title="Overdue Promotions",
        value=f"{overdue_count} Staff",
        delta="Action Required",
        delta_direction="down" if overdue_count > 0 else "neutral",
        subtitle="Rating >= 4.0 & > 3 yrs waiting",
        icon="⏳",
        badge_color="red" if overdue_count > 0 else "emerald"
    )

with kpi3:
    avg_promo_score = float(filtered_df["promotion_readiness_score"].mean()) if not filtered_df.empty else 0.0
    render_kpi_card(
        title="Mean Promo Score",
        value=f"{avg_promo_score:.1f} / 100",
        delta="Enterprise benchmark",
        subtitle="Composite weighted score",
        icon="📈",
        badge_color="blue"
    )

with kpi4:
    high_risk_depts = (attr_risk_df["Risk Tier"] == "High Risk").sum() if not attr_risk_df.empty else 0
    render_kpi_card(
        title="High Flight Risk Units",
        value=f"{high_risk_depts} Depts",
        delta="Retention concern",
        delta_direction="down" if high_risk_depts > 0 else "neutral",
        subtitle="Tier threshold >= 0.65",
        icon="🚪",
        badge_color="amber"
    )

st.markdown("<div style='height: 1.25rem;'></div>", unsafe_allow_html=True)

# 5. Overdue Promotion Urgent Alert Banner
if overdue_count > 0:
    render_alert(
        message=f"CRITICAL TALENT RISK: {overdue_count} high-performing employees have remained in their current grade for over 3.0 years with appraisal ratings >= 4.0. High likelihood of competitor poaching.",
        alert_type="danger"
    )

# 6. Tabbed Presentation: Promotion Engine vs Attrition Matrix
tab_promo, tab_attr = st.tabs(["🚀 Promotion Readiness Engine", "🛡️ Department Attrition Risk Matrix"])

with tab_promo:
    col_p1, col_p2 = st.columns([3, 2])
    with col_p1:
        st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem;'>Promotion Readiness Score Distribution (0–100 Scale)</h3>", unsafe_allow_html=True)
        fig_p = px.histogram(
            filtered_df,
            x="promotion_readiness_score",
            nbins=25,
            color_discrete_sequence=["#3B82F6"],
            labels={"promotion_readiness_score": "Promotion Readiness Score"}
        )
        fig_p.add_vline(x=80.0, line_width=2, line_dash="dash", line_color="#10B981", annotation_text="Threshold (80.0)")
        fig_p.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=10, b=10),
            xaxis=dict(range=[0, 100])
        )
        st.plotly_chart(fig_p, use_container_width=True, key="insights_promo_hist")

    with col_p2:
        st.markdown(
            """
            <div style="background-color: white; border: 1px solid #E2E8F0; border-radius: 0.75rem; padding: 1.25rem;">
                <div style="font-weight: 700; color: #1E3A8A; font-size: 0.95rem; margin-bottom: 0.5rem;">Promotion Scoring Methodology</div>
                <div style="font-size: 0.825rem; color: #475569; line-height: 1.5;">
                    The Promotion Readiness Score is formulated as a multi-factor composite index:
                    <ul style="margin-top: 0.4rem; padding-left: 1.2rem;">
                        <li><strong>40% Performance Rating</strong> (Normalized 1-5 scale)</li>
                        <li><strong>30% Role Stagnation</strong> (Capped at 4 years waiting)</li>
                        <li><strong>20% KPI / Goal Completion %</strong></li>
                        <li><strong>10% Company Tenure</strong> (Scaled to 10 years)</li>
                    </ul>
                    Staff achieving scores >= 80.0 qualify for immediate talent committee review.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-top: 1rem; margin-bottom: 0.5rem;'>Top Promotion Ready Candidates (Score >= 80.0)</h3>", unsafe_allow_html=True)
    p_cols = ["emp_id", "full_name", "dept_name", "job_role", "manager_name", "performance_rating", "years_since_last_promotion", "promotion_readiness_score"]
    valid_p_cols = [c for c in p_cols if c in promo_candidates.columns]
    
    if not promo_candidates.empty:
        st.dataframe(
            promo_candidates[valid_p_cols].sort_values(by="promotion_readiness_score", ascending=False).head(20),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No candidates exceed the 80.0 promotion readiness score threshold.")

    if not overdue_df.empty:
        st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #991B1B; margin-top: 1.5rem; margin-bottom: 0.5rem;'>⚠️ Stagnant High-Performer Alert Roster</h3>", unsafe_allow_html=True)
        st.dataframe(
            overdue_df[valid_p_cols].head(15),
            use_container_width=True,
            hide_index=True
        )

with tab_attr:
    st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem;'>Department Attrition & Flight Risk Scorecard</h3>", unsafe_allow_html=True)
    if not attr_risk_df.empty:
        st.dataframe(
            attr_risk_df,
            use_container_width=True,
            hide_index=True
        )

        fig_risk = px.bar(
            attr_risk_df,
            x="Department",
            y="Avg Risk Score",
            color="Risk Tier",
            color_discrete_map={"Low Risk": "#10B981", "Medium Risk": "#F59E0B", "High Risk": "#EF4444"},
            text="Avg Risk Score"
        )
        fig_risk.update_layout(
            height=340,
            margin=dict(l=10, r=10, t=10, b=10),
            yaxis=dict(range=[0, 1.0], title="Composite Flight Risk Score")
        )
        fig_risk.update_traces(texttemplate="%{text:.3f}", textposition="outside")
        st.plotly_chart(fig_risk, use_container_width=True, key="insights_flight_risk_bar")
    else:
        st.info("Attrition telemetry not available.")

st.markdown("<div style='height: 1.25rem;'></div>", unsafe_allow_html=True)

# 7. One-Click CSV Export for Talent Committee
col_exp1, col_exp2 = st.columns([3, 1])
with col_exp1:
    st.markdown("<div style='font-size: 0.85rem; color: #64748B;'>Generate an executive-ready briefing dossier for the quarterly Talent Succession & Merit Committee review.</div>", unsafe_allow_html=True)

with col_exp2:
    if not promo_candidates.empty:
        csv_bytes = promo_candidates.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download Talent Committee CSV",
            data=csv_bytes,
            file_name="Talent_Committee_Promotion_Roster.csv",
            mime="text/csv",
            use_container_width=True
        )
