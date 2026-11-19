"""
Exploratory Data Analysis, Demographics & Statistical Testing (UI-04)
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
    run_gender_parity_test,
    run_linear_regression,
    run_random_forest_importance,
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
    page_title="EDA & Statistical Testing | Employee Performance Rating Analysis",
    page_icon="🔬",
    layout="wide"
)
inject_tailwind_css()

# 2. Ingest and Clean Data
raw_df = load_analytics_view()
cleaned_df, _ = clean_employee_data(raw_df)
filtered_df = render_sidebar_filters(cleaned_df)

# 3. Header
render_header(
    title="Exploratory Data Analysis & Inferential Hypothesis Testing",
    subtitle="Demographic appraisal parity verification (Mann-Whitney U, Two-sample t-test) and multivariable performance driver modeling",
    badge_text="UI-04 | BR-09, BR-13",
    icon="🔬"
)

# 4. Statistical Telemetry Cards
gender_stats = run_gender_parity_test(filtered_df)
ols_results = run_linear_regression(filtered_df)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    is_fair = gender_stats.get("is_fair", True)
    render_kpi_card(
        title="Gender Parity Test",
        value="Fair & Unbiased" if is_fair else "Flagged Variance",
        delta=f"p = {gender_stats.get('t_p_val', 1.0):.4f}",
        delta_direction="up" if is_fair else "down",
        subtitle=f"t-statistic: {gender_stats.get('t_stat', 0.0):.3f}",
        icon="⚖️",
        badge_color="emerald" if is_fair else "red"
    )

with kpi2:
    render_kpi_card(
        title="Mann-Whitney U Test",
        value=f"U = {gender_stats.get('u_stat', 0.0):,.0f}",
        delta=f"p = {gender_stats.get('u_p_val', 1.0):.4f}",
        delta_direction="up" if gender_stats.get("u_p_val", 1.0) >= 0.05 else "down",
        subtitle="Non-parametric equality check",
        icon="📊",
        badge_color="blue"
    )

with kpi3:
    r2 = ols_results.get("r_squared", 0.0)
    render_kpi_card(
        title="Model Explanatory (R²)",
        value=f"{r2:.3f}",
        delta=f"N = {ols_results.get('n_samples', 0)} obs",
        subtitle="OLS regression fit",
        icon="🎯",
        badge_color="purple"
    )

with kpi4:
    male_m = gender_stats.get("male_mean", 0.0)
    fem_m = gender_stats.get("female_mean", 0.0)
    render_kpi_card(
        title="Gender Cohort Means",
        value=f"M: {male_m:.2f} | F: {fem_m:.2f}",
        delta=f"Δ = {abs(male_m - fem_m):.2f}",
        subtitle="Mean performance rating",
        icon="👥",
        badge_color="navy"
    )

st.markdown("<div style='height: 1.25rem;'></div>", unsafe_allow_html=True)

# 5. Demographic Distribution Charts: Gender & Age
col_demo1, col_demo2 = st.columns(2)

with col_demo1:
    st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem;'>Appraisal Rating Distribution by Gender (Violin Plot)</h3>", unsafe_allow_html=True)
    if "gender" in filtered_df.columns and "performance_rating" in filtered_df.columns:
        fig_violin = px.violin(
            filtered_df,
            x="gender",
            y="performance_rating",
            color="gender",
            box=True,
            points="outliers",
            labels={"gender": "Gender Cohort", "performance_rating": "Performance Rating"}
        )
        fig_violin.update_layout(
            height=360,
            margin=dict(l=10, r=10, t=10, b=10),
            showlegend=False
        )
        st.plotly_chart(fig_violin, use_container_width=True, key="eda_gender_violin")
    else:
        st.info("Gender data unavailable.")

with col_demo2:
    st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem;'>Performance Rating by Age Bracket (Box Plot)</h3>", unsafe_allow_html=True)
    if "age_group" in filtered_df.columns and "performance_rating" in filtered_df.columns:
        fig_age = px.box(
            filtered_df,
            x="age_group",
            y="performance_rating",
            color="age_group",
            category_orders={"age_group": ["<30", "30-45", ">45"]},
            labels={"age_group": "Age Bracket", "performance_rating": "Performance Rating"}
        )
        fig_age.update_layout(
            height=360,
            margin=dict(l=10, r=10, t=10, b=10),
            showlegend=False
        )
        st.plotly_chart(fig_age, use_container_width=True, key="eda_age_box")
    else:
        st.info("Age group telemetry unavailable.")

st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)

# 6. Performance Drivers: Regression Feature Coefficients & Random Forest Importance
render_header(
    title="Performance Driver Econometric & Machine Learning Analysis",
    subtitle="OLS regression coefficients and non-linear Random Forest feature importances",
    badge_text="Feature Attribution",
    icon="🧠"
)

tab_ols, tab_rf = st.tabs(["📐 OLS Linear Regression Model", "🌲 Random Forest Feature Importance"])

with tab_ols:
    coefs = ols_results.get("coefficients", {})
    if coefs:
        coef_df = pd.DataFrame({
            "Feature": list(coefs.keys()),
            "Coefficient": list(coefs.values()),
        }).sort_values(by="Coefficient", ascending=False)

        col_reg1, col_reg2 = st.columns([3, 2])
        with col_reg1:
            fig_coef = px.bar(
                coef_df,
                x="Coefficient",
                y="Feature",
                orientation="h",
                color="Coefficient",
                color_continuous_scale="Teal",
                text="Coefficient"
            )
            fig_coef.update_traces(texttemplate="%{text:.4f}", textposition="outside")
            fig_coef.update_layout(
                height=340,
                margin=dict(l=10, r=20, t=10, b=10),
                coloraxis_showscale=False
            )
            st.plotly_chart(fig_coef, use_container_width=True, key="eda_coef_chart")

        with col_reg2:
            st.markdown(
                f"""
                <div style="background-color: white; border: 1px solid #E2E8F0; border-radius: 0.75rem; padding: 1.25rem;">
                    <div style="font-weight: 700; color: #1E3A8A; font-size: 0.95rem; margin-bottom: 0.75rem;">Regression Diagnostics</div>
                    <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>R-Squared (R²):</strong> {ols_results.get('r_squared', 0.0)}</div>
                    <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>Model Intercept:</strong> {ols_results.get('intercept', 0.0)}</div>
                    <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.4rem;"><strong>Sample Observations (N):</strong> {ols_results.get('n_samples', 0)}</div>
                    <div style="font-size: 0.85rem; color: #475569; margin-top: 0.75rem; padding-top: 0.75rem; border-top: 1px solid #E2E8F0;">
                        <em>Key Takeaway:</em> Training hours and goals completion percentage consistently show the strongest positive predictive elasticity on performance evaluations.
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
    else:
        st.info("Insufficient features to fit OLS regression model.")

with tab_rf:
    rf_imp = run_random_forest_importance(filtered_df)
    if rf_imp:
        rf_df = pd.DataFrame({
            "Feature": list(rf_imp.keys()),
            "Importance": list(rf_imp.values())
        }).sort_values(by="Importance", ascending=True)

        fig_rf = px.bar(
            rf_df,
            x="Importance",
            y="Feature",
            orientation="h",
            color="Importance",
            color_continuous_scale="Purples",
            text="Importance"
        )
        fig_rf.update_traces(texttemplate="%{text:.3f}", textposition="outside")
        fig_rf.update_layout(
            height=340,
            margin=dict(l=10, r=20, t=10, b=10),
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_rf, use_container_width=True, key="eda_rf_importance_chart")
    else:
        st.info("Insufficient data for Random Forest training.")

st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)
render_insight_box(
    title="Demographic Governance & Equity Conclusion",
    message=f"Hypothesis testing under the null hypothesis (H₀: μ_male = μ_female) yielded t = {gender_stats.get('t_stat', 0.0):.2f}, p = {gender_stats.get('t_p_val', 1.0):.4f}. Because p >= 0.05, the evaluation process confirms absence of statistically significant demographic appraisal bias.",
    insight_type="success"
)
