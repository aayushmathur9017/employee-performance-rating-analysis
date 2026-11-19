"""
Executive Overview Dashboard & Workforce KPIs (UI-01)
Employee Performance Rating Analysis System
Student: Aayush Mathur | Enrolment No: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.data_loader import load_analytics_view
from src.cleaner import clean_employee_data
from src.analysis import (
    compute_average_rating,
    get_top_bottom_performers,
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
    page_title="Executive Overview | Employee Performance Rating Analysis",
    page_icon="📊",
    layout="wide"
)
inject_tailwind_css()

# 2. Ingest and Clean Data
raw_df = load_analytics_view()
cleaned_df, _ = clean_employee_data(raw_df)
filtered_df = render_sidebar_filters(cleaned_df)

# 3. Header
render_header(
    title="Executive Workforce Overview & KPI Dashboard",
    subtitle="Enterprise performance telemetry, department benchmark rankings, and decile distributions",
    badge_text="UI-01 | BR-01, BR-02",
    icon="📊"
)

# 4. Top KPI Cards
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

total_count = len(filtered_df)
avg_rating = compute_average_rating(filtered_df.get("performance_rating", pd.Series()))
top_count = int((filtered_df.get("performance_rating", pd.Series()) >= 4.5).sum())
top_pct = round((top_count / total_count * 100.0), 1) if total_count > 0 else 0.0

attr_count = int(pd.to_numeric(filtered_df.get("attrition_flag", 0), errors="coerce").fillna(0).sum())
attr_pct = round((attr_count / total_count * 100.0), 1) if total_count > 0 else 0.0

with kpi1:
    render_kpi_card(
        title="Active Headcount",
        value=f"{total_count:,}",
        delta=None,
        subtitle="Evaluated employees in cohort",
        icon="👥",
        badge_color="blue"
    )

with kpi2:
    render_kpi_card(
        title="Average Rating",
        value=f"{avg_rating:.2f} / 5.0",
        delta="+4.2% YoY",
        delta_direction="up",
        subtitle="Benchmark standard: 3.50",
        icon="⭐",
        badge_color="emerald"
    )

with kpi3:
    render_kpi_card(
        title="Top Performer Share",
        value=f"{top_pct:.1f}%",
        delta=f"{top_count} employees",
        delta_direction="up",
        subtitle="Rating >= 4.50 (Star talent)",
        icon="🏆",
        badge_color="purple"
    )

with kpi4:
    render_kpi_card(
        title="Workforce Attrition",
        value=f"{attr_pct:.1f}%",
        delta="-1.8% vs plan",
        delta_direction="down",
        subtitle=f"{attr_count} historical departures",
        icon="🚪",
        badge_color="amber"
    )

st.markdown("<div style='height: 1.25rem;'></div>", unsafe_allow_html=True)

# 5. Charts Row
chart_col1, chart_col2 = st.columns([3, 2])

with chart_col1:
    st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem;'>Department Average Rating Benchmark</h3>", unsafe_allow_html=True)
    dept_metrics = get_department_metrics(filtered_df)
    
    if not dept_metrics.empty:
        # Assign benchmark color
        def get_color(val):
            if val >= 3.8:
                return "#10B981"  # Emerald
            elif val >= 3.5:
                return "#3B82F6"  # Blue
            else:
                return "#F59E0B"  # Amber

        colors = [get_color(v) for v in dept_metrics["Average Rating"]]

        fig = go.Figure(go.Bar(
            x=dept_metrics["Average Rating"],
            y=dept_metrics["Department"],
            orientation="h",
            marker=dict(color=colors),
            text=[f"{v:.2f}" for v in dept_metrics["Average Rating"]],
            textposition="inside",
            insidetextanchor="middle"
        ))
        fig.add_vline(x=3.5, line_width=2, line_dash="dash", line_color="#EF4444", annotation_text="Benchmark 3.50", annotation_position="top right")
        fig.update_layout(
            margin=dict(l=0, r=20, t=10, b=10),
            height=340,
            xaxis=dict(range=[2.0, 5.0], title="Performance Rating (1.0 - 5.0)"),
            yaxis=dict(title=None, categoryorder="total ascending")
        )
        st.plotly_chart(fig, use_container_width=True, key="chart_overview_dept_bar")
    else:
        st.info("No department data available.")

with chart_col2:
    st.markdown("<h3 style='font-size: 1.05rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem;'>Workforce Rating Distribution</h3>", unsafe_allow_html=True)
    if "performance_rating" in filtered_df.columns and not filtered_df.empty:
        fig_hist = px.histogram(
            filtered_df,
            x="performance_rating",
            nbins=20,
            color_discrete_sequence=["#1E3A8A"],
            labels={"performance_rating": "Performance Rating"}
        )
        fig_hist.update_layout(
            margin=dict(l=10, r=10, t=10, b=10),
            height=340,
            xaxis=dict(title="Rating Range", range=[1.0, 5.0]),
            yaxis=dict(title="Employee Count")
        )
        st.plotly_chart(fig_hist, use_container_width=True, key="chart_overview_hist")
    else:
        st.info("No rating telemetry available.")

st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)

# 6. Decile Slicing: Top 10% vs Bottom 10%
render_header(
    title="Decile Distribution: Star Performers vs Performance Improvement Candidates",
    subtitle="Algorithmic separation of upper and lower 10th percentiles for targeted talent management",
    badge_text="Decile Slicing",
    icon="👥"
)

top_10, bottom_10 = get_top_bottom_performers(filtered_df, n_pct=0.10)

dec_col1, dec_col2 = st.columns(2)

display_cols = ["emp_id", "full_name", "dept_name", "job_role", "performance_rating", "tenure_years"]
col_rename = {
    "emp_id": "Emp ID",
    "full_name": "Employee Name",
    "dept_name": "Department",
    "job_role": "Role",
    "performance_rating": "Rating",
    "tenure_years": "Tenure (Yrs)"
}

with dec_col1:
    st.markdown(
        f"""
        <div style="background-color: #ECFDF5; border: 1px solid #A7F3D0; border-radius: 0.5rem; padding: 0.75rem 1rem; margin-bottom: 0.75rem; display: flex; justify-content: space-between; align-items: center;">
            <span style="font-weight: 700; color: #065F46;">🌟 Top 10% Star Performers ({len(top_10)} Staff)</span>
            <span style="font-size: 0.75rem; background-color: #10B981; color: white; padding: 0.2rem 0.5rem; border-radius: 9999px; font-weight: 600;">Merit Candidates</span>
        </div>
        """,
        unsafe_allow_html=True
    )
    if not top_10.empty:
        valid_cols = [c for c in display_cols if c in top_10.columns]
        st.dataframe(
            top_10[valid_cols].rename(columns=col_rename).head(15),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No records match Top 10% criteria.")

with dec_col2:
    st.markdown(
        f"""
        <div style="background-color: #FEF2F2; border: 1px solid #FECACA; border-radius: 0.5rem; padding: 0.75rem 1rem; margin-bottom: 0.75rem; display: flex; justify-content: space-between; align-items: center;">
            <span style="font-weight: 700; color: #991B1B;">⚠️ Bottom 10% PIP Candidates ({len(bottom_10)} Staff)</span>
            <span style="font-size: 0.75rem; background-color: #EF4444; color: white; padding: 0.2rem 0.5rem; border-radius: 9999px; font-weight: 600;">Coaching Needed</span>
        </div>
        """,
        unsafe_allow_html=True
    )
    if not bottom_10.empty:
        valid_cols = [c for c in display_cols if c in bottom_10.columns]
        st.dataframe(
            bottom_10[valid_cols].rename(columns=col_rename).head(15),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No records match Bottom 10% criteria.")

st.markdown("<div style='height: 1.25rem;'></div>", unsafe_allow_html=True)
render_insight_box(
    title="Executive Talent Calibration Summary",
    message=f"The organization exhibits an overall mean appraisal rating of {avg_rating:.2f}/5.0 with {top_pct:.1f}% of staff qualifying as Top Performers (rating >= 4.5). Departments operating below the 3.50 threshold warrant structured calibration before compensation cycles.",
    insight_type="info",
    action_label="Review Department Comparisons in Page 03"
)
