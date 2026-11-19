"""
Self-Service Data Explorer & Report Export (UI-06)
Employee Performance Rating Analysis System
Student: Aayush Mathur | Enrolment No: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0
"""

import io
import streamlit as st
import pandas as pd

from src.data_loader import load_analytics_view
from src.cleaner import clean_employee_data
from src.ui_components import (
    inject_tailwind_css,
    render_header,
    render_kpi_card,
    render_sidebar_filters,
    render_alert,
)

# 1. Page Configuration
st.set_page_config(
    page_title="Data Explorer & Export | Employee Performance Rating Analysis",
    page_icon="📥",
    layout="wide"
)
inject_tailwind_css()

# 2. Ingest and Clean Data
raw_df = load_analytics_view()
cleaned_df, _ = clean_employee_data(raw_df)
filtered_df = render_sidebar_filters(cleaned_df)

# 3. Header
render_header(
    title="Self-Service Workforce Data Explorer & Export Engine",
    subtitle="Interactive multi-dimensional query grid, dynamic column projection, and dual CSV / Excel report generation",
    badge_text="UI-06 | BR-10, BR-12",
    icon="📥"
)

# 4. Search and Column Projection Controls
ctrl_col1, ctrl_col2 = st.columns([2, 3])

with ctrl_col1:
    search_term = st.text_input("🔍 Free-Text Filter (Search by Name, Role, or ID):", placeholder="e.g. Engineer, John, Sales...")

# Apply search filter
display_df = filtered_df.copy()
if search_term:
    search_lower = search_term.lower()
    mask = display_df.astype(str).apply(lambda row: row.str.lower().str.contains(search_lower).any(), axis=1)
    display_df = display_df[mask]

all_columns = display_df.columns.tolist()
default_cols = [
    c for c in [
        "emp_id", "full_name", "dept_name", "job_role", "gender", "age",
        "current_salary", "tenure_years", "performance_rating", "rating_category",
        "years_since_last_promotion", "promotion_readiness_score", "attrition_type"
    ] if c in all_columns
]

with ctrl_col2:
    selected_cols = st.multiselect(
        "Projected Columns to Display & Export:",
        options=all_columns,
        default=default_cols if default_cols else all_columns[:10]
    )

if not selected_cols:
    selected_cols = default_cols if default_cols else all_columns[:10]

view_df = display_df[selected_cols]

# 5. Telemetry & Download Action Bar
st.markdown("<div style='height: 0.5rem;'></div>", unsafe_allow_html=True)
act_col1, act_col2, act_col3 = st.columns([2, 1, 1])

with act_col1:
    st.markdown(f"<div style='font-size: 0.9rem; color: #1E293B; padding-top: 0.4rem;'>Displaying <strong>{len(view_df):,}</strong> records across <strong>{len(selected_cols)}</strong> attributes</div>", unsafe_allow_html=True)

# Generate CSV bytes
csv_data = view_df.to_csv(index=False).encode("utf-8")

# Generate Excel bytes using openpyxl
excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    view_df.to_excel(writer, index=False, sheet_name="Performance_Data")
excel_data = excel_buffer.getvalue()

with act_col2:
    st.download_button(
        label="📄 Export to CSV",
        data=csv_data,
        file_name="employee_performance_dataset.csv",
        mime="text/csv",
        use_container_width=True
    )

with act_col3:
    st.download_button(
        label="📊 Export to Excel (.xlsx)",
        data=excel_data,
        file_name="employee_performance_dataset.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("<div style='height: 0.75rem;'></div>", unsafe_allow_html=True)

# 6. Interactive Data Grid
st.dataframe(
    view_df,
    use_container_width=True,
    hide_index=True
)

st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)

# 7. Aggregation Pivot Builder
render_header(
    title="Dynamic Cross-Tabulation & Aggregation Pivot",
    subtitle="Summarize key workforce metrics on the fly by custom dimensions",
    badge_text="Pivot Builder",
    icon="🔄"
)

piv_c1, piv_c2, piv_c3 = st.columns(3)
with piv_c1:
    pivot_dim = st.selectbox("Group By Dimension:", options=["dept_name", "job_role", "gender", "age_group", "salary_band", "rating_category"])

with piv_c2:
    pivot_metric = st.selectbox("Aggregate Metric:", options=["performance_rating", "current_salary", "tenure_years", "promotion_readiness_score"])

with piv_c3:
    pivot_agg = st.selectbox("Aggregation Method:", options=["mean", "median", "count", "sum"])

if pivot_dim in filtered_df.columns and pivot_metric in filtered_df.columns:
    if pivot_agg == "mean":
        pivot_res = filtered_df.groupby(pivot_dim)[pivot_metric].mean().round(2).reset_index()
    elif pivot_agg == "median":
        pivot_res = filtered_df.groupby(pivot_dim)[pivot_metric].median().round(2).reset_index()
    elif pivot_agg == "count":
        pivot_res = filtered_df.groupby(pivot_dim)[pivot_metric].count().reset_index()
    else:
        pivot_res = filtered_df.groupby(pivot_dim)[pivot_metric].sum().round(2).reset_index()

    pivot_res.columns = [pivot_dim.replace("_", " ").title(), f"{pivot_agg.title()} of {pivot_metric.replace('_', ' ').title()}"]
    st.dataframe(pivot_res, use_container_width=True, hide_index=True)
