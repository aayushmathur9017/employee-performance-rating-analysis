"""
UI Components & Master SaaS Dashboard Design System
Employee Performance Rating Analysis System
Student: Aayush Mathur | Roll No: 2500320140001 | Section: MCA-B | ABES Engineering College
Version: 8.0 (100% Visual Fidelity to Reference Screenshot + Full Real-Time Functionality)
"""

import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from typing import Optional, Union, Any, List, Dict
import uuid


def inject_custom_theme():
    """
    Injects custom Google Fonts, clean modern styling, and CSS overrides
    to match the enterprise reference dashboard:
    - Deep dark navy/charcoal sidebar (#111827 / #0F172A)
    - Clean white page background (#F8FAFC)
    - Bordered, rounded white metric cards (#FFFFFF)
    - No raw CSS leakages or blank top padding
    """
    css = """<style>
/* Eliminate top collapse button artifact ('keyboard_double.') without touching sidebar content */
[data-testid="stSidebarHeader"],
[data-testid="stSidebarCollapseButton"],
header [data-testid="stSidebarCollapseButton"],
section[data-testid="stSidebar"] [data-testid="stSidebarHeader"],
section[data-testid="stSidebar"] button[data-testid="stSidebarCollapseButton"] {
    display: none !important;
    visibility: hidden !important;
    height: 0px !important;
    padding: 0px !important;
    margin: 0px !important;
    overflow: hidden !important;
}

/* Ensure sidebar container and all content are ALWAYS visible, expanded and properly styled */
section[data-testid="stSidebar"] {
    background-color: #111827 !important;
    border-right: 1px solid #1F2937 !important;
    min-width: 275px !important;
    width: 285px !important;
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
}

section[data-testid="stSidebar"] > div {
    display: flex !important;
    visibility: visible !important;
    opacity: 1 !important;
}

section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
    padding-top: 1.5rem !important;
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
}

/* Base App Background */
.stApp {
    background-color: #F8FAFC !important;
}

/* Eliminate excessive top whitespace */
.main .block-container {
    padding-top: 1rem !important;
    padding-bottom: 2.5rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 1680px !important;
}

header[data-testid="stHeader"] {
    background: transparent !important;
    height: 1.5rem !important;
}

/* Hide default Streamlit multi-page nav list so only custom menu appears */
[data-testid="stSidebarNav"] {
    display: none !important;
}

/* Dark Navy Sidebar Styling */
section[data-testid="stSidebar"] {
    --primary-color: #3B82F6 !important;
    --background-color: #111827 !important;
    --secondary-background-color: #1F2937 !important;
    --text-color: #F8FAFC !important;
    background-color: #111827 !important;
    border-right: 1px solid #1F2937 !important;
}

section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #F8FAFC !important;
}

/* Sidebar Page Links High-Contrast Typography */
section[data-testid="stSidebar"] [data-testid="stPageLink"],
section[data-testid="stSidebar"] div[data-testid="stPageLink"] {
    display: block !important;
    width: 100% !important;
    margin: 0 0 0.15rem 0 !important;
    padding: 0 !important;
}

section[data-testid="stSidebar"] [data-testid="stPageLink"] a,
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"],
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] {
    display: flex !important;
    flex-direction: row !important;
    justify-content: flex-start !important;
    align-items: center !important;
    text-align: left !important;
    width: 100% !important;
    padding: 0.55rem 0.85rem !important;
    border-radius: 8px !important;
    background-color: transparent !important;
    text-decoration: none !important;
    gap: 0.75rem !important;
    transition: all 0.2s ease !important;
}

/* Explicit high-contrast white text for ALL children of stPageLink */
section[data-testid="stSidebar"] [data-testid="stPageLink"] a,
section[data-testid="stSidebar"] [data-testid="stPageLink"] a *,
section[data-testid="stSidebar"] [data-testid="stPageLink"] p,
section[data-testid="stSidebar"] [data-testid="stPageLink"] span,
section[data-testid="stSidebar"] [data-testid="stPageLink"] div,
section[data-testid="stSidebar"] [data-testid="stPageLink"] [data-testid="stMarkdownContainer"] p,
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] [data-testid="stMarkdownContainer"] p,
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] [data-testid="stMarkdownContainer"] p {
    color: #F8FAFC !important;
    font-weight: 600 !important;
    font-size: 0.92rem !important;
    line-height: 1.4 !important;
}

section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] > span:first-child {
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    width: 22px !important;
    min-width: 22px !important;
    font-size: 1.1rem !important;
    text-align: center !important;
    flex-shrink: 0 !important;
    margin: 0 !important;
}

section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] [data-testid="stMarkdownContainer"],
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] p,
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] div {
    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    text-align: left !important;
    flex-grow: 1 !important;
    width: auto !important;
    margin: 0 !important;
    padding: 0 !important;
}

section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]:hover,
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"]:hover,
section[data-testid="stSidebar"] [data-testid="stPageLink"] a:hover {
    background-color: #1F2937 !important;
}

section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]:hover *,
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"]:hover *,
section[data-testid="stSidebar"] [data-testid="stPageLink"] a:hover * {
    color: #FFFFFF !important;
}

section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"][aria-current="page"],
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"][aria-current="page"],
section[data-testid="stSidebar"] [data-testid="stPageLink"] a[aria-current="page"] {
    background: linear-gradient(90deg, rgba(37, 99, 235, 0.4) 0%, rgba(37, 99, 235, 0.12) 100%) !important;
    border-left: 3px solid #3B82F6 !important;
}

section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"][aria-current="page"] *,
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"][aria-current="page"] *,
section[data-testid="stSidebar"] [data-testid="stPageLink"] a[aria-current="page"] * {
    color: #FFFFFF !important;
    font-weight: 700 !important;
}

/* ==========================================================================
   SIDEBAR DATA FILTERS: HIGH-CONTRAST DARK INPUTS & TAGS (STREAMLIT 1.67)
   ========================================================================== */
section[data-testid="stSidebar"] label[data-testid="stWidgetLabel"],
section[data-testid="stSidebar"] div[data-testid="stSelectbox"] label,
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] label {
    color: #E2E8F0 !important;
    font-size: 0.84rem !important;
    font-weight: 600 !important;
    margin-bottom: 0.35rem !important;
}

/* Ensure ALL text inside selectbox and multiselect is pure white */
section[data-testid="stSidebar"] div[data-testid="stSelectbox"] *,
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] * {
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
}

/* Base input container boxes: dark charcoal with subtle border */
section[data-testid="stSidebar"] div[data-testid="stSelectbox"] [class*="equpb2c0"],
section[data-testid="stSidebar"] div[data-testid="stSelectbox"] div[role="combobox"],
section[data-testid="stSidebar"] div[data-testid="stSelectbox"] [data-baseweb="select"] > div,
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [class*="eo2qvtu0"],
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="select"] > div {
    background-color: #1F2937 !important;
    background: #1F2937 !important;
    border: 1px solid #374151 !important;
    border-radius: 8px !important;
}

/* Tags in multiselect */
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-tag],
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] span[data-tag],
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [class*="eo2qvtu3"],
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="tag"] {
    background-color: #2563EB !important;
    background: #2563EB !important;
    border: 1px solid #3B82F6 !important;
    border-radius: 6px !important;
}

section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-tag] *,
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [class*="eo2qvtu3"] * {
    background-color: transparent !important;
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
}

/* Icons */
section[data-testid="stSidebar"] div[data-testid="stSelectbox"] svg,
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] svg {
    fill: #94A3B8 !important;
    color: #94A3B8 !important;
}

section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-tag] svg {
    fill: #FFFFFF !important;
    color: #FFFFFF !important;
}

/* MultiSelect input placeholder */
section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] input::placeholder {
    color: #9CA3AF !important;
    -webkit-text-fill-color: #9CA3AF !important;
}

/* Dropdown popover menu styling */
div[data-testid="stSelectboxVirtualDropdown"],
div[data-testid="stMultiSelectDropdown"],
div[class*="equpb2c7"],
div[data-baseweb="popover"],
div[data-baseweb="popover"] [data-baseweb="menu"],
div[data-baseweb="popover"] ul,
div[data-baseweb="popover"] li {
    background-color: #1F2937 !important;
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
}

div[data-testid="stSelectboxVirtualDropdown"] [role="option"]:hover,
div[data-testid="stMultiSelectDropdown"] [role="option"]:hover,
div[data-baseweb="popover"] li:hover {
    background-color: #374151 !important;
    color: #FFFFFF !important;
}

div[data-testid="stSelectboxVirtualDropdown"] [role="option"][aria-selected="true"],
div[data-testid="stMultiSelectDropdown"] [role="option"][aria-selected="true"],
div[data-baseweb="popover"] li[aria-selected="true"] {
    background-color: #2563EB !important;
    color: #FFFFFF !important;
}

/* Dropdown popover menu styling */
div[data-testid="stSelectboxVirtualDropdown"],
div[data-testid="stMultiSelectDropdown"],
div[class*="equpb2c7"],
div[data-baseweb="popover"],
div[data-baseweb="popover"] [data-baseweb="menu"],
div[data-baseweb="popover"] ul,
div[data-baseweb="popover"] li {
    background-color: #1F2937 !important;
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
}

div[data-testid="stSelectboxVirtualDropdown"] [role="option"]:hover,
div[data-testid="stMultiSelectDropdown"] [role="option"]:hover,
div[data-baseweb="popover"] li:hover {
    background-color: #374151 !important;
    color: #FFFFFF !important;
}

div[data-testid="stSelectboxVirtualDropdown"] [role="option"][aria-selected="true"],
div[data-testid="stMultiSelectDropdown"] [role="option"][aria-selected="true"],
div[data-baseweb="popover"] li[aria-selected="true"] {
    background-color: #2563EB !important;
    color: #FFFFFF !important;
}

/* Dropdown Menu Options (Popover Container) */
div[data-baseweb="popover"],
div[data-baseweb="popover"] ul,
div[data-baseweb="menu"],
ul[role="listbox"] {
    background-color: #1F2937 !important;
    border: 1px solid #374151 !important;
    border-radius: 8px !important;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.4) !important;
}

li[role="option"],
div[role="option"] {
    background-color: #1F2937 !important;
    color: #F3F4F6 !important;
    font-size: 0.88rem !important;
    padding: 0.5rem 0.85rem !important;
}

li[role="option"]:hover,
li[role="option"][aria-selected="true"],
div[role="option"]:hover {
    background-color: #2563EB !important;
    color: #FFFFFF !important;
}

li[role="option"] * {
    color: inherit !important;
}

/* ==========================================================================
   STRICT LINE-BY-LINE LEFT-ALIGNMENT FOR SIDEBAR MAIN MENU
   ========================================================================== */
section[data-testid="stSidebar"] div[data-testid="stPageLink"],
section[data-testid="stSidebar"] [data-testid="stPageLink"] {
    display: block !important;
    width: 100% !important;
    text-align: left !important;
    margin: 0 0 0.15rem 0 !important;
    padding: 0 !important;
}

section[data-testid="stSidebar"] [data-testid="stPageLink"] a,
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"],
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] {
    display: flex !important;
    flex-direction: row !important;
    justify-content: flex-start !important;
    align-items: center !important;
    text-align: left !important;
    width: 100% !important;
    padding: 0.55rem 0.85rem !important;
    border-radius: 8px !important;
    background-color: transparent !important;
    color: #9CA3AF !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
    text-decoration: none !important;
    gap: 0.75rem !important;
    transition: all 0.2s ease !important;
}

section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]:hover,
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"]:hover {
    background-color: #1F2937 !important;
    color: #FFFFFF !important;
}

section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"][aria-current="page"],
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"][aria-current="page"] {
    background-color: #1F2937 !important;
    color: #FFFFFF !important;
    border-left: 3px solid #3B82F6 !important;
}

section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] > span:first-child {
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    width: 22px !important;
    min-width: 22px !important;
    font-size: 1.05rem !important;
    text-align: center !important;
    flex-shrink: 0 !important;
    margin: 0 !important;
}

section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] [data-testid="stMarkdownContainer"],
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] p,
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] div {
    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    text-align: left !important;
    flex-grow: 1 !important;
    width: auto !important;
    margin: 0 !important;
    padding: 0 !important;
    color: inherit !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"],
section[data-testid="stSidebar"] .stButton {
    width: 100% !important;
    margin-bottom: 0.2rem !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"] > button,
section[data-testid="stSidebar"] .stButton > button,
section[data-testid="stSidebar"] button {
    display: flex !important;
    flex-direction: row !important;
    justify-content: flex-start !important;
    align-items: center !important;
    text-align: left !important;
    width: 100% !important;
    padding: 0.55rem 0.95rem !important;
    border-radius: 8px !important;
    background-color: transparent !important;
    border: 1px solid transparent !important;
    color: #9CA3AF !important;
    font-weight: 500 !important;
    font-size: 0.88rem !important;
    transition: all 0.2s ease !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover,
section[data-testid="stSidebar"] .stButton > button:hover,
section[data-testid="stSidebar"] button:hover {
    background-color: #1F2937 !important;
    color: #FFFFFF !important;
    border-color: #374151 !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"] > button > div,
section[data-testid="stSidebar"] .stButton > button > div,
section[data-testid="stSidebar"] button > div,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
    display: flex !important;
    flex-direction: row !important;
    justify-content: flex-start !important;
    align-items: center !important;
    text-align: left !important;
    width: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"] > button p,
section[data-testid="stSidebar"] .stButton > button p,
section[data-testid="stSidebar"] button p {
    text-align: left !important;
    justify-content: flex-start !important;
    align-items: center !important;
    display: flex !important;
    width: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
    color: inherit !important;
    font-size: 0.88rem !important;
    font-weight: 500 !important;
}

/* Streamlit Container Card Polish */
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 14px !important;
    border: 1px solid #E2E8F0 !important;
    background: #FFFFFF !important;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.05) !important;
}

/* Blue Primary Button */
.stButton > button[kind="primary"], .ep-primary-btn {
    background-color: #2563EB !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.1rem !important;
    box-shadow: 0 1px 2px rgba(37, 99, 235, 0.2) !important;
    transition: background-color 0.2s ease !important;
}

.stButton > button[kind="primary"]:hover {
    background-color: #1D4ED8 !important;
}
</style>"""
    st.markdown(css, unsafe_allow_html=True)


def render_top_navbar():
    """
    Renders the top header bar with exact academic capstone title:
    - Left: 'Employee Performance Rating Analysis' with subtitle
    - Right: '● System Active' pill badge
    """
    col1, col2 = st.columns([3.5, 1])
    with col1:
        st.markdown(
            """
            <div>
                <h1 style="font-size: 1.55rem; font-weight: 800; color: #0F172A; margin: 0 0 0.15rem 0; letter-spacing: -0.02em;">
                    Employee Performance Rating Analysis
                </h1>
                <p style="font-size: 0.88rem; color: #64748B; margin: 0; font-weight: 400;">
                    Workforce appraisal ratings, longitudinal performance calibration & predictive analytics
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            """
            <div style="display: flex; justify-content: flex-end; align-items: center; height: 100%; padding-top: 0.35rem;">
                <div style="background: #ECFDF5; border: 1px solid #A7F3D0; color: #059669; font-size: 0.78rem; font-weight: 700;
                            padding: 0.35rem 0.85rem; border-radius: 9999px; display: inline-flex; align-items: center; gap: 0.45rem;">
                    <span style="display: inline-block; width: 8px; height: 8px; background-color: #10B981; border-radius: 50%;"></span>
                    System Active
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )


def render_overview_action_header(on_add_click_label: str = "+ Add Employee"):
    """
    Renders the sub-header row:
    - Left: 'Dashboard Overview' with subtitle 'Employee performance ratings at a glance'
    - Right: Blue '+ Add Employee' action button
    """
    col1, col2 = st.columns([3.5, 1])
    with col1:
        st.markdown(
            """
            <div style="margin-top: 1rem; margin-bottom: 0.85rem;">
                <h2 style="font-size: 1.35rem; font-weight: 700; color: #0F172A; margin: 0 0 0.15rem 0;">
                    Dashboard Overview
                </h2>
                <p style="font-size: 0.86rem; color: #64748B; margin: 0;">
                    Workforce performance ratings & appraisal analytics at a glance
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )
    with col2:
        st.markdown("<div style='height: 1.1rem;'></div>", unsafe_allow_html=True)
        return st.button(on_add_click_label, type="primary", use_container_width=True, key="btn_add_employee_top")


def render_clean_metric_card(
    title: str,
    value: Union[str, int, float],
    subtitle: str,
    icon: str,
    icon_color: str = "#2563EB",
    icon_bg: str = "#EFF6FF"
):
    """
    Renders the exact 4 metric cards from the reference screenshot:
    - Card 1: Total Employees (👥, 516, Employees registered)
    - Card 2: Average Rating (⭐, 3.5 / 5.0, Out of 5.0)
    - Card 3: Average Productivity (📈, 95.0%, Overall productivity)
    - Card 4: Average Attendance (📅, 97.0%, Overall attendance)
    """
    card_html = f"""
    <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 1.25rem 1.35rem;
                box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04); min-height: 135px; display: flex; flex-direction: column; justify-content: space-between;">
        <div style="display: flex; align-items: flex-start; gap: 0.75rem;">
            <div style="background: {icon_bg}; color: {icon_color}; font-size: 1.35rem; width: 36px; height: 36px;
                        border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                {icon}
            </div>
        </div>
        <div>
            <div style="font-size: 0.82rem; font-weight: 500; color: #64748B; margin-top: 0.5rem;">
                {title}
            </div>
            <div style="font-size: 1.75rem; font-weight: 800; color: #0F172A; line-height: 1.2; margin-top: 0.2rem; letter-spacing: -0.02em;">
                {value}
            </div>
            <div style="font-size: 0.75rem; color: #94A3B8; margin-top: 0.25rem;">
                {subtitle}
            </div>
        </div>
    </div>
    """
    if hasattr(st, "html"):
        st.html(card_html)
    else:
        st.markdown(card_html, unsafe_allow_html=True)


def render_sidebar_menu(df):
    """
    Renders the dark charcoal/navy sidebar matching the reference screenshot:
    - Brand Header: [ EP ] Employee | Performance Analytics
    - Section Header: MAIN MENU
    - Menu Items:
      - ▦ Dashboard (active)
      - 👤 Employees (pages/06_Data_Download.py)
      - ➕ Add Employee (interactive form toggle)
      - 📊 Performance Analysis (pages/01_Overview.py)
      - 📄 Reports (pages/06_Data_Download.py)
    - Cohort Filters (Departments, Evaluation Year)
    - Bottom Academic Capstone Attribution (Aayush Mathur, 2500320140001, MCA-B, ABES EC)
    """
    import pandas as pd

    with st.sidebar:
        # Direct scoped stylesheet inside sidebar container
        st.markdown(
            """
            <style>
            /* Eliminate top collapse button artifact ('keyboard_double.') without touching sidebar content */
            [data-testid="stSidebarHeader"],
            [data-testid="stSidebarCollapseButton"],
            header [data-testid="stSidebarCollapseButton"],
            section[data-testid="stSidebar"] [data-testid="stSidebarHeader"],
            section[data-testid="stSidebar"] button[data-testid="stSidebarCollapseButton"] {
                display: none !important;
                visibility: hidden !important;
                height: 0px !important;
                padding: 0px !important;
                margin: 0px !important;
                overflow: hidden !important;
            }

            /* Ensure sidebar container and content are ALWAYS visible, expanded and properly styled */
            section[data-testid="stSidebar"] {
                --primary-color: #3B82F6 !important;
                --background-color: #111827 !important;
                --secondary-background-color: #1F2937 !important;
                --text-color: #F8FAFC !important;
                background-color: #111827 !important;
                border-right: 1px solid #1F2937 !important;
                min-width: 275px !important;
                width: 285px !important;
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
            }

            section[data-testid="stSidebar"] > div {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
            }

            section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
                padding-top: 1.5rem !important;
                display: block !important;
                visibility: visible !important;
                opacity: 1 !important;
            }

            /* Strict line-by-line left-alignment for sidebar menu */
            section[data-testid="stSidebar"] div[data-testid="stPageLink"],
            section[data-testid="stSidebar"] [data-testid="stPageLink"] {
                display: block !important;
                width: 100% !important;
                text-align: left !important;
                margin: 0 0 0.15rem 0 !important;
                padding: 0 !important;
            }

            section[data-testid="stSidebar"] [data-testid="stPageLink"] a,
            section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"],
            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] {
                display: flex !important;
                flex-direction: row !important;
                justify-content: flex-start !important;
                align-items: center !important;
                text-align: left !important;
                width: 100% !important;
                padding: 0.55rem 0.85rem !important;
                border-radius: 8px !important;
                background-color: transparent !important;
                text-decoration: none !important;
                gap: 0.75rem !important;
                transition: all 0.2s ease !important;
            }

            /* Explicit high-contrast white text for ALL children of stPageLink */
            section[data-testid="stSidebar"] [data-testid="stPageLink"] a,
            section[data-testid="stSidebar"] [data-testid="stPageLink"] a *,
            section[data-testid="stSidebar"] [data-testid="stPageLink"] p,
            section[data-testid="stSidebar"] [data-testid="stPageLink"] span,
            section[data-testid="stSidebar"] [data-testid="stPageLink"] div,
            section[data-testid="stSidebar"] [data-testid="stPageLink"] [data-testid="stMarkdownContainer"] p,
            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] [data-testid="stMarkdownContainer"] p,
            section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] [data-testid="stMarkdownContainer"] p {
                color: #F8FAFC !important;
                font-weight: 600 !important;
                font-size: 0.92rem !important;
                line-height: 1.4 !important;
            }

            section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]:hover,
            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"]:hover,
            section[data-testid="stSidebar"] [data-testid="stPageLink"] a:hover {
                background-color: #1F2937 !important;
            }

            section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]:hover *,
            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"]:hover *,
            section[data-testid="stSidebar"] [data-testid="stPageLink"] a:hover * {
                color: #FFFFFF !important;
            }

            section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"][aria-current="page"],
            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"][aria-current="page"],
            section[data-testid="stSidebar"] [data-testid="stPageLink"] a[aria-current="page"] {
                background: linear-gradient(90deg, rgba(37, 99, 235, 0.4) 0%, rgba(37, 99, 235, 0.12) 100%) !important;
                border-left: 3px solid #3B82F6 !important;
            }

            section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"][aria-current="page"] *,
            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"][aria-current="page"] *,
            section[data-testid="stSidebar"] [data-testid="stPageLink"] a[aria-current="page"] * {
                color: #FFFFFF !important;
                font-weight: 700 !important;
            }

            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] > span:first-child {
                display: inline-flex !important;
                align-items: center !important;
                justify-content: center !important;
                width: 22px !important;
                min-width: 22px !important;
                font-size: 1.1rem !important;
                text-align: center !important;
                flex-shrink: 0 !important;
                margin: 0 !important;
            }

            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] [data-testid="stMarkdownContainer"],
            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] p,
            section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] div {
                display: flex !important;
                align-items: center !important;
                justify-content: flex-start !important;
                text-align: left !important;
                flex-grow: 1 !important;
                width: auto !important;
                margin: 0 !important;
                padding: 0 !important;
            }

            /* High Contrast Data Filters in Sidebar (Streamlit 1.67) */
            section[data-testid="stSidebar"] label[data-testid="stWidgetLabel"],
            section[data-testid="stSidebar"] div[data-testid="stSelectbox"] label,
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] label {
                color: #E2E8F0 !important;
                font-size: 0.84rem !important;
                font-weight: 600 !important;
                margin-bottom: 0.35rem !important;
            }

            /* Ensure ALL text inside selectbox and multiselect is pure white */
            section[data-testid="stSidebar"] div[data-testid="stSelectbox"] *,
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] * {
                color: #FFFFFF !important;
                -webkit-text-fill-color: #FFFFFF !important;
            }

            /* Base input container boxes: dark charcoal with subtle border */
            section[data-testid="stSidebar"] div[data-testid="stSelectbox"] [class*="equpb2c0"],
            section[data-testid="stSidebar"] div[data-testid="stSelectbox"] div[role="combobox"],
            section[data-testid="stSidebar"] div[data-testid="stSelectbox"] [data-baseweb="select"] > div,
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [class*="eo2qvtu0"],
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="select"] > div {
                background-color: #1F2937 !important;
                background: #1F2937 !important;
                border: 1px solid #374151 !important;
                border-radius: 8px !important;
            }

            /* Tags in multiselect */
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-tag],
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] span[data-tag],
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [class*="eo2qvtu3"],
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-baseweb="tag"] {
                background-color: #2563EB !important;
                background: #2563EB !important;
                border: 1px solid #3B82F6 !important;
                border-radius: 6px !important;
            }

            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-tag] *,
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [class*="eo2qvtu3"] * {
                background-color: transparent !important;
                color: #FFFFFF !important;
                -webkit-text-fill-color: #FFFFFF !important;
            }

            /* Icons */
            section[data-testid="stSidebar"] div[data-testid="stSelectbox"] svg,
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] svg {
                fill: #94A3B8 !important;
                color: #94A3B8 !important;
            }

            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] [data-tag] svg {
                fill: #FFFFFF !important;
                color: #FFFFFF !important;
            }

            /* MultiSelect input placeholder */
            section[data-testid="stSidebar"] div[data-testid="stMultiSelect"] input::placeholder {
                color: #9CA3AF !important;
                -webkit-text-fill-color: #9CA3AF !important;
            }

            /* Dropdown popover menu styling */
            div[data-testid="stSelectboxVirtualDropdown"],
            div[data-testid="stMultiSelectDropdown"],
            div[class*="equpb2c7"],
            div[data-baseweb="popover"],
            div[data-baseweb="popover"] [data-baseweb="menu"],
            div[data-baseweb="popover"] ul,
            div[data-baseweb="popover"] li {
                background-color: #1F2937 !important;
                color: #FFFFFF !important;
                -webkit-text-fill-color: #FFFFFF !important;
            }

            div[data-testid="stSelectboxVirtualDropdown"] [role="option"]:hover,
            div[data-testid="stMultiSelectDropdown"] [role="option"]:hover,
            div[data-baseweb="popover"] li:hover {
                background-color: #374151 !important;
                color: #FFFFFF !important;
            }

            div[data-testid="stSelectboxVirtualDropdown"] [role="option"][aria-selected="true"],
            div[data-testid="stMultiSelectDropdown"] [role="option"][aria-selected="true"],
            div[data-baseweb="popover"] li[aria-selected="true"] {
                background-color: #2563EB !important;
                color: #FFFFFF !important;
            }
            </style>
            """,
            unsafe_allow_html=True
        )

        # Top Logo Brand
        st.markdown(
            """
            <div style="display: flex; align-items: center; gap: 0.75rem; padding: 0.5rem 0 1.25rem 0; border-bottom: 1px solid #1F2937; margin-bottom: 1.25rem;">
                <div style="background: #2563EB; color: #FFFFFF; font-weight: 800; font-size: 0.95rem; width: 42px; height: 42px;
                            border-radius: 8px; display: flex; align-items: center; justify-content: center; letter-spacing: 0.04em; flex-shrink: 0;">
                    EPRA
                </div>
                <div>
                    <div style="font-size: 0.96rem; font-weight: 700; color: #FFFFFF; line-height: 1.2;">Employee Performance</div>
                    <div style="font-size: 0.73rem; color: #60A5FA; font-weight: 600; margin-top: 1px;">Rating Analysis System</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # Section Header
        st.markdown(
            """
            <div style="font-size: 0.74rem; font-weight: 700; color: #93C5FD; text-transform: uppercase; letter-spacing: 0.09em; margin-bottom: 0.65rem;">
                MAIN MENU
            </div>
            """,
            unsafe_allow_html=True
        )

        # Native Streamlit Page Links (Guaranteed 100% Left Alignment & Active Page Pill)
        st.page_link("app.py", label="Dashboard", icon="🏠")
        st.page_link("pages/06_Data_Download.py", label="Employees", icon="👤")
        st.page_link("pages/01_Overview.py", label="Performance Analysis", icon="📊")
        st.page_link("pages/02_Trend.py", label="Trend Analytics", icon="📈")
        st.page_link("pages/03_Compare.py", label="Department Compare", icon="⚖️")
        st.page_link("pages/04_EDA_Stats.py", label="EDA Statistics", icon="🔬")
        st.page_link("pages/05_Insights.py", label="Executive Insights", icon="💡")
        st.page_link("pages/06_Data_Download.py", label="Reports & Export", icon="📄")
        st.page_link("pages/07_Data_Quality.py", label="Data Quality", icon="🛡️")
        st.page_link("pages/08_About.py", label="About Project", icon="ℹ️")

        st.markdown("<div style='height: 1rem; border-top: 1px solid #1F2937; margin: 1rem 0;'></div>", unsafe_allow_html=True)

        # Real-Time Data Filters
        st.markdown(
            """
            <div style="font-size: 0.74rem; font-weight: 700; color: #93C5FD; text-transform: uppercase; letter-spacing: 0.09em; margin-bottom: 0.55rem;">
                DATA FILTERS
            </div>
            """,
            unsafe_allow_html=True
        )

        all_depts = sorted(df["dept_name"].dropna().unique().tolist()) if "dept_name" in df.columns else []
        selected_depts = st.multiselect(
            "Departments:",
            options=all_depts,
            default=all_depts,
            key="filter_depts"
        )

        all_years = ["All Years"]
        if "review_year" in df.columns:
            years_list = sorted([str(int(y)) for y in df["review_year"].dropna().unique() if pd.notna(y)])
            all_years.extend(years_list)
        selected_year = st.selectbox(
            "Evaluation Year:",
            options=all_years,
            index=0,
            key="filter_eval_year"
        )

        # Real-time visible feedback card so user always knows exactly what is active
        depts_summary = f"All {len(all_depts)} Selected" if len(selected_depts) == len(all_depts) else (f"{len(selected_depts)} Selected" if selected_depts else "None (All Included)")
        if 0 < len(selected_depts) <= 2:
            depts_summary = ", ".join(selected_depts)

        st.markdown(
            f"""
            <div style="background: #1F2937; border: 1px solid #3B82F6; border-radius: 8px; padding: 0.65rem 0.85rem; margin-top: 0.65rem; margin-bottom: 0.5rem;">
                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 0.35rem;">
                    <span style="font-size: 0.72rem; color: #93C5FD; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;">Active Filters</span>
                    <span style="font-size: 0.68rem; background: #2563EB; color: #FFFFFF; font-weight: 700; padding: 2px 7px; border-radius: 9999px;">Real-Time</span>
                </div>
                <div style="font-size: 0.82rem; color: #FFFFFF; font-weight: 600; margin-bottom: 0.2rem; display: flex; align-items: center; gap: 0.4rem;">
                    <span>📅 Year:</span> <span style="color: #60A5FA; font-weight: 700;">{selected_year}</span>
                </div>
                <div style="font-size: 0.82rem; color: #FFFFFF; font-weight: 600; display: flex; align-items: center; gap: 0.4rem;">
                    <span>🏢 Dept:</span> <span style="color: #60A5FA; font-weight: 700;">{depts_summary}</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("<div style='height: 0.5rem;'></div>", unsafe_allow_html=True)

        # Enterprise Platform System Badge
        st.markdown(
            """
            <div style="background: rgba(37, 99, 235, 0.1); border: 1px solid rgba(59, 130, 246, 0.2); border-radius: 10px; padding: 0.85rem; text-align: left;">
                <div style="font-size: 0.68rem; color: #93C5FD; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;">Enterprise System</div>
                <div style="font-size: 0.82rem; font-weight: 700; color: #FFFFFF; margin-top: 0.2rem; line-height: 1.25;">Employee Performance Rating Analysis</div>
                <div style="font-size: 0.72rem; color: #94A3B8; margin-top: 0.35rem;">Workforce Intelligence & Analytics Telemetry</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # Apply real-time filters
    filtered_df = df.copy()
    if selected_depts and "dept_name" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["dept_name"].isin(selected_depts)]
    if selected_year != "All Years" and "review_year" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["review_year"] == int(selected_year)]

    return filtered_df


# Compatible wrapper functions required by tests/test_app.py and existing pages
def inject_tailwind_css() -> None:
    inject_custom_theme()


def render_header(title: str, subtitle: Optional[str] = None, badge_text: Optional[str] = None, icon: Optional[str] = "📊"):
    st.subheader(f"{icon} {title}", divider="blue")
    if subtitle:
        st.caption(subtitle)
    if badge_text:
        st.info(f"🏷️ **{badge_text}**")


def render_kpi_card(
    title: str,
    value: Union[str, int, float],
    delta: Optional[str] = None,
    delta_direction: str = "up",
    subtitle: Optional[str] = None,
    icon: Optional[str] = "📈",
    badge_color: str = "blue",
    key: Optional[str] = None
):
    sub = subtitle if subtitle else (f"{delta} from baseline" if delta else "")
    render_clean_metric_card(
        title=title,
        value=value,
        subtitle=sub,
        icon=icon if icon else "📊"
    )


def render_insight_box(title: str, message: str, insight_type: str = "info", action_label: Optional[str] = None):
    st.info(f"**{title}**\n\n{message}" + (f"\n\n*Action:* **{action_label}**" if action_label else ""))


def render_alert(message: str, alert_type: str = "warning", dismissible: bool = False):
    st.warning(message)


def render_card(title: str, content_html: str, footer: Optional[str] = None):
    with st.container(border=True):
        st.markdown(f"**{title}**")
        st.write(content_html)
        if footer:
            st.caption(footer)


def render_badge(text: str, color: str = "blue") -> str:
    color_map = {
        "blue": "background-color: #DBEAFE; color: #1E40AF;",
        "green": "background-color: #DCFCE7; color: #166534;",
        "amber": "background-color: #FEF3C7; color: #92400E;",
        "red": "background-color: #FEE2E2; color: #991B1B;",
        "purple": "background-color: #F3E8FF; color: #6B21A8;",
        "slate": "background-color: #F1F5F9; color: #475569;",
    }
    style = color_map.get(color, color_map["blue"])
    return f'<span style="display:inline-block; padding:2px 8px; border-radius:9999px; font-size:12px; font-weight:600; {style}">{text}</span>'


def render_sidebar_filters(df):
    return render_sidebar_menu(df)

