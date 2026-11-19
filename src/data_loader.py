"""
Data Ingestion & Caching Layer
Employee Performance Rating Analysis System
Student: Aayush Mathur | Enrolment No: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0

This module provides high-speed, memoized database access to SQLite3 (data/app.db)
with CSV fallback mechanisms, query profiling, and data integrity health checks.
"""

import os
import sqlite3
import time
from typing import Dict, Any, Optional, Tuple
import pandas as pd
import streamlit as st

# Determine base directory
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "app.db")


def get_db_connection(db_path: Optional[str] = None) -> sqlite3.Connection:
    """
    Creates and returns a thread-safe connection to the SQLite database.
    """
    target_path = db_path or DB_PATH
    if not os.path.exists(target_path):
        raise FileNotFoundError(f"SQLite database file not found at: {target_path}")
    conn = sqlite3.connect(target_path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


@st.cache_data(ttl=3600, show_spinner=False)
def load_table(table_name: str) -> pd.DataFrame:
    """
    Loads a specific relational table from SQLite with automatic CSV fallback.
    Supported tables: departments, employees, performance_reviews, training_records, attrition_records.
    """
    try:
        conn = get_db_connection()
        query = f"SELECT * FROM {table_name};"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception as e:
        csv_file = os.path.join(DATA_DIR, f"{table_name}.csv")
        if os.path.exists(csv_file):
            return pd.read_csv(csv_file)
        raise RuntimeError(f"Failed to load table '{table_name}' from SQLite and CSV fallback failed: {e}")


@st.cache_data(ttl=3600, show_spinner=False)
def load_raw_data() -> Dict[str, pd.DataFrame]:
    """
    Loads all core relational tables from SQLite/CSV and returns a dictionary of DataFrames.
    """
    tables = [
        "departments",
        "employees",
        "performance_reviews",
        "training_records",
        "attrition_records",
    ]
    raw_dict = {}
    for tbl in tables:
        raw_dict[tbl] = load_table(tbl)
    return raw_dict


@st.cache_data(ttl=3600, show_spinner=False)
def load_analytics_view() -> pd.DataFrame:
    """
    Loads the primary denormalized analytical dataset (v_employee_analytics) from SQLite.
    If the view is missing, automatically reconstructs the view schema using relational joins.
    """
    try:
        conn = get_db_connection()
        query = "SELECT * FROM v_employee_analytics;"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception:
        # Fallback: Join raw tables programmatically
        raw = load_raw_data()
        emp = raw["employees"]
        dept = raw["departments"]
        rev = raw["performance_reviews"]
        tr = raw["training_records"]
        att = raw["attrition_records"]

        # Aggregate training
        tr_agg = tr.groupby("emp_id").agg(
            total_training_hours=("training_hours", "sum"),
            courses_completed=("training_id", "count"),
            avg_training_score=("training_score", "mean")
        ).reset_index()

        # Merge
        df = emp.merge(dept, on="dept_id", how="inner", suffixes=("", "_dept"))
        df = df.merge(rev, on="emp_id", how="left")
        df = df.merge(tr_agg, on="emp_id", how="left")
        df = df.merge(att, on="emp_id", how="left", suffixes=("", "_att"))

        # Add manager names
        managers = emp[["emp_id", "first_name", "last_name"]].copy()
        managers["manager_name"] = managers["first_name"] + " " + managers["last_name"]
        df = df.merge(
            managers[["emp_id", "manager_name"]],
            left_on="manager_id",
            right_on="emp_id",
            how="left",
            suffixes=("", "_mgr")
        )
        df["manager_name"] = df["manager_name"].fillna("Executive Management")
        df["full_name"] = df["first_name"] + " " + df["last_name"]

        # Age group
        if "age" in df.columns:
            df["age_group"] = pd.cut(
                df["age"],
                bins=[0, 29, 45, 120],
                labels=["<30", "30-45", ">45"]
            ).astype(str)

        return df


def load_analytics_dataset() -> pd.DataFrame:
    """Alias for load_analytics_view() for backward compatibility with BUILD_PLAN."""
    return load_analytics_view()


@st.cache_data(ttl=3600, show_spinner=False)
def load_clean_data() -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Loads and cleans the analytical dataset with memoized high-speed in-memory caching.
    Ensures zero redundant computation on page transitions and filter interactions.
    """
    raw_df = load_analytics_view()
    from src.cleaner import clean_employee_data
    return clean_employee_data(raw_df)


def check_database_health() -> Dict[str, Any]:
    """
    Performs an end-to-end database health check:
    - Verifies file accessibility and size
    - Measures query response latency
    - Checks row counts across all 5 tables and the analytics view
    - Confirms foreign key integrity
    """
    start_time = time.time()
    health_status: Dict[str, Any] = {
        "status": "Healthy",
        "db_path": DB_PATH,
        "db_exists": os.path.exists(DB_PATH),
        "file_size_kb": 0.0,
        "latency_ms": 0.0,
        "tables": {},
        "view_rows": 0,
        "issues": [],
    }

    if not health_status["db_exists"]:
        health_status["status"] = "Critical"
        health_status["issues"].append("SQLite database file not found on disk.")
        return health_status

    health_status["file_size_kb"] = round(os.path.getsize(DB_PATH) / 1024.0, 2)

    try:
        conn = get_db_connection()
        cur = conn.cursor()

        # Check tables
        table_names = [
            "departments",
            "employees",
            "performance_reviews",
            "training_records",
            "attrition_records",
        ]
        for tbl in table_names:
            cur.execute(f"SELECT COUNT(*) FROM {tbl};")
            count = cur.fetchone()[0]
            health_status["tables"][tbl] = count

        # Check view
        cur.execute("SELECT COUNT(*) FROM v_employee_analytics;")
        health_status["view_rows"] = cur.fetchone()[0]

        # Check foreign keys
        cur.execute("PRAGMA foreign_key_check;")
        fk_violations = cur.fetchall()
        if fk_violations:
            health_status["status"] = "Warning"
            health_status["issues"].append(f"Foreign key violations detected: {len(fk_violations)}")

        conn.close()
    except Exception as e:
        health_status["status"] = "Degraded"
        health_status["issues"].append(str(e))

    health_status["latency_ms"] = round((time.time() - start_time) * 1000.0, 2)
    return health_status
