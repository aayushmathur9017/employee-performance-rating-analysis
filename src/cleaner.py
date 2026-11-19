"""
Data Cleansing, Imputation & Quality Audit Pipeline
Employee Performance Rating Analysis System
Student: Aayush Mathur | Enrolment No: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0

This module provides data cleaning, missing value imputation, range enforcement,
IQR outlier detection/clipping, and automated data hygiene audit reporting.
"""

from typing import Dict, Any, Tuple, Optional, List
import numpy as np
import pandas as pd
import streamlit as st


def impute_missing_values(df: pd.DataFrame, numeric_strategy: str = "median") -> pd.DataFrame:
    """
    Imputes missing values:
      - Numeric columns with median (default) or mean.
      - Non-numeric / categorical columns with mode (most frequent) or 'Unknown'.
    Matches TC-T-10 requirements.
    Optimized to only iterate over columns containing missing values.
    """
    df_clean = df.copy()
    cols_with_na = df_clean.columns[df_clean.isna().any()].tolist()
    for col in cols_with_na:
        if df_clean[col].dtype in [np.float64, np.int64, float, int]:
            fill_val = df_clean[col].median() if numeric_strategy == "median" else df_clean[col].mean()
            if pd.isna(fill_val):
                fill_val = 0.0
            df_clean[col] = df_clean[col].fillna(fill_val)
        else:
            mode_vals = df_clean[col].mode()
            fill_val = mode_vals.iloc[0] if not mode_vals.empty else "Unknown"
            df_clean[col] = df_clean[col].fillna(fill_val)
    return df_clean


def detect_outliers_iqr(series: pd.Series, factor: float = 1.5) -> pd.Series:
    """
    Identifies outliers using the Interquartile Range (IQR) rule:
    Returns a boolean Series where True indicates the value lies outside
    [Q1 - factor * IQR, Q3 + factor * IQR].
    """
    numeric_clean = pd.to_numeric(series, errors="coerce").dropna()
    if numeric_clean.empty or len(numeric_clean) < 4:
        return pd.Series(False, index=series.index)
    
    q1 = numeric_clean.quantile(0.25)
    q3 = numeric_clean.quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - factor * iqr
    upper_bound = q3 + factor * iqr
    
    return (series < lower_bound) | (series > upper_bound)


def clip_outliers_iqr(series: pd.Series, factor: float = 1.5) -> pd.Series:
    """
    Clips extreme outliers outside [Q1 - factor*IQR, Q3 + factor*IQR] to the boundary values.
    Matches TC-T-11 requirements.
    """
    numeric_clean = pd.to_numeric(series, errors="coerce").dropna()
    if numeric_clean.empty:
        return series
    
    q1 = numeric_clean.quantile(0.25)
    q3 = numeric_clean.quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - factor * iqr
    upper_bound = q3 + factor * iqr
    return series.clip(lower=lower_bound, upper=upper_bound)


def validate_ranges(df: pd.DataFrame) -> pd.DataFrame:
    """
    Validates and clamps domain-specific numerical ranges:
      - performance_rating: [1.0, 5.0]
      - satisfaction_score, job_satisfaction_score: [1.0, 5.0]
      - work_life_balance_score: [1, 5]
      - goals_completed_pct: [0.0, 100.0]
      - quarterly_kpi_score: [0.0, 100.0]
    """
    df_clean = df.copy()

    if "performance_rating" in df_clean.columns:
        df_clean["performance_rating"] = df_clean["performance_rating"].clip(lower=1.0, upper=5.0)

    for sat_col in ["satisfaction_score", "job_satisfaction_score", "work_life_balance_score"]:
        if sat_col in df_clean.columns:
            df_clean[sat_col] = df_clean[sat_col].clip(lower=1.0, upper=5.0)

    for pct_col in ["goals_completed_pct", "quarterly_kpi_score", "promotion_readiness_score"]:
        if pct_col in df_clean.columns:
            df_clean[pct_col] = df_clean[pct_col].clip(lower=0.0, upper=100.0)

    return df_clean


@st.cache_data(show_spinner=False)
def clean_employee_data(
    df: pd.DataFrame,
    numeric_strategy: str = "median",
    clip_outliers: bool = False
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Comprehensive data cleansing pipeline satisfying BR-11 and UI-07:
      1. Baseline telemetry audit (row count, cell count, missing values per column).
      2. Deduplication on composite keys or entire rows.
      3. Range validation and constraint boundary clamping.
      4. Missing value imputation (median for numerics, mode for categoricals).
      5. IQR outlier audit and optional boundary clipping.
      6. Audit summary report generation with Overall Completeness Health Index (%).
    """
    raw_df = df.copy()
    total_rows_raw = len(raw_df)
    total_cells = raw_df.size
    total_nulls_raw = int(raw_df.isna().sum().sum())
    missing_by_column = raw_df.isna().sum().to_dict()
    actions_log: List[str] = []

    # 1. Deduplication
    # Check for duplicate review events (emp_id + review_year / review_id) or identical rows
    dedup_cols = None
    if "emp_id" in raw_df.columns and "review_id" in raw_df.columns:
        dedup_cols = ["emp_id", "review_id"]
    elif "emp_id" in raw_df.columns and "review_year" in raw_df.columns:
        dedup_cols = ["emp_id", "review_year"]

    if dedup_cols:
        cleaned_df = raw_df.drop_duplicates(subset=dedup_cols).copy()
    else:
        cleaned_df = raw_df.drop_duplicates().copy()

    duplicates_removed = total_rows_raw - len(cleaned_df)
    if duplicates_removed > 0:
        actions_log.append(f"Removed {duplicates_removed} duplicate records using subset {dedup_cols}.")
    else:
        actions_log.append("Deduplication check passed: 0 duplicate records found.")

    # 2. Range validation
    cleaned_df = validate_ranges(cleaned_df)
    actions_log.append("Enforced domain bounds: performance_rating [1.0, 5.0], satisfaction [1.0, 5.0].")

    # 3. Missing Value Imputation
    cleaned_df = impute_missing_values(cleaned_df, numeric_strategy=numeric_strategy)
    total_nulls_clean = int(cleaned_df.isna().sum().sum())
    actions_log.append(f"Imputed missing values using numeric strategy='{numeric_strategy}' and categorical mode.")

    # 4. Outlier Analysis
    outliers_detected: Dict[str, int] = {}
    numeric_cols = cleaned_df.select_dtypes(include=[np.number]).columns.tolist()
    for col in ["current_salary", "monthly_income", "performance_rating", "total_training_hours"]:
        if col in numeric_cols:
            outlier_mask = detect_outliers_iqr(cleaned_df[col])
            count = int(outlier_mask.sum())
            outliers_detected[col] = count
            if clip_outliers and count > 0:
                cleaned_df[col] = clip_outliers_iqr(cleaned_df[col])
                actions_log.append(f"Clipped {count} outliers in '{col}' using 1.5*IQR boundary.")

    # 5. Health Index Calculation
    health_index = round((1.0 - (total_nulls_raw / total_cells if total_cells > 0 else 0.0)) * 100.0, 2)

    audit_report: Dict[str, Any] = {
        "total_rows_raw": total_rows_raw,
        "total_rows_clean": len(cleaned_df),
        "total_cells": total_cells,
        "duplicates_removed": duplicates_removed,
        "total_nulls_raw": total_nulls_raw,
        "total_nulls_clean": total_nulls_clean,
        "health_index_pct": health_index,
        "missing_by_column": missing_by_column,
        "outliers_detected": outliers_detected,
        "cleaning_actions_log": actions_log,
    }

    return cleaned_df, audit_report


def clean_employee_dataset(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Alias for clean_employee_data() for exact compliance with BUILD_PLAN task BUILD-03."""
    return clean_employee_data(df)
