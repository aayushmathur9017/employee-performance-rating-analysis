"""
Statistical Engine, KPI Formulas & Prescriptive Talent Analytics
Employee Performance Rating Analysis System
Student: Aayush Mathur | Enrolment No: 2500320140001 | Section: MCA-B | ABES Engineering College
Date: 2026-11-19 | Version: 1.0

This module implements the core business logic, inferential statistics, predictive modeling,
and algorithmic scoring engines for the Employee Performance Rating Analysis platform.
"""

from typing import Dict, Any, Tuple, List, Optional
import numpy as np
import pandas as pd
from scipy import stats
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
import streamlit as st


# =============================================================================
# 1. CORE KPI & PERFORMANCE AGGREGATIONS (BR-01, BR-02)
# =============================================================================

def compute_average_rating(ratings_series: pd.Series) -> float:
    """
    Computes arithmetic average performance rating rounded to 2 decimal places.
    Handles NaN values safely; returns 0.0 for empty series.
    Matches TC-T-07.
    """
    clean_series = pd.to_numeric(ratings_series, errors="coerce").dropna()
    if clean_series.empty:
        return 0.0
    return round(float(clean_series.mean()), 2)


def get_top_bottom_performers(
    df: pd.DataFrame,
    n_pct: float = 0.10,
    rating_col: str = "performance_rating"
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Identifies Top N% (Star Performers) and Bottom N% (PIP Candidates) based on rating.
    """
    if df.empty or rating_col not in df.columns:
        return pd.DataFrame(), pd.DataFrame()

    valid_df = df.dropna(subset=[rating_col]).sort_values(by=rating_col, ascending=False)
    n_count = max(1, int(np.ceil(len(valid_df) * n_pct)))

    top_performers = valid_df.head(n_count).copy()
    bottom_performers = valid_df.tail(n_count).sort_values(by=rating_col, ascending=True).copy()

    return top_performers, bottom_performers


@st.cache_data(show_spinner=False)
def get_department_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """
    Computes comprehensive statistical aggregates per department:
      - Workforce Count
      - Mean Rating, Median Rating, Std Dev
      - 25th and 75th Percentiles
      - Top Performer Share (Rating >= 4.5 %)
      - PIP / Low Performer Share (Rating < 3.0 %)
      - Average Current Salary
      - Average Tenure Years
    """
    if df.empty or "dept_name" not in df.columns:
        return pd.DataFrame()

    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"
    salary_col = "current_salary" if "current_salary" in df.columns else "monthly_income"
    tenure_col = "tenure_years" if "tenure_years" in df.columns else None

    grouped = df.groupby("dept_name")

    records = []
    for dept_name, group in grouped:
        ratings = pd.to_numeric(group[rating_col], errors="coerce").dropna()
        count = len(group)
        if ratings.empty:
            continue

        mean_val = round(float(ratings.mean()), 2)
        std_val = round(float(ratings.std()), 2) if len(ratings) > 1 else 0.0
        median_val = round(float(ratings.median()), 2)
        p25 = round(float(np.percentile(ratings, 25)), 2)
        p75 = round(float(np.percentile(ratings, 75)), 2)

        top_share = round((float((ratings >= 4.5).sum()) / count) * 100.0, 1)
        pip_share = round((float((ratings < 3.0).sum()) / count) * 100.0, 1)

        avg_salary = 0.0
        if salary_col in group.columns:
            sal_series = pd.to_numeric(group[salary_col], errors="coerce").dropna()
            avg_salary = round(float(sal_series.mean()), 2) if not sal_series.empty else 0.0

        avg_tenure = 0.0
        if tenure_col and tenure_col in group.columns:
            ten_series = pd.to_numeric(group[tenure_col], errors="coerce").dropna()
            avg_tenure = round(float(ten_series.mean()), 1) if not ten_series.empty else 0.0

        records.append({
            "Department": dept_name,
            "Employee Count": count,
            "Average Rating": mean_val,
            "Median Rating": median_val,
            "Std Dev": std_val,
            "25th Percentile": p25,
            "75th Percentile": p75,
            "Top Performers (%)": top_share,
            "PIP Share (%)": pip_share,
            "Average Salary": avg_salary,
            "Average Tenure (Yrs)": avg_tenure,
        })

    dept_df = pd.DataFrame(records)
    if not dept_df.empty:
        dept_df = dept_df.sort_values(by="Average Rating", ascending=False).reset_index(drop=True)
    return dept_df


# =============================================================================
# 2. LEADERSHIP & MANAGER TEAM AGGREGATION (BR-04)
# =============================================================================

@st.cache_data(show_spinner=False)
def get_manager_leaderboard(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ranks managers by team size, team average rating, high performers count,
    average promotion readiness score, and overdue promotions.
    """
    if df.empty or "manager_name" not in df.columns:
        return pd.DataFrame()

    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"
    promo_col = "promotion_readiness_score" if "promotion_readiness_score" in df.columns else None

    records = []
    for manager, group in df.groupby("manager_name"):
        if not manager or manager == "None" or str(manager).strip() == "":
            continue
        team_size = len(group)
        ratings = pd.to_numeric(group[rating_col], errors="coerce").dropna()
        avg_rating = round(float(ratings.mean()), 2) if not ratings.empty else 0.0

        high_performers = int((ratings >= 4.5).sum()) if not ratings.empty else 0
        avg_promo = 0.0
        if promo_col and promo_col in group.columns:
            promos = pd.to_numeric(group[promo_col], errors="coerce").dropna()
            avg_promo = round(float(promos.mean()), 1) if not promos.empty else 0.0

        overdue_count = 0
        if "overdue_promotion_flag" in group.columns:
            overdue_count = int(pd.to_numeric(group["overdue_promotion_flag"], errors="coerce").fillna(0).sum())

        dept_str = group["dept_name"].iloc[0] if "dept_name" in group.columns and not group.empty else "N/A"

        records.append({
            "Manager Name": manager,
            "Department": dept_str,
            "Team Size": team_size,
            "Team Avg Rating": avg_rating,
            "High Performers (>=4.5)": high_performers,
            "Avg Promo Readiness": avg_promo,
            "Overdue Candidates": overdue_count,
        })

    leaderboard = pd.DataFrame(records)
    if not leaderboard.empty:
        leaderboard = leaderboard.sort_values(by=["Team Avg Rating", "Team Size"], ascending=[False, False]).reset_index(drop=True)
    return leaderboard


# =============================================================================
# 3. PRESCRIPTIVE PROMOTION & ATTRITION ENGINES (BR-06, BR-07, BR-14)
# =============================================================================

def compute_promotion_readiness(
    rating: float,
    years_since_promo: float,
    goals_pct: float,
    tenure_years: float
) -> float:
    """
    Computes algorithmic Promotion Readiness Score (0-100 scale):
      - Performance Rating Component: 40% (normalized from 1-5 scale)
      - Promotion Stagnation/Tenure in Role: 30% (capped at 4 years)
      - Goals / Quarterly KPI Completion: 20% (percentage 0-100)
      - Total Company Tenure: 10% (scaled up to 10 years)
    Strictly matches TC-T-08.
    """
    comp_rating = 0.40 * (rating / 5.0) * 100.0
    comp_promo_time = 0.30 * min(years_since_promo / 4.0, 1.0) * 100.0
    comp_goals = 0.20 * goals_pct
    comp_tenure = 0.10 * min(tenure_years * 10.0, 100.0)

    total_score = comp_rating + comp_promo_time + comp_goals + comp_tenure
    return round(float(np.clip(total_score, 0.0, 100.0)), 2)


def compute_promotion_readiness_alt(
    rating_norm: float,
    tenure_norm: float,
    training_norm: float
) -> float:
    """
    Alternative Promotion Readiness Score formula:
      Score = 0.50 * rating_norm + 0.30 * tenure_norm + 0.20 * training_norm
    Input values should be in [0, 100] scale.
    """
    score = 0.50 * rating_norm + 0.30 * tenure_norm + 0.20 * training_norm
    return round(float(np.clip(score, 0.0, 100.0)), 2)


def calculate_promotion_scores(df: pd.DataFrame) -> pd.Series:
    """
    Vectorized computation of Promotion Readiness Scores for a dataframe.
    """
    rating = pd.to_numeric(df.get("performance_rating", 3.0), errors="coerce").fillna(3.0)
    years_promo = pd.to_numeric(df.get("years_since_last_promotion", 1.0), errors="coerce").fillna(1.0)
    goals = pd.to_numeric(df.get("goals_completed_pct", 75.0), errors="coerce").fillna(75.0)
    tenure = pd.to_numeric(df.get("tenure_years", 2.0), errors="coerce").fillna(2.0)

    comp_rating = 0.40 * (rating / 5.0) * 100.0
    comp_promo = 0.30 * np.minimum(years_promo / 4.0, 1.0) * 100.0
    comp_goals = 0.20 * goals
    comp_tenure = 0.10 * np.minimum(tenure * 10.0, 100.0)

    total = comp_rating + comp_promo + comp_goals + comp_tenure
    return np.clip(total, 0.0, 100.0).round(2)


@st.cache_data(show_spinner=False)
def identify_overdue_promotions(
    df: pd.DataFrame,
    rating_threshold: float = 4.0,
    tenure_threshold: float = 3.0
) -> pd.DataFrame:
    """
    Filters for employees with high performance rating (>= 4.0) who have stayed
    in the same role or at the firm for > 3.0 years without a promotion.
    """
    if df.empty:
        return pd.DataFrame()

    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"
    promo_time_col = "years_since_last_promotion" if "years_since_last_promotion" in df.columns else "tenure_years"

    mask = (
        (pd.to_numeric(df[rating_col], errors="coerce") >= rating_threshold) &
        (pd.to_numeric(df[promo_time_col], errors="coerce") > tenure_threshold)
    )
    overdue_df = df[mask].copy()
    return overdue_df.sort_values(by=[rating_col, promo_time_col], ascending=[False, False])


def categorize_attrition_risk(risk_score: float) -> str:
    """
    Categorizes attrition risk score into Low, Medium, or High risk tiers.
    Strictly matches TC-T-09.
    """
    if risk_score >= 0.65:
        return "High Risk"
    elif risk_score >= 0.35:
        return "Medium Risk"
    else:
        return "Low Risk"


@st.cache_data(show_spinner=False)
def compute_department_attrition_risk(df: pd.DataFrame) -> pd.DataFrame:
    """
    Builds a department-level Attrition Risk & Retention Heatmap Index:
      - Historical Attrition Rate (%)
      - Average Attrition Risk Score (0-1)
      - Average Job Satisfaction Score (1-5)
      - Monthly Absenteeism Rate (%)
      - Risk Category Tier
    """
    if df.empty or "dept_name" not in df.columns:
        return pd.DataFrame()

    records = []
    for dept_name, group in df.groupby("dept_name"):
        count = len(group)
        attr_count = 0
        if "attrition_flag" in group.columns:
            attr_count = int(pd.to_numeric(group["attrition_flag"], errors="coerce").fillna(0).sum())
        attr_rate = round((attr_count / count * 100.0), 1) if count > 0 else 0.0

        risk_series = pd.to_numeric(group.get("attrition_risk_score", 0.2), errors="coerce").dropna()
        avg_risk = round(float(risk_series.mean()), 3) if not risk_series.empty else 0.0

        sat_series = pd.to_numeric(group.get("satisfaction_score", group.get("job_satisfaction_score", 3.5)), errors="coerce").dropna()
        avg_sat = round(float(sat_series.mean()), 2) if not sat_series.empty else 3.5

        absent_series = pd.to_numeric(group.get("monthly_absenteeism_rate", 2.0), errors="coerce").dropna()
        avg_absent = round(float(absent_series.mean()), 1) if not absent_series.empty else 2.0

        tier = categorize_attrition_risk(avg_risk)

        records.append({
            "Department": dept_name,
            "Active Headcount": count,
            "Attrition Rate (%)": attr_rate,
            "Avg Risk Score": avg_risk,
            "Avg Satisfaction": avg_sat,
            "Absenteeism (%)": avg_absent,
            "Risk Tier": tier,
        })

    risk_df = pd.DataFrame(records)
    if not risk_df.empty:
        risk_df = risk_df.sort_values(by="Avg Risk Score", ascending=False).reset_index(drop=True)
    return risk_df


# =============================================================================
# 4. INFERENTIAL & ADVANCED STATISTICAL METHODS (BR-03, BR-05, BR-09, BR-13)
# =============================================================================

@st.cache_data(show_spinner=False)
def run_salary_rating_correlation(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Computes Pearson and Spearman correlation coefficients and p-values
    between salary/income and performance rating.
    Matches TC-T-12.
    """
    if df.empty:
        return {"pearson_r": 0.0, "pearson_p": 1.0, "spearman_r": 0.0, "spearman_p": 1.0, "sample_size": 0}

    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"
    sal_col = "current_salary" if "current_salary" in df.columns else "monthly_income"

    sub = df[[rating_col, sal_col]].dropna()
    sub[rating_col] = pd.to_numeric(sub[rating_col], errors="coerce")
    sub[sal_col] = pd.to_numeric(sub[sal_col], errors="coerce")
    sub = sub.dropna()

    if len(sub) < 3:
        return {"pearson_r": 0.0, "pearson_p": 1.0, "spearman_r": 0.0, "spearman_p": 1.0, "sample_size": len(sub)}

    r_p, p_p = stats.pearsonr(sub[rating_col], sub[sal_col])
    r_s, p_s = stats.spearmanr(sub[rating_col], sub[sal_col])

    return {
        "pearson_r": round(float(r_p), 3),
        "pearson_p": float(p_p),
        "spearman_r": round(float(r_s), 3),
        "spearman_p": float(p_s),
        "sample_size": len(sub),
    }


@st.cache_data(show_spinner=False)
def run_gender_parity_test(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Tests demographic equity between Male and Female rating cohorts using
    both Independent Two-Sample t-test and Mann-Whitney U test.
    Matches TC-T-14.
    """
    if df.empty or "gender" not in df.columns:
        return {"t_stat": 0.0, "t_p_val": 1.0, "u_stat": 0.0, "u_p_val": 1.0, "is_fair": True}

    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"
    sub = df[["gender", rating_col]].dropna()
    sub[rating_col] = pd.to_numeric(sub[rating_col], errors="coerce")
    sub = sub.dropna()

    male_ratings = sub[sub["gender"].str.upper().str.startswith("M")][rating_col]
    female_ratings = sub[sub["gender"].str.upper().str.startswith("F")][rating_col]

    if len(male_ratings) < 3 or len(female_ratings) < 3:
        return {
            "t_stat": 0.0, "t_p_val": 1.0, "u_stat": 0.0, "u_p_val": 1.0,
            "male_count": len(male_ratings), "female_count": len(female_ratings),
            "male_mean": 0.0, "female_mean": 0.0, "is_fair": True
        }

    t_stat, t_p = stats.ttest_ind(male_ratings, female_ratings, equal_var=False)
    u_stat, u_p = stats.mannwhitneyu(male_ratings, female_ratings, alternative="two-sided")

    return {
        "t_stat": round(float(t_stat), 3),
        "t_p_val": round(float(t_p), 4),
        "u_stat": round(float(u_stat), 3),
        "u_p_val": round(float(u_p), 4),
        "male_count": len(male_ratings),
        "female_count": len(female_ratings),
        "male_mean": round(float(male_ratings.mean()), 2),
        "female_mean": round(float(female_ratings.mean()), 2),
        "is_fair": bool(t_p >= 0.05),
    }


@st.cache_data(show_spinner=False)
def run_job_role_anova(df: pd.DataFrame, group_col: str = "job_role") -> Dict[str, Any]:
    """
    One-Way ANOVA and Kruskal-Wallis test across job roles or departments.
    Matches TC-T-13.
    """
    if df.empty or group_col not in df.columns:
        return {"f_stat": 0.0, "f_p_val": 1.0, "h_stat": 0.0, "h_p_val": 1.0, "cohorts": 0}

    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"
    sub = df[[group_col, rating_col]].dropna()
    sub[rating_col] = pd.to_numeric(sub[rating_col], errors="coerce")
    sub = sub.dropna()

    groups = [group[rating_col].values for _, group in sub.groupby(group_col) if len(group) >= 3]

    if len(groups) < 2:
        return {"f_stat": 0.0, "f_p_val": 1.0, "h_stat": 0.0, "h_p_val": 1.0, "cohorts": len(groups)}

    f_stat, f_p = stats.f_oneway(*groups)
    h_stat, h_p = stats.kruskal(*groups)

    return {
        "f_stat": round(float(f_stat), 3),
        "f_p_val": round(float(f_p), 5),
        "h_stat": round(float(h_stat), 3),
        "h_p_val": round(float(h_p), 5),
        "cohorts": len(groups),
        "is_significant": bool(f_p < 0.05),
    }


@st.cache_data(show_spinner=False)
def run_linear_regression(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Fits an OLS Linear Regression model predicting performance rating
    from numeric driver variables. Returns R-squared, coefficients, and feature importance.
    Matches TC-T-15.
    """
    if df.empty:
        return {"r_squared": 0.0, "coefficients": {}, "intercept": 0.0}

    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"
    candidate_features = [
        "tenure_years",
        "total_training_hours",
        "satisfaction_score",
        "job_satisfaction_score",
        "projects_completed",
        "goals_completed_pct",
        "monthly_absenteeism_rate",
    ]
    available_features = [col for col in candidate_features if col in df.columns]

    sub = df[[rating_col] + available_features].dropna()
    for col in [rating_col] + available_features:
        sub[col] = pd.to_numeric(sub[col], errors="coerce")
    sub = sub.dropna()

    if len(sub) < 10 or len(available_features) == 0:
        return {"r_squared": 0.0, "coefficients": {}, "intercept": 0.0}

    X = sub[available_features].values
    y = sub[rating_col].values

    model = LinearRegression()
    model.fit(X, y)
    r2 = model.score(X, y)

    coef_dict = {
        feat: round(float(coef), 4)
        for feat, coef in zip(available_features, model.coef_)
    }

    return {
        "r_squared": round(float(r2), 3),
        "coefficients": coef_dict,
        "intercept": round(float(model.intercept_), 3),
        "features": available_features,
        "n_samples": len(sub),
    }


@st.cache_data(show_spinner=False)
def run_random_forest_importance(df: pd.DataFrame) -> Dict[str, float]:
    """
    Fits Random Forest regressor to extract non-linear feature importances.
    """
    if df.empty:
        return {}

    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"
    features = [
        col for col in [
            "tenure_years",
            "total_training_hours",
            "satisfaction_score",
            "projects_completed",
            "goals_completed_pct",
            "current_salary",
        ] if col in df.columns
    ]

    sub = df[[rating_col] + features].dropna()
    for col in [rating_col] + features:
        sub[col] = pd.to_numeric(sub[col], errors="coerce")
    sub = sub.dropna()

    if len(sub) < 20 or len(features) == 0:
        return {}

    rf = RandomForestRegressor(n_estimators=50, random_state=42)
    rf.fit(sub[features], sub[rating_col])

    return {
        feat: round(float(imp), 4)
        for feat, imp in zip(features, rf.feature_importances_)
    }


@st.cache_data(show_spinner=False)
def calc_yoy_performance_trends(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calculates multi-year rating trajectories, rolling averages, and YoY percentage changes.
    """
    if df.empty:
        return pd.DataFrame()

    year_col = "review_year" if "review_year" in df.columns else None
    period_col = "review_period" if "review_period" in df.columns else None
    rating_col = "performance_rating" if "performance_rating" in df.columns else "rating"

    if not year_col or year_col not in df.columns:
        return pd.DataFrame()

    sub = df.copy()
    sub[rating_col] = pd.to_numeric(sub[rating_col], errors="coerce")
    sub[year_col] = pd.to_numeric(sub[year_col], errors="coerce")

    annual_summary = sub.groupby(year_col)[rating_col].agg(["mean", "count"]).reset_index()
    annual_summary.columns = ["Year", "Average Rating", "Review Count"]
    annual_summary["Average Rating"] = annual_summary["Average Rating"].round(2)
    annual_summary["YoY Delta"] = annual_summary["Average Rating"].diff().round(2).fillna(0.0)
    annual_summary["YoY % Change"] = (
        (annual_summary["YoY Delta"] / annual_summary["Average Rating"].shift(1)) * 100.0
    ).round(2).fillna(0.0)

    return annual_summary
