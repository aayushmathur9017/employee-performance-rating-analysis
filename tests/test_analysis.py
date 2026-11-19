"""
Test Suite: Analytical KPI Formulas, Data Cleansing, and Statistical Calculations
Employee Performance Rating Analysis
Student: Aayush Mathur | Enrolment: 2500320140001 | Section: MCA-B | ABES Engineering College
Target: KPI Formulas (BR-01, BR-06, BR-07), Data Cleaning (BR-11), Statistics (BR-03, BR-05, BR-09, BR-13)
"""

import numpy as np
import pandas as pd
import pytest
from scipy import stats
from sklearn.linear_model import LinearRegression


# =============================================================================
# 1. KPI FORMULAS & COMPUTATION TESTS (BR-01, BR-06, BR-07)
# =============================================================================

def compute_average_rating(ratings_series: pd.Series) -> float:
    """Computes arithmetic average performance rating rounded to 2 decimal places."""
    clean_series = ratings_series.dropna()
    if clean_series.empty:
        return 0.0
    return round(float(clean_series.mean()), 2)


def compute_promotion_readiness(
    rating: float,
    years_since_promo: float,
    goals_pct: float,
    tenure_years: float
) -> float:
    """
    Computes algorithmic Promotion Readiness Score (0-100 scale):
      - Performance Rating Component: 40% (normalized from 1-5 scale)
      - Promotion Stagnation/Tenure in Role Component: 30% (capped at 4 years)
      - Goals / Quarterly KPI Completion: 20% (percentage 0-100)
      - Total Company Tenure Component: 10% (scaled up to 10 years)
    """
    comp_rating = 0.40 * (rating / 5.0) * 100.0
    comp_promo_time = 0.30 * min(years_since_promo / 4.0, 1.0) * 100.0
    comp_goals = 0.20 * goals_pct
    comp_tenure = 0.10 * min(tenure_years * 10.0, 100.0)
    
    total_score = comp_rating + comp_promo_time + comp_goals + comp_tenure
    return round(float(np.clip(total_score, 0.0, 100.0)), 2)


def categorize_attrition_risk(risk_score: float) -> str:
    """Categorizes attrition risk score into Low, Medium, or High risk tiers."""
    if risk_score >= 0.65:
        return "High Risk"
    elif risk_score >= 0.35:
        return "Medium Risk"
    else:
        return "Low Risk"


def test_kpi_average_rating_calculation():
    """TC-T-07: Test arithmetic and department-grouped average rating calculations."""
    mock_ratings = pd.Series([4.5, 3.8, 4.2, 2.9, 5.0, 3.6, np.nan])
    avg_rating = compute_average_rating(mock_ratings)
    assert avg_rating == 4.00, f"Expected 4.00, got {avg_rating}"

    # Department-wise aggregation test
    df = pd.DataFrame({
        "dept_name": ["Engineering", "Engineering", "Sales", "Sales", "HR"],
        "rating": [4.5, 3.5, 4.0, 3.0, 5.0]
    })
    dept_avgs = df.groupby("dept_name")["rating"].mean().round(2).to_dict()
    assert dept_avgs["Engineering"] == 4.00
    assert dept_avgs["Sales"] == 3.50
    assert dept_avgs["HR"] == 5.00


def test_promotion_readiness_score_formula():
    """TC-T-08: Test composite Promotion Readiness Score bounds and formula weighting."""
    # Top performer scenario: Rating 5.0, 4 years waiting, 100% goals, 10 years tenure
    score_max = compute_promotion_readiness(5.0, 4.0, 100.0, 10.0)
    assert score_max == 100.0, f"Expected 100.0 for max performer, got {score_max}"

    # Low performer scenario: Rating 1.0, 0 years waiting, 10% goals, 0.5 years tenure
    score_min = compute_promotion_readiness(1.0, 0.0, 10.0, 0.5)
    assert 0.0 <= score_min <= 20.0, f"Expected low score, got {score_min}"

    # Mid-career solid performer: Rating 4.0, 3.0 years waiting, 85% goals, 4.0 years tenure
    score_mid = compute_promotion_readiness(4.0, 3.0, 85.0, 4.0)
    expected_mid = round(
        0.40 * (4.0 / 5.0) * 100 +
        0.30 * (3.0 / 4.0) * 100 +
        0.20 * 85.0 +
        0.10 * 40.0,
        2
    )
    assert score_mid == expected_mid, f"Expected {expected_mid}, got {score_mid}"
    assert 0.0 <= score_mid <= 100.0


def test_attrition_risk_categorization():
    """TC-T-09: Test threshold classification of attrition risk into actionable tiers."""
    assert categorize_attrition_risk(0.85) == "High Risk"
    assert categorize_attrition_risk(0.65) == "High Risk"
    assert categorize_attrition_risk(0.50) == "Medium Risk"
    assert categorize_attrition_risk(0.35) == "Medium Risk"
    assert categorize_attrition_risk(0.20) == "Low Risk"
    assert categorize_attrition_risk(0.0) == "Low Risk"


# =============================================================================
# 2. DATA CLEANSING & TRANSFORMATION TESTS (BR-11)
# =============================================================================

def impute_missing_values(df: pd.DataFrame, numeric_strategy: str = "median") -> pd.DataFrame:
    """Imputes missing values: numeric cols with median/mean, categoricals with mode."""
    df_clean = df.copy()
    for col in df_clean.columns:
        if df_clean[col].dtype in [np.float64, np.int64, float, int]:
            fill_val = df_clean[col].median() if numeric_strategy == "median" else df_clean[col].mean()
            df_clean[col] = df_clean[col].fillna(fill_val)
        else:
            mode_vals = df_clean[col].mode()
            fill_val = mode_vals.iloc[0] if not mode_vals.empty else "Unknown"
            df_clean[col] = df_clean[col].fillna(fill_val)
    return df_clean


def clip_outliers_iqr(series: pd.Series, factor: float = 1.5) -> pd.Series:
    """Clips extreme outliers outside [Q1 - factor*IQR, Q3 + factor*IQR] to the boundary."""
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    lower_bound = q1 - factor * iqr
    upper_bound = q3 + factor * iqr
    return series.clip(lower=lower_bound, upper=upper_bound)


def test_data_cleaning_missing_value_imputation():
    """TC-T-10: Test imputation of injected missing values preserving record count and non-null values."""
    raw_df = pd.DataFrame({
        "emp_id": ["E1", "E2", "E3", "E4", "E5"],
        "rating": [4.0, np.nan, 3.5, 4.5, np.nan],
        "feedback_score": [80.0, 90.0, np.nan, 85.0, 95.0],
        "department": ["Eng", "Sales", np.nan, "Eng", "Eng"]
    })
    
    assert raw_df.isna().sum().sum() == 4, "Raw dataframe should have 4 missing values"
    
    cleaned_df = impute_missing_values(raw_df, numeric_strategy="median")
    
    assert cleaned_df.isna().sum().sum() == 0, "Cleaned dataframe must have 0 missing values"
    assert len(cleaned_df) == 5, "Row count must remain intact after imputation"
    assert cleaned_df.loc[0, "rating"] == 4.0, "Existing non-null values must remain unchanged"
    assert cleaned_df.loc[1, "rating"] == 4.0, "Median of [4.0, 3.5, 4.5] is 4.0"
    assert cleaned_df.loc[2, "department"] == "Eng", "Mode of ['Eng', 'Sales', 'Eng', 'Eng'] is 'Eng'"


def test_data_cleaning_outlier_clipping():
    """TC-T-11: Test IQR outlier detection and boundary clipping on abnormal metric values."""
    # Base normal distribution with deliberate extreme anomalies
    data = [20, 22, 21, 23, 22, 20, 24, 21, 22, 23, 100, -50]
    series = pd.Series(data)
    
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    lower_limit = q1 - 1.5 * iqr
    upper_limit = q3 + 1.5 * iqr
    
    clipped = clip_outliers_iqr(series)
    
    assert clipped.max() <= upper_limit, f"Clipped max {clipped.max()} exceeds upper bound {upper_limit}"
    assert clipped.min() >= lower_limit, f"Clipped min {clipped.min()} below lower bound {lower_limit}"
    assert len(clipped) == len(series), "Series length must not change after clipping"


# =============================================================================
# 3. STATISTICAL FUNCTIONS TESTS (BR-03, BR-05, BR-09, BR-13)
# =============================================================================

def test_statistical_pearson_correlation():
    """TC-T-12: Test Pearson correlation r and p-value calculation between income and performance rating."""
    # Synthetic strongly positively correlated data
    np.random.seed(42)
    ratings = np.linspace(2.0, 5.0, 50)
    incomes = ratings * 25000 + np.random.normal(0, 5000, 50)
    
    r_val, p_val = stats.pearsonr(ratings, incomes)
    
    assert -1.0 <= r_val <= 1.0, f"Correlation coefficient {r_val} outside [-1, 1]"
    assert r_val > 0.80, f"Expected strong positive correlation, got {r_val}"
    assert p_val < 0.001, f"Expected statistically significant p-value (<0.001), got {p_val}"


def test_statistical_anova_job_roles():
    """TC-T-13: Test One-Way ANOVA across distinct job role cohorts to detect significant variance."""
    np.random.seed(42)
    # 3 cohorts with distinctly different means
    cohort_dev = np.random.normal(loc=4.2, scale=0.3, size=30)
    cohort_sales = np.random.normal(loc=3.5, scale=0.4, size=30)
    cohort_support = np.random.normal(loc=3.1, scale=0.3, size=30)
    
    f_stat, p_val = stats.f_oneway(cohort_dev, cohort_sales, cohort_support)
    
    assert f_stat > 0, f"F-statistic must be strictly positive, got {f_stat}"
    assert 0.0 <= p_val <= 1.0, f"p-value must be between 0 and 1, got {p_val}"
    assert p_val < 0.05, f"Expected significant difference among cohorts, got p={p_val}"


def test_statistical_two_sample_t_test():
    """TC-T-14: Test independent two-sample t-test for demographic equity / gender bias verification."""
    np.random.seed(42)
    # Fair evaluation simulation: Male and Female cohorts drawn from same distribution
    male_ratings = np.random.normal(loc=3.75, scale=0.5, size=60)
    female_ratings = np.random.normal(loc=3.76, scale=0.5, size=60)
    
    t_stat, p_val = stats.ttest_ind(male_ratings, female_ratings, equal_var=True)
    
    assert isinstance(t_stat, (float, np.floating)), "T-statistic must be a float"
    assert 0.0 <= p_val <= 1.0, f"p-value must be between 0 and 1, got {p_val}"
    # Under null hypothesis of no bias, p-value should be >= 0.05
    assert p_val >= 0.05, f"Expected no significant bias (p >= 0.05), got p={p_val}"


def test_ols_regression_performance_predictors():
    """TC-T-15 (Aux): Test multi-variable OLS regression model fit, R-squared and feature coefficients."""
    np.random.seed(42)
    n = 100
    tenure = np.random.uniform(1.0, 10.0, n)
    training_hrs = np.random.uniform(10.0, 80.0, n)
    # Synthetic performance rating with known positive influence
    rating = 2.0 + 0.15 * tenure + 0.02 * training_hrs + np.random.normal(0, 0.2, n)
    
    X = np.column_stack([tenure, training_hrs])
    y = rating
    
    model = LinearRegression()
    model.fit(X, y)
    r_sq = model.score(X, y)
    
    assert 0.0 <= r_sq <= 1.0, f"R-squared {r_sq} must be between 0 and 1"
    assert r_sq > 0.60, f"Expected explanatory power R^2 > 0.60, got {r_sq}"
    assert model.coef_[0] > 0, "Expected positive coefficient for tenure"
    assert model.coef_[1] > 0, "Expected positive coefficient for training hours"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
