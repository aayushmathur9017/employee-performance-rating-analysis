# Employee Performance Rating Analysis (EPRA)
### Enterprise Workforce Intelligence, Longitudinal Calibration & Predictive Analytics Platform

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://streamlit.io)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![Tests Passing](https://img.shields.io/badge/tests-20%2F20%20passed-brightgreen.svg)](tests/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 🎓 Capstone Project Attribution
- **Student Name**: Aayush Mathur
- **University Enrolment No**: `2500320140001`
- **Program & Section**: Master of Computer Applications (MCA) | Section: MCA-B
- **Institution**: ABES Engineering College, Ghaziabad
- **Academic Capstone**: Major Project (Semester III/IV)
- **Evaluation Period**: 2021 – 2026 (Extended 6-Year Longitudinal Dataset)

---

## 📌 Project Overview
**Employee Performance Rating Analysis (EPRA)** is an enterprise-grade workforce telemetry and appraisal intelligence platform. Built with Python, Streamlit, and Plotly, EPRA provides real-time data cleansing, statistical inference (ANOVA, Two-Sample *t*-tests, OLS regression), machine learning feature importance (Random Forest), and prescriptive talent decision support (promotion readiness index & flight risk calibration).

### Key Highlights
- **Active Workforce Telemetry**: Tracks 575+ active personnel across 7 departments with 2,051 longitudinal performance appraisal records spanning 2021 through 2026.
- **Microsecond Cached Execution**: Implements `@st.cache_data` memoization yielding a **55x speedup** (cached query latency ~8.3ms).
- **Statistical Rigor**: Integrated Scipy & Statsmodels engines executing hypothesis testing, gender parity audits, and salary-to-rating elasticity.
- **Enterprise UI/UX**: Custom SaaS design system with dark navy sidebar navigation, royal blue interactive tags, and high-contrast telemetry cards.
- **Zero-Failure Test Suite**: 20/20 automated unit and integration tests passing (`pytest -v`).

---

## 🏛️ Application Architecture & Navigation

```
├── app.py                      # Master Entrypoint & Executive Gateway (Dashboard)
├── pages/
│   ├── 01_Overview.py          # UI-01: Descriptive Performance & Appraisal Calibration
│   ├── 02_Trend.py             # UI-02: Multi-Year Longitudinal Trends (2021-2026)
│   ├── 03_Compare.py           # UI-03: Cross-Departmental Benchmarking & Radar Analysis
│   ├── 04_EDA_Stats.py         # UI-04: Inferential Statistics, ANOVA & Random Forest ML
│   ├── 05_Insights.py          # UI-05: Prescriptive Promotion Readiness & Flight Risk Matrix
│   ├── 06_Data_Download.py     # UI-06: Employee Directory, Filtering & Multi-Format Export
│   ├── 07_Data_Quality.py      # UI-07: Real-Time Data Hygiene, Audit Trail & Anomaly Detection
│   └── 08_About.py             # UI-08: Academic Capstone Blueprint & Traceability Matrix
├── src/
│   ├── analysis.py             # Statistical & ML engines (RF, OLS, ANOVA, t-test, scoring)
│   ├── cleaner.py              # Automated data cleaning pipeline & anomaly remediation
│   ├── data_loader.py          # SQLite database connection & cached analytics loader
│   ├── generate_data.py        # Realistic enterprise data synthesis generator (2021-2026)
│   └── ui_components.py        # Custom styling, dark sidebar, KPI cards & active filters
├── data/
│   ├── app.db                  # Relational SQLite database
│   ├── employee_performance_data.xlsx  # Multi-tab canonical Excel workbook
│   └── *.csv                   # Normalized CSV tables (employees, reviews, depts, training)
├── docs/                       # Project documentation, BRD, TRD, ERD, and capstone presentation
└── tests/
    ├── test_analysis.py        # Statistical calculations & ML model verification
    ├── test_app.py             # UI components, config integrity, and module imports
    └── test_data.py            # Ingestion, schema validity, and cleaning pipeline assertions
```

---

## 🛠️ Technology Stack

| Layer | Technologies |
|---|---|
| **Frontend & UI** | Streamlit (v1.67), Custom CSS Design System, Responsive Flexbox |
| **Data Processing** | Python 3.13, Pandas, NumPy |
| **Statistical Analysis** | SciPy (Stats), Statsmodels (OLS Regression) |
| **Machine Learning** | Scikit-learn (RandomForestRegressor for feature importance) |
| **Data Visualizations** | Plotly Express, Plotly Graph Objects (Donut, Radar, Box, Bar, Scatter) |
| **Database & Persistence** | SQLite3, OpenPyXL, SQLite relational schema |
| **Testing & Quality** | Pytest, Pytest-Mock, Pytest-Cov (20/20 Tests Passing) |

---

## 🚀 Getting Started

### 1. Prerequisites
- Python 3.10, 3.11, 3.12, or 3.13 installed
- Git CLI

### 2. Clone the Repository
```bash
git clone https://github.com/aayushmathur9017/employee-performance-rating-analysis.git
cd employee-performance-rating-analysis
```

### 3. Create a Virtual Environment (Recommended)
```bash
python -m venv .venv

# Windows:
.venv\Scripts\activate

# macOS / Linux:
source .venv/bin/activate
```

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

### 5. Launch the Application
```bash
streamlit run app.py
```
The application will launch automatically in your default browser at `http://localhost:8501`.

---

## 🧪 Running Automated Tests
The repository includes an exhaustive test suite covering data integrity, statistical methods, ML models, and configuration integrity:

```bash
pytest -v
```

Expected Output:
```text
tests/test_analysis.py .........                                         [ 45%]
tests/test_app.py .....                                                  [ 70%]
tests/test_data.py ......                                                [100%]
======================= 20 passed in 10.39s ========================
```

---

## 📊 Analytical Methodology

1. **Random Forest Feature Importance**:
   - Trains an ensemble of 100 decision trees to determine non-linear predictors of performance rating (e.g. project completion rate, attendance, peer reviews).
2. **One-Way ANOVA (Analysis of Variance)**:
   - Tests whether mean performance rating varies significantly across job roles and departmental tiers ($F$-statistic & $p$-value calibration).
3. **Gender Parity Two-Sample Independent $t$-Test**:
   - Assesses demographic calibration equity to verify no systemic rating disparity exists across genders.
4. **OLS Linear Regression**:
   - Models salary-to-rating elasticity, determining the expected appraisal rating impact per compensation increment.
5. **Promotion Readiness Index (0–100)**:
   - Composite scoring combining tenure, appraisal ratings, project execution velocity, and training certifications completed.

---

## 📄 License & Academic Integrity
This project is developed as part of the MCA Major Capstone curriculum at **ABES Engineering College, Ghaziabad** by **Aayush Mathur (2500320140001)**. Distributed under the MIT License.
