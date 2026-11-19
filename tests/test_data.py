"""
Test Suite: Data Integrity & Database Verification
Employee Performance Rating Analysis
Student: Aayush Mathur | Enrolment: 2500320140001 | Section: MCA-B | ABES Engineering College
Target: data/app.db, SQLite Tables, Primary & Foreign Keys, Analytics Views
"""

import os
import sqlite3
import pytest

# Determine absolute path to data/app.db
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DB_PATH = os.path.join(BASE_DIR, "data", "app.db")


@pytest.fixture(scope="module")
def db_conn():
    """Provides a SQLite connection fixture for test functions."""
    assert os.path.exists(DB_PATH), f"Database file does not exist at {DB_PATH}"
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    yield conn
    conn.close()


# -----------------------------------------------------------------------------
# TC-T-01: Database File & Connection Verification
# -----------------------------------------------------------------------------
def test_db_file_and_connection(db_conn):
    """TC-T-01: Verify SQLite database file exists and opens a valid connection."""
    assert os.path.isfile(DB_PATH), f"DB file missing at {DB_PATH}"
    assert os.path.getsize(DB_PATH) > 0, "DB file is empty (0 bytes)"
    cur = db_conn.cursor()
    cur.execute("SELECT 1;")
    res = cur.fetchone()[0]
    assert res == 1, "Failed to execute basic SELECT query on app.db"


# -----------------------------------------------------------------------------
# TC-T-02: Table Existence Verification
# -----------------------------------------------------------------------------
def test_table_existence(db_conn):
    """TC-T-02: Verify all 5 core normalized relational tables exist in SQLite master."""
    cur = db_conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
    existing_tables = {row[0] for row in cur.fetchall()}
    
    expected_tables = {
        "departments",
        "employees",
        "performance_reviews",
        "training_records",
        "attrition_records",
    }
    missing_tables = expected_tables - existing_tables
    assert not missing_tables, f"Missing required tables in SQLite: {missing_tables}"


# -----------------------------------------------------------------------------
# TC-T-03: Row Counts Threshold Verification
# -----------------------------------------------------------------------------
def test_row_counts_threshold(db_conn):
    """TC-T-03: Verify record volume thresholds (employees > 150, reviews > 300)."""
    cur = db_conn.cursor()
    
    # Employees count
    cur.execute("SELECT COUNT(*) FROM employees;")
    emp_count = cur.fetchone()[0]
    assert emp_count > 150, f"Expected >150 employees, found {emp_count}"
    
    # Performance reviews count
    cur.execute("SELECT COUNT(*) FROM performance_reviews;")
    review_count = cur.fetchone()[0]
    assert review_count > 300, f"Expected >300 performance reviews, found {review_count}"
    
    # Departments count
    cur.execute("SELECT COUNT(*) FROM departments;")
    dept_count = cur.fetchone()[0]
    assert dept_count >= 5, f"Expected >=5 departments, found {dept_count}"
    
    # Training records count
    cur.execute("SELECT COUNT(*) FROM training_records;")
    training_count = cur.fetchone()[0]
    assert training_count >= 150, f"Expected >=150 training records, found {training_count}"
    
    # Attrition records count
    cur.execute("SELECT COUNT(*) FROM attrition_records;")
    attrition_count = cur.fetchone()[0]
    assert attrition_count >= 150, f"Expected >=150 attrition records, found {attrition_count}"


# -----------------------------------------------------------------------------
# TC-T-04: Primary Key Uniqueness & Not-Null Integrity
# -----------------------------------------------------------------------------
def test_primary_key_uniqueness(db_conn):
    """TC-T-04: Verify Primary Key uniqueness across all core entity tables."""
    cur = db_conn.cursor()
    tables_pk = [
        ("departments", "dept_id"),
        ("employees", "emp_id"),
        ("performance_reviews", "review_id"),
        ("training_records", "training_id"),
        ("attrition_records", "attrition_id"),
    ]
    
    for table, pk in tables_pk:
        # Check total rows vs distinct PK count
        cur.execute(f"SELECT COUNT(*), COUNT(DISTINCT {pk}) FROM {table};")
        total, distinct_pk = cur.fetchone()
        assert total == distinct_pk, (
            f"Primary key {pk} violation in table {table}: {total} rows vs {distinct_pk} distinct IDs"
        )
        
        # Check for NULL primary keys
        cur.execute(f"SELECT COUNT(*) FROM {table} WHERE {pk} IS NULL;")
        null_count = cur.fetchone()[0]
        assert null_count == 0, f"Table {table} contains {null_count} NULL primary keys in column {pk}"


# -----------------------------------------------------------------------------
# TC-T-05: Foreign Key Referential Integrity
# -----------------------------------------------------------------------------
def test_foreign_key_referential_integrity(db_conn):
    """TC-T-05: Verify referential integrity between child tables and parent tables."""
    cur = db_conn.cursor()
    
    # 1. Employees dept_id references departments.dept_id
    cur.execute("""
        SELECT COUNT(*) FROM employees e
        WHERE e.dept_id IS NOT NULL 
          AND e.dept_id NOT IN (SELECT dept_id FROM departments);
    """)
    orphan_emp_dept = cur.fetchone()[0]
    assert orphan_emp_dept == 0, f"Found {orphan_emp_dept} employees referencing non-existent departments"
    
    # 2. Performance reviews emp_id references employees.emp_id
    cur.execute("""
        SELECT COUNT(*) FROM performance_reviews pr
        WHERE pr.emp_id NOT IN (SELECT emp_id FROM employees);
    """)
    orphan_reviews = cur.fetchone()[0]
    assert orphan_reviews == 0, f"Found {orphan_reviews} performance reviews referencing non-existent employees"
    
    # 3. Training records emp_id references employees.emp_id
    cur.execute("""
        SELECT COUNT(*) FROM training_records tr
        WHERE tr.emp_id NOT IN (SELECT emp_id FROM employees);
    """)
    orphan_trainings = cur.fetchone()[0]
    assert orphan_trainings == 0, f"Found {orphan_trainings} training records referencing non-existent employees"
    
    # 4. Attrition records emp_id references employees.emp_id
    cur.execute("""
        SELECT COUNT(*) FROM attrition_records ar
        WHERE ar.emp_id NOT IN (SELECT emp_id FROM employees);
    """)
    orphan_attrition = cur.fetchone()[0]
    assert orphan_attrition == 0, f"Found {orphan_attrition} attrition records referencing non-existent employees"


# -----------------------------------------------------------------------------
# TC-T-06: Analytical View Accessibility & Structure
# -----------------------------------------------------------------------------
def test_analytics_view_accessibility(db_conn):
    """TC-T-06: Verify v_employee_analytics view accessibility and presence of essential schema columns."""
    cur = db_conn.cursor()
    
    # Verify view exists
    cur.execute("SELECT name FROM sqlite_master WHERE type='view' AND name='v_employee_analytics';")
    view_name = cur.fetchone()
    assert view_name is not None, "View 'v_employee_analytics' does not exist in SQLite catalog"
    
    # Verify view returns records
    cur.execute("SELECT COUNT(*) FROM v_employee_analytics;")
    view_row_count = cur.fetchone()[0]
    assert view_row_count > 0, "View 'v_employee_analytics' returned 0 records"
    
    # Verify essential columns exist in view projection
    cur.execute("SELECT * FROM v_employee_analytics LIMIT 1;")
    sample_row = cur.fetchone()
    column_names = sample_row.keys()
    
    essential_columns = [
        "emp_id",
        "dept_name",
        "performance_rating",
        "review_year",
    ]
    for col in essential_columns:
        assert col in column_names, f"Expected column '{col}' missing from v_employee_analytics view"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
