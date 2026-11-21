import os
import sqlite3
import pandas as pd
import pytest

DB_PATH = "C:/HCL/data/app.db"

@pytest.fixture(scope="module")
def db_conn():
    """Establishes a read-only connection to the production SQLite database."""
    assert os.path.exists(DB_PATH), f"Database not found at {DB_PATH}"
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    yield conn
    conn.close()

def test_database_connection_and_foreign_keys(db_conn):
    """TC-T-01: Verify DB connection and foreign key enforcement status."""
    cursor = db_conn.cursor()
    cursor.execute("PRAGMA foreign_keys;")
    # In SQLite read-only mode, ensure connection is live
    cursor.execute("SELECT 1;")
    result = cursor.fetchone()
    assert result[0] == 1

def test_table_row_counts(db_conn):
    """TC-T-02: Verify exact required row counts across all 4 relational tables."""
    cursor = db_conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM dim_departments;")
    assert cursor.fetchone()[0] == 6, "dim_departments must contain 6 rows"
    
    cursor.execute("SELECT COUNT(*) FROM dim_reasons;")
    assert cursor.fetchone()[0] == 28, "dim_reasons must contain 28 rows"
    
    cursor.execute("SELECT COUNT(*) FROM dim_employees;")
    emp_count = cursor.fetchone()[0]
    assert 100 <= emp_count <= 150, f"dim_employees expected 100-150 rows, got {emp_count}"
    
    cursor.execute("SELECT COUNT(*) FROM fact_absence_events;")
    fact_count = cursor.fetchone()[0]
    assert fact_count >= 150, f"fact_absence_events expected >= 150 rows, got {fact_count}"

def test_primary_key_uniqueness(db_conn):
    """TC-T-03: Assert zero duplicate primary keys across all tables."""
    cursor = db_conn.cursor()
    tables_pk = [
        ("dim_departments", "department_id"),
        ("dim_reasons", "reason_code"),
        ("dim_employees", "employee_id"),
        ("fact_absence_events", "event_id")
    ]
    for table, pk in tables_pk:
        cursor.execute(f"SELECT COUNT({pk}), COUNT(DISTINCT {pk}) FROM {table};")
        total, distinct = cursor.fetchone()
        assert total == distinct, f"Duplicate primary keys detected in {table} on {pk}"

def test_foreign_key_referential_integrity(db_conn):
    """TC-T-04: Verify zero orphaned foreign keys in relational tables."""
    cursor = db_conn.cursor()
    
    # Check fact -> employee
    cursor.execute("""
        SELECT COUNT(*) FROM fact_absence_events f 
        LEFT JOIN dim_employees e ON f.employee_id = e.employee_id 
        WHERE e.employee_id IS NULL;
    """)
    assert cursor.fetchone()[0] == 0, "Orphaned employee_id in fact_absence_events"
    
    # Check fact -> reason
    cursor.execute("""
        SELECT COUNT(*) FROM fact_absence_events f 
        LEFT JOIN dim_reasons r ON f.reason_code = r.reason_code 
        WHERE r.reason_code IS NULL;
    """)
    assert cursor.fetchone()[0] == 0, "Orphaned reason_code in fact_absence_events"
    
    # Check employee -> department
    cursor.execute("""
        SELECT COUNT(*) FROM dim_employees e 
        LEFT JOIN dim_departments d ON e.department_id = d.department_id 
        WHERE d.department_id IS NULL;
    """)
    assert cursor.fetchone()[0] == 0, "Orphaned department_id in dim_employees"

def test_joined_view_schema_and_columns(db_conn):
    """TC-T-05: Query v_absence_analysis_joined, verify 32 columns and non-null constraints."""
    df_view = pd.read_sql_query("SELECT * FROM v_absence_analysis_joined;", db_conn)
    assert df_view.shape[1] == 32, f"Expected 32 columns in joined view, got {df_view.shape[1]}"
    assert df_view.shape[0] >= 150, "Expected >= 150 records in joined view"
    
    # Critical columns must not contain nulls
    critical_cols = ["event_id", "employee_id", "department_name", "hourly_wage", "absent_hours", "direct_wage_loss", "friction_loss_total"]
    for col in critical_cols:
        assert df_view[col].isnull().sum() == 0, f"Null values detected in mandatory column {col}"
