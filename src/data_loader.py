import os
import sqlite3
from pathlib import Path
import pandas as pd
import streamlit as st

# Portable path resolution for local Windows and Streamlit Community Cloud (Linux)
BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "app.db"

def get_database_path() -> str:
    """Resolves absolute path to SQLite database across environments."""
    if DB_PATH.exists():
        return str(DB_PATH)
    alt_path = Path("C:/HCL/data/app.db")
    if alt_path.exists():
        return str(alt_path)
    return str(DB_PATH)

def get_db_connection():
    """Establishes a read-only SQLite database connection."""
    db_path = get_database_path()
    return sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)

@st.cache_data(ttl=3600, show_spinner=False)
def load_clean_data(departments=None, reason_categories=None, date_range=None) -> pd.DataFrame:
    """
    Loads analytical view v_absence_analysis_joined from SQLite with optional filtering.
    Cached for high performance (< 100ms on warm cache).
    """
    conn = get_db_connection()
    query = "SELECT * FROM v_absence_analysis_joined WHERE 1=1"
    params = []

    if departments:
        placeholders = ",".join(["?"] * len(departments))
        query += f" AND department_name IN ({placeholders})"
        params.extend(departments)

    if reason_categories:
        placeholders = ",".join(["?"] * len(reason_categories))
        query += f" AND category_type IN ({placeholders})"
        params.extend(reason_categories)

    if date_range and len(date_range) == 2:
        start_date, end_date = date_range
        query += " AND absence_date BETWEEN ? AND ?"
        params.extend([str(start_date), str(end_date)])

    query += " ORDER BY absence_date ASC;"

    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    if "absence_date" in df.columns:
        df["absence_date"] = pd.to_datetime(df["absence_date"])

    return df

@st.cache_data(ttl=3600, show_spinner=False)
def get_filter_options():
    """Retrieves unique filter options for Streamlit sidebar controls."""
    conn = get_db_connection()
    depts = pd.read_sql_query("SELECT DISTINCT department_name FROM dim_departments ORDER BY department_name;", conn)["department_name"].tolist()
    cats = pd.read_sql_query("SELECT DISTINCT category_type FROM dim_reasons ORDER BY category_type;", conn)["category_type"].tolist()
    dates = pd.read_sql_query("SELECT MIN(absence_date) as min_d, MAX(absence_date) as max_d FROM fact_absence_events;", conn)
    conn.close()

    min_date = pd.to_datetime(dates["min_d"].iloc[0]).date()
    max_date = pd.to_datetime(dates["max_d"].iloc[0]).date()

    return {
        "departments": depts,
        "categories": cats,
        "min_date": min_date,
        "max_date": max_date
    }

@st.cache_data(ttl=3600, show_spinner=False)
def load_raw_table(table_name: str) -> pd.DataFrame:
    """Loads raw relational table for auditing and inspection."""
    valid_tables = ["dim_departments", "dim_reasons", "dim_employees", "fact_absence_events"]
    if table_name not in valid_tables:
        raise ValueError(f"Invalid table name {table_name}. Allowed: {valid_tables}")
    conn = get_db_connection()
    df = pd.read_sql_query(f"SELECT * FROM {table_name};", conn)
    conn.close()
    return df
