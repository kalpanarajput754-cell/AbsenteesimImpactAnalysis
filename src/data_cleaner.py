import pandas as pd
import numpy as np

EXPECTED_COLUMNS = [
    "event_id", "employee_id", "employee_full_name", "department_id", "department_name",
    "cost_center_code", "job_title", "hourly_wage", "age", "service_tenure_years",
    "distance_to_work_km", "daily_transit_cost", "bmi", "bmi_category", "is_smoker",
    "is_drinker", "children_count", "pet_count", "reason_code", "category_type",
    "icd_chapter_code", "reason_description", "is_preventable", "absence_date",
    "day_of_week", "month_of_year", "quarter", "year", "absent_hours", "absent_days",
    "direct_wage_loss", "friction_loss_total"
]

def validate_schema_assertions(df: pd.DataFrame) -> dict:
    """
    Validates structural schema, required column names, and critical non-null constraints.
    Returns status report dictionary.
    """
    issues = []
    
    # 1. Column presence check
    missing_cols = [c for c in EXPECTED_COLUMNS if c not in df.columns]
    if missing_cols:
        issues.append(f"Missing expected columns: {missing_cols}")
        
    # 2. Critical non-null constraints
    critical_cols = ["event_id", "employee_id", "hourly_wage", "absent_hours", "direct_wage_loss", "friction_loss_total"]
    for c in critical_cols:
        if c in df.columns and df[c].isnull().sum() > 0:
            issues.append(f"Nulls detected in mandatory column: {c} ({df[c].isnull().sum()} records)")

    # 3. Numeric bounds validation
    if "absent_hours" in df.columns:
        if (df["absent_hours"] <= 0).any() or (df["absent_hours"] > 160).any():
            issues.append("absent_hours contains out-of-bound values (<= 0 or > 160)")

    if "hourly_wage" in df.columns:
        if (df["hourly_wage"] <= 0).any():
            issues.append("hourly_wage contains non-positive values")

    return {
        "is_valid": len(issues) == 0,
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "issues": issues
    }

def audit_dataset_health(df: pd.DataFrame) -> dict:
    """
    Conducts comprehensive data health audit: missing rates, duplicates, and anomaly counts.
    """
    total_cells = df.size
    null_cells = df.isnull().sum().sum()
    duplicate_events = df.duplicated(subset=["event_id"]).sum() if "event_id" in df.columns else 0
    
    # Outlier detection via IQR on absent_hours
    if "absent_hours" in df.columns:
        q1 = df["absent_hours"].quantile(0.25)
        q3 = df["absent_hours"].quantile(0.75)
        iqr = q3 - q1
        outlier_count = ((df["absent_hours"] < (q1 - 1.5 * iqr)) | (df["absent_hours"] > (q3 + 1.5 * iqr))).sum()
    else:
        outlier_count = 0
        
    return {
        "completeness_score": round((1.0 - (null_cells / max(1, total_cells))) * 100.0, 2),
        "null_cells": int(null_cells),
        "duplicate_events": int(duplicate_events),
        "outlier_events": int(outlier_count),
        "total_records": len(df)
    }
