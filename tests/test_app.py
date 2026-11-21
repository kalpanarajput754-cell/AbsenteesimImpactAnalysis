import sys
import os
sys.path.insert(0, "C:/HCL")

import time
import sqlite3
import pandas as pd
import pytest

DB_PATH = "C:/HCL/data/app.db"

def test_data_loading_performance_latency():
    """TC-T-12: Measure database query response time, must execute in under 3.0 seconds."""
    start_time = time.time()
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    df = pd.read_sql_query("SELECT * FROM v_absence_analysis_joined;", conn)
    conn.close()
    elapsed = time.time() - start_time
    
    assert len(df) >= 150, "Dataframe must contain >= 150 rows"
    assert elapsed < 3.0, f"Query execution latency ({elapsed:.3f}s) exceeded 3.0s ceiling"
    print(f"\nQuery benchmark passed in {elapsed*1000:.1f} ms")

def test_ui_module_imports():
    """Verify src/ui.py imports and exposes required Tailwind component functions."""
    from src import ui
    assert hasattr(ui, "render_header")
    assert hasattr(ui, "render_kpi_card")
    assert hasattr(ui, "render_kpi_grid_4")
    assert hasattr(ui, "render_insight_box")

def test_metrics_integrity_aggregates():
    """Verify key financial aggregates in v_absence_analysis_joined."""
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    df = pd.read_sql_query("SELECT SUM(direct_wage_loss) as base_loss, SUM(friction_loss_total) as total_loss FROM v_absence_analysis_joined;", conn)
    conn.close()
    
    base_loss = df["base_loss"].iloc[0]
    total_loss = df["total_loss"].iloc[0]
    assert base_loss > 100000.0, "Base payroll loss should exceed $100k for 320 events"
    assert round(base_loss * 1.20, 0) == round(total_loss, 0), "1.20x friction multiplier must balance"
