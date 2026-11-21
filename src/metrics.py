import pandas as pd
import numpy as np

def calculate_bradford(spells: int, total_days: float) -> float:
    """
    Computes Bradford Factor index: B = S^2 * D
    Where S is number of absence spells/instances, and D is total days absent.
    """
    return float((spells ** 2) * total_days)

def classify_bradford_tier(score: float) -> str:
    """
    Categorizes Bradford Factor into HR organizational risk tiers:
    - Low: < 50 (Informal monitoring)
    - Medium: 50 - 249 (Line manager check-in)
    - High: 250 - 749 (Formal HR intervention / review)
    - Severe: >= 750 (Disciplinary review / corrective action plan)
    """
    if score < 50:
        return "Low"
    elif score <= 249:
        return "Medium"
    elif score <= 749:
        return "High"
    else:
        return "Severe"

def calculate_bradford_scores(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregates absence events per employee and calculates Bradford Factor and Risk Tier.
    """
    if df.empty:
        return pd.DataFrame(columns=[
            "employee_id", "employee_full_name", "department_name", "job_title",
            "spells", "total_days", "total_hours", "total_cost", "bradford_score", "bradford_tier"
        ])

    grouped = df.groupby(["employee_id", "employee_full_name", "department_name", "job_title"]).agg(
        spells=("event_id", "count"),
        total_days=("absent_days", "sum"),
        total_hours=("absent_hours", "sum"),
        total_cost=("friction_loss_total", "sum")
    ).reset_index()

    grouped["bradford_score"] = grouped.apply(
        lambda r: calculate_bradford(r["spells"], r["total_days"]), axis=1
    )
    grouped["bradford_tier"] = grouped["bradford_score"].apply(classify_bradford_tier)
    
    return grouped.sort_values(by="bradford_score", ascending=False).reset_index(drop=True)

def calculate_payroll_loss(df: pd.DataFrame, multiplier: float = 1.20) -> tuple:
    """
    Computes total direct wage loss and total friction loss using the 1.20x multiplier.
    Returns (direct_wage_loss, total_friction_loss).
    """
    if df.empty:
        return 0.0, 0.0
    direct = float(df["direct_wage_loss"].sum()) if "direct_wage_loss" in df.columns else float((df["absent_hours"] * df["hourly_wage"]).sum())
    total = float(direct * multiplier)
    return round(direct, 2), round(total, 2)

def simulate_savings(df: pd.DataFrame, reduction_pct: float, preventable_only: bool = True) -> dict:
    """
    Simulates financial recovery from What-If absentee reduction scenarios (5% - 25%).
    If preventable_only=True, applies reduction specifically to preventable/lifestyle absence reasons.
    """
    if df.empty:
        return {
            "baseline_direct_loss": 0.0,
            "baseline_friction_loss": 0.0,
            "target_pool_loss": 0.0,
            "direct_savings": 0.0,
            "friction_savings": 0.0,
            "total_net_savings": 0.0,
            "projected_loss": 0.0
        }

    direct_total = float(df["direct_wage_loss"].sum())
    friction_total = float(df["friction_loss_total"].sum())

    if preventable_only and "is_preventable" in df.columns:
        target_df = df[df["is_preventable"] == 1]
    else:
        target_df = df

    target_direct = float(target_df["direct_wage_loss"].sum())
    target_friction = float(target_df["friction_loss_total"].sum())

    factor = reduction_pct / 100.0
    direct_savings = target_direct * factor
    friction_savings = (target_friction - target_direct) * factor
    total_net_savings = direct_savings + friction_savings

    return {
        "baseline_direct_loss": round(direct_total, 2),
        "baseline_friction_loss": round(friction_total, 2),
        "target_pool_loss": round(target_friction, 2),
        "direct_savings": round(direct_savings, 2),
        "friction_savings": round(friction_savings, 2),
        "total_net_savings": round(total_net_savings, 2),
        "projected_loss": round(friction_total - total_net_savings, 2)
    }
