import numpy as np
import pandas as pd
from scipy import stats

def compute_commute_correlations(df: pd.DataFrame) -> dict:
    """
    Computes Pearson and Spearman correlation between distance_to_work_km and absent_hours.
    Also fits an OLS regression line (slope, intercept).
    """
    if df.empty or "distance_to_work_km" not in df.columns or "absent_hours" not in df.columns:
        return {
            "pearson_r": 0.0, "pearson_p": 1.0,
            "spearman_rho": 0.0, "spearman_p": 1.0,
            "slope": 0.0, "intercept": 0.0, "sample_size": 0
        }

    clean_df = df[["distance_to_work_km", "absent_hours"]].dropna()
    if len(clean_df) < 5:
        return {
            "pearson_r": 0.0, "pearson_p": 1.0,
            "spearman_rho": 0.0, "spearman_p": 1.0,
            "slope": 0.0, "intercept": 0.0, "sample_size": len(clean_df)
        }

    x = clean_df["distance_to_work_km"].values
    y = clean_df["absent_hours"].values

    r_val, p_pearson = stats.pearsonr(x, y)
    rho_val, p_spearman = stats.spearmanr(x, y)

    # OLS linear regression slope & intercept
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)

    return {
        "pearson_r": round(float(r_val), 4),
        "pearson_p": float(p_pearson),
        "spearman_rho": round(float(rho_val), 4),
        "spearman_p": float(p_spearman),
        "slope": round(float(slope), 4),
        "intercept": round(float(intercept), 4),
        "r_squared": round(float(r_val ** 2), 4),
        "sample_size": int(len(clean_df))
    }

def run_bmi_anova(df: pd.DataFrame) -> dict:
    """
    Executes One-Way ANOVA across BMI Categories ('Normal', 'Overweight', 'Obese') on absent_hours.
    """
    if df.empty or "bmi_category" not in df.columns or "absent_hours" not in df.columns:
        return {"f_stat": 0.0, "p_value": 1.0, "is_significant": False, "cohort_summary": {}}

    cohorts = ["Normal", "Overweight", "Obese"]
    groups = []
    cohort_summary = {}

    for c in cohorts:
        subset = df[df["bmi_category"] == c]["absent_hours"].dropna()
        if len(subset) >= 2:
            groups.append(subset.values)
            cohort_summary[c] = {
                "count": int(len(subset)),
                "mean_hours": round(float(subset.mean()), 2),
                "std_hours": round(float(subset.std()), 2)
            }
        else:
            cohort_summary[c] = {"count": len(subset), "mean_hours": 0.0, "std_hours": 0.0}

    if len(groups) >= 2:
        f_stat, p_val = stats.f_oneway(*groups)
        f_stat = float(f_stat) if not np.isnan(f_stat) else 0.0
        p_val = float(p_val) if not np.isnan(p_val) else 1.0
    else:
        f_stat, p_val = 0.0, 1.0

    return {
        "f_stat": round(f_stat, 3),
        "p_value": float(p_val),
        "is_significant": bool(p_val < 0.05),
        "cohort_summary": cohort_summary
    }

def run_day_of_week_chisquare(df: pd.DataFrame) -> dict:
    """
    Runs Chi-Square goodness-of-fit on weekday absence counts against equal expected frequency.
    Calculates Monday/Friday surge percentage over Tuesday-Thursday midweek average.
    """
    weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    if df.empty or "day_of_week" not in df.columns:
        return {
            "chi2_stat": 0.0, "p_value": 1.0, "is_significant": False,
            "surge_pct": 0.0, "counts": {d: 0 for d in weekdays}
        }

    counts_series = df["day_of_week"].value_counts()
    observed = [int(counts_series.get(d, 0)) for d in weekdays]
    total_events = sum(observed)

    if total_events == 0:
        return {
            "chi2_stat": 0.0, "p_value": 1.0, "is_significant": False,
            "surge_pct": 0.0, "counts": {d: 0 for d in weekdays}
        }

    expected = [total_events / 5.0] * 5
    chi2_stat, p_val = stats.chisquare(observed, expected)

    # Surge Calculation: (Mon + Fri)/2 vs (Tue + Wed + Thu)/3
    mon_fri_avg = (observed[0] + observed[4]) / 2.0
    midweek_avg = max(1.0, (observed[1] + observed[2] + observed[3]) / 3.0)
    surge_pct = round(((mon_fri_avg - midweek_avg) / midweek_avg) * 100.0, 1)

    return {
        "chi2_stat": round(float(chi2_stat), 3),
        "p_value": float(p_val),
        "is_significant": bool(p_val < 0.05),
        "surge_pct": surge_pct,
        "counts": dict(zip(weekdays, observed))
    }

def compute_pareto_curve(df: pd.DataFrame) -> pd.DataFrame:
    """
    Computes Pareto 80/20 Lorenz curve data: cumulative percentage of employees
    versus cumulative percentage of total lost hours.
    """
    if df.empty or "employee_id" not in df.columns or "absent_hours" not in df.columns:
        return pd.DataFrame(columns=["rank", "employee_id", "hours", "cum_hours", "cum_hours_pct", "cum_emp_pct"])

    emp_hours = df.groupby(["employee_id", "employee_full_name"])["absent_hours"].sum().reset_index()
    emp_hours = emp_hours.sort_values(by="absent_hours", ascending=False).reset_index(drop=True)

    total_hours = emp_hours["absent_hours"].sum()
    n_employees = len(emp_hours)

    if total_hours == 0 or n_employees == 0:
        return pd.DataFrame()

    emp_hours["rank"] = emp_hours.index + 1
    emp_hours["cum_hours"] = emp_hours["absent_hours"].cumsum()
    emp_hours["cum_hours_pct"] = (emp_hours["cum_hours"] / total_hours) * 100.0
    emp_hours["cum_emp_pct"] = (emp_hours["rank"] / n_employees) * 100.0

    return emp_hours
