import numpy as np
import pandas as pd
from scipy import stats
import pytest

def calculate_bradford(spells, total_days):
    """B = S^2 * D"""
    return (spells ** 2) * total_days

def classify_bradford_tier(score):
    if score < 50:
        return "Low"
    elif score <= 249:
        return "Medium"
    elif score <= 749:
        return "High"
    else:
        return "Severe"

def calculate_total_loss(absent_hours, hourly_wage, friction_factor=1.20):
    direct_loss = absent_hours * hourly_wage
    total_loss = direct_loss * friction_factor
    return round(direct_loss, 2), round(total_loss, 2)

def test_bradford_factor_formula_known_dataset():
    """TC-T-06: Verify Bradford Factor on deterministic known benchmark scenarios."""
    # Scenario A: Single long-term illness (1 spell of 10 days) -> Low Disruption
    score_a = calculate_bradford(spells=1, total_days=10.0)
    assert score_a == 10.0
    assert classify_bradford_tier(score_a) == "Low"
    
    # Scenario B: Multiple short leaves (5 spells of 2 days = 10 days) -> High Disruption
    score_b = calculate_bradford(spells=5, total_days=10.0)
    assert score_b == 250.0
    assert classify_bradford_tier(score_b) == "High"
    
    # Scenario C: Chronic repeated casual absents (8 spells of 2 days = 16 days) -> Severe
    score_c = calculate_bradford(spells=8, total_days=16.0)
    assert score_c == 1024.0
    assert classify_bradford_tier(score_c) == "Severe"

def test_financial_loss_formula_known_dataset():
    """TC-T-07: Verify direct wage loss and 1.20x friction multiplier on known numbers."""
    hours = 8.0
    wage = 50.00
    direct, total = calculate_total_loss(hours, wage, friction_factor=1.20)
    assert direct == 400.00
    assert total == 480.00
    
    # Verify decimal precision on odd wage
    hours_2 = 16.0
    wage_2 = 41.25
    direct_2, total_2 = calculate_total_loss(hours_2, wage_2, friction_factor=1.20)
    assert direct_2 == 660.00
    assert total_2 == 792.00

def test_bivariate_correlation_statistics():
    """TC-T-08: Verify Pearson and Spearman correlation functions."""
    np.random.seed(42)
    x = np.array([5.0, 10.0, 15.0, 25.0, 35.0, 45.0, 55.0])
    y = np.array([1.0, 2.0, 2.0, 3.0, 5.0, 6.0, 7.0]) # Positive trend
    
    r_val, p_val = stats.pearsonr(x, y)
    rho_val, sp_pval = stats.spearmanr(x, y)
    
    assert 0.8 <= r_val <= 1.0, f"Expected high positive correlation, got {r_val}"
    assert p_val < 0.05, f"Expected significant p-value, got {p_val}"
    assert 0.8 <= rho_val <= 1.0

def test_one_way_anova_bmi():
    """TC-T-09: Verify one-way ANOVA F-statistic calculation."""
    # Known cohorts
    normal_bmi_hours = [8.0, 8.0, 4.0, 8.0, 16.0]
    overweight_hours = [8.0, 16.0, 16.0, 24.0, 8.0]
    obese_hours = [24.0, 32.0, 40.0, 16.0, 24.0]
    
    f_stat, p_val = stats.f_oneway(normal_bmi_hours, overweight_hours, obese_hours)
    assert f_stat > 0, "F-statistic must be positive"
    assert p_val < 0.05, "Variance across cohorts should be statistically significant"

def test_chi_square_goodness_of_fit():
    """TC-T-10: Verify Chi-Square test detecting Monday/Friday surges."""
    # Observed vs uniform expected across 5 weekdays
    observed_counts = [141, 35, 34, 37, 73] # Mon, Tue, Wed, Thu, Fri
    total_events = sum(observed_counts)
    expected_counts = [total_events / 5] * 5
    
    chi2_stat, p_val = stats.chisquare(observed_counts, expected_counts)
    assert chi2_stat > 10.0, "Chi-Square statistic should be large due to spike"
    assert p_val < 0.001, "Rejects null hypothesis of uniform distribution"

def test_pareto_concentration_calculation():
    """TC-T-11: Verify 80/20 concentration ranking logic."""
    hours_per_emp = pd.Series([120, 95, 80, 75, 40, 20, 15, 10, 8, 5]) # 10 employees
    total_hours = hours_per_emp.sum() # 468
    top_2_hours = hours_per_emp.iloc[:2].sum() # 215 (45.9% by top 20%)
    
    cum_pct_hours = (hours_per_emp.cumsum() / total_hours).values
    assert cum_pct_hours[1] > 0.40, "Top 20% should concentrate over 40% of hours"
