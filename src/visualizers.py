import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np

# Corporate Design Palette
PALETTE = {
    "primary": "#2563EB",
    "secondary": "#4F46E5",
    "accent": "#0D9488",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "success": "#10B981",
    "neutral_dark": "#1E293B",
    "neutral_light": "#F8FAFC",
    "grid": "#E2E8F0"
}

TIER_COLORS = {
    "Low": "#10B981",      # Emerald
    "Medium": "#F59E0B",   # Amber
    "High": "#F97316",     # Orange
    "Severe": "#EF4444"    # Rose Red
}

def standard_layout(fig, title: str, x_title: str = "", y_title: str = ""):
    """Applies a clean, responsive layout template to Plotly figures."""
    fig.update_layout(
        title={"text": f"<b>{title}</b>", "font": {"size": 15, "color": PALETTE["neutral_dark"]}},
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=40, r=40, t=50, b=40),
        font=dict(family="system-ui, -apple-system, sans-serif", size=12, color=PALETTE["neutral_dark"]),
        xaxis=dict(
            title=x_title,
            showgrid=True,
            gridcolor=PALETTE["grid"],
            gridwidth=1,
            zeroline=False
        ),
        yaxis=dict(
            title=y_title,
            showgrid=True,
            gridcolor=PALETTE["grid"],
            gridwidth=1,
            zeroline=False
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    return fig

def plot_monthly_trend(df: pd.DataFrame) -> go.Figure:
    """Monthly aggregate lost hours and financial impact trend line."""
    if df.empty or "absence_date" not in df.columns:
        return go.Figure()

    df_monthly = df.copy()
    df_monthly["month_period"] = df_monthly["absence_date"].dt.to_period("M").astype(str)
    grouped = df_monthly.groupby("month_period").agg(
        total_hours=("absent_hours", "sum"),
        total_cost=("friction_loss_total", "sum")
    ).reset_index()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=grouped["month_period"],
        y=grouped["total_hours"],
        name="Lost Hours",
        marker_color="rgba(37, 99, 235, 0.75)",
        yaxis="y"
    ))
    fig.add_trace(go.Scatter(
        x=grouped["month_period"],
        y=grouped["total_cost"],
        name="Total Financial Loss ($)",
        mode="lines+markers",
        line=dict(color=PALETTE["danger"], width=3),
        marker=dict(size=7),
        yaxis="y2"
    ))

    fig.update_layout(
        title="<b>Monthly Absence Trajectory & Financial Friction</b>",
        xaxis=dict(title="Month"),
        yaxis=dict(title="Lost Hours", showgrid=True, gridcolor=PALETTE["grid"]),
        yaxis2=dict(
            title="Total Financial Loss ($)",
            overlaying="y",
            side="right",
            showgrid=False
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=40, r=50, t=50, b=40)
    )
    return fig

def plot_department_hotspots(df: pd.DataFrame) -> go.Figure:
    """Horizontal bar chart ranking departments by total hours lost."""
    if df.empty or "department_name" not in df.columns:
        return go.Figure()

    dept_agg = df.groupby("department_name").agg(
        lost_hours=("absent_hours", "sum"),
        loss_cost=("friction_loss_total", "sum")
    ).reset_index().sort_values(by="lost_hours", ascending=True)

    fig = go.Figure(go.Bar(
        x=dept_agg["lost_hours"],
        y=dept_agg["department_name"],
        orientation="h",
        marker=dict(
            color=dept_agg["lost_hours"],
            colorscale="Blues",
            showscale=False
        ),
        text=dept_agg.apply(lambda r: f"{r['lost_hours']:.0f}h (${r['loss_cost']:,.0f})", axis=1),
        textposition="auto"
    ))
    return standard_layout(fig, "Departmental Absence Hotspots & Financial Friction", "Lost Hours", "Department")

def plot_bradford_scatter(df_bradford: pd.DataFrame) -> go.Figure:
    """2D bubble scatter plot: Spells (X) vs Days (Y), sized by Bradford Score."""
    if df_bradford.empty:
        return go.Figure()

    fig = go.Figure()
    for tier, color in TIER_COLORS.items():
        subset = df_bradford[df_bradford["bradford_tier"] == tier]
        if not subset.empty:
            fig.add_trace(go.Scatter(
                x=subset["spells"],
                y=subset["total_days"],
                mode="markers",
                name=f"{tier} Tier",
                marker=dict(
                    size=np.clip(subset["bradford_score"] / 20.0, 10, 50),
                    color=color,
                    opacity=0.75,
                    line=dict(width=1.5, color="#FFFFFF")
                ),
                text=subset.apply(
                    lambda r: f"<b>{r['employee_full_name']}</b><br>Score: {r['bradford_score']:.0f}<br>Spells: {r['spells']}<br>Days: {r['total_days']:.1f}<br>Dept: {r['department_name']}",
                    axis=1
                ),
                hoverinfo="text"
            ))

    return standard_layout(fig, "Bradford Disruption Landscape (Spells vs Total Days)", "Absence Spells (Frequency)", "Total Absent Days (Duration)")

def plot_pareto_curve(pareto_df: pd.DataFrame) -> go.Figure:
    """Lorenz curve plotting cumulative % of workforce vs cumulative % of lost hours."""
    if pareto_df.empty:
        return go.Figure()

    fig = go.Figure()
    # 80/20 Equality baseline line
    fig.add_trace(go.Scatter(
        x=[0, 100], y=[0, 100],
        mode="lines",
        name="Perfect Equality",
        line=dict(color="#94A3B8", dash="dash", width=1.5)
    ))
    # Actual cumulative curve
    fig.add_trace(go.Scatter(
        x=[0] + pareto_df["cum_emp_pct"].tolist(),
        y=[0] + pareto_df["cum_hours_pct"].tolist(),
        mode="lines+markers",
        name="Workforce Distribution",
        line=dict(color=PALETTE["primary"], width=3),
        marker=dict(size=5)
    ))

    # Add 20% vertical marker
    fig.add_vline(x=20, line_width=1.5, line_dash="dot", line_color=PALETTE["danger"],
                  annotation_text="Top 20% Workforce", annotation_position="top left")

    fig = standard_layout(fig, "Pareto Lorenz Concentration Analysis (80/20 Rule)", "Cumulative % of Workforce", "Cumulative % of Lost Hours")
    fig.update_xaxes(range=[0, 100])
    fig.update_yaxes(range=[0, 100])
    return fig

def plot_icd10_treemap(df: pd.DataFrame) -> go.Figure:
    """Hierarchical treemap of absence medical diagnostic reasons."""
    if df.empty or "category_type" not in df.columns:
        return go.Figure()

    treemap_df = df.groupby(["category_type", "reason_description"]).agg(
        hours=("absent_hours", "sum"),
        cost=("friction_loss_total", "sum")
    ).reset_index()

    fig = px.treemap(
        treemap_df,
        path=["category_type", "reason_description"],
        values="hours",
        color="cost",
        color_continuous_scale="Reds",
        title="<b>ICD-10 Diagnostic Classification Treemap</b>"
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=40, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="system-ui, -apple-system, sans-serif")
    )
    return fig

def plot_commute_scatter(df: pd.DataFrame, reg_stats: dict) -> go.Figure:
    """Scatter plot of commute distance vs absent hours with OLS trendline."""
    if df.empty or "distance_to_work_km" not in df.columns:
        return go.Figure()

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df["distance_to_work_km"],
        y=df["absent_hours"],
        mode="markers",
        name="Absence Events",
        marker=dict(color=PALETTE["accent"], size=7, opacity=0.6)
    ))

    # OLS Trendline
    if reg_stats.get("sample_size", 0) > 5:
        x_min = df["distance_to_work_km"].min()
        x_max = df["distance_to_work_km"].max()
        y_min = reg_stats["slope"] * x_min + reg_stats["intercept"]
        y_max = reg_stats["slope"] * x_max + reg_stats["intercept"]
        fig.add_trace(go.Scatter(
            x=[x_min, x_max],
            y=[y_min, y_max],
            mode="lines",
            name=f"OLS Trendline (r={reg_stats.get('pearson_r', 0):.2f})",
            line=dict(color=PALETTE["danger"], width=2.5)
        ))

    return standard_layout(fig, "Commute Distance Burden vs Absence Duration", "Distance to Work (km)", "Absent Hours")

def plot_bmi_box(df: pd.DataFrame) -> go.Figure:
    """Boxplot of absent hours across BMI categories."""
    if df.empty or "bmi_category" not in df.columns:
        return go.Figure()

    order = ["Normal", "Overweight", "Obese"]
    filtered_df = df[df["bmi_category"].isin(order)]

    fig = px.box(
        filtered_df,
        x="bmi_category",
        y="absent_hours",
        category_orders={"bmi_category": order},
        color="bmi_category",
        color_discrete_map={"Normal": "#10B981", "Overweight": "#F59E0B", "Obese": "#EF4444"},
        points="all"
    )
    fig.update_layout(showlegend=False)
    return standard_layout(fig, "Absence Duration Distribution across BMI Cohorts", "BMI Category", "Absent Hours")

def plot_loss_waterfall(direct_loss: float, friction_loss: float) -> go.Figure:
    """Waterfall chart visualizing direct payroll wage loss and operational overtime friction."""
    friction_delta = round(friction_loss - direct_loss, 2)
    fig = go.Figure(go.Waterfall(
        name="Financial Loss Breakdown",
        orientation="v",
        measure=["relative", "relative", "total"],
        x=["Direct Base Wage Loss", "Overtime & Friction (+20%)", "Total Organization Loss"],
        y=[direct_loss, friction_delta, friction_loss],
        text=[f"${direct_loss:,.0f}", f"+${friction_delta:,.0f}", f"${friction_loss:,.0f}"],
        textposition="outside",
        connector={"line": {"color": "#94A3B8"}},
        decreasing={"marker": {"color": PALETTE["success"]}},
        increasing={"marker": {"color": PALETTE["danger"]}},
        totals={"marker": {"color": PALETTE["primary"]}}
    ))
    return standard_layout(fig, "Financial Loss Waterfall: Wage Loss vs Operational Friction", "Cost Component", "Loss ($)")
