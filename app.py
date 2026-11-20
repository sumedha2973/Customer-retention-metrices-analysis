"""
app.py — Screen UI-01: Executive Overview
Customer Retention Metrics Analysis (CRMA)
Sumedha | Roll 2503201450146 | MCA 2nd Yr | ABES
"""
import streamlit as st
import plotly.graph_objects as go
from src.data_loader import load_analytical_view, load_table
from src.ui import render_header, render_kpi_row
from src.components import render_methodology_drawer, render_chart_wrapper
from src.filters import render_sidebar_filters
from src.metrics import calculate_churn_retention_rates, calculate_monthly_retention_trend, calculate_mrr_movements

st.set_page_config(
    page_title="CRMA — Executive Overview",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 1. Render Header
render_header(
    title="Customer Retention Metrics Analysis (CRMA)",
    subtitle="Executive Retention & Recurring Revenue Cockpit | NoidaSoft Technologies Pvt. Ltd."
)

# 2. Load Data & Filters
raw_view = load_analytical_view()
df_subs = load_table("fact_subscriptions")
filtered_view = render_sidebar_filters(raw_view)

# Get filtered subscription IDs
active_cust_ids = filtered_view["customer_id"].unique()
filtered_subs = df_subs[df_subs["customer_id"].isin(active_cust_ids)].copy()

# 3. Compute Top Headline KPIs for Analysis Month
curr_month = st.session_state.get("filter_month", "2025-12")
# Previous month for deltas
yr, mo = map(int, curr_month.split("-"))
prev_mo = 12 if mo == 1 else mo - 1
prev_yr = yr - 1 if mo == 1 else yr
prev_month = f"{prev_yr}-{prev_mo:02d}"

kpi_curr = calculate_churn_retention_rates(filtered_subs, curr_month)
kpi_prev = calculate_churn_retention_rates(filtered_subs, prev_month)

delta_ccr = round(kpi_curr["ccr"] - kpi_prev["ccr"], 2)
delta_crr = round(kpi_curr["crr"] - kpi_prev["crr"], 2)

mrr_df = calculate_mrr_movements(filtered_subs)
curr_mrr_row = mrr_df[mrr_df["month"] == curr_month]
churned_mrr = curr_mrr_row["churned_mrr"].values[0] if not curr_mrr_row.empty else 0.0
curr_nrr = curr_mrr_row["nrr"].values[0] if not curr_mrr_row.empty else 100.0

# Render 5 Tailwind Cards
metrics_cards = [
    {
        "title": "Customer Churn (CCR)",
        "value": f"{kpi_curr['ccr']:.1f}%",
        "delta": f"{'▲' if delta_ccr > 0 else '▼'} {abs(delta_ccr):.1f} pp",
        "status": "negative" if delta_ccr > 0 else "positive"
    },
    {
        "title": "Customer Retention (CRR)",
        "value": f"{kpi_curr['crr']:.1f}%",
        "delta": f"{'▲' if delta_crr >= 0 else '▼'} {abs(delta_crr):.1f} pp",
        "status": "positive" if delta_crr >= 0 else "negative"
    },
    {
        "title": "MRR Lost to Churn",
        "value": f"₹{churned_mrr:,.0f}",
        "delta": f"{kpi_curr['churned']} accounts lost",
        "status": "negative" if churned_mrr > 0 else "positive"
    },
    {
        "title": "Net Revenue Retention",
        "value": f"{curr_nrr:.1f}%",
        "delta": "Net expansion" if curr_nrr >= 100 else "Net contraction",
        "status": "positive" if curr_nrr >= 100 else "negative"
    },
    {
        "title": "Active Subscriptions",
        "value": f"{kpi_curr['starting_active'] - kpi_curr['churned']}",
        "delta": f"of {len(active_cust_ids)} filtered",
        "status": "neutral"
    }
]
render_kpi_row(metrics_cards)

st.markdown("---")

# 4. Charts Section
col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("📈 24-Month Churn vs Retention Trajectory")
    trend_df = calculate_monthly_retention_trend(filtered_subs)

    fig_ret = go.Figure()
    fig_ret.add_trace(go.Scatter(
        x=trend_df["month"], y=trend_df["crr"],
        name="Retention Rate (CRR %)",
        line=dict(color="#059669", width=3),
        mode="lines+markers"
    ))
    fig_ret.add_trace(go.Scatter(
        x=trend_df["month"], y=trend_df["ccr"],
        name="Churn Rate (CCR %)",
        line=dict(color="#E11D48", width=2, dash="dash"),
        mode="lines+markers"
    ))
    fig_ret.update_layout(
        yaxis=dict(title="Percentage (%)", range=[0, 105]),
        xaxis=dict(title="Calendar Month", tickangle=-45),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    render_chart_wrapper(fig_ret)
    render_methodology_drawer(
        title="CRR & CCR Computation",
        formula="CRR % = 100.0 - (COUNT(churned_in_month) / COUNT(started_active_before_month) * 100.0)",
        source="fact_subscriptions (subscription_status, start_date, churn_date)",
        sample_size=f"{len(filtered_subs)} subscription contract terms",
        assumptions="Evaluated on closed cohorts at calendar month boundaries matching RULE-01/02."
    )

with col_right:
    st.subheader("💰 Monthly Recurring Revenue (MRR) Trends")
    fig_mrr = go.Figure()
    fig_mrr.add_trace(go.Scatter(
        x=mrr_df["month"], y=mrr_df["gross_mrr"],
        name="Gross Active MRR",
        line=dict(color="#2B4C7E", width=3),
        mode="lines"
    ))
    fig_mrr.add_trace(go.Bar(
        x=mrr_df["month"], y=mrr_df["churned_mrr"],
        name="Churned MRR (Lost)",
        marker_color="#E11D48",
        opacity=0.75
    ))
    fig_mrr.add_trace(go.Scatter(
        x=mrr_df["month"], y=mrr_df["net_mrr"],
        name="Net MRR",
        line=dict(color="#059669", width=2, dash="dot"),
        mode="lines"
    ))
    fig_mrr.update_layout(
        yaxis=dict(title="Revenue in INR (₹)"),
        xaxis=dict(title="Calendar Month", tickangle=-45),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    render_chart_wrapper(fig_mrr)
    render_methodology_drawer(
        title="MRR Revenue Dynamics",
        formula="Net MRR = Gross Active MRR - Churned MRR; NRR % = (Net MRR / Starting MRR) * 100",
        source="fact_subscriptions.monthly_fee_inr",
        sample_size=f"₹{mrr_df['gross_mrr'].iloc[-1]:,.0f} active revenue pool",
        assumptions="Contract prices normalized to monthly recurring equivalents matching RULE-03."
    )
