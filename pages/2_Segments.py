"""
pages/2_Segments.py — Screen UI-03: Segment Comparison
Customer Retention Metrics Analysis (CRMA)
"""
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from src.data_loader import load_analytical_view
from src.filters import render_sidebar_filters
from src.components import render_methodology_drawer, render_chart_wrapper
from src.metrics import calculate_clv_by_segment
from src.stats import run_channel_contingency_test

st.set_page_config(page_title="CRMA — Segment Comparison", page_icon="🔀", layout="wide")

st.title("🔀 Multi-Dimensional Segment Comparison")
st.caption("Compare customer retention rates, lifetime values, and acquisition channel longevity across customer profiles.")

raw_view = load_analytical_view()
filtered_view = render_sidebar_filters(raw_view)

if filtered_view.empty:
    st.warning("⚠️ No customer records match the active filter criteria. Please broaden your selection.")
    st.stop()

col_seg1, col_seg2 = st.columns([1, 1])

with col_seg1:
    st.subheader("📊 Retention Rate by Plan Tier")
    plan_ret = filtered_view.groupby("plan_name", observed=True).agg(
        total=("customer_id", "count"),
        churned=("is_churned", "sum")
    ).reset_index()
    plan_ret["retention_rate"] = ((plan_ret["total"] - plan_ret["churned"]) / plan_ret["total"] * 100).round(1)

    fig_plan = px.bar(
        plan_ret, x="plan_name", y="retention_rate",
        color="plan_name",
        text="retention_rate",
        color_discrete_sequence=["#2B4C7E", "#059669", "#D97706"],
        labels={"plan_name": "Plan Tier", "retention_rate": "Retention Rate (%)"}
    )
    fig_plan.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig_plan.update_layout(yaxis=dict(range=[0, 115]), showlegend=False)
    render_chart_wrapper(fig_plan)

with col_seg2:
    st.subheader("🏭 Retention Rate by Industry")
    ind_ret = filtered_view.groupby("industry", observed=True).agg(
        total=("customer_id", "count"),
        churned=("is_churned", "sum")
    ).reset_index()
    ind_ret["retention_rate"] = ((ind_ret["total"] - ind_ret["churned"]) / ind_ret["total"] * 100).round(1)

    fig_ind = px.bar(
        ind_ret, x="industry", y="retention_rate",
        text="retention_rate",
        color_discrete_sequence=["#1B365D"],
        labels={"industry": "Industry Sector", "retention_rate": "Retention Rate (%)"}
    )
    fig_ind.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig_ind.update_layout(yaxis=dict(range=[0, 115]))
    render_chart_wrapper(fig_ind)

st.markdown("---")

col_clv, col_chan = st.columns([1, 1])

with col_clv:
    st.subheader("💎 Customer Lifetime Value (CLV ₹) by Plan & Industry")
    clv_df = filtered_view.groupby(["plan_name", "industry"], observed=True).agg(
        avg_fee=("monthly_fee_inr", "mean"),
        avg_tenure=("tenure_months", "mean")
    ).reset_index()
    clv_df["clv"] = (clv_df["avg_fee"] * clv_df["avg_tenure"]).round(0)

    fig_clv = px.bar(
        clv_df, x="plan_name", y="clv", color="industry",
        barmode="group",
        labels={"plan_name": "Plan Tier", "clv": "Estimated CLV (₹)", "industry": "Industry"}
    )
    render_chart_wrapper(fig_clv)
    render_methodology_drawer(
        title="Customer Lifetime Value (CLV)",
        formula="CLV = AVG(monthly_fee_inr) * AVG(tenure_months)",
        source="fact_subscriptions (monthly_fee_inr, tenure_months)",
        sample_size=f"{len(filtered_view)} customer accounts",
        assumptions="Normalized across completed historical tenures and active contracts (RULE-04)."
    )

with col_chan:
    st.subheader("📡 Acquisition Channel Longevity")
    chan_df = filtered_view.groupby("acquisition_channel", observed=True).agg(
        total=("customer_id", "count"),
        churned=("is_churned", "sum")
    ).reset_index()
    chan_df["retention_rate"] = ((chan_df["total"] - chan_df["churned"]) / chan_df["total"] * 100).round(1)
    chan_df = chan_df.sort_values("retention_rate", ascending=True)

    fig_chan = px.bar(
        chan_df, y="acquisition_channel", x="retention_rate",
        orientation="h",
        text="retention_rate",
        color="retention_rate",
        color_continuous_scale="Blues",
        labels={"acquisition_channel": "Channel", "retention_rate": "Retention Rate (%)"}
    )
    fig_chan.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig_chan.update_layout(xaxis=dict(range=[0, 115]), coloraxis_showscale=False)
    render_chart_wrapper(fig_chan)

    # Statistical Significance Badge
    chi_res = run_channel_contingency_test(filtered_view)
    badge_color = "green" if chi_res["significant"] else "orange"
    st.markdown(
        f":{badge_color}[**Chi-Square Independence Test:** Chi2 = {chi_res['chi2']}, "
        f"p-value = {chi_res['p_value']} ({'Significant' if chi_res['significant'] else 'Not Significant'})]"
    )
    st.caption(chi_res["interpretation"])
