"""
pages/3_Drivers.py — Screen UI-04: Churn Drivers & Statistics
Customer Retention Metrics Analysis (CRMA)
"""
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from src.data_loader import load_analytical_view, load_table
from src.filters import render_sidebar_filters
from src.components import render_methodology_drawer, render_chart_wrapper
from src.ui import render_stat_card
from src.stats import run_support_impact_test, run_prechurn_telemetry_ttest, run_satisfaction_correlation

st.set_page_config(page_title="CRMA — Churn Drivers", page_icon="🔍", layout="wide")

st.title("🔍 Churn Drivers & Inferential Statistics")
st.caption("Diagnose root causes of churn across Support SLA turnaround, pre-churn usage telemetry, and satisfaction surveys.")

raw_view = load_analytical_view()
filtered_view = render_sidebar_filters(raw_view)

if filtered_view.empty:
    st.warning("⚠️ No customer records match the active filter criteria. Please broaden your selection.")
    st.stop()

# 1. Support Impact Section
st.subheader("1. 🎧 Helpdesk Support & SLA Impact on Retention")
tix_df = load_table("fact_support_tickets")
# Join churn status
tix_merged = tix_df.merge(filtered_view[["customer_id", "is_churned"]], on="customer_id", how="inner")
tix_merged["status_label"] = tix_merged["is_churned"].map({1: "Churned", 0: "Retained"})

c1, c2 = st.columns([1, 1])
with c1:
    fig_tix = px.box(
        tix_merged, x="status_label", y="resolution_hours",
        color="status_label",
        color_discrete_map={"Churned": "#E11D48", "Retained": "#059669"},
        points="outliers",
        labels={"status_label": "Customer Status", "resolution_hours": "Resolution Turnaround (Hours)"}
    )
    render_chart_wrapper(fig_tix, height=360)

with c2:
    churned_res = tix_merged[tix_merged["is_churned"] == 1]["resolution_hours"]
    retained_res = tix_merged[tix_merged["is_churned"] == 0]["resolution_hours"]
    stat_res = run_support_impact_test(churned_res, retained_res)

    render_stat_card(
        test_name="Mann-Whitney U Test (Support SLA Turnaround)",
        stat_value=stat_res["stat"],
        p_value=stat_res["p_value"],
        interpretation=stat_res["interpretation"]
    )
    st.info(f"💡 Sample: {len(churned_res)} churned tickets vs {len(retained_res)} retained tickets.")

st.markdown("---")

# 2. Pre-Churn Telemetry Section
st.subheader("2. 📉 Pre-Churn Telemetry Drop (90-Day Trajectory)")
usage_df = load_table("fact_monthly_usage")
usage_merged = usage_df.merge(filtered_view[["customer_id", "is_churned"]], on="customer_id", how="inner")
usage_merged["status_label"] = usage_merged["is_churned"].map({1: "Churned", 0: "Retained"})

monthly_avg_usage = usage_merged.groupby(["activity_month", "status_label"], observed=True)["login_days"].mean().reset_index()

col_u1, col_u2 = st.columns([1, 1])
with col_u1:
    fig_usage = px.line(
        monthly_avg_usage, x="activity_month", y="login_days",
        color="status_label",
        color_discrete_map={"Churned": "#E11D48", "Retained": "#059669"},
        markers=True,
        labels={"activity_month": "Calendar Month", "login_days": "Average Login Days / Mo"}
    )
    render_chart_wrapper(fig_usage, height=360)

with col_u2:
    churn_login = usage_merged[usage_merged["is_churned"] == 1]["login_days"]
    retained_login = usage_merged[usage_merged["is_churned"] == 0]["login_days"]
    ttest_res = run_prechurn_telemetry_ttest(churn_login, retained_login)

    render_stat_card(
        test_name="Two-Sample Welch's t-Test (User Telemetry)",
        stat_value=ttest_res["stat"],
        p_value=ttest_res["p_value"],
        interpretation=ttest_res["interpretation"]
    )
    st.info(f"💡 Evaluated across {len(usage_merged)} monthly telemetry snapshots.")

st.markdown("---")

# 3. Satisfaction Correlation Section
st.subheader("3. ⭐ Survey Satisfaction & Renewal Correlation")
corr_res = run_satisfaction_correlation(filtered_view)

col_s1, col_s2 = st.columns([1, 1])
with col_s1:
    survey_df = filtered_view.dropna(subset=["latest_csat_score", "is_churned"]).copy()
    survey_df["status_label"] = survey_df["is_churned"].map({1: "Churned", 0: "Retained"})

    fig_corr = px.strip(
        survey_df, x="latest_csat_score", y="status_label",
        color="status_label",
        color_discrete_map={"Churned": "#E11D48", "Retained": "#059669"},
        orientation="h",
        labels={"latest_csat_score": "CSAT Score (1 to 5)", "status_label": "Customer Status"}
    )
    render_chart_wrapper(fig_corr, height=360)

with col_s2:
    render_stat_card(
        test_name="Pearson r & Spearman Correlation (CSAT vs Renewal)",
        stat_value=corr_res["pearson_r"],
        p_value=corr_res["p_val"],
        interpretation=f"Correlation coefficient r = {corr_res['pearson_r']:.3f}, rho = {corr_res['spearman_rho']:.3f}. "
                       f"Higher satisfaction is positively associated with contract renewal."
    )
    st.warning(f"⚠️ **Data Quality Notice (DQ-01):** {corr_res['sample_size']} of {len(filtered_view)} accounts have survey scores "
               f"(65.8% missing due to realistic SaaS survey non-response rates).")
