"""
pages/1_Cohorts.py — Screen UI-02: Cohort Analysis
Customer Retention Metrics Analysis (CRMA)
"""
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from src.data_loader import load_analytical_view
from src.filters import render_sidebar_filters
from src.components import render_methodology_drawer, render_chart_wrapper
from src.metrics import build_cohort_survival_matrix

st.set_page_config(page_title="CRMA — Cohort Analysis", page_icon="👥", layout="wide")

st.title("👥 Cohort Survival & Retention Heatmap")
st.caption("Track signup-month customer cohorts across 1, 3, 6, and 12-month lifecycle decay curves.")

raw_view = load_analytical_view()
filtered_view = render_sidebar_filters(raw_view)

if filtered_view.empty:
    st.warning("⚠️ No customer records match the active filter criteria. Please broaden your selection.")
    st.stop()

# Build Cohort Matrix
matrix_df = build_cohort_survival_matrix(filtered_view)

if not matrix_df.empty:
    # Prepare matrix for heatmap
    heatmap_data = matrix_df.set_index("cohort")[["M+1", "M+3", "M+6", "M+12"]]

    col1, col2 = st.columns([3, 2])

    with col1:
        st.subheader("🗓 Triangular Survival Heatmap (%)")
        fig_heat = px.imshow(
            heatmap_data,
            labels=dict(x="Lifecycle Interval", y="Signup Cohort Month", color="Survival %"),
            x=["Month 1", "Month 3", "Month 6", "Month 12"],
            y=heatmap_data.index,
            color_continuous_scale="RdYlGn",
            range_color=[40, 100],
            text_auto=True,
            aspect="auto"
        )
        render_chart_wrapper(fig_heat, height=480)
        render_methodology_drawer(
            title="Cohort Decay Matrix",
            formula="Survival Rate % = (COUNT(customers_active_at_tenure_k) / COUNT(initial_cohort_size)) * 100",
            source="dim_customers.signup_cohort_month, fact_subscriptions.tenure_months",
            sample_size=f"{len(filtered_view)} customers across {len(matrix_df)} signup cohorts",
            assumptions="Closed cohort analysis; customers categorized by calendar month of initial signup."
        )

    with col2:
        st.subheader("📉 Cohort Decay Curves")
        fig_curve = go.Figure()
        # Plot top 5 largest cohorts
        sample_cohorts = matrix_df.sort_values("initial_size", ascending=False).head(5)
        for _, row in sample_cohorts.iterrows():
            fig_curve.add_trace(go.Scatter(
                x=[1, 3, 6, 12],
                y=[row["M+1"], row["M+3"], row["M+6"], row["M+12"]],
                mode="lines+markers",
                name=f"Cohort {row['cohort']} (N={row['initial_size']})"
            ))
        fig_curve.update_layout(
            xaxis=dict(title="Lifecycle Month", tickvals=[1, 3, 6, 12]),
            yaxis=dict(title="Survival Rate (%)", range=[30, 105]),
        )
        render_chart_wrapper(fig_curve, height=480)

    st.markdown("### 📋 Cohort Retention Summary Table")
    st.dataframe(matrix_df, use_container_width=True)
