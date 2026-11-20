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
    # ── Triangular Cohort Retention Heatmap (%) Matching Reference Design ──
    cohorts_available = sorted(filtered_view["signup_cohort_month"].dropna().unique())
    # Display 8 cohorts for clean triangular visualization
    display_cohorts = [c for c in cohorts_available if c.startswith("2024")][:8]
    if len(display_cohorts) < 3:
        display_cohorts = cohorts_available[:8]
    
    n_cohorts = len(display_cohorts)
    x_labels = [f"M{i}" for i in range(n_cohorts)]

    z_vals = []
    text_vals = []
    annotations = []

    for i, c in enumerate(display_cohorts):
        c_df = filtered_view[filtered_view["signup_cohort_month"] == c]
        tot = len(c_df)
        row_z = []
        row_txt = []
        avail = n_cohorts - i
        for m in range(n_cohorts):
            if m < avail:
                if m == 0:
                    pct = 100
                else:
                    surv = c_df[c_df["tenure_months"] >= m]
                    pct = int(round(len(surv) / tot * 100)) if tot > 0 else 0
                row_z.append(pct)
                row_txt.append(str(pct))
                
                # High-contrast cell annotations
                text_color = "white" if pct >= 65 else "#111827"
                annotations.append(dict(
                    x=x_labels[m],
                    y=c,
                    text=f"<b>{pct}</b>",
                    font=dict(color=text_color, size=13, family="Arial, sans-serif"),
                    showarrow=False
                ))
            else:
                row_z.append(None)
                row_txt.append("")
        z_vals.append(row_z)
        text_vals.append(row_txt)

    # Reference Greens color palette (pale mint green to deep forest green)
    greens_scale = [
        [0.0, "#E8F5E9"],
        [0.2, "#C8E6C9"],
        [0.4, "#81C784"],
        [0.6, "#4CAF50"],
        [0.8, "#2E7D32"],
        [1.0, "#1B5E20"]
    ]

    fig_heat = go.Figure(data=go.Heatmap(
        z=z_vals,
        x=x_labels,
        y=display_cohorts,
        colorscale=greens_scale,
        zmin=0,
        zmax=100,
        showscale=False,
        xgap=2,
        ygap=2,
        hoverongaps=False
    ))

    fig_heat.update_layout(
        title=dict(
            text="<b>Cohort Retention Heatmap (%)</b>",
            x=0.5,
            xanchor="center",
            font=dict(size=18, family="Arial, sans-serif", color="#111827")
        ),
        xaxis=dict(
            title=dict(text="<b>Months Since Signup</b>", font=dict(size=14, color="#111827")),
            side="bottom",
            showgrid=False,
            linecolor="#111827",
            linewidth=1,
            mirror=True,
            ticks="outside"
        ),
        yaxis=dict(
            title=dict(text="<b>Cohort</b>", font=dict(size=14, color="#111827")),
            autorange="reversed",
            showgrid=False,
            linecolor="#111827",
            linewidth=1,
            mirror=True,
            ticks="outside"
        ),
        annotations=annotations,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=70, r=40, t=70, b=60),
        height=480
    )

    col1, col2 = st.columns([3, 2])

    with col1:
        st.subheader("🗓 Triangular Retention Heatmap (%)")
        render_chart_wrapper(fig_heat, height=480)
        render_methodology_drawer(
            title="Cohort Decay Matrix",
            formula="Retention Rate % at Mk = (COUNT(active at tenure >= k) / initial cohort size) * 100",
            source="dim_customers.signup_cohort_month, fact_subscriptions.tenure_months",
            sample_size=f"{len(filtered_view)} customers across {len(display_cohorts)} signup cohorts",
            assumptions="Triangular survival matrix; future lifecycle months are masked out as cohorts mature."
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
