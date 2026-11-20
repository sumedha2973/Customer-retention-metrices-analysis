"""
src/components.py — Reusable UI Widgets & Chart Wrappers
Customer Retention Metrics Analysis (CRMA)
Provides methodology expander drawers, Plotly chart layout wrappers, and filter badges.
"""
import streamlit as st
import plotly.graph_objects as go


def render_methodology_drawer(title: str, formula: str, source: str, sample_size: str, assumptions: str):
    """Render a consistent, collapsible methodology expander beneath any chart."""
    with st.expander(f"ℹ️ Methodology & Technical Details — {title}", expanded=False):
        c1, c2 = st.columns([1, 1])
        with c1:
            st.markdown(f"**Mathematical / SQL Formula:**")
            st.code(formula, language="sql")
            st.markdown(f"**Data Lineage / Source Tables:** `{source}`")
        with c2:
            st.markdown(f"**Active Sample Size:** `{sample_size}`")
            st.markdown(f"**Filtering & Assumptions:** {assumptions}")


def render_chart_wrapper(fig: go.Figure, height: int = 420):
    """Standardize Plotly chart margins, background, and responsive sizing."""
    fig.update_layout(
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=height,
        font=dict(family="sans-serif", size=12, color="#1B365D"),
        hoverlabel=dict(bgcolor="#1B365D", font_color="#FFFFFF")
    )
    st.plotly_chart(fig, use_container_width=True)
