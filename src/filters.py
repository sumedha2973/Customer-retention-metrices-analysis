"""
src/filters.py — Shared Sidebar Filter State Manager
Customer Retention Metrics Analysis (CRMA)
Persists user filter selections across all multipage views via st.session_state.
"""
import streamlit as st
import pandas as pd


def init_global_filters():
    """Ensure filter state keys exist in st.session_state."""
    defaults = {
        "filter_plan": [],
        "filter_industry": [],
        "filter_region": [],
        "filter_channel": [],
        "filter_month": "2025-12",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render_sidebar_filters(df_view: pd.DataFrame) -> pd.DataFrame:
    """Render multiselect filter controls in sidebar and return filtered slice."""
    init_global_filters()

    with st.sidebar:
        st.markdown("### 🎛 Global Filter Slicers")

        # Month selector for overview
        all_months = sorted(list(set(
            [d[:7] for d in df_view["signup_cohort_month"].dropna()]
        )))
        if "2025-12" not in all_months:
            all_months.append("2025-12")
        all_months = sorted(all_months)

        default_idx = all_months.index("2025-12") if "2025-12" in all_months else len(all_months) - 1
        st.session_state.filter_month = st.selectbox(
            "📅 Analysis Month",
            options=all_months,
            index=default_idx
        )

        # Plan filter
        plans = sorted(df_view["plan_name"].dropna().unique().tolist())
        st.session_state.filter_plan = st.multiselect(
            "🏷 Plan Tier",
            options=plans,
            default=st.session_state.filter_plan
        )

        # Industry filter
        industries = sorted(df_view["industry"].dropna().unique().tolist())
        st.session_state.filter_industry = st.multiselect(
            "🏭 Industry",
            options=industries,
            default=st.session_state.filter_industry
        )

        # Region filter
        regions = sorted(df_view["region"].dropna().unique().tolist())
        st.session_state.filter_region = st.multiselect(
            "🌍 Region",
            options=regions,
            default=st.session_state.filter_region
        )

        # Reset button
        if st.button("🔄 Reset All Filters", use_container_width=True):
            st.session_state.filter_plan = []
            st.session_state.filter_industry = []
            st.session_state.filter_region = []
            st.session_state.filter_channel = []
            st.session_state.filter_month = "2025-12"
            st.rerun()

        st.markdown("---")
        st.caption("Customer Retention Metrics Analysis v1.0 | ABES MCA")

    # Apply filters
    filtered_df = df_view.copy()
    if st.session_state.filter_plan:
        filtered_df = filtered_df[filtered_df["plan_name"].isin(st.session_state.filter_plan)]
    if st.session_state.filter_industry:
        filtered_df = filtered_df[filtered_df["industry"].isin(st.session_state.filter_industry)]
    if st.session_state.filter_region:
        filtered_df = filtered_df[filtered_df["region"].isin(st.session_state.filter_region)]

    return filtered_df
