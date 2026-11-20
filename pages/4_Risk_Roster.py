"""
pages/4_Risk_Roster.py — Screen UI-05: Proactive Health Risk Roster
Customer Retention Metrics Analysis (CRMA)
"""
import streamlit as st
import datetime
import pandas as pd
from src.data_loader import load_analytical_view
from src.filters import render_sidebar_filters
from src.risk_scoring import compute_composite_risk_score
from src.ui import render_kpi_row, render_risk_badge

st.set_page_config(page_title="CRMA — Risk Roster", page_icon="⚠️", layout="wide")

st.title("⚠️ Proactive Customer Health Risk Roster")
st.caption("Actionable account roster prioritizing CS intervention by composite risk score (RULE-05).")

raw_view = load_analytical_view()
filtered_view = render_sidebar_filters(raw_view)

# Compute real-time risk scores
scored_df = compute_composite_risk_score(filtered_view)

# Filter for active subscriptions only
active_roster = scored_df[scored_df["is_churned"] == 0].copy()

# Summary Count Cards
high_count = len(active_roster[active_roster["calc_risk_tier"] == "High"])
med_count = len(active_roster[active_roster["calc_risk_tier"] == "Medium"])
low_count = len(active_roster[active_roster["calc_risk_tier"] == "Low"])

cards = [
    {"title": "High Risk Accounts (>= 70)", "value": str(high_count), "delta": "Urgent intervention", "status": "negative"},
    {"title": "Medium Risk Accounts (40-69)", "value": str(med_count), "delta": "Monitor closely", "status": "neutral"},
    {"title": "Healthy / Low Risk (< 40)", "value": str(low_count), "delta": "Stable accounts", "status": "positive"},
    {"title": "Total Active Roster", "value": str(len(active_roster)), "delta": "Filtered cohort", "status": "neutral"}
]
render_kpi_row(cards)

st.markdown("---")

# Slicers within page
col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
with col_f1:
    tier_choice = st.multiselect("Filter Risk Tier", ["High", "Medium", "Low"], default=["High", "Medium"])
with col_f2:
    min_score = st.slider("Minimum Risk Score", 0.0, 100.0, 40.0, 5.0)

# Apply page filters
display_df = active_roster[
    (active_roster["calc_risk_tier"].isin(tier_choice)) &
    (active_roster["calc_risk_score"] >= min_score)
].sort_values("calc_risk_score", ascending=False)

# Export Button
with col_f3:
    st.write("")
    st.write("")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_bytes = display_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Download At-Risk Roster (CSV)",
        data=csv_bytes,
        file_name=f"at_risk_customers_{timestamp}.csv",
        mime="text/csv",
        use_container_width=True
    )

st.subheader(f"📋 Prioritized Intervention List ({len(display_df)} Accounts)")

# Display polished table
show_cols = [
    "customer_id", "company_name", "plan_name", "industry", "region",
    "calc_risk_score", "calc_risk_tier", "tenure_months", "monthly_fee_inr",
    "total_support_tickets", "avg_monthly_login_days", "latest_csat_score"
]
rename_map = {
    "customer_id": "Customer ID",
    "company_name": "Company",
    "plan_name": "Plan",
    "industry": "Industry",
    "region": "Region",
    "calc_risk_score": "Risk Score",
    "calc_risk_tier": "Risk Tier",
    "tenure_months": "Tenure (Mo)",
    "monthly_fee_inr": "Monthly Fee (₹)",
    "total_support_tickets": "Tickets",
    "avg_monthly_login_days": "Avg Logins/Mo",
    "latest_csat_score": "Latest CSAT"
}
table_data = display_df[show_cols].rename(columns=rename_map)

st.dataframe(
    table_data.style.background_gradient(subset=["Risk Score"], cmap="Reds"),
    use_container_width=True,
    height=450
)
