"""
pages/6_Methodology.py — Screen UI-08: Methodology & Data Dictionary
Customer Retention Metrics Analysis (CRMA)
"""
import streamlit as st
import pandas as pd

st.set_page_config(page_title="CRMA — Methodology", page_icon="📖", layout="wide")

st.title("📖 Analytical Methodology & Data Dictionary")
st.caption("Formal specifications, mathematical formulas, and relational data dictionaries for audit and verification.")

tab1, tab2, tab3 = st.tabs(["📐 KPI Formula Specifications", "🧪 Statistical Test Protocols", "🗄 Relational Data Dictionary"])

with tab1:
    st.subheader("Formal Business Metric Formulas (MSR-01 to MSR-11)")
    kpi_specs = [
        {"ID": "MSR-01", "Metric": "Customer Churn Rate (CCR %)", "Formula": "(COUNT(churned_customers) / COUNT(starting_active)) * 100", "Target": "< 2.5% / month", "OQ": "OQ-01"},
        {"ID": "MSR-02", "Metric": "Customer Retention Rate (CRR %)", "Formula": "100.0 - CCR %", "Target": "> 97.5% / month", "OQ": "OQ-01"},
        {"ID": "MSR-03", "Metric": "Cohort Survival Rate (%)", "Formula": "(COUNT(active_at_M+k) / COUNT(cohort_size)) * 100", "Target": "> 85% at M+12", "OQ-03": "OQ-03"},
        {"ID": "MSR-04", "Metric": "Customer Lifetime Value (CLV ₹)", "Formula": "AVG(monthly_fee_inr) * AVG(tenure_months)", "Target": "> ₹75,000", "OQ": "OQ-04"},
        {"ID": "MSR-05", "Metric": "Channel Longevity Rate (%)", "Formula": "(COUNT(retained_channel) / COUNT(total_channel)) * 100", "Target": "Referral > 90%", "OQ": "OQ-05"},
        {"ID": "MSR-06", "Metric": "Mean Ticket Resolution Hours", "Formula": "AVG(resolution_hours) grouped by churn status", "Target": "< 24.0 hours", "OQ": "OQ-06"},
        {"ID": "MSR-07", "Metric": "Pre-Churn Login Activity Drop", "Formula": "((AVG(logins_prev_60d) - AVG(logins_last_30d)) / AVG(logins_prev_60d)) * 100", "Target": "Significant (p < 0.05)", "OQ": "OQ-07"},
        {"ID": "MSR-08", "Metric": "Satisfaction Correlation (CSAT/NPS)", "Formula": "Pearson r / Spearman rho (CSAT vs Renewed)", "Target": "r > 0.40", "OQ": "OQ-08"},
        {"ID": "MSR-09", "Metric": "MRR Churn in INR (₹)", "Formula": "SUM(monthly_fee_inr) for churned contracts in month M", "Target": "< ₹50,000 / month", "OQ": "OQ-09"},
        {"ID": "MSR-10", "Metric": "Net Revenue Retention (NRR %)", "Formula": "((Starting MRR + Expansion - Contraction - Churned) / Starting MRR) * 100", "Target": "> 105%", "OQ": "OQ-09"},
        {"ID": "MSR-11", "Metric": "Composite Customer Risk Score", "Formula": "0.35*LoginDrop + 0.25*OpenTickets + 0.20*LowCSAT + 0.20*RenewalProximity", "Target": "High Risk >= 70", "OQ": "OQ-10"},
    ]
    st.table(pd.DataFrame(kpi_specs))

with tab2:
    st.subheader("Inferential Statistical Protocols & Decision Thresholds")
    stat_specs = [
        {"Test": "Mann-Whitney U Test", "Application": "Support Ticket Volumes & SLA Turnaround", "Null Hypothesis (H0)": "Distributions of tickets are identical between churned and retained", "Threshold": "Alpha = 0.05", "OQ": "OQ-06"},
        {"Test": "Two-Sample Welch's t-Test", "Application": "90-Day Pre-Churn Usage Telemetry Decline", "Null Hypothesis (H0)": "Mean login days in t-30d equal mean login days in t-60d", "Threshold": "Alpha = 0.05", "OQ": "OQ-07"},
        {"Test": "Chi-Square Independence Test", "Application": "Acquisition Channel vs Customer Retention", "Null Hypothesis (H0)": "Channel type and churn status are statistically independent", "Threshold": "Alpha = 0.05", "OQ": "OQ-05"},
        {"Test": "Pearson r & Spearman rho", "Application": "CSAT Survey Scores vs Subscription Renewal", "Null Hypothesis (H0)": "Zero linear / monotonic correlation between satisfaction and renewal", "Threshold": "Alpha = 0.05", "OQ": "OQ-08"},
    ]
    st.table(pd.DataFrame(stat_specs))

with tab3:
    st.subheader("Database Schema & Entity Specifications")
    st.markdown("""
    * **`dim_customers` (8 cols, 200 rows):** Customer firmographic attributes (`customer_id` PK, `company_name`, `industry`, `company_size`, `region`, `acquisition_channel`, `signup_date`, `signup_cohort_month`).
    * **`dim_plans` (6 cols, 3 rows):** Subscription plan catalog (`plan_id` PK, `plan_name`, `plan_tier_level`, `monthly_base_price`, `max_user_seats`, `sla_support_level`).
    * **`fact_subscriptions` (11 cols, 229 rows):** Contract terms (`subscription_id` PK, `customer_id` FK, `plan_id` FK, `start_date`, `end_date`, `monthly_fee_inr`, `is_churned`, `churn_date`, `tenure_months`).
    * **`fact_support_tickets` (8 cols, 704 clean rows):** Helpdesk tickets (`ticket_id` PK, `customer_id` FK, `created_at`, `resolved_at`, `resolution_hours`, `priority`, `category`, `ticket_csat`).
    * **`fact_monthly_usage` (8 cols, 2,489 rows):** Monthly product usage telemetry (`usage_id` PK, `customer_id` FK, `activity_month`, `login_days`, `core_feature_clicks`, `active_users_count`, `csat_score`, `nps_score`).
    * **`vw_customer_retention_analytics` (33 cols, 229 rows):** Denormalized star schema analytical view pre-joining all master dimensions and computing account-level risk scores.
    """)
