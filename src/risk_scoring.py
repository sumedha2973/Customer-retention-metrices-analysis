"""
src/risk_scoring.py — Composite Customer Health Risk Engine
Customer Retention Metrics Analysis (CRMA)
Implements RULE-05: 4-factor weighted linear combination score.
Score = 0.35*LoginDrop + 0.25*OpenTickets + 0.20*LowCSAT + 0.20*RenewalProximity
"""
import pandas as pd
import numpy as np


def compute_composite_risk_score(df_view: pd.DataFrame) -> pd.DataFrame:
    """
    Compute real-time health risk score [0, 100] and assign risk tiers (RULE-05).
    High Risk >= 70, Medium Risk 40-69, Low Risk < 40.
    """
    df = df_view.copy()

    # Factor 1: Login drop (normalized 0-100)
    # Higher score if avg_monthly_login_days is low
    login_factor = np.clip(100 - (df["avg_monthly_login_days"].fillna(15) / 25.0 * 100), 0, 100)

    # Factor 2: Tickets factor (high tickets = high risk)
    ticket_factor = np.clip((df["total_support_tickets"].fillna(2) / 8.0) * 100, 0, 100)

    # Factor 3: Low CSAT factor (low CSAT = high risk)
    csat = df["latest_csat_score"].fillna(3.5)
    csat_factor = np.clip((5.0 - csat) / 4.0 * 100, 0, 100)

    # Factor 4: Renewal proximity / tenure (shorter tenure = higher initial risk)
    tenure = df["tenure_months"].fillna(12)
    tenure_factor = np.clip((24 - tenure) / 24.0 * 100, 0, 100)

    # Weighted linear combination
    raw_score = (
        0.35 * login_factor +
        0.25 * ticket_factor +
        0.20 * csat_factor +
        0.20 * tenure_factor
    )

    df["calc_risk_score"] = np.round(np.clip(raw_score, 0, 100), 1)

    # Assign risk tier
    conditions = [
        df["calc_risk_score"] >= 70.0,
        df["calc_risk_score"] >= 40.0,
    ]
    choices = ["High", "Medium"]
    df["calc_risk_tier"] = np.select(conditions, choices, default="Low")

    return df
