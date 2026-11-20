"""
src/metrics.py — Core Business Calculation Engine
Customer Retention Metrics Analysis (CRMA)
Implements RULE-01, RULE-02, RULE-03, RULE-04, and Cohort Survival Matrices.
"""
import pandas as pd
import numpy as np


def calculate_churn_retention_rates(df_subs: pd.DataFrame, month_str: str) -> dict:
    """
    Calculate Monthly CCR % and CRR % for a given YYYY-MM.
    RULE-01/02: CCR = churned_in_month / started_active_before_month * 100
    """
    if df_subs.empty:
        return {"ccr": 0.0, "crr": 100.0, "starting_active": 0, "churned": 0}

    # Started active before or at month boundary
    started = df_subs[df_subs["start_date"] < f"{month_str}-31"]
    if len(started) == 0:
        return {"ccr": 0.0, "crr": 100.0, "starting_active": 0, "churned": 0}

    churned = started[
        (started["is_churned"] == 1) &
        (started["churn_date"].str.startswith(month_str, na=False))
    ]

    ccr = (len(churned) / len(started)) * 100.0 if len(started) > 0 else 0.0
    crr = max(0.0, 100.0 - ccr)

    return {
        "ccr": round(ccr, 2),
        "crr": round(crr, 2),
        "starting_active": len(started),
        "churned": len(churned),
    }


def calculate_monthly_retention_trend(df_subs: pd.DataFrame) -> pd.DataFrame:
    """Compute 24-month time-series dataframe of CCR %, CRR %, and counts."""
    months = sorted(list(set(
        [d[:7] for d in df_subs["start_date"].dropna()] +
        [d[:7] for d in df_subs["churn_date"].dropna() if d]
    )))
    # Keep standard 24 months 2024-01 to 2025-12
    months = [m for m in months if "2024-01" <= m <= "2025-12"]
    if not months:
        months = [f"2024-{i:02d}" for i in range(1, 13)] + [f"2025-{i:02d}" for i in range(1, 13)]

    records = []
    for m in months:
        res = calculate_churn_retention_rates(df_subs, m)
        records.append({
            "month": m,
            "ccr": res["ccr"],
            "crr": res["crr"],
            "starting_active": res["starting_active"],
            "churned": res["churned"]
        })
    return pd.DataFrame(records)


def calculate_mrr_movements(df_subs: pd.DataFrame) -> pd.DataFrame:
    """
    Compute Gross MRR, Churned MRR, and Net MRR over 24 months (RULE-03).
    """
    months = [f"2024-{i:02d}" for i in range(1, 13)] + [f"2025-{i:02d}" for i in range(1, 13)]
    records = []

    for m in months:
        # Active in this month
        active = df_subs[
            (df_subs["start_date"] <= f"{m}-31") &
            ((df_subs["churn_date"].isna()) | (df_subs["churn_date"] > f"{m}-01"))
        ]
        gross_mrr = active["monthly_fee_inr"].sum()

        # Churned in this month
        churned = df_subs[
            (df_subs["is_churned"] == 1) &
            (df_subs["churn_date"].str.startswith(m, na=False))
        ]
        churned_mrr = churned["monthly_fee_inr"].sum()
        net_mrr = gross_mrr - churned_mrr

        # Approximation of NRR %
        nrr = ((gross_mrr - churned_mrr) / gross_mrr * 100.0) if gross_mrr > 0 else 100.0

        records.append({
            "month": m,
            "gross_mrr": round(gross_mrr, 2),
            "churned_mrr": round(churned_mrr, 2),
            "net_mrr": round(net_mrr, 2),
            "nrr": round(nrr, 1),
        })
    return pd.DataFrame(records)


def calculate_clv_by_segment(df_view: pd.DataFrame, segment_col: str) -> pd.DataFrame:
    """
    Calculate Customer Lifetime Value (CLV) per segment (RULE-04):
    CLV = AVG(monthly_fee_inr) * AVG(tenure_months)
    """
    if segment_col not in df_view.columns or df_view.empty:
        return pd.DataFrame(columns=[segment_col, "avg_mrr", "avg_tenure", "clv"])

    grouped = df_view.groupby(segment_col, observed=True).agg(
        avg_mrr=("monthly_fee_inr", "mean"),
        avg_tenure=("tenure_months", "mean"),
        customer_count=("customer_id", "count")
    ).reset_index()

    grouped["clv"] = (grouped["avg_mrr"] * grouped["avg_tenure"]).round(2)
    grouped["avg_mrr"] = grouped["avg_mrr"].round(2)
    grouped["avg_tenure"] = grouped["avg_tenure"].round(1)
    return grouped


def build_cohort_survival_matrix(df_view: pd.DataFrame) -> pd.DataFrame:
    """
    Construct triangular cohort survival matrix (%) across lifecycle intervals (1, 3, 6, 12 months).
    """
    if df_view.empty:
        return pd.DataFrame()

    cohorts = sorted(df_view["signup_cohort_month"].dropna().unique())
    intervals = [1, 3, 6, 12]
    matrix = []

    for c in cohorts:
        cohort_data = df_view[df_view["signup_cohort_month"] == c]
        total_in_cohort = len(cohort_data)
        if total_in_cohort == 0:
            continue

        row = {"cohort": c, "initial_size": total_in_cohort}
        for k in intervals:
            # Check tenure >= k
            survived = cohort_data[cohort_data["tenure_months"] >= k]
            pct = round(len(survived) / total_in_cohort * 100.0, 1)
            row[f"M+{k}"] = pct
        matrix.append(row)

    return pd.DataFrame(matrix)
