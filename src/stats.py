"""
src/stats.py — Inferential Statistics Engine
Customer Retention Metrics Analysis (CRMA)
Wraps SciPy statistical hypothesis tests (Mann-Whitney U, t-test, Chi-Square, Pearson/Spearman).
"""
import numpy as np
import pandas as pd
from scipy import stats


def run_support_impact_test(churned_tickets: pd.Series, retained_tickets: pd.Series) -> dict:
    """
    Mann-Whitney U test comparing support ticket volumes or SLA between churned and retained.
    """
    c_clean = churned_tickets.dropna()
    r_clean = retained_tickets.dropna()

    if len(c_clean) == 0 or len(r_clean) == 0:
        return {"stat": 0.0, "p_value": 1.0, "significant": False, "interpretation": "Insufficient data"}

    stat, p_val = stats.mannwhitneyu(c_clean, r_clean, alternative="two-sided")
    # Effect size r = Z / sqrt(N) approx
    n_total = len(c_clean) + len(r_clean)
    z_approx = (stat - (len(c_clean) * len(r_clean) / 2)) / np.sqrt(len(c_clean) * len(r_clean) * (n_total + 1) / 12)
    effect_r = abs(z_approx) / np.sqrt(n_total) if n_total > 0 else 0.0

    is_sig = p_val < 0.05
    interp = (
        f"Statistically significant difference (p = {p_val:.4f} < 0.05, effect size r = {effect_r:.2f}). "
        f"Churned accounts experienced a higher ticket load."
        if is_sig else
        f"No statistically significant difference observed (p = {p_val:.4f} >= 0.05)."
    )

    return {
        "stat": round(stat, 2),
        "p_value": round(p_val, 4),
        "effect_r": round(effect_r, 2),
        "significant": is_sig,
        "interpretation": interp,
    }


def run_prechurn_telemetry_ttest(churned_usage: pd.Series, retained_usage: pd.Series) -> dict:
    """
    Two-sample Welch's t-test comparing pre-churn usage telemetry (login days / clicks).
    """
    c_clean = churned_usage.dropna()
    r_clean = retained_usage.dropna()

    if len(c_clean) < 2 or len(r_clean) < 2:
        return {"stat": 0.0, "p_value": 1.0, "significant": False, "interpretation": "Insufficient data"}

    stat, p_val = stats.ttest_ind(c_clean, r_clean, equal_var=False)
    is_sig = p_val < 0.05
    diff = r_clean.mean() - c_clean.mean()

    interp = (
        f"Statistically significant drop in engagement (t = {stat:.2f}, p = {p_val:.4f} < 0.05). "
        f"Retained accounts averaged {diff:.1f} more active units than churned accounts."
        if is_sig else
        f"Engagement difference was not statistically significant (p = {p_val:.4f})."
    )

    return {
        "stat": round(stat, 2),
        "p_value": round(p_val, 4),
        "significant": is_sig,
        "interpretation": interp,
    }


def run_channel_contingency_test(df: pd.DataFrame) -> dict:
    """
    Chi-Square test of independence between Acquisition Channel and Churn Status.
    """
    contingency = pd.crosstab(df["acquisition_channel"], df["is_churned"])
    if contingency.empty or contingency.shape[0] < 2:
        return {"chi2": 0.0, "p_value": 1.0, "significant": False, "dof": 0, "interpretation": "Insufficient categories"}

    chi2, p_val, dof, _ = stats.chi2_contingency(contingency)
    is_sig = p_val < 0.05
    interp = (
        f"Acquisition channel and retention are significantly dependent (chi2 = {chi2:.2f}, p = {p_val:.4f} < 0.05). "
        f"Customer retention varies meaningfully across acquisition channels."
        if is_sig else
        f"Retention differences across channels are not statistically significant (p = {p_val:.4f})."
    )

    return {
        "chi2": round(chi2, 2),
        "p_value": round(p_val, 4),
        "dof": dof,
        "significant": is_sig,
        "interpretation": interp,
    }


def run_satisfaction_correlation(df: pd.DataFrame) -> dict:
    """
    Pearson r and Spearman rho correlations between CSAT score and Renewal Status.
    """
    valid = df.dropna(subset=["latest_csat_score", "is_churned"]).copy()
    # Renewed = 1 - is_churned
    valid["renewed"] = 1 - valid["is_churned"]

    if len(valid) < 5:
        return {"pearson_r": 0.0, "spearman_rho": 0.0, "p_val": 1.0, "sample_size": len(valid), "significant": False}

    r, p_pearson = stats.pearsonr(valid["latest_csat_score"], valid["renewed"])
    rho, p_spearman = stats.spearmanr(valid["latest_csat_score"], valid["renewed"])

    return {
        "pearson_r": round(r, 3),
        "spearman_rho": round(rho, 3),
        "p_val": round(min(p_pearson, p_spearman), 4),
        "sample_size": len(valid),
        "significant": (p_pearson < 0.05 or p_spearman < 0.05),
    }
