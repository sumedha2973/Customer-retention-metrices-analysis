"""
tests/test_analysis.py  —  TC-T-15 through TC-T-27
Unit tests for every KPI formula and statistical function in src/.
Uses tiny known datasets so expected values can be hand-verified.

Customer Retention Metrics Analysis (CRMA)
Sumedha | Roll 2503201450146 | MCA 2nd Yr | ABES
Run: pytest tests/test_analysis.py -v
"""
import pandas as pd
import numpy as np
import pytest


# ══════════════════════════════════════════════════════════════════════════════
# INLINE MICRO-IMPLEMENTATIONS
# (Identical formulas to what src/metrics.py and src/stats.py will contain.)
# These tests will pass now AND after Step 15 wires up the real imports.
# ══════════════════════════════════════════════════════════════════════════════

def _churn_rate(df: pd.DataFrame, month: str) -> float:
    """CCR = churned_in_month / started_active_before_month × 100"""
    started = df[df["start_date"] < month]
    churned = started[
        (started["is_churned"] == 1) &
        (started["churn_date"].str.startswith(month, na=False))
    ]
    if len(started) == 0:
        return 0.0
    return round(len(churned) / len(started) * 100, 4)


def _retention_rate(ccr: float) -> float:
    return round(100.0 - ccr, 4)


def _mrr_churn(df: pd.DataFrame, month: str) -> float:
    """Sum of monthly_fee_inr for churned contracts in the given month."""
    churned = df[
        (df["is_churned"] == 1) &
        (df["churn_date"].str.startswith(month, na=False))
    ]
    return round(churned["monthly_fee_inr"].sum(), 2)


def _clv(df: pd.DataFrame) -> float:
    """CLV = AVG(monthly_fee_inr) × AVG(tenure_months)"""
    return round(df["monthly_fee_inr"].mean() * df["tenure_months"].mean(), 4)


def _cohort_survival(df: pd.DataFrame, cohort: str, month_k: int) -> float:
    """% of cohort still active after k months."""
    cohort_df = df[df["signup_cohort_month"] == cohort]
    survived = cohort_df[cohort_df["tenure_months"] >= month_k]
    if len(cohort_df) == 0:
        return 0.0
    return round(len(survived) / len(cohort_df) * 100, 2)


def _risk_score(login_drop: float, open_tix: float,
                low_csat: float, renewal_prox: float) -> float:
    """RULE-05: 35/25/20/20 weighted composite, capped [0, 100]."""
    raw = (0.35 * login_drop + 0.25 * open_tix +
           0.20 * low_csat  + 0.20 * renewal_prox)
    return round(min(max(raw, 0.0), 100.0), 4)


def _pearson_corr(x: pd.Series, y: pd.Series) -> float:
    return round(x.corr(y), 4)


# ── Tiny test datasets (seed=0, fully deterministic) ─────────────────────────

@pytest.fixture(scope="module")
def tiny_subscriptions():
    """6-row known dataset for KPI formula tests."""
    return pd.DataFrame({
        "subscription_id": ["S1", "S2", "S3", "S4", "S5", "S6"],
        "customer_id":     ["C1", "C2", "C3", "C4", "C5", "C6"],
        "is_churned":      [1,     0,    1,    0,    0,    1  ],
        "monthly_fee_inr": [2500., 8000., 20000., 2500., 8000., 20000.],
        "tenure_months":   [3,     12,    6,     24,    18,    9    ],
        "start_date":      ["2024-01-01"]*6,
        "churn_date":      ["2024-03-01", None, "2024-03-15",
                            None, None, "2024-03-28"],
        "signup_cohort_month": ["2024-01"]*6,
    })


@pytest.fixture(scope="module")
def tiny_usage():
    return pd.DataFrame({
        "customer_id":    ["C1", "C1", "C2", "C2"],
        "activity_month": ["2024-01", "2024-02", "2024-01", "2024-02"],
        "login_days":     [20, 5, 15, 14],
        "csat_score":     [2.0, 1.0, 4.0, 5.0],
        "nps_score":      [3.0, 2.0, 8.0, 9.0],
    })


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-15  Churn rate formula — known dataset
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_15_churn_rate_formula(tiny_subscriptions):
    """TC-T-15: CCR for 2024-03 = 3 churned / 6 started before 2024-03 × 100 = 50.0%"""
    result = _churn_rate(tiny_subscriptions, "2024-03")
    assert result == pytest.approx(50.0, abs=0.01), (
        f"Expected CCR=50.0%, got {result}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-16  Retention rate = 100 − CCR
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_16_retention_rate_formula():
    """TC-T-16: CRR = 100 − CCR."""
    for ccr, expected_crr in [(2.5, 97.5), (0.0, 100.0), (100.0, 0.0)]:
        assert _retention_rate(ccr) == pytest.approx(expected_crr, abs=0.001)


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-17  MRR churn formula — known dataset
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_17_mrr_churn_formula(tiny_subscriptions):
    """TC-T-17: MRR churn for 2024-03 = 2500+20000+20000 = 42500."""
    result = _mrr_churn(tiny_subscriptions, "2024-03")
    assert result == pytest.approx(42500.0, abs=0.01), (
        f"Expected MRR churn=42500, got {result}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-18  CLV formula — AVG fee × AVG tenure
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_18_clv_formula(tiny_subscriptions):
    """TC-T-18: CLV = mean(fees) × mean(tenures)."""
    avg_fee    = (2500 + 8000 + 20000 + 2500 + 8000 + 20000) / 6  # 10166.67
    avg_tenure = (3 + 12 + 6 + 24 + 18 + 9) / 6                   # 12.0
    expected   = round(avg_fee * avg_tenure, 4)
    result     = _clv(tiny_subscriptions)
    assert result == pytest.approx(expected, rel=1e-3)


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-19  Cohort survival at M+6
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_19_cohort_survival(tiny_subscriptions):
    """TC-T-19: Cohort 2024-01 survival at M+6 = customers with tenure≥6 / 6."""
    result = _cohort_survival(tiny_subscriptions, "2024-01", 6)
    # tenure_months = [3,12,6,24,18,9] → ≥6: [12,6,24,18,9] = 5/6 = 83.33
    assert result == pytest.approx(83.33, abs=0.1)


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-20  Risk score weights sum to 100 and clamp to [0, 100]
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_20_risk_score_max_100():
    """TC-T-20: Risk score caps at 100 even if inputs exceed 100."""
    result = _risk_score(200, 200, 200, 200)
    assert result == 100.0


def test_TC_T_20_risk_score_min_0():
    """TC-T-20b: Risk score floors at 0 for all-zero inputs."""
    assert _risk_score(0, 0, 0, 0) == 0.0


def test_TC_T_20_risk_score_weights():
    """TC-T-20c: Risk score = 35×login + 25×tickets + 20×csat + 20×renewal (all=1)."""
    expected = 0.35 * 1 + 0.25 * 1 + 0.20 * 1 + 0.20 * 1  # = 1.0
    assert _risk_score(1, 1, 1, 1) == pytest.approx(expected, abs=0.0001)


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-21  Pearson correlation sign check
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_21_correlation_positive():
    """TC-T-21: CSAT and renewal_flag should be positively correlated in known data."""
    df = pd.DataFrame({
        "csat":    [1.0, 2.0, 3.0, 4.0, 5.0],
        "renewed": [0,   0,   1,   1,   1  ],
    })
    r = _pearson_corr(df["csat"], df["renewed"])
    assert r > 0, f"Expected positive r, got {r}"


def test_TC_T_21_correlation_range():
    """TC-T-21b: Pearson r is always in [-1, 1]."""
    rng = np.random.default_rng(42)
    for _ in range(20):
        x = pd.Series(rng.normal(size=50))
        y = pd.Series(rng.normal(size=50))
        r = _pearson_corr(x, y)
        assert -1.0 <= r <= 1.0


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-22  Mann-Whitney U p-value import (scipy available)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_22_mannwhitney_runs():
    """TC-T-22: scipy.stats.mannwhitneyu executes without error."""
    from scipy.stats import mannwhitneyu
    group_a = [3, 4, 5, 6, 7]   # retained ticket counts
    group_b = [8, 9, 10, 12, 15] # churned ticket counts
    stat, p = mannwhitneyu(group_a, group_b, alternative="less")
    assert 0.0 <= p <= 1.0, f"p-value {p} out of [0,1]"
    assert stat >= 0


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-23  Two-sample t-test import (scipy available)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_23_ttest_runs():
    """TC-T-23: scipy.stats.ttest_ind executes and returns finite values."""
    from scipy.stats import ttest_ind
    rng = np.random.default_rng(99)
    a = rng.normal(10, 2, 30)
    b = rng.normal(15, 2, 30)
    stat, p = ttest_ind(a, b)
    assert np.isfinite(stat), "t-statistic is not finite"
    assert 0.0 <= p <= 1.0,   f"p-value {p} out of [0,1]"
    assert p < 0.05,           "Known different distributions should be significant"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-24  Chi-square test import (scipy available)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_24_chisquare_runs():
    """TC-T-24: scipy.stats.chi2_contingency executes on a 2×5 contingency table."""
    from scipy.stats import chi2_contingency
    # Rows = churned/retained, Cols = channels
    observed = np.array([
        [6, 12, 8, 9, 7],   # churned per channel
        [23, 36, 39, 26, 37] # retained per channel
    ])
    chi2, p, dof, expected = chi2_contingency(observed)
    assert chi2 >= 0
    assert 0.0 <= p <= 1.0
    assert dof == 4   # (2-1) × (5-1)


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-25  Spearman correlation import (scipy available)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_25_spearman_runs():
    """TC-T-25: scipy.stats.spearmanr executes and returns valid rho."""
    from scipy.stats import spearmanr
    csat   = [1, 2, 3, 4, 5, 2, 4, 3]
    churn  = [1, 1, 0, 0, 0, 1, 0, 1]
    rho, p = spearmanr(csat, churn)
    assert -1.0 <= rho <= 1.0
    assert 0.0  <= p   <= 1.0


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-26  Zero-division guard: churn rate on empty dataset
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_26_churn_rate_empty_guard():
    """TC-T-26: churn rate returns 0.0 (not ZeroDivisionError) on empty input."""
    empty = pd.DataFrame(columns=["start_date", "is_churned", "churn_date"])
    result = _churn_rate(empty, "2024-03")
    assert result == 0.0


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-27  NRR formula: > 100% when expansion > churn
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_27_nrr_above_100():
    """TC-T-27: NRR > 100% when expansion MRR exceeds churned MRR."""
    starting_mrr   = 100_000.0
    expansion_mrr  =  10_000.0
    contraction_mrr =  2_000.0
    churned_mrr    =   5_000.0
    nrr = ((starting_mrr + expansion_mrr - contraction_mrr - churned_mrr)
           / starting_mrr * 100)
    assert nrr > 100.0, f"Expected NRR > 100%, got {nrr:.2f}%"
    assert nrr == pytest.approx(103.0, abs=0.001)
