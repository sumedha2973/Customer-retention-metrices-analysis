"""
tests/test_data.py  —  TC-T-01 through TC-T-14
Data-layer tests: DB connectivity, schema, row counts, PK/FK integrity,
domain ranges, and data-quality issues.

Customer Retention Metrics Analysis (CRMA)
Sumedha | Roll 2503201450146 | MCA 2nd Yr | ABES
Run: pytest tests/test_data.py -v
"""
import os
import sqlite3
import pandas as pd
import pytest

# ── Fixture: shared in-memory connection ──────────────────────────────────────
DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "app.db")


@pytest.fixture(scope="module")
def conn():
    """Open a read-only SQLite connection for the whole module."""
    con = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    con.execute("PRAGMA foreign_keys = ON")
    yield con
    con.close()


@pytest.fixture(scope="module")
def tables(conn):
    """Load all five tables into a dict of DataFrames once per module."""
    return {
        "dim_customers":        pd.read_sql("SELECT * FROM dim_customers",        conn),
        "dim_plans":            pd.read_sql("SELECT * FROM dim_plans",            conn),
        "fact_subscriptions":   pd.read_sql("SELECT * FROM fact_subscriptions",   conn),
        "fact_support_tickets": pd.read_sql("SELECT * FROM fact_support_tickets", conn),
        "fact_monthly_usage":   pd.read_sql("SELECT * FROM fact_monthly_usage",   conn),
    }


@pytest.fixture(scope="module")
def view_df(conn):
    """Load the analytical view."""
    return pd.read_sql("SELECT * FROM vw_customer_retention_analytics", conn)


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-01  DB file exists and is readable
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_01_db_file_exists():
    """TC-T-01: data/app.db exists and SQLite can open it."""
    assert os.path.exists(DB_PATH), f"DB not found at {DB_PATH}"
    con = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    result = con.execute("SELECT sqlite_version()").fetchone()
    con.close()
    assert result is not None


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-02  All five tables exist
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_02_all_tables_exist(conn):
    """TC-T-02: Five expected tables are present in the database."""
    expected = {
        "dim_customers", "dim_plans", "fact_subscriptions",
        "fact_support_tickets", "fact_monthly_usage",
    }
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    ).fetchall()
    actual = {r[0] for r in rows}
    assert expected.issubset(actual), f"Missing tables: {expected - actual}"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-03  Row counts within expected ranges
# ─────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("table,min_rows,max_rows", [
    ("dim_customers",        150,  5000),
    ("dim_plans",              3,     5),
    ("fact_subscriptions",   150, 10000),
    ("fact_support_tickets", 100, 20000),
    ("fact_monthly_usage",   500, 80000),
])
def test_TC_T_03_row_counts(conn, table, min_rows, max_rows):
    """TC-T-03: Every table has rows within the designed range."""
    count = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
    assert min_rows <= count <= max_rows, (
        f"{table}: row count {count} not in [{min_rows}, {max_rows}]"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-04  Analytical view exists with ≥ 30 columns
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_04_view_columns(view_df):
    """TC-T-04: vw_customer_retention_analytics has ≥ 30 columns."""
    assert len(view_df.columns) >= 30, (
        f"View only has {len(view_df.columns)} columns (expected ≥ 30)"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-05  Primary key uniqueness
# ─────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("table,pk_col", [
    ("dim_customers",        "customer_id"),
    ("dim_plans",            "plan_id"),
    ("fact_subscriptions",   "subscription_id"),
    ("fact_support_tickets", "ticket_id"),
    ("fact_monthly_usage",   "usage_id"),
])
def test_TC_T_05_pk_uniqueness(tables, table, pk_col):
    """TC-T-05: Every primary key column has strictly unique values."""
    df = tables[table]
    assert df[pk_col].is_unique, (
        f"{table}.{pk_col}: found {df[pk_col].duplicated().sum()} duplicates"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-06  Foreign key integrity
# ─────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("child_table,fk_col,parent_table,pk_col", [
    ("fact_subscriptions",   "customer_id", "dim_customers", "customer_id"),
    ("fact_subscriptions",   "plan_id",     "dim_plans",     "plan_id"),
    ("fact_support_tickets", "customer_id", "dim_customers", "customer_id"),
    ("fact_monthly_usage",   "customer_id", "dim_customers", "customer_id"),
])
def test_TC_T_06_foreign_key_integrity(tables, child_table, fk_col, parent_table, pk_col):
    """TC-T-06: Every FK value in child tables resolves to a valid PK in the parent."""
    child  = tables[child_table]
    parent = tables[parent_table]
    orphans = child[~child[fk_col].isin(parent[pk_col])]
    assert len(orphans) == 0, (
        f"{child_table}.{fk_col}: {len(orphans)} orphaned FK values"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-07  Domain / range checks
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_07_monthly_fee_non_negative(tables):
    """TC-T-07a: All monthly_fee_inr values are ≥ 0."""
    df = tables["fact_subscriptions"]
    bad = df[df["monthly_fee_inr"] < 0]
    assert len(bad) == 0, f"Found {len(bad)} rows with negative monthly_fee_inr"


def test_TC_T_07_login_days_range(tables):
    """TC-T-07b: login_days is between 0 and 31 (inclusive)."""
    df = tables["fact_monthly_usage"]
    bad = df[(df["login_days"] < 0) | (df["login_days"] > 31)]
    assert len(bad) == 0, f"Found {len(bad)} rows with invalid login_days"


def test_TC_T_07_csat_score_range(tables):
    """TC-T-07c: csat_score (non-null) is between 1 and 5."""
    df = tables["fact_monthly_usage"].dropna(subset=["csat_score"])
    bad = df[(df["csat_score"] < 1) | (df["csat_score"] > 5)]
    assert len(bad) == 0, f"Found {len(bad)} rows with CSAT out of [1,5]"


def test_TC_T_07_nps_score_range(tables):
    """TC-T-07d: nps_score (non-null) is between 0 and 10."""
    df = tables["fact_monthly_usage"].dropna(subset=["nps_score"])
    bad = df[(df["nps_score"] < 0) | (df["nps_score"] > 10)]
    assert len(bad) == 0, f"Found {len(bad)} rows with NPS out of [0,10]"


def test_TC_T_07_tenure_months_positive(tables):
    """TC-T-07e: tenure_months ≥ 1 for all subscriptions."""
    df = tables["fact_subscriptions"]
    bad = df[df["tenure_months"] < 1]
    assert len(bad) == 0, f"Found {len(bad)} rows with tenure_months < 1"


def test_TC_T_07_resolution_hours_non_negative(tables):
    """TC-T-07f: resolution_hours ≥ 0."""
    df = tables["fact_support_tickets"]
    bad = df[df["resolution_hours"] < 0]
    assert len(bad) == 0, f"Found {len(bad)} rows with negative resolution_hours"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-08  Required columns present
# ─────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("table,required_cols", [
    ("dim_customers", [
        "customer_id", "company_name", "industry", "company_size",
        "region", "acquisition_channel", "signup_date", "signup_cohort_month"
    ]),
    ("dim_plans", [
        "plan_id", "plan_name", "plan_tier_level",
        "monthly_base_price", "max_user_seats", "sla_support_level"
    ]),
    ("fact_subscriptions", [
        "subscription_id", "customer_id", "plan_id", "start_date", "end_date",
        "billing_cadence", "monthly_fee_inr", "subscription_status",
        "is_churned", "churn_date", "tenure_months"
    ]),
    ("fact_support_tickets", [
        "ticket_id", "customer_id", "created_at", "resolved_at",
        "resolution_hours", "priority", "category", "ticket_csat"
    ]),
    ("fact_monthly_usage", [
        "usage_id", "customer_id", "activity_month", "login_days",
        "core_feature_clicks", "active_users_count", "csat_score", "nps_score"
    ]),
])
def test_TC_T_08_required_columns(tables, table, required_cols):
    """TC-T-08: Every required column is present in each table."""
    df = tables[table]
    missing = [c for c in required_cols if c not in df.columns]
    assert not missing, f"{table}: missing columns {missing}"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-09  Mandatory non-null fields
# ─────────────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("table,non_null_col", [
    ("dim_customers",        "customer_id"),
    ("dim_customers",        "company_name"),
    ("dim_customers",        "signup_date"),
    ("fact_subscriptions",   "subscription_id"),
    ("fact_subscriptions",   "is_churned"),
    ("fact_subscriptions",   "monthly_fee_inr"),
    ("fact_subscriptions",   "start_date"),
    ("fact_support_tickets", "ticket_id"),
    ("fact_support_tickets", "customer_id"),
    ("fact_monthly_usage",   "usage_id"),
    ("fact_monthly_usage",   "activity_month"),
])
def test_TC_T_09_mandatory_non_null(tables, table, non_null_col):
    """TC-T-09: Mandatory columns must have zero null values."""
    df = tables[table]
    null_count = df[non_null_col].isna().sum()
    assert null_count == 0, (
        f"{table}.{non_null_col}: found {null_count} null values"
    )


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-10  Duplicate ticket check (DB should be clean, CSV may have dups)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_10_no_duplicate_tickets_in_db(tables):
    """TC-T-10: DB fact_support_tickets has no duplicate ticket_ids (DQ-02 cleaned)."""
    df = tables["fact_support_tickets"]
    dup_count = df["ticket_id"].duplicated().sum()
    assert dup_count == 0, f"Found {dup_count} duplicate ticket_ids in DB"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-11  Dim_plans has exactly 3 tiers
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_11_exactly_three_plans(tables):
    """TC-T-11: dim_plans contains exactly 3 subscription tiers."""
    df = tables["dim_plans"]
    assert len(df) == 3, f"Expected 3 plans, got {len(df)}"
    assert set(df["plan_name"]) == {"Starter", "Professional", "Enterprise"}


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-12  is_churned is binary (0 or 1 only)
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_12_is_churned_binary(tables):
    """TC-T-12: fact_subscriptions.is_churned contains only 0 or 1."""
    df = tables["fact_subscriptions"]
    invalid = df[~df["is_churned"].isin([0, 1])]
    assert len(invalid) == 0, f"Found {len(invalid)} non-binary is_churned values"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-13  health_risk_score within [0, 100]
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_13_risk_score_range(view_df):
    """TC-T-13: health_risk_score is between 0 and 100."""
    bad = view_df[
        (view_df["health_risk_score"] < 0) | (view_df["health_risk_score"] > 100)
    ]
    assert len(bad) == 0, f"Found {len(bad)} risk scores outside [0, 100]"


# ─────────────────────────────────────────────────────────────────────────────
# TC-T-14  risk_tier values are valid
# ─────────────────────────────────────────────────────────────────────────────
def test_TC_T_14_risk_tier_values(view_df):
    """TC-T-14: risk_tier column only contains 'High', 'Medium', or 'Low'."""
    valid = {"High", "Medium", "Low"}
    actual = set(view_df["risk_tier"].dropna().unique())
    invalid = actual - valid
    assert not invalid, f"Unexpected risk_tier values: {invalid}"
