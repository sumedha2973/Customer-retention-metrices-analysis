"""
src/data_loader.py — Data loading, caching, and validation layer.
Reads from embedded SQLite DB `data/app.db` with read-only connection.
Decorated with @st.cache_data for instant interactive performance.
"""
import os
import sqlite3
import pandas as pd
import streamlit as st

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "app.db")


def get_connection():
    """Establish a read-only SQLite connection using URI mode."""
    db_abs = os.path.abspath(DB_PATH)
    uri = f"file:{db_abs}?mode=ro"
    conn = sqlite3.connect(uri, uri=True)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@st.cache_data(ttl=3600, show_spinner=False)
def load_table(table_name: str) -> pd.DataFrame:
    """Load a specific table from SQLite into a typed pandas DataFrame."""
    conn = get_connection()
    try:
        df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
    finally:
        conn.close()

    # Downcast known categorical string columns
    cat_cols = ["industry", "company_size", "region", "acquisition_channel", "plan_name", "priority", "category"]
    for col in cat_cols:
        if col in df.columns:
            df[col] = df[col].astype("category")

    return df


@st.cache_data(ttl=3600, show_spinner=False)
def load_analytical_view() -> pd.DataFrame:
    """Load the pre-computed 33-column analytical view with validation."""
    conn = get_connection()
    try:
        df = pd.read_sql("SELECT * FROM vw_customer_retention_analytics", conn)
    finally:
        conn.close()

    # Ensure validation integrity
    assert df["customer_id"].is_unique, "Validation Error: customer_id in view must be unique!"
    assert (df["monthly_fee_inr"] >= 0).all(), "Validation Error: monthly_fee_inr cannot be negative!"

    # Optimize categorical types
    for col in ["industry", "company_size", "region", "acquisition_channel", "plan_name", "subscription_status", "risk_tier"]:
        if col in df.columns:
            df[col] = df[col].astype("category")

    return df


def get_all_data() -> dict:
    """Convenience function to return all tables and view in a dictionary."""
    return {
        "dim_customers": load_table("dim_customers"),
        "dim_plans": load_table("dim_plans"),
        "fact_subscriptions": load_table("fact_subscriptions"),
        "fact_support_tickets": load_table("fact_support_tickets"),
        "fact_monthly_usage": load_table("fact_monthly_usage"),
        "view": load_analytical_view(),
    }
