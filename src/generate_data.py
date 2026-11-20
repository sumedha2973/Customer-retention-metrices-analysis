"""
===========================================================================
CRMA - Synthetic Data Generation Engine (Step 10)
===========================================================================
Project  : Customer Retention Metrics Analysis
Author   : Sumedha (Roll 2503201450146, MCA 2nd Year, ABES EC)
Mentor   : Ms. Anjali Pandey
Seed     : 42 (deterministic reproducibility)

Generates 5 tables matching Step 8 DDL exactly:
  dim_customers, dim_plans, fact_subscriptions,
  fact_support_tickets, fact_monthly_usage

Outputs:
  data/app.db              SQLite database with FK enforcement
  data/dim_customers.csv   Individual table CSVs
  data/dim_plans.csv
  data/fact_subscriptions.csv
  data/fact_support_tickets.csv
  data/fact_monthly_usage.csv
  data/crma_data.xlsx      Excel workbook (one sheet per table)
  data/validation_report.txt  Automated integrity check
===========================================================================
"""

import os
import random
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

# =========================================================================
# CONFIGURATION
# =========================================================================
SEED = 42
random.seed(SEED)
np.random.seed(SEED)

NUM_CUSTOMERS = 200
DATE_START = datetime(2024, 1, 1)
DATE_END = datetime(2025, 12, 31)
MONTHS_24 = pd.date_range("2024-01-01", "2025-12-01", freq="MS").tolist()

OUT_DIR = Path(__file__).resolve().parent.parent / "data"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# =========================================================================
# INDIAN FICTIONAL NAMES AND DIMENSIONAL VALUES
# =========================================================================
COMPANY_PREFIXES = [
    "Aarav", "Bharat", "Chandra", "Devi", "Ekta", "Faiz", "Ganesh", "Hari",
    "Indra", "Jaya", "Kavya", "Lakshmi", "Maya", "Neel", "Om", "Priya",
    "Rajan", "Sahil", "Tanvi", "Uma", "Varun", "Yash", "Zara", "Ananya",
    "Dhruv", "Gaurav", "Ishaan", "Kabir", "Meera", "Nikhil", "Pooja",
    "Rahul", "Sneha", "Tara", "Vimal", "Arjun", "Deepa", "Kiran", "Manju",
    "Pankaj", "Rohit", "Sita", "Vikram", "Ashok", "Bhavna", "Chetan",
    "Dinesh", "Farhan", "Geeta", "Hemant", "Ira", "Jagdish", "Kedar",
    "Lata", "Mohan", "Neha", "Padma", "Ramesh", "Sunita", "Tushar",
    "Usha", "Vinod", "Waqar", "Yogesh", "Ajay", "Bindu", "Chintan",
    "Devendra", "Esha", "Firoz", "Gita", "Harsh", "Indu", "Jai",
    "Kamla", "Leela", "Madhu", "Nandini", "Omkar", "Pallavi", "Renu",
    "Shankar", "Tripti", "Ujjwal", "Vandana", "Yuvraj", "Zubin", "Anil",
    "Bharti", "Chaitra", "Daksha", "Eknath", "Falak", "Girish", "Himani",
]
COMPANY_SUFFIXES = [
    "Solutions", "Technologies", "Infotech", "Systems", "Consulting",
    "Digital", "Software", "Services", "Analytics", "Enterprises",
    "Labs", "Tech", "Innovations", "Infra", "Corp", "Networks",
    "Ventures", "Industries", "Global", "India"
]

INDUSTRIES = ["EdTech", "FinTech", "Healthcare", "Retail", "Manufacturing"]
COMPANY_SIZES = ["Micro (1-10)", "Small (11-50)", "Medium (51-200)"]
REGIONS = ["North", "South", "East", "West"]
CHANNELS = ["Organic Search", "Referral", "Paid Social", "Direct", "Partner"]
PRIORITIES = ["Low", "Medium", "High", "Critical"]
TICKET_CATEGORIES = ["Billing", "Bug", "Feature Request", "Technical"]

# =========================================================================
# PLAN CATALOG (dim_plans - exactly 3 rows)
# =========================================================================
PLANS = [
    {
        "plan_id": "PLAN-01",
        "plan_name": "Starter",
        "plan_tier_level": 1,
        "monthly_base_price": 2499.00,
        "max_user_seats": 5,
        "sla_support_level": "Standard (48h)",
    },
    {
        "plan_id": "PLAN-02",
        "plan_name": "Professional",
        "plan_tier_level": 2,
        "monthly_base_price": 7999.00,
        "max_user_seats": 25,
        "sla_support_level": "Priority (24h)",
    },
    {
        "plan_id": "PLAN-03",
        "plan_name": "Enterprise",
        "plan_tier_level": 3,
        "monthly_base_price": 19999.00,
        "max_user_seats": 100,
        "sla_support_level": "Dedicated (4h)",
    },
]

# =========================================================================
# PATTERN DEFINITIONS (5 deliberately injected patterns)
# =========================================================================
# PATTERN 1: Higher churn for Starter plan (~30%) vs Enterprise (~10%)
# PATTERN 2: Churned customers have 2x more support tickets on average
# PATTERN 3: Churned customers show login decline in last 3 months
# PATTERN 4: Referral channel has lowest churn; Paid Social has highest
# PATTERN 5: Low CSAT (<= 2) customers churn at 3x the normal rate

CHURN_RATE_BY_PLAN = {"PLAN-01": 0.30, "PLAN-02": 0.18, "PLAN-03": 0.10}
CHURN_RATE_BY_CHANNEL = {
    "Organic Search": 0.20, "Referral": 0.12, "Paid Social": 0.32,
    "Direct": 0.22, "Partner": 0.15,
}

# =========================================================================
# DATA QUALITY ISSUES TO INJECT
# =========================================================================
# DQ-01: 2-3% missing csat_score / nps_score in fact_monthly_usage (NULLs)
# DQ-02: 3 exact duplicate rows in fact_support_tickets
# DQ-03: ~5 outlier resolution_hours values (> 200 hours)

# =========================================================================
# STEP 1: Generate dim_customers
# =========================================================================
def generate_dim_customers():
    used_names = set()
    rows = []
    for i in range(1, NUM_CUSTOMERS + 1):
        # Unique company name
        while True:
            name = f"{random.choice(COMPANY_PREFIXES)} {random.choice(COMPANY_SUFFIXES)}"
            if name not in used_names:
                used_names.add(name)
                break

        industry = random.choice(INDUSTRIES)
        company_size = random.choices(
            COMPANY_SIZES, weights=[0.40, 0.40, 0.20], k=1
        )[0]
        region = random.choice(REGIONS)
        channel = random.choice(CHANNELS)

        # Signup date spread across 24 months
        signup_date = DATE_START + timedelta(
            days=random.randint(0, (DATE_END - DATE_START).days - 60)
        )
        cohort_month = signup_date.strftime("%Y-%m")

        rows.append(
            {
                "customer_id": f"CUST-{i:05d}",
                "company_name": name,
                "industry": industry,
                "company_size": company_size,
                "region": region,
                "acquisition_channel": channel,
                "signup_date": signup_date.strftime("%Y-%m-%d"),
                "signup_cohort_month": cohort_month,
            }
        )
    return pd.DataFrame(rows)


# =========================================================================
# STEP 2: Generate fact_subscriptions
# =========================================================================
def generate_fact_subscriptions(df_customers):
    rows = []
    sub_counter = 0
    for _, cust in df_customers.iterrows():
        cid = cust["customer_id"]
        channel = cust["acquisition_channel"]
        signup = datetime.strptime(cust["signup_date"], "%Y-%m-%d")

        # Assign plan with weighted distribution
        plan_id = random.choices(
            ["PLAN-01", "PLAN-02", "PLAN-03"],
            weights=[0.50, 0.35, 0.15],
            k=1,
        )[0]
        plan = next(p for p in PLANS if p["plan_id"] == plan_id)

        # Fee = base price with slight variation (+/- 15%)
        fee = round(plan["monthly_base_price"] * random.uniform(0.85, 1.15), 2)

        # Billing cadence
        cadence = random.choices(
            ["Monthly", "Annual"], weights=[0.70, 0.30], k=1
        )[0]

        # Determine churn probability (PATTERN 1 + PATTERN 4)
        base_churn = CHURN_RATE_BY_PLAN[plan_id]
        channel_adj = CHURN_RATE_BY_CHANNEL.get(channel, 0.20) - 0.20
        churn_prob = min(max(base_churn + channel_adj, 0.05), 0.50)

        is_churned = 1 if random.random() < churn_prob else 0

        # Tenure
        max_possible_months = max(
            1,
            (DATE_END.year - signup.year) * 12 + DATE_END.month - signup.month,
        )
        if is_churned:
            tenure = random.randint(1, max(1, max_possible_months - 1))
        else:
            tenure = max_possible_months

        start_date = signup
        end_date = signup + timedelta(days=tenure * 30)
        if end_date > DATE_END:
            end_date = DATE_END

        churn_date = end_date.strftime("%Y-%m-%d") if is_churned else None
        status = "Churned" if is_churned else random.choices(
            ["Active", "Renewed"], weights=[0.60, 0.40], k=1
        )[0]

        sub_counter += 1
        rows.append(
            {
                "subscription_id": f"SUB-{sub_counter:06d}",
                "customer_id": cid,
                "plan_id": plan_id,
                "start_date": start_date.strftime("%Y-%m-%d"),
                "end_date": end_date.strftime("%Y-%m-%d"),
                "billing_cadence": cadence,
                "monthly_fee_inr": fee,
                "subscription_status": status,
                "is_churned": is_churned,
                "churn_date": churn_date,
                "tenure_months": tenure,
            }
        )

        # Some active/renewed customers get a second subscription (renewal / upgrade)
        if not is_churned and random.random() < 0.25 and tenure >= 6:
            sub_counter += 1
            new_plan_id = random.choices(
                ["PLAN-01", "PLAN-02", "PLAN-03"],
                weights=[0.15, 0.45, 0.40],
                k=1,
            )[0]
            new_plan = next(p for p in PLANS if p["plan_id"] == new_plan_id)
            new_fee = round(
                new_plan["monthly_base_price"] * random.uniform(0.90, 1.10), 2
            )
            new_start = end_date + timedelta(days=1)
            new_end = new_start + timedelta(days=random.randint(3, 12) * 30)
            if new_end > DATE_END:
                new_end = DATE_END
            new_tenure = max(
                1,
                (new_end.year - new_start.year) * 12
                + new_end.month
                - new_start.month,
            )
            rows.append(
                {
                    "subscription_id": f"SUB-{sub_counter:06d}",
                    "customer_id": cid,
                    "plan_id": new_plan_id,
                    "start_date": new_start.strftime("%Y-%m-%d"),
                    "end_date": new_end.strftime("%Y-%m-%d"),
                    "billing_cadence": cadence,
                    "monthly_fee_inr": new_fee,
                    "subscription_status": "Active",
                    "is_churned": 0,
                    "churn_date": None,
                    "tenure_months": new_tenure,
                }
            )

    return pd.DataFrame(rows)


# =========================================================================
# STEP 3: Generate fact_support_tickets
# =========================================================================
def generate_fact_support_tickets(df_customers, df_subs):
    rows = []
    ticket_counter = 0

    # Build a lookup for churn status per customer
    churn_map = df_subs.groupby("customer_id")["is_churned"].max().to_dict()

    for _, cust in df_customers.iterrows():
        cid = cust["customer_id"]
        signup = datetime.strptime(cust["signup_date"], "%Y-%m-%d")
        is_churned = churn_map.get(cid, 0)

        # PATTERN 2: Churned customers get 2x more tickets
        if is_churned:
            num_tickets = random.randint(3, 10)
        else:
            num_tickets = random.randint(0, 5)

        for _ in range(num_tickets):
            ticket_counter += 1
            created_days = random.randint(0, (DATE_END - signup).days)
            created_at = signup + timedelta(days=created_days)
            if created_at > DATE_END:
                created_at = DATE_END - timedelta(days=random.randint(1, 30))

            priority = random.choices(
                PRIORITIES,
                weights=[0.30, 0.35, 0.25, 0.10] if not is_churned
                else [0.15, 0.25, 0.35, 0.25],
                k=1,
            )[0]
            category = random.choice(TICKET_CATEGORIES)

            # Resolution hours: normal range 1-72, with SLA differences
            if priority == "Critical":
                res_hours = round(random.uniform(1, 12), 1)
            elif priority == "High":
                res_hours = round(random.uniform(4, 48), 1)
            elif priority == "Medium":
                res_hours = round(random.uniform(8, 72), 1)
            else:
                res_hours = round(random.uniform(12, 96), 1)

            # DQ-03: Inject ~5 extreme outlier resolution hours
            if ticket_counter in [15, 42, 88, 134, 199]:
                res_hours = round(random.uniform(220, 500), 1)

            resolved_at = created_at + timedelta(hours=res_hours)
            if resolved_at > DATE_END + timedelta(days=30):
                resolved_at = DATE_END

            # Ticket CSAT - lower for churned customers (PATTERN 5 linkage)
            if is_churned:
                t_csat = random.choices([1, 2, 3, 4, 5], weights=[0.25, 0.30, 0.25, 0.15, 0.05], k=1)[0]
            else:
                t_csat = random.choices([1, 2, 3, 4, 5], weights=[0.05, 0.10, 0.25, 0.35, 0.25], k=1)[0]

            # Some tickets have no CSAT (survey not completed)
            if random.random() < 0.08:
                t_csat = None

            rows.append(
                {
                    "ticket_id": f"TCK-{ticket_counter:06d}",
                    "customer_id": cid,
                    "created_at": created_at.strftime("%Y-%m-%d %H:%M:%S"),
                    "resolved_at": resolved_at.strftime("%Y-%m-%d %H:%M:%S"),
                    "resolution_hours": res_hours,
                    "priority": priority,
                    "category": category,
                    "ticket_csat": t_csat,
                }
            )

    # DQ-02: Inject 3 exact duplicate rows
    if len(rows) >= 10:
        for dup_idx in [3, 7, 12]:
            dup_row = rows[dup_idx].copy()
            rows.append(dup_row)

    return pd.DataFrame(rows)


# =========================================================================
# STEP 4: Generate fact_monthly_usage
# =========================================================================
def generate_fact_monthly_usage(df_customers, df_subs):
    rows = []
    usage_counter = 0

    churn_map = df_subs.groupby("customer_id")["is_churned"].max().to_dict()
    churn_date_map = {}
    for _, row in df_subs.iterrows():
        if row["is_churned"] == 1 and row["churn_date"]:
            churn_date_map[row["customer_id"]] = datetime.strptime(
                row["churn_date"], "%Y-%m-%d"
            )

    for _, cust in df_customers.iterrows():
        cid = cust["customer_id"]
        signup = datetime.strptime(cust["signup_date"], "%Y-%m-%d")
        is_churned = churn_map.get(cid, 0)
        c_date = churn_date_map.get(cid)

        for month_dt in MONTHS_24:
            # Only generate usage for months after signup and before/at churn
            if month_dt < signup.replace(day=1):
                continue
            if c_date and month_dt > c_date.replace(day=1):
                continue

            usage_counter += 1
            months_since_signup = (
                (month_dt.year - signup.year) * 12 + month_dt.month - signup.month
            )

            # Base engagement
            base_logins = random.randint(10, 25)
            base_clicks = random.randint(50, 300)
            base_users = random.randint(2, 15)

            # PATTERN 3: Churned customers show declining logins in last 3 months
            if is_churned and c_date:
                months_to_churn = (
                    (c_date.year - month_dt.year) * 12
                    + c_date.month
                    - month_dt.month
                )
                if months_to_churn <= 3:
                    decay = 0.3 + 0.2 * (3 - months_to_churn)
                    base_logins = max(1, int(base_logins * (1 - decay)))
                    base_clicks = max(5, int(base_clicks * (1 - decay)))
                    base_users = max(1, int(base_users * (1 - decay)))

            login_days = min(31, max(0, base_logins + random.randint(-3, 3)))
            feature_clicks = max(0, base_clicks + random.randint(-20, 20))
            active_users = max(1, base_users + random.randint(-1, 2))

            # CSAT & NPS (quarterly-ish, so ~33% of months have scores)
            csat_score = None
            nps_score = None
            if random.random() < 0.35:
                # PATTERN 5: Churned customers get lower CSAT
                if is_churned:
                    csat_score = random.choices(
                        [1, 2, 3, 4, 5], weights=[0.20, 0.30, 0.30, 0.15, 0.05], k=1
                    )[0]
                    nps_score = random.choices(
                        list(range(0, 11)),
                        weights=[0.08, 0.10, 0.12, 0.12, 0.10, 0.10, 0.10, 0.08, 0.08, 0.07, 0.05],
                        k=1,
                    )[0]
                else:
                    csat_score = random.choices(
                        [1, 2, 3, 4, 5], weights=[0.03, 0.07, 0.20, 0.40, 0.30], k=1
                    )[0]
                    nps_score = random.choices(
                        list(range(0, 11)),
                        weights=[0.02, 0.03, 0.03, 0.05, 0.05, 0.08, 0.10, 0.15, 0.18, 0.16, 0.15],
                        k=1,
                    )[0]

            # DQ-01: Inject ~2-3% extra NULLs into csat_score / nps_score
            if random.random() < 0.025:
                csat_score = None
            if random.random() < 0.025:
                nps_score = None

            rows.append(
                {
                    "usage_id": f"USG-{usage_counter:06d}",
                    "customer_id": cid,
                    "activity_month": month_dt.strftime("%Y-%m"),
                    "login_days": login_days,
                    "core_feature_clicks": feature_clicks,
                    "active_users_count": active_users,
                    "csat_score": csat_score,
                    "nps_score": nps_score,
                }
            )

    return pd.DataFrame(rows)


# =========================================================================
# STEP 5: Create SQLite Database with Full DDL
# =========================================================================
DDL_SCRIPT = """
PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS dim_customers (
    customer_id         TEXT PRIMARY KEY,
    company_name        TEXT NOT NULL,
    industry            TEXT NOT NULL CHECK(industry IN ('EdTech', 'FinTech', 'Healthcare', 'Retail', 'Manufacturing')),
    company_size        TEXT NOT NULL CHECK(company_size IN ('Micro (1-10)', 'Small (11-50)', 'Medium (51-200)')),
    region              TEXT NOT NULL CHECK(region IN ('North', 'South', 'East', 'West')),
    acquisition_channel TEXT NOT NULL CHECK(acquisition_channel IN ('Organic Search', 'Referral', 'Paid Social', 'Direct', 'Partner')),
    signup_date         DATE NOT NULL,
    signup_cohort_month TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS dim_plans (
    plan_id             TEXT PRIMARY KEY,
    plan_name           TEXT NOT NULL UNIQUE CHECK(plan_name IN ('Starter', 'Professional', 'Enterprise')),
    plan_tier_level     INTEGER NOT NULL CHECK(plan_tier_level IN (1, 2, 3)),
    monthly_base_price  REAL NOT NULL CHECK(monthly_base_price >= 0.0),
    max_user_seats      INTEGER NOT NULL CHECK(max_user_seats > 0),
    sla_support_level   TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS fact_subscriptions (
    subscription_id     TEXT PRIMARY KEY,
    customer_id         TEXT NOT NULL,
    plan_id             TEXT NOT NULL,
    start_date          DATE NOT NULL,
    end_date            DATE NOT NULL,
    billing_cadence     TEXT NOT NULL CHECK(billing_cadence IN ('Monthly', 'Annual')),
    monthly_fee_inr     REAL NOT NULL CHECK(monthly_fee_inr >= 0.0),
    subscription_status TEXT NOT NULL CHECK(subscription_status IN ('Active', 'Churned', 'Renewed')),
    is_churned          INTEGER NOT NULL CHECK(is_churned IN (0, 1)),
    churn_date          DATE,
    tenure_months       INTEGER NOT NULL CHECK(tenure_months >= 1),
    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON DELETE RESTRICT ON UPDATE CASCADE,
    FOREIGN KEY (plan_id)     REFERENCES dim_plans(plan_id)         ON DELETE RESTRICT ON UPDATE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_sub_cust ON fact_subscriptions(customer_id);
CREATE INDEX IF NOT EXISTS idx_sub_plan ON fact_subscriptions(plan_id);

CREATE TABLE IF NOT EXISTS fact_support_tickets (
    ticket_id           TEXT PRIMARY KEY,
    customer_id         TEXT NOT NULL,
    created_at          DATETIME NOT NULL,
    resolved_at         DATETIME,
    resolution_hours    REAL CHECK(resolution_hours >= 0.0),
    priority            TEXT NOT NULL CHECK(priority IN ('Low', 'Medium', 'High', 'Critical')),
    category            TEXT NOT NULL CHECK(category IN ('Billing', 'Bug', 'Feature Request', 'Technical')),
    ticket_csat         INTEGER CHECK(ticket_csat BETWEEN 1 AND 5),
    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_ticket_cust ON fact_support_tickets(customer_id);

CREATE TABLE IF NOT EXISTS fact_monthly_usage (
    usage_id            TEXT PRIMARY KEY,
    customer_id         TEXT NOT NULL,
    activity_month      TEXT NOT NULL,
    login_days          INTEGER NOT NULL CHECK(login_days BETWEEN 0 AND 31),
    core_feature_clicks INTEGER NOT NULL CHECK(core_feature_clicks >= 0),
    active_users_count  INTEGER NOT NULL CHECK(active_users_count >= 0),
    csat_score          INTEGER CHECK(csat_score BETWEEN 1 AND 5),
    nps_score           INTEGER CHECK(nps_score BETWEEN 0 AND 10),
    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON DELETE CASCADE ON UPDATE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_usage_cust_month ON fact_monthly_usage(customer_id, activity_month);
"""

VIEW_SQL = """
CREATE VIEW IF NOT EXISTS vw_customer_retention_analytics AS
SELECT 
    c.customer_id,
    c.company_name,
    c.industry,
    c.company_size,
    c.region,
    c.acquisition_channel,
    c.signup_date,
    c.signup_cohort_month,
    p.plan_id,
    p.plan_name,
    p.plan_tier_level,
    p.monthly_base_price,
    s.subscription_id,
    s.start_date AS sub_start_date,
    s.end_date AS sub_end_date,
    s.billing_cadence,
    s.monthly_fee_inr,
    s.subscription_status,
    s.is_churned,
    s.churn_date,
    s.tenure_months,
    (s.monthly_fee_inr * s.tenure_months) AS lifetime_revenue_inr,
    CASE WHEN s.is_churned = 1 THEN s.monthly_fee_inr ELSE 0.0 END AS churned_mrr_inr,
    COALESCE(t.total_tickets, 0) AS total_support_tickets,
    COALESCE(t.avg_res_hours, 0.0) AS avg_resolution_hours,
    COALESCE(t.critical_tickets, 0) AS critical_tickets_count,
    COALESCE(u.avg_login_days, 0) AS avg_monthly_login_days,
    COALESCE(u.total_feature_clicks, 0) AS total_feature_clicks,
    u.latest_csat_score,
    u.latest_nps_score,
    CASE 
        WHEN u.latest_nps_score >= 9 THEN 'Promoter'
        WHEN u.latest_nps_score BETWEEN 7 AND 8 THEN 'Passive'
        WHEN u.latest_nps_score IS NOT NULL THEN 'Detractor'
        ELSE 'No Survey'
    END AS nps_category,
    ROUND(
        (CASE WHEN COALESCE(u.avg_login_days, 0) < 5 THEN 35.0 WHEN COALESCE(u.avg_login_days, 0) < 12 THEN 18.0 ELSE 0.0 END) +
        (CASE WHEN COALESCE(t.critical_tickets, 0) >= 2 THEN 25.0 WHEN COALESCE(t.critical_tickets, 0) = 1 THEN 12.0 ELSE 0.0 END) +
        (CASE WHEN u.latest_csat_score <= 2 THEN 20.0 WHEN u.latest_csat_score = 3 THEN 10.0 ELSE 0.0 END) +
        (CASE WHEN s.subscription_status = 'Active' AND julianday(s.end_date) - julianday('now') BETWEEN 0 AND 30 THEN 20.0 ELSE 0.0 END),
        1
    ) AS health_risk_score,
    CASE 
        WHEN ROUND(
            (CASE WHEN COALESCE(u.avg_login_days, 0) < 5 THEN 35.0 WHEN COALESCE(u.avg_login_days, 0) < 12 THEN 18.0 ELSE 0.0 END) +
            (CASE WHEN COALESCE(t.critical_tickets, 0) >= 2 THEN 25.0 WHEN COALESCE(t.critical_tickets, 0) = 1 THEN 12.0 ELSE 0.0 END) +
            (CASE WHEN u.latest_csat_score <= 2 THEN 20.0 WHEN u.latest_csat_score = 3 THEN 10.0 ELSE 0.0 END) +
            (CASE WHEN s.subscription_status = 'Active' AND julianday(s.end_date) - julianday('now') BETWEEN 0 AND 30 THEN 20.0 ELSE 0.0 END),
            1
        ) >= 70.0 THEN 'High'
        WHEN ROUND(
            (CASE WHEN COALESCE(u.avg_login_days, 0) < 5 THEN 35.0 WHEN COALESCE(u.avg_login_days, 0) < 12 THEN 18.0 ELSE 0.0 END) +
            (CASE WHEN COALESCE(t.critical_tickets, 0) >= 2 THEN 25.0 WHEN COALESCE(t.critical_tickets, 0) = 1 THEN 12.0 ELSE 0.0 END) +
            (CASE WHEN u.latest_csat_score <= 2 THEN 20.0 WHEN u.latest_csat_score = 3 THEN 10.0 ELSE 0.0 END) +
            (CASE WHEN s.subscription_status = 'Active' AND julianday(s.end_date) - julianday('now') BETWEEN 0 AND 30 THEN 20.0 ELSE 0.0 END),
            1
        ) >= 40.0 THEN 'Medium'
        ELSE 'Low'
    END AS risk_tier
FROM dim_customers c
JOIN fact_subscriptions s ON c.customer_id = s.customer_id
JOIN dim_plans p ON s.plan_id = p.plan_id
LEFT JOIN (
    SELECT customer_id, COUNT(*) AS total_tickets, AVG(resolution_hours) AS avg_res_hours, SUM(CASE WHEN priority IN ('High', 'Critical') THEN 1 ELSE 0 END) AS critical_tickets
    FROM fact_support_tickets GROUP BY customer_id
) t ON c.customer_id = t.customer_id
LEFT JOIN (
    SELECT customer_id, AVG(login_days) AS avg_login_days, SUM(core_feature_clicks) AS total_feature_clicks, MAX(csat_score) AS latest_csat_score, MAX(nps_score) AS latest_nps_score
    FROM fact_monthly_usage GROUP BY customer_id
) u ON c.customer_id = u.customer_id;
"""


def create_sqlite_db(df_customers, df_plans, df_subs, df_tickets, df_usage):
    db_path = OUT_DIR / "app.db"
    if db_path.exists():
        db_path.unlink()

    conn = sqlite3.connect(str(db_path))
    conn.executescript(DDL_SCRIPT)

    df_plans.to_sql("dim_plans", conn, if_exists="append", index=False)
    df_customers.to_sql("dim_customers", conn, if_exists="append", index=False)
    df_subs.to_sql("fact_subscriptions", conn, if_exists="append", index=False)

    # For tickets, we must handle duplicates - SQLite PK will reject exact dups
    # Insert with OR IGNORE to skip the 3 duplicate-PK rows
    for _, row in df_tickets.iterrows():
        try:
            conn.execute(
                "INSERT INTO fact_support_tickets VALUES (?,?,?,?,?,?,?,?)",
                (row["ticket_id"], row["customer_id"], row["created_at"],
                 row["resolved_at"], row["resolution_hours"], row["priority"],
                 row["category"],
                 int(row["ticket_csat"]) if pd.notna(row["ticket_csat"]) else None),
            )
        except sqlite3.IntegrityError:
            pass  # Skip duplicate PKs

    df_usage.to_sql("fact_monthly_usage", conn, if_exists="append", index=False)
    conn.executescript(VIEW_SQL)
    conn.commit()

    # Verify view column count
    cur = conn.execute("SELECT * FROM vw_customer_retention_analytics LIMIT 1")
    view_cols = [desc[0] for desc in cur.description]

    conn.close()
    return db_path, view_cols


# =========================================================================
# STEP 6: Validation Report
# =========================================================================
def run_validation(db_path, df_customers, df_plans, df_subs, df_tickets, df_usage):
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA foreign_keys = ON")

    lines = []
    lines.append("=" * 72)
    lines.append("CRMA SYNTHETIC DATA VALIDATION REPORT")
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Random Seed: {SEED}")
    lines.append("=" * 72)

    # Row counts
    lines.append("\n--- ROW COUNTS ---")
    tables = [
        ("dim_customers", df_customers),
        ("dim_plans", df_plans),
        ("fact_subscriptions", df_subs),
        ("fact_support_tickets (CSV incl. dups)", df_tickets),
        ("fact_monthly_usage", df_usage),
    ]
    for name, df in tables:
        db_count = conn.execute(
            f"SELECT COUNT(*) FROM {name.split(' ')[0]}"
        ).fetchone()[0]
        lines.append(f"  {name:<45} CSV: {len(df):>6}   DB: {db_count:>6}")

    view_count = conn.execute(
        "SELECT COUNT(*) FROM vw_customer_retention_analytics"
    ).fetchone()[0]
    view_cols = conn.execute(
        "SELECT * FROM vw_customer_retention_analytics LIMIT 1"
    ).description
    lines.append(f"  {'vw_customer_retention_analytics':<45} DB: {view_count:>6}   Columns: {len(view_cols)}")

    # FK Integrity
    lines.append("\n--- FOREIGN KEY INTEGRITY ---")
    fk_checks = [
        ("fact_subscriptions.customer_id -> dim_customers",
         "SELECT COUNT(*) FROM fact_subscriptions WHERE customer_id NOT IN (SELECT customer_id FROM dim_customers)"),
        ("fact_subscriptions.plan_id -> dim_plans",
         "SELECT COUNT(*) FROM fact_subscriptions WHERE plan_id NOT IN (SELECT plan_id FROM dim_plans)"),
        ("fact_support_tickets.customer_id -> dim_customers",
         "SELECT COUNT(*) FROM fact_support_tickets WHERE customer_id NOT IN (SELECT customer_id FROM dim_customers)"),
        ("fact_monthly_usage.customer_id -> dim_customers",
         "SELECT COUNT(*) FROM fact_monthly_usage WHERE customer_id NOT IN (SELECT customer_id FROM dim_customers)"),
    ]
    all_fk_pass = True
    for desc, sql in fk_checks:
        orphans = conn.execute(sql).fetchone()[0]
        status = "PASS" if orphans == 0 else f"FAIL ({orphans} orphans)"
        if orphans > 0:
            all_fk_pass = False
        lines.append(f"  {desc:<55} {status}")

    # Range checks
    lines.append("\n--- RANGE / DOMAIN CHECKS ---")
    range_checks = [
        ("monthly_fee_inr >= 0",
         "SELECT COUNT(*) FROM fact_subscriptions WHERE monthly_fee_inr < 0"),
        ("tenure_months >= 1",
         "SELECT COUNT(*) FROM fact_subscriptions WHERE tenure_months < 1"),
        ("login_days BETWEEN 0 AND 31",
         "SELECT COUNT(*) FROM fact_monthly_usage WHERE login_days < 0 OR login_days > 31"),
        ("csat_score BETWEEN 1 AND 5 (or NULL)",
         "SELECT COUNT(*) FROM fact_monthly_usage WHERE csat_score IS NOT NULL AND (csat_score < 1 OR csat_score > 5)"),
        ("nps_score BETWEEN 0 AND 10 (or NULL)",
         "SELECT COUNT(*) FROM fact_monthly_usage WHERE nps_score IS NOT NULL AND (nps_score < 0 OR nps_score > 10)"),
        ("resolution_hours >= 0",
         "SELECT COUNT(*) FROM fact_support_tickets WHERE resolution_hours < 0"),
    ]
    for desc, sql in range_checks:
        violations = conn.execute(sql).fetchone()[0]
        status = "PASS" if violations == 0 else f"FAIL ({violations} violations)"
        lines.append(f"  {desc:<55} {status}")

    # Data Quality Issues Summary
    lines.append("\n--- INJECTED DATA QUALITY ISSUES ---")
    null_csat = conn.execute(
        "SELECT COUNT(*) FROM fact_monthly_usage WHERE csat_score IS NULL"
    ).fetchone()[0]
    null_nps = conn.execute(
        "SELECT COUNT(*) FROM fact_monthly_usage WHERE nps_score IS NULL"
    ).fetchone()[0]
    total_usage = conn.execute("SELECT COUNT(*) FROM fact_monthly_usage").fetchone()[0]
    lines.append(f"  DQ-01: NULL csat_score rows:     {null_csat:>6} / {total_usage} ({null_csat/total_usage*100:.1f}%)")
    lines.append(f"  DQ-01: NULL nps_score rows:      {null_nps:>6} / {total_usage} ({null_nps/total_usage*100:.1f}%)")
    
    dup_tickets = len(df_tickets) - conn.execute("SELECT COUNT(*) FROM fact_support_tickets").fetchone()[0]
    lines.append(f"  DQ-02: Duplicate ticket rows:    {dup_tickets:>6} (removed in DB, present in CSV)")

    outliers = conn.execute(
        "SELECT COUNT(*) FROM fact_support_tickets WHERE resolution_hours > 200"
    ).fetchone()[0]
    lines.append(f"  DQ-03: Outlier resolution_hours: {outliers:>6} (> 200 hours)")

    # Pattern Verification
    lines.append("\n--- PATTERN VERIFICATION ---")

    # P1: Churn rate by plan
    lines.append("  PATTERN 1: Higher churn for Starter vs Enterprise")
    for pid in ["PLAN-01", "PLAN-02", "PLAN-03"]:
        total = conn.execute(f"SELECT COUNT(DISTINCT customer_id) FROM fact_subscriptions WHERE plan_id = '{pid}'").fetchone()[0]
        churned = conn.execute(f"SELECT COUNT(DISTINCT customer_id) FROM fact_subscriptions WHERE plan_id = '{pid}' AND is_churned = 1").fetchone()[0]
        rate = (churned / total * 100) if total > 0 else 0
        pname = next(p["plan_name"] for p in PLANS if p["plan_id"] == pid)
        lines.append(f"    {pname:<15} Churned: {churned:>4}/{total:<4} ({rate:.1f}%)")

    # P2: Avg tickets churned vs retained
    lines.append("  PATTERN 2: Churned customers have more support tickets")
    avg_tick_churned = conn.execute("""
        SELECT AVG(cnt) FROM (
            SELECT s.customer_id, COUNT(t.ticket_id) as cnt
            FROM fact_subscriptions s
            LEFT JOIN fact_support_tickets t ON s.customer_id = t.customer_id
            WHERE s.is_churned = 1
            GROUP BY s.customer_id
        )
    """).fetchone()[0] or 0
    avg_tick_retained = conn.execute("""
        SELECT AVG(cnt) FROM (
            SELECT s.customer_id, COUNT(t.ticket_id) as cnt
            FROM fact_subscriptions s
            LEFT JOIN fact_support_tickets t ON s.customer_id = t.customer_id
            WHERE s.is_churned = 0
            GROUP BY s.customer_id
        )
    """).fetchone()[0] or 0
    lines.append(f"    Avg tickets (churned):  {avg_tick_churned:.1f}")
    lines.append(f"    Avg tickets (retained): {avg_tick_retained:.1f}")

    # P4: Churn by channel
    lines.append("  PATTERN 4: Channel churn rates (Referral lowest, Paid Social highest)")
    for ch in CHANNELS:
        total = conn.execute(f"SELECT COUNT(DISTINCT c.customer_id) FROM dim_customers c JOIN fact_subscriptions s ON c.customer_id = s.customer_id WHERE c.acquisition_channel = '{ch}'").fetchone()[0]
        churned = conn.execute(f"SELECT COUNT(DISTINCT c.customer_id) FROM dim_customers c JOIN fact_subscriptions s ON c.customer_id = s.customer_id WHERE c.acquisition_channel = '{ch}' AND s.is_churned = 1").fetchone()[0]
        rate = (churned / total * 100) if total > 0 else 0
        lines.append(f"    {ch:<18} {churned:>4}/{total:<4} ({rate:.1f}%)")

    lines.append("\n" + "=" * 72)
    lines.append("OVERALL RESULT: ALL CHECKS PASSED" if all_fk_pass else "OVERALL RESULT: SOME CHECKS FAILED")
    lines.append("=" * 72)

    conn.close()
    report = "\n".join(lines)
    report_path = OUT_DIR / "validation_report.txt"
    report_path.write_text(report, encoding="utf-8")
    print(report)
    return report_path


# =========================================================================
# MAIN EXECUTION
# =========================================================================
def main():
    print("=" * 60)
    print("CRMA Synthetic Data Generation Engine - Step 10")
    print("=" * 60)

    # Generate tables
    print("\n[1/6] Generating dim_customers...")
    df_customers = generate_dim_customers()
    print(f"      -> {len(df_customers)} rows")

    df_plans = pd.DataFrame(PLANS)
    print(f"[2/6] dim_plans (catalog)... -> {len(df_plans)} rows")

    print("[3/6] Generating fact_subscriptions...")
    df_subs = generate_fact_subscriptions(df_customers)
    print(f"      -> {len(df_subs)} rows")

    print("[4/6] Generating fact_support_tickets...")
    df_tickets = generate_fact_support_tickets(df_customers, df_subs)
    print(f"      -> {len(df_tickets)} rows (incl. 3 intentional duplicates)")

    print("[5/6] Generating fact_monthly_usage...")
    df_usage = generate_fact_monthly_usage(df_customers, df_subs)
    print(f"      -> {len(df_usage)} rows")

    # Save CSVs
    print("\n[6/6] Saving outputs...")
    df_customers.to_csv(OUT_DIR / "dim_customers.csv", index=False)
    df_plans.to_csv(OUT_DIR / "dim_plans.csv", index=False)
    df_subs.to_csv(OUT_DIR / "fact_subscriptions.csv", index=False)
    df_tickets.to_csv(OUT_DIR / "fact_support_tickets.csv", index=False)
    df_usage.to_csv(OUT_DIR / "fact_monthly_usage.csv", index=False)
    print("      CSVs saved.")

    # Save Excel (one sheet per table)
    xlsx_path = OUT_DIR / "crma_data.xlsx"
    with pd.ExcelWriter(str(xlsx_path), engine="openpyxl") as writer:
        df_customers.to_excel(writer, sheet_name="dim_customers", index=False)
        df_plans.to_excel(writer, sheet_name="dim_plans", index=False)
        df_subs.to_excel(writer, sheet_name="fact_subscriptions", index=False)
        df_tickets.to_excel(writer, sheet_name="fact_support_tickets", index=False)
        df_usage.to_excel(writer, sheet_name="fact_monthly_usage", index=False)
    print(f"      Excel saved: {xlsx_path}")

    # Create SQLite DB
    db_path, view_cols = create_sqlite_db(
        df_customers, df_plans, df_subs, df_tickets, df_usage
    )
    print(f"      SQLite DB saved: {db_path}")
    print(f"      Analytical View columns: {len(view_cols)}")

    # Validation
    print("\n" + "=" * 60)
    print("RUNNING VALIDATION...")
    print("=" * 60)
    report_path = run_validation(
        db_path, df_customers, df_plans, df_subs, df_tickets, df_usage
    )
    print(f"\nValidation report saved: {report_path}")
    print("\nStep 10 COMPLETE.")


if __name__ == "__main__":
    main()
