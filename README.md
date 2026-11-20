# Customer Retention Metrics Analysis (CRMA)

[![Python 3.12+](https://img.shields.io/badge/python-3.12%2B-blue.svg)](https://www.python.org/)
[![Streamlit 1.32](https://img.shields.io/badge/streamlit-1.32.2-red.svg)](https://streamlit.io/)
[![Database SQLite](https://img.shields.io/badge/database-SQLite%203-lightgrey.svg)](https://sqlite.org/)
[![Tests Passing](https://img.shields.io/badge/tests-35%20passed-brightgreen.svg)]()
[![Academic MCA](https://img.shields.io/badge/MCA%20Capstone-ABES%20Engineering%20College-orange.svg)]()

> **Enterprise B2B SaaS Customer Retention, Cohort Survival & Churn Diagnostics Platform**  
> Developed by **Sumedha** (Roll: 2503201450146), MCA 2nd Year, ABES Engineering College  
> Faculty Mentor: **Ms. Anjali Pandey** | Fictitious Enterprise: **NoidaSoft Technologies Pvt. Ltd.**

---

## 📌 Executive Summary

The Customer Retention Metrics Analysis (CRMA) platform is an enterprise-grade analytics application designed to track, diagnose, and forecast customer retention, logo churn, and recurring revenue decay over a 24-month horizon (January 2024 to December 2025). Built on an embedded SQLite database (`data/app.db`) containing 200 customer accounts, 229 contract lifecycles, 704 customer service tickets, and 2,489 monthly product telemetry records, the application empowers executive leadership and customer success teams with real-time operational insights.

### Key Capabilities
* **Executive Retention Cockpit:** Instant tracking of Customer Churn Rate (CCR %), Customer Retention Rate (CRR %), Churned MRR (₹), and Net Revenue Retention (NRR %) with month-over-month delta badges.
* **Triangular Cohort Survival Analysis:** Visualizing signup-month customer retention decay at months 1, 3, 6, and 12 via Plotly heatmaps.
* **Firmographic Segmentation:** Dynamic multi-attribute slicing across Plan Tier (Starter, Pro, Enterprise), Industry, Company Size, Region, and Acquisition Channel.
* **Inferential Statistical Hypothesis Testing:** Rigorous validation of churn drivers using Mann-Whitney U tests for support SLA turnaround, two-sample t-tests for 90-day pre-churn telemetry drop, Chi-Square tests of independence for acquisition channels, and Pearson/Spearman satisfaction correlations.
* **Proactive Customer Health Risk Roster:** Automated 4-factor composite risk scoring (RULE-05: 35% login drop + 25% tickets + 20% low CSAT + 20% renewal proximity) ranking at-risk accounts with client-side CSV export.
* **Transparent Data Governance:** Complete analytical audit, methodology expander drawers beneath every visualization, and interactive data dictionaries.

---

## 🗄 Relational Schema Architecture

The platform operates over a star-schema SQLite relational database (`data/app.db`):
* `dim_customers` (8 columns, 200 rows): Customer firmographic profiles and acquisition dates.
* `dim_plans` (6 columns, 3 rows): Product subscription tiers (Starter, Professional, Enterprise).
* `fact_subscriptions` (11 columns, 229 rows): Subscription contracts, billing cadences, and churn dates.
* `fact_support_tickets` (8 columns, 704 clean rows): Support SLA turnaround hours and priority rankings.
* `fact_monthly_usage` (8 columns, 2,489 rows): Monthly login activity, feature clicks, and survey scores.
* `vw_customer_retention_analytics` (33 columns, 229 rows): Analytical view pre-joining all entities and computing health risk scores.

---

## 🚀 Quickstart & Installation

### 1. Clone & Set Up Virtual Environment
```bash
git clone https://github.com/<your-username>/<rollno>-crma.git
cd <rollno>-crma

python -m venv .venv
# On Windows:
.venv\Scripts\activate
# On Linux / macOS:
source .venv/bin/activate

pip install --upgrade pip
pip install -r requirements.txt
```

### 2. Launch Local Streamlit Application
```bash
streamlit run app.py
```
Open your browser at `http://localhost:8501`.

### 3. Run Automated Pytest Test Suite
```bash
# Run all 35 unit, integration, and smoke tests:
pytest tests/ -v

# Run with test coverage audit:
pytest --cov=src tests/ --cov-report=term-missing
```

---

## 📂 Project Structure

```text
├── .streamlit/config.toml          # Theme tokens (Navy #1B365D, Steel #2B4C7E)
├── data/
│   ├── app.db                     # Embedded SQLite database (read-only mode)
│   ├── crma_data.xlsx             # Multi-tab data workbook
│   └── *.csv                      # Raw source tables
├── docs/                          # All 14 project deliverables (Word, Excel, ER Diagrams)
├── pages/
│   ├── 1_Cohorts.py               # UI-02: Cohort Survival Heatmap
│   ├── 2_Segments.py              # UI-03: Multi-attribute Slicing & CLV
│   ├── 3_Drivers.py               # UI-04: Support SLA & Statistical Testing
│   ├── 4_Risk_Roster.py           # UI-05: Health Risk Roster & CSV Export
│   ├── 5_Data_Quality.py          # UI-06: Missingness Heatmaps & Profiling
│   ├── 6_Methodology.py           # UI-08: Formal Formulas & Data Dictionary
│   └── 7_About.py                 # UI-07: Academic Attribution
├── src/
│   ├── components.py              # Reusable UI widgets & methodology drawers
│   ├── data_loader.py             # SQLite connection, @st.cache_data, & validation
│   ├── filters.py                 # Shared st.session_state sidebar controls
│   ├── metrics.py                 # CCR, CRR, NRR, CLV, & Cohort calculations
│   ├── risk_scoring.py            # RULE-05 4-factor composite risk scoring engine
│   ├── stats.py                   # SciPy statistical testing implementations
│   └── ui.py                      # Tailwind HTML components via iframe bridge
├── tests/
│   ├── test_data.py               # TC-T-01..14 Data layer assertions
│   ├── test_analysis.py           # TC-T-15..27 Analytical formulas
│   └── test_app.py                # TC-T-28..35 Streamlit AppTest smoke tests
├── app.py                         # UI-01: Executive Overview entrypoint
├── BUILD_PLAN.md                  # Task-by-task implementation blueprint
├── README.md                      # This document
└── requirements.txt               # Pinned dependencies
```

---

## 🎓 Academic Attribution

* **Learner:** Sumedha | Roll: 2503201450146
* **Program:** MCA 2nd Year (Data Analytics), ABES Engineering College
* **Mentor:** Ms. Anjali Pandey
* **Course:** Final Year Capstone Project
