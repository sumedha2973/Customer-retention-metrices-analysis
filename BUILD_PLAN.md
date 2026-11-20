# BUILD PLAN: Customer Retention Metrics Analysis (CRMA)

**Project:** Customer Retention Metrics Analysis (CRMA)  
**Learner:** Sumedha | Roll: 2503201450146 | MCA 2nd Year, ABES Engineering College  
**Mentor:** Ms. Anjali Pandey | Domain: B2B SaaS Customer Retention Analytics  
**Product:** Streamlit Multipage App (Python 3.12, SQLite, Plotly, Tailwind CSS via components, Pytest)  
**Status:** Baselined for Step 15 Execution  

---

## 1. Project Summary & Document Registry

The Customer Retention Metrics Analysis (CRMA) system is an enterprise-grade analytics platform designed for NoidaSoft Technologies Pvt. Ltd., a fictitious mid-market B2B SaaS enterprise. The application tracks, diagnoses, and forecasts customer retention, subscription churn, and revenue decay over a 24-month horizon (January 2024 to December 2025). Built on an embedded SQLite database (`data/app.db`) containing 200 customer accounts, 229 subscription lifecycles, 704 customer service tickets, and 2,489 monthly usage records, the platform provides real-time visibility into customer health. It features interactive cohort survival heatmaps, multi-attribute segment drill-downs, statistical hypothesis testing (Mann-Whitney U, Student's t-test, Chi-Square, Pearson/Spearman correlations), an automated 4-factor customer health risk scoring model, and full data governance.

### Project Document Registry
* [01_Project_Understanding.docx](file:///c:/3rd%20sem/HCL-anti/docs/01_Project_Understanding.docx) — Problem statement, 10 Stakeholders (SH-01..10), 10 Operational Questions (OQ-01..10), SMART objectives, and measurable success criteria.
* [02_HMW_Statements.docx](file:///c:/3rd%20sem/HCL-anti/docs/02_HMW_Statements.docx) — Design thinking synthesis, stakeholder pain points, and 10 How Might We statements (HMW-01..10).
* [03_Needs_Analysis.docx](file:///c:/3rd%20sem/HCL-anti/docs/03_Needs_Analysis.docx) — 12 Business Requirements (BR-01..12) with MoSCoW prioritization, measurement criteria, and gap analysis.
* [04_Traceability_Matrix.xlsx](file:///c:/3rd%20sem/HCL-anti/docs/04_Traceability_Matrix.xlsx) — Complete bidirectional RTM (v2.0 Final) tracing OQ $\to$ HMW $\to$ BR $\to$ DB Column $\to$ US $\to$ UI $\to$ TC.
* [05_BRD.docx](file:///c:/3rd%20sem/HCL-anti/docs/05_BRD.docx) — Formal Business Requirements Document, RACI matrix, 7 business calculation rules (RULE-01..07), and data governance.
* [07_TRD.docx](file:///c:/3rd%20sem/HCL-anti/docs/07_TRD.docx) — Technical Architecture, 12 Functional Requirements (FR-01..12), 5 NFRs, Streamlit caching strategy, and deployment spec.
* [08_Database_Design.xlsx](file:///c:/3rd%20sem/HCL-anti/docs/08_Database_Design.xlsx) — Relational schema (5 tables, 1 analytical view `vw_customer_retention_analytics`, 33 columns), DDL, and 11 KPI definitions.
* [09_ER_Diagram.docx](file:///c:/3rd%20sem/HCL-anti/docs/09_ER_Diagram.docx) & [09_ER_Diagram.png](file:///c:/3rd%20sem/HCL-anti/docs/09_ER_Diagram.png) — Relational Entity-Relationship diagram, cardinality definitions, and schema normalization narrative.
* [crma_data.xlsx](file:///c:/3rd%20sem/HCL-anti/data/crma_data.xlsx) & [validation_report.txt](file:///c:/3rd%20sem/HCL-anti/data/validation_report.txt) — Verified synthetic dataset and data quality audit report (Seed 42).
* [11_User_Stories.docx](file:///c:/3rd%20sem/HCL-anti/docs/11_User_Stories.docx) — 16 Agile User Stories (US-01..16) across 5 Epics with Given/When/Then acceptance criteria and 5-sprint plan.
* [12_UI_Framework.docx](file:///c:/3rd%20sem/HCL-anti/docs/12_UI_Framework.docx) — Multipage UI layout, ASCII wireframes for 8 screens (UI-01..08), Tailwind design tokens, and best-chart rationale.
* [13_Test_Cases.docx](file:///c:/3rd%20sem/HCL-anti/docs/13_Test_Cases.docx) — Test specification with 20 Functional Tests (TC-F-01..20) and 35 Automated Technical Tests (TC-T-01..35).

---

## 2. Final Project Folder Structure

```text
c:\3rd sem\HCL-anti\
├── .streamlit/
│   └── config.toml               # Streamlit theme (primary: #2B4C7E, background: #FFFFFF)
├── data/
│   ├── app.db                    # Embedded SQLite database (clean analytical view + 5 tables)
│   ├── crma_data.xlsx            # Master multi-tab Excel workbook
│   ├── dim_customers.csv         # 200 customer accounts
│   ├── dim_plans.csv             # 3 subscription plans (Starter, Pro, Enterprise)
│   ├── fact_subscriptions.csv    # 229 contracts (active, churned, renewed)
│   ├── fact_support_tickets.csv  # 707 tickets (including 3 raw injection duplicates)
│   ├── fact_monthly_usage.csv    # 2,489 monthly product telemetry records
│   └── validation_report.txt     # Data generation integrity report
├── docs/                         # All 14 project stage deliverables (Word, Excel, Images)
├── pages/
│   ├── 1_Cohorts.py              # UI-02: Cohort Survival Heatmap & Decay Curves
│   ├── 2_Segments.py             # UI-03: Multi-attribute Slicing, CLV, & Channel Retention
│   ├── 3_Drivers.py              # UI-04: Support Impact, Pre-churn Usage Drop, CSAT Correlations
│   ├── 4_Risk_Roster.py          # UI-05: Customer Health Risk Roster & CSV Export
│   ├── 5_Data_Quality.py         # UI-06: Data Profiling, Missing Values Heatmap, & DQ Audits
│   ├── 6_Methodology.py          # UI-08: Formal Formulas, Statistical Specs, & Data Dictionary
│   └── 7_About.py                # UI-07: Academic & Project Metadata, Team, Architecture
├── src/
│   ├── __init__.py
│   ├── components.py             # Reusable UI widgets (Tailwind KPI cards, Stat badges, Drawers)
│   ├── data_loader.py            # SQLite read-only connection, @st.cache_data, validation assertions
│   ├── filters.py                # Shared st.session_state sidebar filter controls & reset logic
│   ├── metrics.py                # Core formulas: CCR, CRR, NRR, CLV, Cohort Survival Matrix
│   ├── risk_scoring.py           # RULE-05 4-factor composite weighted risk score engine
│   ├── stats.py                  # Hypothesis testing: Mann-Whitney U, t-test, Chi-Square, Pearson/Spearman
│   └── ui.py                     # Tailwind HTML component templates via st.components.v1.html
├── tests/
│   ├── __init__.py
│   ├── test_data.py              # TC-T-01..14: Schema validation, row counts, PK/FK, range bounds
│   ├── test_analysis.py          # TC-T-15..27: Formula accuracy, known datasets, statistical rigor
│   └── test_app.py               # TC-T-28..35: Streamlit AppTest smoke tests, latency < 3s, error bounds
├── app.py                        # UI-01: Executive Overview landing page & multi-tab navigation
├── BUILD_PLAN.md                 # This implementation runbook
├── generate_doc.py               # Automation helper
├── README.md                     # Comprehensive project documentation & local deployment runbook
└── requirements.txt              # Pinned deterministic dependencies
```

---

## 3. Step-by-Step Build Tasks (Step 15 Execution Blueprint)

### BUILD-01: Setup, Environment & Pinned Dependencies
* **Goal:** Establish deterministic build environment, pin requirements, verify SQLite database, and configure Streamlit defaults.
* **Files Affected:** `requirements.txt`, `.streamlit/config.toml`
* **Linked Stories & FRs:** Technical Foundation, NFR-05
* **Exact Prompt for AI:**
  > "Create a strictly pinned `requirements.txt` containing streamlit==1.32.2, pandas==2.2.1, numpy==1.26.4, scipy==1.12.0, statsmodels==0.14.1, plotly==5.20.0, openpyxl==3.1.5, pytest==8.1.1, pytest-cov==5.0.0. Ensure `.streamlit/config.toml` has headless=true, primaryColor='#2B4C7E', backgroundColor='#FFFFFF', and secondaryBackgroundColor='#F0F2F6'."
* **Verification Check:** Run `pip list` or import packages in Python; check config.toml format.
* **Must-Pass Tests:** `tests/test_data.py::test_TC_T_01_db_file_exists`

---

### BUILD-02: Data Layer, Validation Engine & In-Memory Cache
* **Goal:** Create `src/data_loader.py` to establish read-only SQLite connections, implement `@st.cache_data` caching with type optimization, and execute schema validation assertions on load.
* **Files Affected:** `src/data_loader.py`
* **Linked Stories & FRs:** US-01, FR-01, FR-02, NFR-01, NFR-03
* **Exact Prompt for AI:**
  > "Implement `src/data_loader.py`. Define functions `load_table(table_name)` and `load_analytical_view()` decorated with `@st.cache_data(ttl=3600)`. Connect to `data/app.db` via URI `file:data/app.db?mode=ro`. Downcast string categorical columns (`plan_name`, `industry`, `region`, `acquisition_channel`) to `category` dtype. Execute validation assertions: assert `customer_id` is unique, assert no negative subscription fees, assert 100% foreign key integrity. Provide clean try-except error logging using `st.error()` without raw tracebacks."
* **Verification Check:** Import `src/data_loader.py` and invoke `load_analytical_view()`; confirm it returns a 229-row, 33-column DataFrame within 200ms.
* **Must-Pass Tests:** `TC-T-01` through `TC-T-14`

---

### BUILD-03: Business Metrics & Cohort Engine
* **Goal:** Implement the mathematical calculation engine in `src/metrics.py` for CCR %, CRR %, Gross MRR Churn, Net Revenue Retention (NRR %), Customer Lifetime Value (CLV), and the $24 \times 12$ triangular cohort survival matrix.
* **Files Affected:** `src/metrics.py`
* **Linked Stories & FRs:** US-03, US-04, US-06, US-07, US-15, FR-01, FR-03, FR-06, FR-08
* **Exact Prompt for AI:**
  > "Write `src/metrics.py`. Implement vectorized pandas functions:
  > 1. `calculate_churn_retention_rates(df, date_col)`: computes monthly and quarterly CCR and CRR matching RULE-01 and RULE-02. Include zero-division guards returning 0.0.
  > 2. `calculate_mrr_movements(df)`: computes Gross Starting MRR, Expansion MRR, Churned MRR, and NRR % matching RULE-03.
  > 3. `calculate_clv_by_segment(df, group_col)`: computes CLV = AVG(monthly_fee_inr) * AVG(tenure_months) matching RULE-04.
  > 4. `build_cohort_survival_matrix(df)`: groups subscriptions by `signup_cohort_month` and computes % survival at lifecycle intervals (1, 3, 6, 12 months)."
* **Verification Check:** Execute functions against sample DataFrames; verify returned numbers match hand calculations.
* **Must-Pass Tests:** `TC-T-15`, `TC-T-16`, `TC-T-17`, `TC-T-18`, `TC-T-19`, `TC-T-26`, `TC-T-27`

---

### BUILD-04: Statistical Testing & Risk Scoring Engines
* **Goal:** Build `src/stats.py` for inferential hypothesis testing and `src/risk_scoring.py` for the 4-factor composite customer health risk algorithm.
* **Files Affected:** `src/stats.py`, `src/risk_scoring.py`
* **Linked Stories & FRs:** US-08, US-09, US-10, US-11, US-12, FR-04, FR-05, FR-09, FR-10, FR-11
* **Exact Prompt for AI:**
  > "1. In `src/stats.py`, implement functions wrapping SciPy:
  >    - `run_support_impact_test(churned_tix, retained_tix)`: Mann-Whitney U test with effect size $r$.
  >    - `run_prechurn_telemetry_ttest(churned_usage, retained_usage)`: Two-sample Welch's t-test with 95% confidence intervals.
  >    - `run_channel_contingency_test(df)`: Chi-square test of independence on a $2 \times 5$ contingency matrix.
  >    - `run_satisfaction_correlation(df)`: Pearson $r$ and Spearman $\rho$ between CSAT/NPS and renewal flag.
  > 2. In `src/risk_scoring.py`, implement `compute_composite_risk_score(df)` applying RULE-05:
  >    $\text{Score} = 0.35 \times \text{LoginDrop} + 0.25 \times \text{OpenTickets} + 0.20 \times \text{LowCSAT} + 0.20 \times \text{RenewalProximity}$.
  >    Clamp scores to $[0, 100]$ and assign tiers: High ($\ge 70$), Medium ($40-69$), Low ($< 40$)."
* **Verification Check:** Confirm p-values and statistics return valid floating-point numbers; confirm risk tiers correctly partition customers.
* **Must-Pass Tests:** `TC-T-20`, `TC-T-21`, `TC-T-22`, `TC-T-23`, `TC-T-24`, `TC-T-25`

---

### BUILD-05: Reusable UI Components & Shared Global Filters
* **Goal:** Expand `src/ui.py` and `src/components.py` with Tailwind HTML templates, chart container wrappers, stat result cards, and implement `src/filters.py` for shared `st.session_state` sidebar controls.
* **Files Affected:** `src/ui.py`, `src/components.py`, `src/filters.py`
* **Linked Stories & FRs:** US-02, US-05, FR-02, FR-07, NFR-04
* **Exact Prompt for AI:**
  > "1. In `src/components.py`, implement `render_methodology_drawer(title, formula, source, sample_size, assumptions)` using `st.expander()`. Implement `render_chart_wrapper(fig, title, height)` using `st.plotly_chart(use_container_width=True)`.
  > 2. In `src/filters.py`, write `init_global_filters()` to register `plan_filter`, `industry_filter`, `size_filter`, `region_filter`, and `date_range` in `st.session_state`. Write `render_sidebar_filters(df)` rendering multiselects and a 'Reset Filters' button that triggers `st.rerun()`. Write `apply_filters(df)` to return the filtered DataFrame."
* **Verification Check:** Verify session state keys initialize properly; test filter reset button functionality.
* **Must-Pass Tests:** `TC-F-07`, `TC-F-08`, `TC-F-19`, `TC-T-34`

---

### BUILD-06: Screen UI-01 — Executive Overview (`app.py`)
* **Goal:** Construct the root landing page featuring top Tailwind KPI metric cards, 24-month CCR/CRR dual-axis trend line, and 24-month MRR revenue waterfall/line chart.
* **Files Affected:** `app.py`
* **Linked Stories & FRs:** US-03, US-04, US-15, FR-01, FR-06, UI-01
* **Exact Prompt for AI:**
  > "Construct `app.py`. Call `st.set_page_config(page_title='CRMA - Executive Overview', layout='wide')`. Render the Tailwind header. Instantiate global sidebar filters. Compute latest month's KPIs (CCR %, CRR %, Churned MRR ₹, NRR %, Active Customers) with MoM delta badges rendered via `render_kpi_row()`. Render two Plotly charts:
  > 1. Dual-axis line chart for 24-month Churn Rate vs Retention Rate.
  > 2. Multi-line chart showing Gross MRR, Churned MRR, and Net MRR over 24 months.
  > Attach `render_methodology_drawer()` under each chart."
* **Verification Check:** Run `streamlit run app.py`; verify page renders in browser within 2 seconds.
* **Must-Pass Tests:** `TC-F-01`, `TC-F-02`, `TC-F-03`, `TC-F-04`, `TC-T-28`, `TC-T-30`, `TC-T-31`, `TC-T-33`

---

### BUILD-07: Screen UI-02 — Cohort Analysis (`pages/1_Cohorts.py`)
* **Goal:** Build the cohort lifecycle analysis page displaying the triangular survival heatmap and multi-line cohort decay curves.
* **Files Affected:** `pages/1_Cohorts.py`
* **Linked Stories & FRs:** US-06, FR-03, UI-02
* **Exact Prompt for AI:**
  > "Implement `pages/1_Cohorts.py`. Connect to data via `src/data_loader.py`. Render section header. Build a triangular survival heatmap using Plotly (`px.imshow` or `go.Heatmap`) showing Signup Cohort (rows) vs Lifecycle Month (M1, M3, M6, M12) with green-to-red color scale and percentage labels. Add a secondary Plotly chart displaying survival decay curves for selected cohorts. Include methodology expander and summary interpretation box."
* **Verification Check:** Open Cohorts page; verify heatmap grid matches expected retention rates (Month 1 > Month 12).
* **Must-Pass Tests:** `TC-F-05`, `TC-F-06`, `TC-T-32`

---

### BUILD-08: Screen UI-03 — Segment Comparison (`pages/2_Segments.py`)
* **Goal:** Build multidimensional slicing page comparing retention rates across Plan, Industry, Company Size, Region, and Acquisition Channel, alongside CLV analysis.
* **Files Affected:** `pages/2_Segments.py`
* **Linked Stories & FRs:** US-05, US-07, US-08, FR-02, FR-08, FR-09, UI-03
* **Exact Prompt for AI:**
  > "Implement `pages/2_Segments.py`. Apply active global filters. Build:
  > 1. Grouped bar chart comparing Retention Rate (%) across selected segment dimensions.
  > 2. Customer Lifetime Value (CLV ₹) grouped bar chart by Plan Tier and Industry.
  > 3. Horizontal bar chart displaying 12-month retention by Acquisition Channel, annotated with Chi-Square test p-value badge and statistical significance note.
  > Handle empty filter results gracefully with an informative warning banner."
* **Verification Check:** Test slicing by different industries; confirm CLV for Enterprise > Starter.
* **Must-Pass Tests:** `TC-F-07`, `TC-F-08`, `TC-F-09`, `TC-F-10`, `TC-F-19`, `TC-T-32`

---

### BUILD-09: Screen UI-04 — Churn Drivers (`pages/3_Drivers.py`)
* **Goal:** Build the statistical insights page analyzing the root causes of churn across Support SLA turnaround, pre-churn usage telemetry drop, and CSAT/NPS survey correlations.
* **Files Affected:** `pages/3_Drivers.py`
* **Linked Stories & FRs:** US-09, US-10, US-11, FR-05, FR-10, FR-11, UI-04
* **Exact Prompt for AI:**
  > "Implement `pages/3_Drivers.py`. Display three analytical sections:
  > 1. Support Impact: Side-by-side Plotly box plots for ticket counts and resolution hours (Churned vs Retained) accompanied by Mann-Whitney U test result card (`render_stat_card()`).
  > 2. Pre-Churn Telemetry: Dual-line chart showing 90-day trajectory of login days and feature clicks prior to churn, paired with two-sample t-test result card.
  > 3. Satisfaction Correlation: Scatter plot with OLS trendline showing CSAT vs renewal flag, reporting Pearson $r$, Spearman $\rho$, and sample size notice (noting 65.8% null survey rate)."
* **Verification Check:** Confirm all three statistical test cards display formatted p-values and plain-English interpretations.
* **Must-Pass Tests:** `TC-F-11`, `TC-F-12`, `TC-F-13`, `TC-T-32`

---

### BUILD-10: Screen UI-05 — Risk Roster & Data Export (`pages/4_Risk_Roster.py`)
* **Goal:** Build the proactive customer retention roster ranking accounts by health risk score, complete with risk tier badges and client-side CSV download.
* **Files Affected:** `pages/4_Risk_Roster.py`
* **Linked Stories & FRs:** US-12, US-13, US-16, FR-04, FR-12, UI-05
* **Exact Prompt for AI:**
  > "Implement `pages/4_Risk_Roster.py`. Calculate real-time health risk scores for all active accounts. Display top KPI summary cards for High, Medium, and Low risk account counts. Provide sidebar filters for Risk Tier and Minimum Risk Score slider. Render interactive table sorted descending by `health_risk_score`, formatting risk tiers with semantic badges. Implement client-side `st.download_button()` exporting the filtered roster to CSV with timestamped filename `at_risk_customers_YYYYMMDD_HHMMSS.csv`."
* **Verification Check:** Sort table by risk score; verify clicking 'Download CSV' triggers browser file download.
* **Must-Pass Tests:** `TC-F-14`, `TC-F-15`, `TC-F-16`, `TC-T-32`

---

### BUILD-11: Screen UI-06 — Data Quality & EDA (`pages/5_Data_Quality.py`)
* **Goal:** Construct transparent data profiling page displaying dataset statistics, missing value heatmaps, distributions, and documentation of injected DQ issues.
* **Files Affected:** `pages/5_Data_Quality.py`
* **Linked Stories & FRs:** US-01, UI-06
* **Exact Prompt for AI:**
  > "Implement `pages/5_Data_Quality.py`. Display dataset overview cards (total tables, records, completeness %). Render a missing-values heatmap showing null percentages per column across all tables (highlighting DQ-01 CSAT/NPS nulls). Plot distribution histograms for `monthly_fee_inr`, `tenure_months`, and `resolution_hours`. Present a structured audit table explaining the three intentional DQ issues (DQ-01 survey nulls, DQ-02 duplicate tickets, DQ-03 resolution outliers) and how the pipeline handled each."
* **Verification Check:** Check that DQ-01 (65.8% nulls) and DQ-02 (deduplication) are clearly documented with visualizations.
* **Must-Pass Tests:** `TC-T-10`, `TC-T-32`

---

### BUILD-12: Screens UI-07 & UI-08 — About & Methodology (`pages/6_Methodology.py`, `pages/7_About.py`)
* **Goal:** Build the academic governance, methodology, and project attribution pages.
* **Files Affected:** `pages/6_Methodology.py`, `pages/7_About.py`
* **Linked Stories & FRs:** US-02, US-14, FR-07, UI-07, UI-08
* **Exact Prompt for AI:**
  > "1. In `pages/6_Methodology.py`, render:
  >    - Comprehensive KPI Catalog detailing all 11 measures (MSR-01..11) with formulas matching Step 8.
  >    - Statistical Test Specifications (hypotheses, test statistic, $\alpha = 0.05$).
  >    - Complete Data Dictionary tables for all 5 entities and the analytical view.
  > 2. In `pages/7_About.py`, render project attribution (Learner: Sumedha, Roll: 2503201450146, MCA 2nd Year, ABES Engineering College; Mentor: Ms. Anjali Pandey), B2B SaaS business context, system architecture diagram, and GitHub repo link."
* **Verification Check:** Verify all 11 KPI definitions match Step 8 `KPIs_Measures` sheet verbatim.
* **Must-Pass Tests:** `TC-F-17`, `TC-F-18`, `TC-T-32`

---

### BUILD-13: Test Suite Execution & Coverage Verification
* **Goal:** Execute the full automated pytest suite across all test modules; verify zero failures and benchmark render performance.
* **Files Affected:** `tests/test_data.py`, `tests/test_analysis.py`, `tests/test_app.py`
* **Linked Stories & FRs:** NFR-01, NFR-02, NFR-05
* **Exact Prompt for AI:**
  > "Execute `pytest tests/ -v --cov=src`. Ensure all 35 technical tests pass with zero errors. Verify that Streamlit AppTest tests (TC-T-28..35) run cleanly against `app.py` and all 7 subpages. Assert that initial page render latency is strictly under 3.0 seconds."
* **Verification Check:** Terminal output reports `35 passed in X.XXs` with $\ge 70\%$ code coverage on `src/`.
* **Must-Pass Tests:** All 35 tests (`TC-T-01` through `TC-T-35`)

---

### BUILD-14: Documentation, README & Git Packaging
* **Goal:** Create a production-ready `README.md` with application screenshots, architecture diagram, setup instructions, and git commit history.
* **Files Affected:** `README.md`, `.gitignore`
* **Linked Stories & FRs:** Project Handover & Academic Evaluation
* **Exact Prompt for AI:**
  > "Author an enterprise `README.md`. Include:
  > 1. Badges (Python 3.12, Streamlit 1.32, Tests Passing, License MIT).
  > 2. Executive project summary & business value narrative.
  > 3. Architecture diagram and relational schema overview.
  > 4. Step-by-step local setup commands (`python -m venv venv`, `pip install -r requirements.txt`, `streamlit run app.py`).
  > 5. Pytest execution guide.
  > 6. Author attribution and ABES academic metadata."
* **Verification Check:** Preview README.md formatting; ensure all local markdown links resolve.
* **Must-Pass Tests:** Verification check of file existence.

---

### BUILD-15: Streamlit Community Cloud Deployment
* **Goal:** Deploy the application to Streamlit Community Cloud and verify live cloud URL.
* **Files Affected:** Repository Git root
* **Linked Stories & FRs:** Cloud Deployment Runbook (TRD Section 9)
* **Exact Prompt for AI:**
  > "Prepare repository for Streamlit Community Cloud deployment: ensure `requirements.txt` is in root, `data/app.db` is tracked in git (LFS not required as DB is < 1MB), verify relative paths use `os.path.join()`. Provide deployment checklist and live sanity test instructions."
* **Verification Check:** App loads successfully on `https://<user>-crma.streamlit.app` with HTTP 200.
* **Must-Pass Tests:** Smoke test on cloud URL.

---

## 4. Definition of Done (DoD)

The CRMA project is formally declared **DONE** when all of the following criteria are met:
1. **Requirements Coverage:** 100% of Must-Have Business Requirements (BR-01 through BR-08) and Functional Requirements (FR-01 through FR-12) are fully implemented and verified in the UI.
2. **Automated Test Suite:** All 35 automated technical tests in `tests/test_data.py`, `tests/test_analysis.py`, and `tests/test_app.py` execute with zero failures (`35 passed`).
3. **Manual Functional Tests:** All 20 functional test cases (TC-F-01 through TC-F-20) are executed and marked `Passed`.
4. **Data Integrity:** Embedded database `data/app.db` satisfies 100% of foreign key constraints, zero duplicate primary keys, and no domain range violations.
5. **UI & Styling:** All 8 pages render with consistent Navy/Steel branding, responsive Tailwind KPI cards via iframe bridge, and interactive Plotly visualizations.
6. **Error Handling:** Graceful error boundary handling across all pages — zero raw Python tracebacks displayed under empty or invalid filter selections (NFR-02).
7. **Performance:** App initial cold start loads in $< 5.0$ seconds; cached page switches and filter re-renders complete in $< 2.0$ seconds (NFR-01).
8. **Documentation:** Comprehensive `README.md` in repository root, complete project Word/Excel deliverables baselined in `docs/`, and live Streamlit Cloud deployment accessible.

---

## 5. Execution Commands Quick Reference

```bash
# 1. Environment Setup & Dependency Installation
python -m venv venv
venv\Scripts\activate       # On Windows
pip install --upgrade pip
pip install -r requirements.txt

# 2. Dataset Generation & Validation Check (if regenerating)
python src/generate_data.py

# 3. Launch Local Streamlit Application
streamlit run app.py

# 4. Execute Automated Pytest Test Suite
pytest tests/ -v

# 5. Run Specific Test Modules
pytest tests/test_data.py -v
pytest tests/test_analysis.py -v
pytest tests/test_app.py -v

# 6. Generate Test Coverage Report
pytest --cov=src tests/ --cov-report=term-missing --cov-report=html
```

---

## 6. Viva Preparation: 10 Critical Questions & Expert Answers

1. **Q: What is the business difference between Customer Churn Rate (CCR) and Net Revenue Retention (NRR)?**  
   *Answer:* CCR measures the percentage of customer *logos/accounts* lost over a period, treating a ₹2,500/mo Starter account identically to a ₹20,000/mo Enterprise account. NRR measures the percentage of *recurring revenue* retained from existing customers, accounting for account expansion (upgrades) and contraction (downgrades) alongside churn. An enterprise can have a 5% account churn rate but an NRR of 105%, indicating expansion revenue from surviving accounts more than offset lost logos.

2. **Q: Why did you use an embedded SQLite database rather than keeping data in CSV or pandas in memory?**  
   *Answer:* SQLite enforces strict relational schema integrity, primary key uniqueness, foreign key referential integrity, and domain `CHECK` constraints at the storage engine level. Furthermore, it allows SQL pre-aggregation via views (`vw_customer_retention_analytics`) with B-tree indexing on foreign keys, minimizing in-memory memory footprint on cloud containers while maintaining zero external network credential dependencies.

3. **Q: Explain the statistical justification for choosing the Mann-Whitney U test for support ticket analysis (OQ-06).**  
   *Answer:* Support ticket volume and resolution hours are heavily right-skewed count/duration metrics that violate the normality assumption required by Student's t-test. The Mann-Whitney U test is a non-parametric rank-sum test that evaluates whether the distribution of tickets for churned customers stochastically exceeds that of retained customers without assuming Gaussian normality.

4. **Q: How does the 4-factor composite customer health risk score work (RULE-05)?**  
   *Answer:* The model computes a weighted index $[0, 100]$: $35\%$ login activity decline (comparing trailing 30 days vs prior 60 days), $25\%$ open ticket count, $20\%$ low CSAT survey score (inverted), and $20\%$ contract renewal proximity ($\le 60$ days). Scores $\ge 70$ trigger High Risk alerts, allowing CS teams to intervene before contract expiration.

5. **Q: How does Streamlit's `@st.cache_data` decorator improve performance, and how do you prevent stale data?**  
   *Answer:* `@st.cache_data` serializes the output of expensive data loading and aggregation functions into an in-memory cache keyed on function input parameters. On user filter interactions, Streamlit re-executes the script from top to bottom; caching prevents re-reading SQLite disk files, achieving sub-second render speeds. Stale data is prevented by setting `ttl=3600` (1 hour) and using immutable read-only DB connections.

6. **Q: Why did you inject intentional data quality issues (DQ-01, DQ-02, DQ-03) during data generation?**  
   *Answer:* Real-world SaaS data is never pristine. Injecting 65.8% missing survey scores (DQ-01) replicates realistic 20-35% survey response rates; injecting 3 duplicate tickets (DQ-02) tests database deduplication logic; and injecting 5 resolution turnaround outliers $> 200$ hours (DQ-03) ensures the analytics layer handles skewness using robust medians and non-parametric statistics.

7. **Q: Why is Tailwind CSS loaded via `st.components.v1.html()` instead of standard Streamlit markdown?**  
   *Answer:* Standard Streamlit KPI metric cards have rigid layout and styling capabilities. Loading Tailwind CDN inside an isolated iframe allows us to build responsive, micro-styled enterprise cards featuring subtle box shadows, custom hex gradients, rounded pills, and dynamic semantic delta badges matching professional design standards.

8. **Q: How does your application prevent SQL injection vulnerabilities?**  
   *Answer:* SQLite database connections are opened in strict read-only mode using the URI parameter `file:data/app.db?mode=ro`. In addition, all data querying is restricted to pre-defined view selections or parameterized queries in `src/data_loader.py`; no dynamic string concatenation of user inputs is ever executed against the SQL engine.

9. **Q: How do you interpret a triangular cohort survival heatmap?**  
   *Answer:* Rows represent customer signup cohorts (grouped by YYYY-MM); columns represent lifecycle tenure milestones ($M+1, M+3, M+6, M+12$). The triangular shape naturally occurs because newer cohorts have not yet reached mature lifecycle milestones. Reading across a row reveals the retention decay curve of a specific cohort; reading down a column compares whether newer cohorts retain better at the same tenure mark than historical cohorts.

10. **Q: What is the significance of the Gherkin Given/When/Then syntax in your test cases?**  
    *Answer:* Gherkin syntax provides an unambiguous contract between business requirements and technical verification. 'Given' specifies the precondition and system state; 'When' isolates the specific user or system action; and 'Then' asserts the observable outcome. This structure translates directly into automated test assertions in pytest and manual QA validation steps.
