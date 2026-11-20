import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_step8_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default

    # Palette
    NAVY_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    STEEL_FILL = PatternFill(start_color="2B4C7E", end_color="2B4C7E", fill_type="solid")
    LBL_FILL = PatternFill(start_color="EAEEF3", end_color="EAEEF3", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F4F7FB", end_color="F4F7FB", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    PK_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    FK_FILL = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")

    WHITE_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    BOLD_FONT = Font(name="Calibri", size=10, bold=True, color="1B365D")
    REG_FONT = Font(name="Calibri", size=9.5, color="222222")
    CODE_FONT = Font(name="Consolas", size=9, color="1B365D")
    TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1B365D")
    SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="555555")
    SEC_HDR_FONT = Font(name="Calibri", size=11, bold=True, color="1B365D")

    THIN_BORDER = Border(
        left=Side(style='thin', color='D0D7DE'),
        right=Side(style='thin', color='D0D7DE'),
        top=Side(style='thin', color='D0D7DE'),
        bottom=Side(style='thin', color='D0D7DE')
    )

    def style_header_row(ws, row_idx, headers, fill=NAVY_FILL):
        ws.row_dimensions[row_idx].height = 26
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx, value=h)
            cell.font = WHITE_FONT
            cell.fill = fill
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def style_data_table(ws, start_row, data, align_center_cols=[], code_cols=[]):
        for r_idx, row_vals in enumerate(data):
            curr_row = start_row + r_idx
            ws.row_dimensions[curr_row].height = 22
            bg = ZEBRA_FILL if r_idx % 2 == 1 else WHITE_FILL
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=curr_row, column=c_idx, value=val)
                cell.font = REG_FONT
                cell.border = THIN_BORDER
                cell.fill = bg
                if c_idx in align_center_cols:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                if c_idx in code_cols:
                    cell.font = CODE_FONT

                # Highlight PK / FK tags
                if str(val) == "PK":
                    cell.fill = PK_FILL
                    cell.font = Font(name="Calibri", size=9.5, bold=True, color="B25900")
                elif str(val) == "FK":
                    cell.fill = FK_FILL
                    cell.font = Font(name="Calibri", size=9.5, bold=True, color="276A3C")

    # =========================================================================
    # SHEET 1: Overview
    # =========================================================================
    ws1 = wb.create_sheet(title="Overview")
    ws1.views.sheetView[0].showGridLines = True
    ws1.merge_cells("A1:B1")
    ws1["A1"] = "Database Design Specification — Customer Retention Metrics Analysis"
    ws1["A1"].font = TITLE_FONT
    ws1.row_dimensions[1].height = 24

    ws1.merge_cells("A2:B2")
    ws1["A2"] = "Relational Star/Snowflake Schema Architecture | Target: SQLite3 (data/app.db) | MCA Capstone, ABES"
    ws1["A2"].font = SUBTITLE_FONT
    ws1.row_dimensions[2].height = 18

    overview_items = [
        ("Project Name", "Customer Retention Metrics Analysis (CRMA)"),
        ("Database Engine", "SQLite 3.42+ (Embedded file-based relational storage at data/app.db)"),
        ("Schema Architecture", "Normalized 3NF Relational Star-Schema: 2 Master Dimensions + 3 Core Fact Tables"),
        ("Primary Goal", "Serve high-performance analytical queries for 10 Operational Questions (OQ-01 to OQ-10) and Streamlit multipage views without data redundancy."),
        ("Entity Count", "5 Relational Tables: dim_customers, dim_plans, fact_subscriptions, fact_support_tickets, fact_monthly_usage"),
        ("Analytical View", "vw_customer_retention_analytics (Joined comprehensive analytical view with 36 curated columns)"),
        ("Sizing Baseline", "Design Plan / Seed Sample: 100+ master rows, 250+ fact rows. Full Synthetic Scale (Step 10): 5,000 customers, 24 historical months (~70,000 telemetry rows)."),
        ("Constraint Rules", "PRAGMA foreign_keys = ON; strict NOT NULL on identifiers/dates; CHECK constraints on price, ratings, and statuses; B-Tree indexes on all FKs."),
        ("Referential Integrity", "ON DELETE RESTRICT and ON UPDATE CASCADE to guarantee zero orphan records."),
        ("Governance & Audit", "Every attribute mapped directly to supported OQs and Business Requirements (BR-01 to BR-13).")
    ]

    for idx, (k, v) in enumerate(overview_items, start=4):
        c1 = ws1.cell(row=idx, column=1, value=k)
        c2 = ws1.cell(row=idx, column=2, value=v)
        c1.font = BOLD_FONT
        c1.fill = LBL_FILL
        c1.border = THIN_BORDER
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.font = REG_FONT
        c2.border = THIN_BORDER
        c2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws1.row_dimensions[idx].height = 24

    ws1.column_dimensions["A"].width = 24
    ws1.column_dimensions["B"].width = 90

    # =========================================================================
    # SHEET 2: Tables_Summary
    # =========================================================================
    ws2 = wb.create_sheet(title="Tables_Summary")
    ws2.views.sheetView[0].showGridLines = True
    tbl_summary_headers = ["Table Name", "Type", "Table Grain", "Primary Key (PK)", "Foreign Keys (FK)", "Columns", "Design Rows", "Description"]
    style_header_row(ws2, 1, tbl_summary_headers)
    ws2.auto_filter.ref = "A1:H6"
    ws2.freeze_panes = "A2"

    tbl_summary_data = [
        ("dim_customers", "Master", "One row per customer account", "customer_id", "None", 8, "100 - 5,000", "Customer profile firmographics: industry, size, region, channel, signup date."),
        ("dim_plans", "Master", "One row per subscription tier", "plan_id", "None", 6, "3 - 5", "Catalog of product tiers: Starter, Professional, Enterprise, base price, user limits."),
        ("fact_subscriptions", "Fact", "One row per subscription contract term", "subscription_id", "customer_id, plan_id", 11, "250 - 7,500", "Core subscription contracts, active/churn status, renewal dates, MRR fees."),
        ("fact_support_tickets", "Fact", "One row per helpdesk support ticket", "ticket_id", "customer_id", 8, "200 - 15,000", "Customer service tickets, category, SLA resolution turnaround, CSAT score."),
        ("fact_monthly_usage", "Fact", "One row per customer per calendar month", "usage_id", "customer_id", 8, "300 - 70,000", "Monthly application telemetry: logins, feature clicks, storage, NPS score.")
    ]
    style_data_table(ws2, 2, tbl_summary_data, align_center_cols=[2, 4, 6, 7], code_cols=[1, 4, 5])
    
    col_w2 = {"A": 22, "B": 12, "C": 35, "D": 18, "E": 24, "F": 12, "G": 16, "H": 45}
    for c_l, w in col_w2.items(): ws2.column_dimensions[c_l].width = w

    # =========================================================================
    # DATA DICTIONARIES (Sheets 3 to 7)
    # =========================================================================
    dict_headers = ["Column Name", "Data Type", "Key", "Nullable", "Allowed Values / Format", "Description & Business Semantics", "Supported OQ / BR"]

    # Table 1: dim_customers
    ws_dc = wb.create_sheet(title="dim_customers")
    ws_dc.views.sheetView[0].showGridLines = True
    style_header_row(ws_dc, 1, dict_headers)
    ws_dc.auto_filter.ref = "A1:G9"
    ws_dc.freeze_panes = "A2"
    dc_data = [
        ("customer_id", "TEXT", "PK", "N", "CUST-XXXXX (e.g. CUST-00001)", "Unique alphanumeric identifier for business account.", "All OQs, BR-01..13"),
        ("company_name", "TEXT", "None", "N", "Fictitious SMB corporate name", "Registered corporate entity name.", "OQ-10, BR-05"),
        ("industry", "TEXT", "None", "N", "EdTech, FinTech, Healthcare, Retail, Manufacturing", "Industry vertical classification for segmentation.", "OQ-02, BR-03"),
        ("company_size", "TEXT", "None", "N", "Micro (1-10), Small (11-50), Medium (51-200)", "Employee headcount band for firmographic analysis.", "OQ-02, BR-03"),
        ("region", "TEXT", "None", "N", "North, South, East, West", "Geographic operating territory within India.", "OQ-02, BR-03"),
        ("acquisition_channel", "TEXT", "None", "N", "Organic Search, Referral, Paid Social, Direct, Partner", "Marketing acquisition source attribution.", "OQ-05, BR-10"),
        ("signup_date", "DATE", "None", "N", "YYYY-MM-DD (2024-01-01 to 2025-12-31)", "Original account creation date at NoidaSoft.", "OQ-01, OQ-03, BR-04"),
        ("signup_cohort_month", "TEXT", "None", "N", "YYYY-MM (e.g. 2024-03)", "Calendar month of initial signup for cohort analysis.", "OQ-03, BR-04")
    ]
    style_data_table(ws_dc, 2, dc_data, align_center_cols=[2, 3, 4], code_cols=[1])

    # Table 2: dim_plans
    ws_dp = wb.create_sheet(title="dim_plans")
    ws_dp.views.sheetView[0].showGridLines = True
    style_header_row(ws_dp, 1, dict_headers)
    ws_dp.auto_filter.ref = "A1:G7"
    ws_dp.freeze_panes = "A2"
    dp_data = [
        ("plan_id", "TEXT", "PK", "N", "PLAN-01, PLAN-02, PLAN-03", "Unique plan tier identifier.", "OQ-02, OQ-04, BR-03"),
        ("plan_name", "TEXT", "None", "N", "Starter, Professional, Enterprise", "Commercial product tier name.", "OQ-02, OQ-04, BR-03"),
        ("plan_tier_level", "INTEGER", "None", "N", "1, 2, 3 (Ordinal tier ranking)", "Hierarchical level used to compute upgrades/downgrades.", "OQ-09, BR-07"),
        ("monthly_base_price", "REAL", "None", "N", ">= 0.00 (INR 2,499.00, 7,999.00, 19,999.00)", "Official catalog list price per month in Indian Rupees.", "OQ-04, OQ-09, BR-09"),
        ("max_user_seats", "INTEGER", "None", "N", "5, 25, 100", "Maximum licensed user seats permitted under plan tier.", "OQ-02, BR-03"),
        ("sla_support_level", "TEXT", "None", "N", "Standard (48h), Priority (24h), Dedicated (4h)", "Contractual support ticket turnaround SLA guarantee.", "OQ-06, BR-06")
    ]
    style_data_table(ws_dp, 2, dp_data, align_center_cols=[2, 3, 4], code_cols=[1])

    # Table 3: fact_subscriptions
    ws_fs = wb.create_sheet(title="fact_subscriptions")
    ws_fs.views.sheetView[0].showGridLines = True
    style_header_row(ws_fs, 1, dict_headers)
    ws_fs.auto_filter.ref = "A1:G12"
    ws_fs.freeze_panes = "A2"
    fs_data = [
        ("subscription_id", "TEXT", "PK", "N", "SUB-XXXXXX (e.g. SUB-000001)", "Unique contract subscription record identifier.", "All OQs, BR-01..13"),
        ("customer_id", "TEXT", "FK", "N", "References dim_customers(customer_id)", "Foreign key pointing to subscriber account.", "All OQs, BR-01..13"),
        ("plan_id", "TEXT", "FK", "N", "References dim_plans(plan_id)", "Foreign key pointing to subscription product tier.", "OQ-02, OQ-04, BR-03"),
        ("start_date", "DATE", "None", "N", "YYYY-MM-DD", "Contract activation start date.", "OQ-01, OQ-03, BR-01"),
        ("end_date", "DATE", "None", "N", "YYYY-MM-DD (start_date + duration)", "Contract expiration date (renewal decision point).", "OQ-01, OQ-10, BR-05"),
        ("billing_cadence", "TEXT", "None", "N", "Monthly, Annual", "Payment billing cycle term.", "OQ-09, BR-07"),
        ("monthly_fee_inr", "REAL", "None", "N", ">= 0.00 (Actual contracted MRR)", "Contracted monthly recurring revenue in INR.", "OQ-04, OQ-09, BR-07"),
        ("subscription_status", "TEXT", "None", "N", "Active, Churned, Renewed", "Current operational state of contract per RULE-01.", "OQ-01, OQ-02, BR-01"),
        ("is_churned", "INTEGER", "None", "N", "0 (Active/Renewed), 1 (Churned)", "Binary flag indicating customer attrition.", "All OQs, BR-01..12"),
        ("churn_date", "DATE", "None", "Y", "YYYY-MM-DD (Null for active accounts)", "Official date cancellation took effect.", "OQ-01, OQ-02, BR-02"),
        ("tenure_months", "INTEGER", "None", "N", ">= 1 (Calculated elapsed duration)", "Total continuous active lifespan in months.", "OQ-04, BR-04, BR-09")
    ]
    style_data_table(ws_fs, 2, fs_data, align_center_cols=[2, 3, 4], code_cols=[1])

    # Table 4: fact_support_tickets
    ws_ft = wb.create_sheet(title="fact_support_tickets")
    ws_ft.views.sheetView[0].showGridLines = True
    style_header_row(ws_ft, 1, dict_headers)
    ws_ft.auto_filter.ref = "A1:G9"
    ws_ft.freeze_panes = "A2"
    ft_data = [
        ("ticket_id", "TEXT", "PK", "N", "TCK-XXXXXX (e.g. TCK-000001)", "Unique helpdesk ticket identifier.", "OQ-06, BR-06"),
        ("customer_id", "TEXT", "FK", "N", "References dim_customers(customer_id)", "Foreign key linking ticket to customer account.", "OQ-06, BR-06"),
        ("created_at", "DATETIME", "None", "N", "YYYY-MM-DD HH:MM:SS", "Timestamp when support request was logged.", "OQ-06, BR-06"),
        ("resolved_at", "DATETIME", "None", "Y", "YYYY-MM-DD HH:MM:SS (Null if open)", "Timestamp when ticket was officially closed.", "OQ-06, BR-06"),
        ("resolution_hours", "REAL", "None", "Y", ">= 0.0 (Elapsed resolution turnaround)", "Total hours elapsed from creation to resolution.", "OQ-06, BR-06"),
        ("priority", "TEXT", "None", "N", "Low, Medium, High, Critical", "Assigned support ticket urgency level.", "OQ-06, OQ-10, BR-05"),
        ("category", "TEXT", "None", "N", "Billing, Bug, Feature Request, Technical", "Functional area of helpdesk inquiry.", "OQ-06, BR-06"),
        ("ticket_csat", "INTEGER", "None", "Y", "1 to 5 (Null if survey uncompleted)", "Transaction-level post-resolution satisfaction.", "OQ-06, OQ-08, BR-08")
    ]
    style_data_table(ws_ft, 2, ft_data, align_center_cols=[2, 3, 4], code_cols=[1])

    # Table 5: fact_monthly_usage
    ws_fu = wb.create_sheet(title="fact_monthly_usage")
    ws_fu.views.sheetView[0].showGridLines = True
    style_header_row(ws_fu, 1, dict_headers)
    ws_fu.auto_filter.ref = "A1:G9"
    ws_fu.freeze_panes = "A2"
    fu_data = [
        ("usage_id", "TEXT", "PK", "N", "USG-XXXXXX (e.g. USG-000001)", "Unique monthly usage summary record identifier.", "OQ-07, BR-11"),
        ("customer_id", "TEXT", "FK", "N", "References dim_customers(customer_id)", "Foreign key linking telemetry to customer account.", "OQ-07, BR-11"),
        ("activity_month", "TEXT", "None", "N", "YYYY-MM (e.g. 2024-06)", "Calendar month for aggregated usage metrics.", "OQ-07, BR-11"),
        ("login_days", "INTEGER", "None", "N", "0 to 31 (Active days logged in)", "Number of distinct calendar days with >= 1 login.", "OQ-07, OQ-10, BR-05"),
        ("core_feature_clicks", "INTEGER", "None", "N", ">= 0 (Count of critical tool events)", "Total engagement actions on core SaaS features.", "OQ-07, BR-11"),
        ("active_users_count", "INTEGER", "None", "N", ">= 0 (Count of distinct team users)", "Distinct user accounts active during the month.", "OQ-07, BR-03"),
        ("csat_score", "INTEGER", "None", "Y", "1 to 5 (Quarterly feedback rating)", "Customer satisfaction rating (1=Poor, 5=Excellent).", "OQ-08, OQ-10, BR-08"),
        ("nps_score", "INTEGER", "None", "Y", "0 to 10 (Net Promoter Score rating)", "Willingness to recommend rating (0 to 10).", "OQ-08, BR-12")
    ]
    style_data_table(ws_fu, 2, fu_data, align_center_cols=[2, 3, 4], code_cols=[1])

    for ws_d in [ws_dc, ws_dp, ws_fs, ws_ft, ws_fu]:
        col_w_dict = {"A": 22, "B": 14, "C": 10, "D": 10, "E": 30, "F": 45, "G": 22}
        for c_l, w in col_w_dict.items(): ws_d.column_dimensions[c_l].width = w

    # =========================================================================
    # SHEET 8: Relationships
    # =========================================================================
    ws_rel = wb.create_sheet(title="Relationships")
    ws_rel.views.sheetView[0].showGridLines = True
    rel_headers = ["Relationship ID", "Parent Table", "Child Table", "Parent PK", "Child FK", "Cardinality", "On Delete / Update", "Relationship Description & Business Rule"]
    style_header_row(ws_rel, 1, rel_headers)
    ws_rel.auto_filter.ref = "A1:H5"
    ws_rel.freeze_panes = "A2"

    rel_data = [
        ("REL-01", "dim_customers", "fact_subscriptions", "customer_id", "customer_id", "1 : N", "RESTRICT / CASCADE", "One customer account can hold multiple historical subscription terms over time."),
        ("REL-02", "dim_plans", "fact_subscriptions", "plan_id", "plan_id", "1 : N", "RESTRICT / CASCADE", "One catalog plan tier is subscribed to by many customer contract instances."),
        ("REL-03", "dim_customers", "fact_support_tickets", "customer_id", "customer_id", "1 : N", "CASCADE / CASCADE", "One customer logs zero, one, or many support helpdesk tickets."),
        ("REL-04", "dim_customers", "fact_monthly_usage", "customer_id", "customer_id", "1 : N", "CASCADE / CASCADE", "One customer generates exactly one usage and sentiment summary record per calendar month.")
    ]
    style_data_table(ws_rel, 2, rel_data, align_center_cols=[1, 4, 5, 6, 7], code_cols=[2, 3, 4, 5])
    col_w_rel = {"A": 16, "B": 20, "C": 22, "D": 16, "E": 16, "F": 14, "G": 20, "H": 50}
    for c_l, w in col_w_rel.items(): ws_rel.column_dimensions[c_l].width = w

    # =========================================================================
    # SHEET 9: KPIs_Measures
    # =========================================================================
    ws_kpi = wb.create_sheet(title="KPIs_Measures")
    ws_kpi.views.sheetView[0].showGridLines = True
    kpi_headers = ["Measure ID", "Business Metric Name", "Level / Grain", "SQL / pandas Computational Formula", "Filter & Business Conditions", "Supported OQ", "Target / Benchmark"]
    style_header_row(ws_kpi, 1, kpi_headers)
    ws_kpi.auto_filter.ref = "A1:G11"
    ws_kpi.freeze_panes = "A2"

    kpi_data = [
        ("MSR-01", "Customer Churn Rate (CCR %)", "Monthly", "(COUNT(churned_customers) / COUNT(starting_active_customers)) * 100", "is_churned = 1 within calendar month M; excludes newly added accounts in M", "OQ-01", "< 2.5% per month"),
        ("MSR-02", "Customer Retention Rate (CRR %)", "Monthly", "100.0 - CCR_pct (for closed cohorts)", "Closed cohort calculation matching RULE-02", "OQ-01", "> 97.5% per month"),
        ("MSR-03", "Cohort Survival Rate (%)", "Cohort / Lifecycle Month", "(COUNT(active_in_month_k) / COUNT(original_cohort_size)) * 100", "signup_cohort_month = C, evaluated at k in [1, 3, 6, 12]", "OQ-03", "> 85% at M+12"),
        ("MSR-04", "Customer Lifetime Value (CLV)", "Segment / Plan", "AVG(monthly_fee_inr) * AVG(tenure_months)", "Calculated per plan_type or industry segment", "OQ-04", "> INR 75,000 avg"),
        ("MSR-05", "Channel Retention Rate (%)", "Acquisition Channel", "(COUNT(retained_by_channel) / COUNT(total_by_channel)) * 100", "Evaluated at 12-month tenure boundary across acquisition_channel", "OQ-05", "Referral > 90%"),
        ("MSR-06", "Mean Ticket Resolution Hours", "Overall / Churn status", "AVG(resolution_hours) grouped by is_churned", "Excludes tickets with status = 'Open' or resolved_at IS NULL", "OQ-06", "< 24.0 hours"),
        ("MSR-07", "Pre-Churn Login Activity Drop", "Customer / 90d window", "((AVG(logins_prev_60d) - AVG(logins_last_30d)) / AVG(logins_prev_60d)) * 100", "Evaluated for churned customers in the 90 days preceding churn_date", "OQ-07", "Statistically lower (p<0.01)"),
        ("MSR-08", "Satisfaction Correlation (CSAT/NPS)", "Customer Level", "CORR(latest_csat_score, renewed_flag)", "Pearson r / Spearman rho computed across accounts with valid survey scores", "OQ-08", "r > 0.45 (positive)"),
        ("MSR-09", "Monthly Recurring Revenue Churn (INR)", "Monthly", "SUM(monthly_fee_inr) for all contracts where is_churned = 1 in month M", "Captures lost MRR in rupees rather than raw account count", "OQ-09", "< INR 50,000 / mo"),
        ("MSR-10", "Net Revenue Retention (NRR %)", "Monthly / Annual", "((Starting_MRR + Expansion_MRR - Contraction_MRR - Churned_MRR) / Starting_MRR) * 100", "Evaluated over trailing 12-month window per RULE-03", "OQ-09", "> 105% (Net expansion)"),
        ("MSR-11", "Composite Customer Risk Score", "Active Account", "(0.35 * login_drop_score) + (0.25 * open_tickets_score) + (0.20 * low_csat_score) + (0.20 * renewal_score)", "Evaluated dynamically for active accounts with end_date <= today + 60 days", "OQ-10", "High Risk >= 70")
    ]
    style_data_table(ws_kpi, 2, kpi_data, align_center_cols=[1, 3, 6], code_cols=[4])
    col_w_kpi = {"A": 14, "B": 28, "C": 18, "D": 45, "E": 45, "F": 14, "G": 20}
    for c_l, w in col_w_kpi.items(): ws_kpi.column_dimensions[c_l].width = w

    # =========================================================================
    # SHEET 10: SQLite_DDL
    # =========================================================================
    ws_ddl = wb.create_sheet(title="SQLite_DDL")
    ws_ddl.views.sheetView[0].showGridLines = True
    ws_ddl.merge_cells("A1:B1")
    ws_ddl["A1"] = "SQLite DDL Implementation Script (Schema & Views)"
    ws_ddl["A1"].font = TITLE_FONT
    ws_ddl.row_dimensions[1].height = 24

    ws_ddl.merge_cells("A2:B2")
    ws_ddl["A2"] = "Copy and execute directly in SQLite3 / Python sqlite3 module | Includes full PK, FK, CHECK, and View statements."
    ws_ddl["A2"].font = SUBTITLE_FONT
    ws_ddl.row_dimensions[2].height = 18

    ddl_statements = [
        ("-- 1. Enable Foreign Key Enforcement", "PRAGMA foreign_keys = ON;\nPRAGMA journal_mode = WAL;"),
        
        ("-- 2. Master Table: dim_customers",
         "CREATE TABLE IF NOT EXISTS dim_customers (\n"
         "    customer_id         TEXT PRIMARY KEY,\n"
         "    company_name        TEXT NOT NULL,\n"
         "    industry            TEXT NOT NULL CHECK(industry IN ('EdTech', 'FinTech', 'Healthcare', 'Retail', 'Manufacturing')),\n"
         "    company_size        TEXT NOT NULL CHECK(company_size IN ('Micro (1-10)', 'Small (11-50)', 'Medium (51-200)')),\n"
         "    region              TEXT NOT NULL CHECK(region IN ('North', 'South', 'East', 'West')),\n"
         "    acquisition_channel TEXT NOT NULL CHECK(acquisition_channel IN ('Organic Search', 'Referral', 'Paid Social', 'Direct', 'Partner')),\n"
         "    signup_date         DATE NOT NULL,\n"
         "    signup_cohort_month TEXT NOT NULL\n"
         ");"),
        
        ("-- 3. Master Table: dim_plans",
         "CREATE TABLE IF NOT EXISTS dim_plans (\n"
         "    plan_id             TEXT PRIMARY KEY,\n"
         "    plan_name           TEXT NOT NULL UNIQUE CHECK(plan_name IN ('Starter', 'Professional', 'Enterprise')),\n"
         "    plan_tier_level     INTEGER NOT NULL CHECK(plan_tier_level IN (1, 2, 3)),\n"
         "    monthly_base_price  REAL NOT NULL CHECK(monthly_base_price >= 0.0),\n"
         "    max_user_seats      INTEGER NOT NULL CHECK(max_user_seats > 0),\n"
         "    sla_support_level   TEXT NOT NULL\n"
         ");"),
        
        ("-- 4. Fact Table: fact_subscriptions",
         "CREATE TABLE IF NOT EXISTS fact_subscriptions (\n"
         "    subscription_id     TEXT PRIMARY KEY,\n"
         "    customer_id         TEXT NOT NULL,\n"
         "    plan_id             TEXT NOT NULL,\n"
         "    start_date          DATE NOT NULL,\n"
         "    end_date            DATE NOT NULL,\n"
         "    billing_cadence     TEXT NOT NULL CHECK(billing_cadence IN ('Monthly', 'Annual')),\n"
         "    monthly_fee_inr     REAL NOT NULL CHECK(monthly_fee_inr >= 0.0),\n"
         "    subscription_status TEXT NOT NULL CHECK(subscription_status IN ('Active', 'Churned', 'Renewed')),\n"
         "    is_churned          INTEGER NOT NULL CHECK(is_churned IN (0, 1)),\n"
         "    churn_date          DATE,\n"
         "    tenure_months       INTEGER NOT NULL CHECK(tenure_months >= 1),\n"
         "    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON DELETE RESTRICT ON UPDATE CASCADE,\n"
         "    FOREIGN KEY (plan_id)     REFERENCES dim_plans(plan_id)         ON DELETE RESTRICT ON UPDATE CASCADE\n"
         ");\n"
         "CREATE INDEX IF NOT EXISTS idx_sub_cust ON fact_subscriptions(customer_id);\n"
         "CREATE INDEX IF NOT EXISTS idx_sub_plan ON fact_subscriptions(plan_id);"),

        ("-- 5. Fact Table: fact_support_tickets",
         "CREATE TABLE IF NOT EXISTS fact_support_tickets (\n"
         "    ticket_id           TEXT PRIMARY KEY,\n"
         "    customer_id         TEXT NOT NULL,\n"
         "    created_at          DATETIME NOT NULL,\n"
         "    resolved_at         DATETIME,\n"
         "    resolution_hours    REAL CHECK(resolution_hours >= 0.0),\n"
         "    priority            TEXT NOT NULL CHECK(priority IN ('Low', 'Medium', 'High', 'Critical')),\n"
         "    category            TEXT NOT NULL CHECK(category IN ('Billing', 'Bug', 'Feature Request', 'Technical')),\n"
         "    ticket_csat         INTEGER CHECK(ticket_csat BETWEEN 1 AND 5),\n"
         "    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON DELETE CASCADE ON UPDATE CASCADE\n"
         ");\n"
         "CREATE INDEX IF NOT EXISTS idx_ticket_cust ON fact_support_tickets(customer_id);"),

        ("-- 6. Fact Table: fact_monthly_usage",
         "CREATE TABLE IF NOT EXISTS fact_monthly_usage (\n"
         "    usage_id            TEXT PRIMARY KEY,\n"
         "    customer_id         TEXT NOT NULL,\n"
         "    activity_month      TEXT NOT NULL,\n"
         "    login_days          INTEGER NOT NULL CHECK(login_days BETWEEN 0 AND 31),\n"
         "    core_feature_clicks INTEGER NOT NULL CHECK(core_feature_clicks >= 0),\n"
         "    active_users_count  INTEGER NOT NULL CHECK(active_users_count >= 0),\n"
         "    csat_score          INTEGER CHECK(csat_score BETWEEN 1 AND 5),\n"
         "    nps_score           INTEGER CHECK(nps_score BETWEEN 0 AND 10),\n"
         "    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON DELETE CASCADE ON UPDATE CASCADE,\n"
         "    UNIQUE(customer_id, activity_month)\n"
         ");\n"
         "CREATE INDEX IF NOT EXISTS idx_usage_cust_month ON fact_monthly_usage(customer_id, activity_month);"),

        ("-- 7. Joined Analysis View: vw_customer_retention_analytics (36 Curated Columns)",
         "CREATE VIEW IF NOT EXISTS vw_customer_retention_analytics AS\n"
         "SELECT \n"
         "    c.customer_id,\n"
         "    c.company_name,\n"
         "    c.industry,\n"
         "    c.company_size,\n"
         "    c.region,\n"
         "    c.acquisition_channel,\n"
         "    c.signup_date,\n"
         "    c.signup_cohort_month,\n"
         "    p.plan_id,\n"
         "    p.plan_name,\n"
         "    p.plan_tier_level,\n"
         "    p.monthly_base_price,\n"
         "    s.subscription_id,\n"
         "    s.start_date AS sub_start_date,\n"
         "    s.end_date AS sub_end_date,\n"
         "    s.billing_cadence,\n"
         "    s.monthly_fee_inr,\n"
         "    s.subscription_status,\n"
         "    s.is_churned,\n"
         "    s.churn_date,\n"
         "    s.tenure_months,\n"
         "    (s.monthly_fee_inr * s.tenure_months) AS lifetime_revenue_inr,\n"
         "    CASE WHEN s.is_churned = 1 THEN s.monthly_fee_inr ELSE 0.0 END AS churned_mrr_inr,\n"
         "    COALESCE(t.total_tickets, 0) AS total_support_tickets,\n"
         "    COALESCE(t.avg_res_hours, 0.0) AS avg_resolution_hours,\n"
         "    COALESCE(t.critical_tickets, 0) AS critical_tickets_count,\n"
         "    COALESCE(u.avg_login_days, 0) AS avg_monthly_login_days,\n"
         "    COALESCE(u.total_feature_clicks, 0) AS total_feature_clicks,\n"
         "    u.latest_csat_score,\n"
         "    u.latest_nps_score,\n"
         "    CASE \n"
         "        WHEN u.latest_nps_score >= 9 THEN 'Promoter'\n"
         "        WHEN u.latest_nps_score BETWEEN 7 AND 8 THEN 'Passive'\n"
         "        WHEN u.latest_nps_score IS NOT NULL THEN 'Detractor'\n"
         "        ELSE 'No Survey'\n"
         "    END AS nps_category,\n"
         "    -- Multi-Factor Rule-Based Risk Score Calculation (0-100)\n"
         "    ROUND(\n"
         "        (CASE WHEN COALESCE(u.avg_login_days, 0) < 5 THEN 35.0 WHEN COALESCE(u.avg_login_days, 0) < 12 THEN 18.0 ELSE 0.0 END) +\n"
         "        (CASE WHEN COALESCE(t.critical_tickets, 0) >= 2 THEN 25.0 WHEN COALESCE(t.critical_tickets, 0) = 1 THEN 12.0 ELSE 0.0 END) +\n"
         "        (CASE WHEN u.latest_csat_score <= 2 THEN 20.0 WHEN u.latest_csat_score = 3 THEN 10.0 ELSE 0.0 END) +\n"
         "        (CASE WHEN s.subscription_status = 'Active' AND julianday(s.end_date) - julianday('now') BETWEEN 0 AND 30 THEN 20.0 ELSE 0.0 END),\n"
         "        1\n"
         "    ) AS health_risk_score,\n"
         "    CASE \n"
         "        WHEN ROUND(\n"
         "            (CASE WHEN COALESCE(u.avg_login_days, 0) < 5 THEN 35.0 WHEN COALESCE(u.avg_login_days, 0) < 12 THEN 18.0 ELSE 0.0 END) +\n"
         "            (CASE WHEN COALESCE(t.critical_tickets, 0) >= 2 THEN 25.0 WHEN COALESCE(t.critical_tickets, 0) = 1 THEN 12.0 ELSE 0.0 END) +\n"
         "            (CASE WHEN u.latest_csat_score <= 2 THEN 20.0 WHEN u.latest_csat_score = 3 THEN 10.0 ELSE 0.0 END) +\n"
         "            (CASE WHEN s.subscription_status = 'Active' AND julianday(s.end_date) - julianday('now') BETWEEN 0 AND 30 THEN 20.0 ELSE 0.0 END),\n"
         "            1\n"
         "        ) >= 70.0 THEN 'High Risk'\n"
         "        WHEN ROUND(\n"
         "            (CASE WHEN COALESCE(u.avg_login_days, 0) < 5 THEN 35.0 WHEN COALESCE(u.avg_login_days, 0) < 12 THEN 18.0 ELSE 0.0 END) +\n"
         "            (CASE WHEN COALESCE(t.critical_tickets, 0) >= 2 THEN 25.0 WHEN COALESCE(t.critical_tickets, 0) = 1 THEN 12.0 ELSE 0.0 END) +\n"
         "            (CASE WHEN u.latest_csat_score <= 2 THEN 20.0 WHEN u.latest_csat_score = 3 THEN 10.0 ELSE 0.0 END) +\n"
         "            (CASE WHEN s.subscription_status = 'Active' AND julianday(s.end_date) - julianday('now') BETWEEN 0 AND 30 THEN 20.0 ELSE 0.0 END),\n"
         "            1\n"
         "        ) >= 40.0 THEN 'Medium Risk'\n"
         "        ELSE 'Low Risk'\n"
         "    END AS risk_tier\n"
         "FROM dim_customers c\n"
         "JOIN fact_subscriptions s ON c.customer_id = s.customer_id\n"
         "JOIN dim_plans p ON s.plan_id = p.plan_id\n"
         "LEFT JOIN (\n"
         "    SELECT customer_id, COUNT(*) AS total_tickets, AVG(resolution_hours) AS avg_res_hours, SUM(CASE WHEN priority IN ('High', 'Critical') THEN 1 ELSE 0 END) AS critical_tickets\n"
         "    FROM fact_support_tickets GROUP BY customer_id\n"
         ") t ON c.customer_id = t.customer_id\n"
         "LEFT JOIN (\n"
         "    SELECT customer_id, AVG(login_days) AS avg_login_days, SUM(core_feature_clicks) AS total_feature_clicks, MAX(csat_score) AS latest_csat_score, MAX(nps_score) AS latest_nps_score\n"
         "    FROM fact_monthly_usage GROUP BY customer_id\n"
         ") u ON c.customer_id = u.customer_id;")
    ]

    curr_r = 4
    for section_title, sql_text in ddl_statements:
        c1 = ws_ddl.cell(row=curr_r, column=1, value=section_title)
        c1.font = BOLD_FONT
        c1.fill = LBL_FILL
        c1.border = THIN_BORDER
        ws_ddl.row_dimensions[curr_r].height = 22

        c2 = ws_ddl.cell(row=curr_r + 1, column=1, value=sql_text)
        c2.font = CODE_FONT
        c2.border = THIN_BORDER
        c2.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        lines = sql_text.count("\n") + 1
        ws_ddl.row_dimensions[curr_r + 1].height = max(lines * 14 + 10, 30)

        curr_r += 3

    ws_ddl.column_dimensions["A"].width = 110

    os.makedirs("docs", exist_ok=True)
    out_path = os.path.join("docs", "08_Database_Design.xlsx")
    wb.save(out_path)
    print(f"Successfully generated {out_path}")

if __name__ == "__main__":
    build_step8_workbook()
