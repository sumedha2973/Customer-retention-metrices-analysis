import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="D3D3D3", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def build_step5_brd():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(5)

    NAVY = RGBColor(0x1B, 0x36, 0x5D)
    STEEL = RGBColor(0x2B, 0x4C, 0x7E)

    # Document Header
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run("BUSINESS REQUIREMENTS DOCUMENT (BRD)")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Customer Retention Metrics Analysis — Enterprise Specifications")
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = STEEL

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = NAVY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = STEEL
        return p

    def format_table(table, headers, data, widths, align_center_cols=[]):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table, color="B0C4DE", sz="4")
        hdr_row = table.rows[0]
        for j, h in enumerate(headers):
            cell = hdr_row.cells[j]
            cell.width = widths[j]
            set_cell_background(cell, "1B365D")
            set_cell_margins(cell, 70, 70, 90, 90)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if j in align_center_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            r.font.size = Pt(9)
        
        for i, row_data in enumerate(data):
            row = table.rows[i + 1]
            bg_color = "F7F9FC" if i % 2 == 1 else "FFFFFF"
            for j, val in enumerate(row_data):
                cell = row.cells[j]
                cell.width = widths[j]
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, 60, 60, 90, 90)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                if j in align_center_cols:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(str(val))
                r.font.size = Pt(8.5)
                if j == 0 and ("BR-" in str(val) or "RULE-" in str(val) or "RSK-" in str(val) or "SH-" in str(val)):
                    r.font.bold = True
                    r.font.color.rgb = STEEL
                elif "Must" in str(val) or "High" in str(val) or "A" in str(val) or "R" in str(val):
                    if str(val) in ["Must", "High", "Critical"]:
                        r.font.bold = True

    # 1. Document Control
    add_h1("1. Document Control")
    doc_control_headers = ["Attribute", "Specification Details"]
    doc_control_data = [
        ("Project Name", "Customer Retention Metrics Analysis"),
        ("Project Code", "CRMA-DA-2026"),
        ("Document Title", "Business Requirements Document (BRD)"),
        ("Document Version", "Version 1.0 (Final Approved Baseline)"),
        ("Release Date", "20 November 2026"),
        ("Primary Author", "Sumedha (Roll No: 2503201450146, MCA 2nd Year, ABES Engineering College)"),
        ("Academic Mentor & Reviewer", "Ms. Anjali Pandey, Assistant Professor, ABES Engineering College"),
        ("Assumed Corporate Sponsor", "NoidaSoft Technologies Pvt. Ltd. (Executive Leadership Team)"),
        ("Document Sensitivity", "Internal Academic & Research Project (100% Synthetic Data)")
    ]
    t_meta = doc.add_table(rows=len(doc_control_data) + 1, cols=2)
    format_table(t_meta, doc_control_headers, doc_control_data, [Inches(2.2), Inches(4.3)], align_center_cols=[])

    # 2. Executive Summary
    add_h1("2. Executive Summary")
    p_exec1 = doc.add_paragraph()
    p_exec1.add_run("Customer retention is the primary determinant of long-term economic viability in subscription software organizations. NoidaSoft Technologies Pvt. Ltd.—a business-to-business (B2B) SaaS company serving approximately 5,000 small and medium-sized business (SMB) accounts across India—currently struggles with high reporting latency, inconsistent retention calculations, and uncoordinated account management due to fragmented data storage across disparate operational silos.")

    p_exec2 = doc.add_paragraph()
    p_exec2.add_run("This Business Requirements Document (BRD) specifies the functional, operational, and governance requirements for the Customer Retention Metrics Analysis system. By centralizing multi-table customer telemetry, billing histories, and service interactions into an automated analytical engine, the application eliminates 3 to 4 days of manual monthly Excel reporting, establishes an authoritative single source of truth for retention KPIs, delivers granular cohort survival tracking, and introduces a transparent early-warning customer health risk score to enable proactive account interventions.")

    # 3. Business Problem and Opportunity
    add_h1("3. Business Problem and Opportunity")
    add_h2("3.1 The Business Problem")
    p_prob = doc.add_paragraph()
    p_prob.add_run("In the current operating model, data resides in five disconnected systems: CRM records, subscription billing logs, helpdesk ticket archives, application login logs, and quarterly customer satisfaction surveys. Every month-end, analysts spend 3 to 4 working days performing fragile VLOOKUP operations to consolidate datasets. This causes significant business friction:\n"
                   "1. Information Latency: Leadership receives churn figures up to 30 days after customer attrition occurs.\n"
                   "2. Calculation Discrepancies: Finance and Customer Success compute churn using differing formulas, eroding executive trust.\n"
                   "3. Hidden Attrition Patterns: Aggregate churn hides underlying weakness in specific pricing plans, regions, or onboarding cohorts.\n"
                   "4. Reactive Firefighting: Account managers discover customer cancellations only after contract renewal dates have expired.")

    add_h2("3.2 The Business Opportunity")
    p_opp = doc.add_paragraph()
    p_opp.add_run("Research in subscription software economics confirms that acquiring a new customer costs 5 to 7 times more than retaining an existing account. By implementing an automated, unified retention analytics platform, NoidaSoft can:\n"
                  "• Reduce manual reporting overhead from 4 days to near-instantaneous query execution.\n"
                  "• Protect Monthly Recurring Revenue (MRR) through early detection of accounts exhibiting churn risk indicators.\n"
                  "• Optimize acquisition marketing spend by directing budget toward customer acquisition channels with superior long-term survival rates.")

    # 4. Objectives and Success Metrics
    add_h1("4. Business Objectives & Key Performance Indicators (KPIs)")
    p_obj_desc = doc.add_paragraph()
    p_obj_desc.add_run("The project is governed by measurable target KPIs directly supporting the SMART objectives established in Step 1:")

    kpi_headers = ["Objective ID", "Strategic Objective", "Target KPI / SLA", "Measurement & Verification Method"]
    kpi_data = [
        ("OBJ-1", "Synthetic Relational Database", ">= 5,000 accounts, 24 historical months, 6 normalized 3NF tables", "Automated SQL integrity validation script (0 duplicate PKs, 0 orphan FKs)."),
        ("OBJ-2", "Core Retention Metric Engine", ">= 8 core KPIs matching benchmark definitions at 100% parity", "Automated pytest parity tests comparing pandas computations against SQL benchmarks."),
        ("OBJ-3", "Interactive Analytical Visualizations", ">= 10 interactive Plotly charts covering all 10 OQs", "Interactive walkthrough of time-series, cohort heatmap, and risk distributions."),
        ("OBJ-4", "Inferential Statistical Driver Analysis", ">= 3 hypothesis tests (Chi-Square, t-test, correlation) with p < 0.05 reporting", "SciPy automated statistical test output verification on platform dashboard."),
        ("OBJ-5", "Software Quality Assurance", ">= 25 automated pytest test cases; >= 70% code coverage", "Automated test suite execution report and pytest-cov coverage metrics."),
        ("OBJ-6", "Cloud Hosting & Academic Handover", "App launch < 5.0s, filter latency < 3.0s; complete ABES document sign-off", "Streamlit Community Cloud deployment audit and academic mentor review.")
    ]
    t_kpi = doc.add_table(rows=len(kpi_data) + 1, cols=4)
    format_table(t_kpi, kpi_headers, kpi_data, [Inches(1.0), Inches(2.2), Inches(1.8), Inches(1.5)], align_center_cols=[0])

    # 5. Stakeholder RACI Matrix
    add_h1("5. Stakeholder Governance (RACI Matrix)")
    p_raci_intro = doc.add_paragraph()
    p_raci_intro.add_run("The RACI matrix delineates organizational accountability across the project lifecycle: Responsible (R), Accountable (A), Consulted (C), Informed (I):")

    raci_headers = ["Project Lifecycle Phase", "SH-01\nCEO", "SH-02\nHead CS", "SH-03\nCSMs", "SH-04\nMktg", "SH-05\nProd", "SH-06\nSupp", "SH-07\nFin", "SH-08\nAnalyst", "SH-09\nMentor", "SH-10\nStudent"]
    raci_data = [
        ("1. Requirements Definition (BRD/PRD)", "I", "C", "C", "C", "C", "C", "C", "C", "A", "R"),
        ("2. Data Modeling & Synthesis", "I", "I", "I", "I", "I", "I", "I", "C", "A", "R"),
        ("3. Retention Metric & Algorithm Design", "C", "A", "C", "I", "C", "I", "A", "C", "A", "R"),
        ("4. Dashboard Wireframing & UI Build", "I", "C", "C", "I", "C", "I", "I", "C", "A", "R"),
        ("5. QA Testing & Statistical Validation", "I", "I", "I", "I", "I", "I", "I", "C", "A", "R"),
        ("6. Cloud Deployment & Runbook", "I", "I", "I", "I", "I", "I", "I", "I", "A", "R"),
        ("7. Academic Review & Presentation", "I", "I", "I", "I", "I", "I", "I", "I", "A", "R")
    ]
    t_raci = doc.add_table(rows=len(raci_data) + 1, cols=11)
    format_table(t_raci, raci_headers, raci_data, [Inches(2.0), Inches(0.45), Inches(0.45), Inches(0.45), Inches(0.45), Inches(0.45), Inches(0.45), Inches(0.45), Inches(0.45), Inches(0.45), Inches(0.45)], align_center_cols=[1,2,3,4,5,6,7,8,9,10])

    # 6. Scope
    add_h1("6. Project Scope")
    scope_headers = ["In Scope (Phase 1 Baseline)", "Out of Scope (Deferred / Excluded)"]
    scope_data = [
        ("Synthetic B2B SaaS data modeling (24 months, 5,000 customer accounts)", "Real customer personally identifiable information (PII) or commercial databases"),
        ("SQLite relational database design and 3NF physical schema", "Live production API integrations (Salesforce CRM, Zendesk, Stripe, Jira)"),
        ("Core retention calculations: Churn %, Retention %, Cohorts, CLV, MRR Churn, NRR", "Complex machine learning predictive churn models (rule-based risk model used)"),
        ("Interactive Plotly visual analytics: time-series, cohort heatmap, waterfall charts", "Automated email/SMS dispatch or external webhook alerting engines"),
        ("Inferential statistical hypothesis testing (Chi-square, t-test, Mann-Whitney U)", "A/B testing experimentation platform or marketing campaign automation"),
        ("Streamlit interactive web dashboard styled with custom Tailwind CSS utility blocks", "User authentication, multi-tenant RBAC, and login session management"),
        ("Filtered CSV export for operational customer success follow-ups", "Native mobile operating system applications (iOS / Android)"),
        ("Automated pytest test suite (>= 25 test cases) covering logic parity and integrity", "High-concurrency distributed load testing or network security penetration tests"),
        ("Public deployment on Streamlit Community Cloud with public URL access", "Dedicated private cloud infrastructure (AWS/Azure/GCP) or custom domain SSL")
    ]
    t_scope = doc.add_table(rows=len(scope_data) + 1, cols=2)
    format_table(t_scope, scope_headers, scope_data, [Inches(3.25), Inches(3.25)])

    # 7. Business Requirements
    add_h1("7. Detailed Business Requirements (BR-01 through BR-13)")
    p_br_intro = doc.add_paragraph()
    p_br_intro.add_run("The functional requirements below define the complete business capabilities authorized for the CRMA platform, tracing directly to the RTM established in Step 4:")

    br_specs = [
        ("BR-01", "Official Retention & Churn Rate Standards", "HMW-01, HMW-11", "SH-01, SH-08", "Must",
         "The business needs to measure customer churn rate and retention rate for any month and quarter, using one agreed written definition, so that leadership and every team rely on the same official number.",
         "Conflicting churn definitions between finance and customer success create executive reporting gridlock and prevent effective resource allocation.",
         "GIVEN a selected historical calendar month, WHEN churn and retention rates are calculated, THEN Customer Churn Rate = (Lost Accounts / Starting Active Accounts) * 100 and Retention Rate = 100 - Churn Rate, matching financial audit figures with 0.0% variance."),
        
        ("BR-02", "Historical 24-Month Trend Analysis", "HMW-01", "SH-01", "Must",
         "The business needs to compare churn and retention over the full 24 months, so that leadership can tell whether retention is improving or getting worse.",
         "Single-month snapshots fail to reveal seasonality or whether strategic product improvements are successfully bending the customer retention curve.",
         "GIVEN 24 consecutive calendar months of data, WHEN the time-series analytics view is generated, THEN monthly retention rates, churn rates, and active account trajectories are plotted continuously with trailing 3-month moving averages."),
        
        ("BR-03", "Multi-Dimensional Customer Segmentation", "HMW-02", "SH-02", "Must",
         "The business needs to compare retention across customer segments (plan type, industry, company size and region), so that retention effort goes first to the weakest segments.",
         "Aggregate metrics mask severe attrition in specific customer niches, preventing customer success from deploying targeted playbooks.",
         "GIVEN active customer records, WHEN filtered across Plan Type (Starter/Pro/Enterprise), Industry, Company Size, or Region, THEN the system recalculates segment-specific retention rates and highlights underperforming segments exceeding the corporate churn threshold."),
        
        ("BR-04", "12-Month Cohort Survival Matrix", "HMW-03", "SH-02", "Must",
         "The business needs to follow each signup-month group of customers through its first 12 months, so that it can tell whether newer customers stay as long as earlier ones.",
         "Cohort analysis is the definitive metric for SaaS product health, revealing whether onboarding enhancements improve customer longevity over time.",
         "GIVEN all accounts grouped by original signup month, WHEN the cohort matrix is displayed, THEN each row represents a monthly cohort and columns display percentage retention at Month 1, Month 3, Month 6, and Month 12, visualized as a triangular heatmap."),
        
        ("BR-05", "Proactive At-Risk Account Identification", "HMW-04", "SH-02, SH-03", "Must",
         "The business needs to identify active customers who show warning signs (falling usage, open tickets, low satisfaction scores, renewal approaching) and see the reasons, so that customer success managers can act before the renewal date.",
         "Reactive churn management occurs after contracts lapse; proactive early warnings allow CSMs to intervene during the critical 30-to-60-day pre-renewal window.",
         "GIVEN all active customer accounts, WHEN the health score is evaluated, THEN every account is assigned an objective composite risk score (0-100) and categorized into High (>=70), Medium (40-69), or Low (<40) with specific contributing risk drivers listed."),
        
        ("BR-06", "Support SLA & Ticket Churn Impact", "HMW-07", "SH-06", "Must",
         "The business needs to find out whether customers with more support tickets or slower resolution leave more often, so that support effort can be directed at issues that affect retention.",
         "Support leadership lacks quantitative evidence to justify helpdesk staffing increases or revised SLA escalation pathways.",
         "GIVEN support ticket logs and customer churn records, WHEN statistical comparison is executed, THEN ticket volume per customer and resolution turnaround hours are compared between retained and churned accounts, reporting mean differences and statistical significance (p-value)."),
        
        ("BR-07", "Monthly Recurring Revenue (MRR) Churn & NRR", "HMW-08, HMW-09", "SH-07", "Must",
         "The business needs to know how much monthly recurring revenue is lost to churn and how revenue from existing customers changes each month, so that finance can see the money impact and forecast it.",
         "Customer count churn distorts financial reality; losing five low-tier accounts does not equal losing one top-tier Enterprise account.",
         "GIVEN historical subscription fee records, WHEN the financial revenue engine runs, THEN Gross MRR Churn, Expansion MRR, and Net Revenue Retention (NRR %) are calculated per month, where NRR > 100% indicates net expansion."),
        
        ("BR-08", "Academic Metric Traceability & Governance", "HMW-12", "SH-09, SH-10", "Must",
         "The business needs every reported metric and test result to be linked to a written definition and a named data source, so that the mentor and reviewers can verify each number.",
         "Academic compliance and enterprise governance require that all presented metrics are fully auditable, reproducible, and verifiable.",
         "GIVEN any KPI card or statistical visual, WHEN inspected by a reviewer, THEN an accompanying definition panel states the mathematical formula, underlying tables/columns, sample size, and assumptions."),
        
        ("BR-09", "Customer Lifetime Value (CLV) & Tenure", "HMW-09", "SH-07, SH-01", "Should",
         "The business needs to know the average customer lifetime and the lifetime value of a customer for each segment, so that finance and leadership can judge what a customer is worth.",
         "Customer acquisition budgets cannot be rationally planned without knowing the historical economic return generated across plan tiers.",
         "GIVEN payment histories and subscription durations, WHEN CLV analytics are requested, THEN average tenure in months and historical Customer Lifetime Value are computed per plan tier and visualized in comparison tables."),
        
        ("BR-10", "Acquisition Channel Retention Comparison", "HMW-05", "SH-04", "Should",
         "The business needs to compare how long customers stay by acquisition channel, so that Sales and Marketing can favour channels that bring lasting customers.",
         "Marketing teams frequently optimize for low customer acquisition costs (CAC) on channels that yield fragile accounts that churn immediately.",
         "GIVEN customer attribution data across Organic, Referral, Paid Ads, and Partners, WHEN channel longevity is analyzed, THEN average retention rates and Kaplan-Meier style survival rates are compared across acquisition channels."),
        
        ("BR-11", "Pre-Churn Usage Telemetry Comparison", "HMW-06", "SH-05", "Should",
         "The business needs to compare the product usage of customers who left with those who stayed, in the months before leaving, so that product managers can see where engagement drops.",
         "Product managers must determine whether feature abandonment or declining login frequency serves as an early behavioral signal of cancellation.",
         "GIVEN monthly application telemetry (logins and core feature clicks), WHEN usage is evaluated across the 90 days preceding churn, THEN a statistically significant decline in active days is displayed relative to benchmark retained accounts."),
        
        ("BR-12", "Customer Satisfaction (CSAT/NPS) Correlation", "HMW-10", "SH-05", "Could",
         "The business needs to see whether customers who give low satisfaction scores are less likely to renew, so that the product team can decide how far to rely on these scores.",
         "Survey response rates are voluntary; leadership must evaluate whether CSAT/NPS ratings accurately predict renewal outcomes or merely reflect survey bias.",
         "GIVEN quarterly survey records, WHEN correlation analysis is conducted, THEN the Pearson/Spearman correlation coefficient between satisfaction scores and renewal status is computed with clear distribution visualizations."),
        
        ("BR-13", "Automated Live System Data Refresh", "HMW-11", "SH-08", "Won’t",
         "The business needs to refresh its retention figures automatically from live company systems, so that analysts no longer prepare data files by hand.",
         "Live API connectors require ongoing authentication infrastructure, credential management, and transactional overhead inappropriate for a static analytical project.",
         "Formally deferred to Phase 2 roadmap; verified that Phase 1 utilizes high-performance indexed SQLite file-based ingestion.")
    ]

    for req in br_specs:
        add_h2(f"{req[0]}: {req[1]}")
        p_r = doc.add_paragraph()
        p_r.add_run(f"• Priority: ").bold = True
        r_prio = p_r.add_run(f"{req[4]}\n")
        if req[4] == "Must":
            r_prio.bold = True
            r_prio.font.color.rgb = RGBColor(0x8B, 0x00, 0x00)
        p_r.add_run(f"• Source HMW: ").bold = True
        p_r.add_run(f"{req[2]}  |  ")
        p_r.add_run(f"Stakeholders: ").bold = True
        p_r.add_run(f"{req[3]}\n")
        p_r.add_run(f"• Requirement Statement: ").bold = True
        p_r.add_run(f"{req[5]}\n")
        p_r.add_run(f"• Business Rationale: ").bold = True
        p_r.add_run(f"{req[6]}\n")
        p_r.add_run(f"• Acceptance Criteria: ").bold = True
        p_r.add_run(f"{req[7]}")

    # 8. Business Rules
    add_h1("8. Core Business Rules & Metric Calculation Logic")
    p_rules_intro = doc.add_paragraph()
    p_rules_intro.add_run("To eliminate ambiguity and resolve historical calculation disputes, the following enterprise business rules are codified for all analytical calculations:")

    rules_headers = ["Rule ID", "Governing Business Rule", "Mathematical Formulation / Logic", "Operational Enforcement"]
    rules_data = [
        ("RULE-01", "Standard Churn Definition", 
         "An account is classified as Churned in month M if its subscription expiration date occurred in month M and no renewal payment or active contract is recorded within a 30-day grace period.", 
         "Enforced across all customer status categorization queries."),
        
        ("RULE-02", "Customer Churn & Retention Rates", 
         "Customer Churn Rate (CCR) = (Customers Churned in Month M / Active Customers at Start of Month M) * 100.\nCustomer Retention Rate (CRR) = 100 - CCR (for closed cohorts).", 
         "Evaluated monthly; starting customers excludes new accounts acquired within month M."),
        
        ("RULE-03", "Net Revenue Retention (NRR)", 
         "NRR (%) = [(Starting MRR + Expansion MRR - Contraction MRR - Churned MRR) / Starting MRR] * 100.", 
         "Measures the net change in recurring revenue from retained accounts over a 12-month trailing window."),
        
        ("RULE-04", "Customer Lifetime Value (CLV)", 
         "Historical CLV = Average Monthly Revenue Per Account (ARPU) * Average Active Lifespan in Months.\nSegment CLV = Cumulative Billing Revenue for Segment / Total Unique Customers in Segment.", 
         "Enables direct ROI comparison across Starter, Pro, and Enterprise tiers."),
        
        ("RULE-05", "Multi-Factor Account Health Score", 
         "Risk Score (0 to 100) = (Login Decline * 0.35) + (Unresolved Tickets * 0.25) + (Low CSAT Rating * 0.20) + (Renewal Proximity * 0.20).\nClassification: High Risk (>=70), Medium Risk (40-69), Low Risk (<40).", 
         "Enforced dynamically across all active accounts approaching contract renewal."),
        
        ("RULE-06", "Cohort Grouping Convention", 
         "A customer cohort comprises all accounts whose initial subscription activation occurred within the same calendar Year-Month (YYYY-MM).", 
         "Monitored longitudinally at discrete lifecycle intervals: M+1, M+3, M+6, M+12."),
        
        ("RULE-07", "Statistical Significance Standard", 
         "All inferential hypothesis tests (Chi-Square test of independence, two-sample t-test, Mann-Whitney U test) require an alpha threshold of 0.05.", 
         "P-values < 0.05 reject the null hypothesis and confirm statistically significant driver relationships.")
    ]
    t_rules = doc.add_table(rows=len(rules_data) + 1, cols=4)
    format_table(t_rules, rules_headers, rules_data, [Inches(1.0), Inches(1.8), Inches(2.2), Inches(1.5)], align_center_cols=[0])

    # 9. Data Requirements
    add_h1("9. Data Architecture & Entity Requirements")
    p_data_desc = doc.add_paragraph()
    p_data_desc.add_run("The analytical engine requires six normalized entity tables within an embedded SQLite database:")

    data_headers = ["Entity Name", "Business Description", "Key Attributes Required", "Update Frequency"]
    data_entities = [
        ("Customers", "Core firmographic account profile for all 5,000 SMB accounts.", "customer_id (PK), company_name, industry, company_size, region, channel, signup_date.", "Static 24-Month Snapshot"),
        ("Subscriptions", "Contract lifecycle, tier plans, monthly fees, and renewal outcomes.", "subscription_id (PK), customer_id (FK), plan_type, monthly_fee, start_date, end_date, status, renewed_flag.", "Static 24-Month Snapshot"),
        ("Payments", "Historical transactional billing ledgers and monthly fee collections.", "payment_id (PK), customer_id (FK), payment_date, amount_inr, payment_method, payment_status.", "Static 24-Month Snapshot"),
        ("Support Tickets", "Customer service interactions, issue categories, and SLA turnaround.", "ticket_id (PK), customer_id (FK), created_date, resolved_date, resolution_hours, priority, status.", "Static 24-Month Snapshot"),
        ("Product Usage", "Monthly telemetry tracking platform logins and core feature actions.", "usage_id (PK), customer_id (FK), activity_month, login_days, core_feature_events.", "Static 24-Month Snapshot"),
        ("Surveys", "Quarterly customer feedback ratings (CSAT and Net Promoter Score).", "survey_id (PK), customer_id (FK), survey_date, csat_score (1-5), nps_score (0-10).", "Static 24-Month Snapshot")
    ]
    t_data = doc.add_table(rows=len(data_entities) + 1, cols=4)
    format_table(t_data, data_headers, data_entities, [Inches(1.2), Inches(2.1), Inches(2.2), Inches(1.0)], align_center_cols=[0, 3])

    # 10. Assumptions, Constraints, Dependencies
    add_h1("10. Assumptions, Constraints, and Dependencies")
    add_h2("10.1 Business Assumptions")
    p_a = doc.add_paragraph()
    p_a.add_run("• A-01: NoidaSoft Technologies is a simulated corporate environment representing an Indian B2B SaaS organization.\n"
                "• A-02: All underlying data is synthetic, deterministic, and generated using fixed random seeds to ensure 100% reproducibility.\n"
                "• A-03: The commercial model is contract-based recurring subscription; customers renew or churn at contract expiration.\n"
                "• A-04: The historical observation window spans exactly 24 calendar months (January 2024 to December 2025).\n"
                "• A-05: All financial figures are denominated in Indian Rupees (INR, ₹) with no multi-currency exchange rate adjustments.\n"
                "• A-06: A transparent, rule-based risk scoring algorithm satisfies business needs in place of an opaque ML model.\n"
                "• A-16: HMW-12 is cross-cutting and governs institutional traceability across all requirements.\n"
                "• A-17: BR-13 (Won't Have) is retained in the RTM to maintain formal record of deferred scope.")

    add_h2("10.2 Technical & Operational Constraints")
    p_c = doc.add_paragraph()
    p_c.add_run("• C-01: Single developer (Sumedha) executing full-stack development within the academic semester timeframe.\n"
                "• C-02: Open-source and free-tier software stack: Python 3.10+, SQLite3, Streamlit, Plotly, SciPy, pytest, openpyxl, python-docx.\n"
                "• C-03: Streamlit Community Cloud hosting limits memory consumption (~1 GB); database footprint must remain under 50 MB.\n"
                "• C-04: Development lifecycle discipline: Zero application or synthesis code written before Step 10.")

    add_h2("10.3 Project Dependencies")
    p_d = doc.add_paragraph()
    p_d.add_run("• DEP-01: Completion and approval of Step 4 Traceability Matrix prior to technical architecture modeling.\n"
                "• DEP-02: Synthetic generation engine producing statistically coherent behavioral distributions (Step 10).\n"
                "• DEP-03: Streamlit Community Cloud GitHub continuous deployment integration.")

    # 11. Risks and Mitigation
    add_h1("11. Risk Management Matrix")
    risk_headers = ["Risk ID", "Risk Description", "Likelihood", "Impact", "Mitigation Strategy"]
    risk_data = [
        ("RSK-01", "Synthetic data lacks realistic behavioral correlations, causing statistical tests to fail.", "Medium", "High", 
         "Inject deliberate correlation rules during dataset generation (e.g., ticket surges and login declines preceding churn events)."),
        ("RSK-02", "Streamlit Community Cloud instance crashes due to memory overhead during cohort heatmap rendering.", "Low", "High", 
         "Pre-aggregate cohort matrices and indexed SQL query results; avoid in-memory raw row recalculations."),
        ("RSK-03", "Discrepancies arise between pandas KPI logic and SQL benchmark queries.", "Medium", "Medium", 
         "Implement automated pytest regression test suite asserting exact parity (Delta = 0.0) between pandas and SQL outputs."),
        ("RSK-04", "Scope creep introduced through automated email alerting or machine learning requests.", "Low", "Medium", 
         "Strict enforcement of MoSCoW prioritization: BR-12 (alerting) and ML modeling formally classified as Won't Have in BRD.")
    ]
    t_risk = doc.add_table(rows=len(risk_data) + 1, cols=5)
    format_table(t_risk, risk_headers, risk_data, [Inches(0.8), Inches(2.2), Inches(0.9), Inches(0.8), Inches(1.8)], align_center_cols=[0, 2, 3])

    # 12. Glossary
    add_h1("12. Glossary of Terms")
    gloss_headers = ["Term / Acronym", "Enterprise Definition"]
    gloss_data = [
        ("ARR / MRR", "Annual Recurring Revenue / Monthly Recurring Revenue — predictable recurring subscription revenue normalized to annual or monthly periods."),
        ("CCR (Customer Churn Rate)", "Percentage of active accounts at the start of a period that cancel or fail to renew their subscription during that period."),
        ("CRR (Customer Retention Rate)", "Percentage of active accounts that remain active throughout the period, calculated as 100% - CCR for a closed base."),
        ("Cohort Analysis", "Analytical technique that separates users into related groups based on common activation dates to observe behavioral decay longitudinally."),
        ("CLV (Customer Lifetime Value)", "Total gross economic revenue generated by a customer account throughout its entire active relationship with the firm."),
        ("NRR (Net Revenue Retention)", "Percentage of recurring revenue retained from existing customers over a period, incorporating expansion, contraction, and churn."),
        ("CSAT (Customer Satisfaction)", "Transaction-level sentiment score rating service experience from 1 (Very Unsatisfied) to 5 (Very Satisfied)."),
        ("NPS (Net Promoter Score)", "Customer loyalty metric based on willingness to recommend on a scale of 0 to 10; categorized into Promoters, Passives, and Detractors."),
        ("3NF (Third Normal Form)", "Relational database schema design standard that eliminates data redundancy and prevents insertion, update, and deletion anomalies.")
    ]
    t_gloss = doc.add_table(rows=len(gloss_data) + 1, cols=2)
    format_table(t_gloss, gloss_headers, gloss_data, [Inches(1.8), Inches(4.7)], align_center_cols=[0])

    # 13. Sign-off
    add_h1("13. Formal Requirements Sign-Off")
    p_sign_desc = doc.add_paragraph()
    p_sign_desc.add_run("By signing below, the stakeholders confirm that this Business Requirements Document accurately reflects business objectives, operational definitions, and project scope. The agreed baseline authorizes progression to Step 6 (Product Requirements Document).")

    sign_headers = ["Stakeholder Role", "Designated Name & Title", "Signature", "Approval Date"]
    sign_data = [
        ("Academic Mentor & Reviewer", "Ms. Anjali Pandey\nAssistant Professor, ABES EC", "____________________", "20 November 2026"),
        ("Student Lead & Analyst", "Sumedha\nMCA 2nd Year, Roll: 2503201450146", "____________________", "20 November 2026"),
        ("Head of Customer Success", "Corporate Sponsor Representative\nNoidaSoft Technologies Pvt. Ltd.", "____________________", "20 November 2026"),
        ("Executive Leadership", "Chief Executive Officer\nNoidaSoft Technologies Pvt. Ltd.", "____________________", "20 November 2026")
    ]
    t_sign = doc.add_table(rows=len(sign_data) + 1, cols=4)
    format_table(t_sign, sign_headers, sign_data, [Inches(1.8), Inches(2.2), Inches(1.3), Inches(1.2)], align_center_cols=[2, 3])

    os.makedirs("docs", exist_ok=True)
    out_path = os.path.join("docs", "05_BRD.docx")
    doc.save(out_path)
    print(f"Successfully generated {out_path}")

if __name__ == "__main__":
    build_step5_brd()
