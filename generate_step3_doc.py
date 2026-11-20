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

def build_step3_document():
    doc = docx.Document()
    
    # 1-inch margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    NAVY = RGBColor(0x1B, 0x36, 0x5D)
    STEEL = RGBColor(0x2B, 0x4C, 0x7E)

    # Title & Subtitle
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run("NEEDS ANALYSIS & BUSINESS REQUIREMENTS")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Customer Retention Metrics Analysis — Business Capability Formulation")
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(14)
    run_sub.font.color.rgb = STEEL

    # Metadata Table
    meta_data = [
        ("Project name", "Customer Retention Metrics Analysis"),
        ("Domain", "Data Analytics (Python)"),
        ("Student", "Sumedha"),
        ("Roll number", "2503201450146"),
        ("Programme", "MCA, 2nd Year, ABES Engineering College"),
        ("Mentor", "Ms. Anjali Pandey"),
        ("Deliverable", "Step 3: Needs Analysis & Business Requirements"),
        ("Version / date", "v1.0 | 19 November 2026")
    ]
    t_meta = doc.add_table(rows=len(meta_data), cols=2)
    t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_meta, color="CCCCCC", sz="4")
    col_widths = [Inches(1.8), Inches(4.7)]
    for i, (k, v) in enumerate(meta_data):
        row = t_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)
        set_cell_background(c0, "F0F4F8")
        set_cell_background(c1, "FFFFFF")
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = NAVY

        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = NAVY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = STEEL
        return p

    def format_table(table, headers, data, widths, align_right_cols=[]):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table, color="B0C4DE", sz="4")
        hdr_row = table.rows[0]
        for j, h in enumerate(headers):
            cell = hdr_row.cells[j]
            cell.width = widths[j]
            set_cell_background(cell, "1B365D")
            set_cell_margins(cell, 80, 80, 100, 100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if j in align_right_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(h)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            r.font.size = Pt(9.5)
        
        for i, row_data in enumerate(data):
            row = table.rows[i + 1]
            bg_color = "F7F9FC" if i % 2 == 1 else "FFFFFF"
            for j, val in enumerate(row_data):
                cell = row.cells[j]
                cell.width = widths[j]
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, 70, 70, 100, 100)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                if j in align_right_cols:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r = p.add_run(str(val))
                r.font.size = Pt(9)
                if j == 0 and ("BR-" in str(val) or "HMW-" in str(val)):
                    r.font.bold = True
                    r.font.color.rgb = STEEL
                elif "Must" in str(val):
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(0x8B, 0x00, 0x00)

    # 1. Introduction
    add_h1("1. Needs Analysis Overview")
    p_intro = doc.add_paragraph()
    p_intro.add_run("The Needs Analysis translates the design-thinking problem statements (HMWs) formulated in Step 2 into explicit, testable Business Requirements (BRs). In strict adherence to professional Business Analysis standards, each requirement defines WHAT capability the business requires and WHY (the organizational benefit), without prescribing HOW the capability will be implemented (avoiding references to SQL queries, Streamlit widgets, or specific programming constructs).")

    # 2. Master Table of Business Requirements
    add_h1("2. Master Business Requirements Matrix (BR-01 to BR-12)")
    p_tbl_note = doc.add_paragraph()
    p_tbl_note.add_run("Each requirement follows the standard enterprise syntax: 'The business needs to <capability> so that <benefit>.' Prioritization adheres to MoSCoW methodology (Must, Should, Could, Won't), maintaining a 58.3% (~60%) ratio for core Must-Have capabilities.")

    br_headers = ["BR ID", "Business Requirement Statement", "Stakeholder", "Source HMW", "Priority", "Measurement & Verification Method"]
    br_data = [
        ("BR-01", 
         "The business needs to compute standardized monthly and quarterly customer retention and churn rates across historical periods so that executive leadership can evaluate overall business health from a single source of truth without reporting delays.", 
         "SH-01, SH-02", 
         "HMW-01", 
         "Must", 
         "100% mathematical parity against verified historical benchmark calculations; zero cross-departmental discrepancy."),
        
        ("BR-02", 
         "The business needs to segment customer retention and churn metrics by subscription plan, industry vertical, company size, and geographic region so that customer success teams can direct targeted retention programs to high-risk business segments.", 
         "SH-02, SH-03", 
         "HMW-02", 
         "Must", 
         "Verification of multi-dimensional filtering across all 4 firmographic attributes with accurate subgroup retention rates."),
        
        ("BR-03", 
         "The business needs to analyze customer cohort survival rates at 1, 3, 6, and 12 months from original signup so that product and operational leadership can evaluate whether customer retention improves over successive onboarding generations.", 
         "SH-02, SH-05", 
         "HMW-03", 
         "Must", 
         "Validated cohort survival matrix displaying retention percentages at exact monthly lifecycle intervals for all 24 signup cohorts."),
        
        ("BR-04", 
         "The business needs to calculate average customer tenure and Customer Lifetime Value (CLV) across plan tiers so that leadership can prioritize customer retention investments toward segments with the highest long-term economic return.", 
         "SH-01, SH-07", 
         "HMW-04", 
         "Should", 
         "Audited calculation of average active months and cumulative gross revenue per segment matching billing ledger totals."),
        
        ("BR-05", 
         "The business needs to compare customer retention rates and average lifespan across different acquisition channels so that marketing leadership can reallocate acquisition budgets toward channels delivering the most sustainable customer relationships.", 
         "SH-04", 
         "HMW-05", 
         "Should", 
         "Channel retention breakdown and survival comparison verified with chi-square statistical significance reporting."),
        
        ("BR-06", 
         "The business needs to evaluate the relationship between support ticket volume, resolution turnaround times, and contract cancellation rates so that customer support leads can establish SLA thresholds that minimize churn risk.", 
         "SH-06", 
         "HMW-06", 
         "Must", 
         "Hypothesis test and distribution comparison showing statistically significant difference in support metrics between retained and churned accounts."),
        
        ("BR-07", 
         "The business needs to track customer platform login frequency and core feature utilization over time so that product managers and account teams can identify leading behavioral indicators of account abandonment.", 
         "SH-05, SH-03", 
         "HMW-07", 
         "Must", 
         "Trend comparison demonstrating statistically lower engagement activity in pre-churn windows compared to retained accounts."),
        
        ("BR-08", 
         "The business needs to analyze the correlation between customer satisfaction survey ratings (CSAT and NPS) and renewal outcomes so that customer success leadership can determine how customer sentiment influences contract renewals.", 
         "SH-02, SH-03", 
         "HMW-08", 
         "Should", 
         "Calculated correlation coefficients and renewal rate distributions across CSAT and NPS score bands."),
        
        ("BR-09", 
         "The business needs to measure Monthly Recurring Revenue (MRR) lost to churn and calculate Net Revenue Retention (NRR) across monthly cycles so that finance leadership can monitor recurring revenue health and forecast cash-flow stability.", 
         "SH-07, SH-01", 
         "HMW-09", 
         "Must", 
         "Monthly MRR churn amount and NRR percentage calculated and reconciled exactly against total active recurring subscription revenue."),
        
        ("BR-10", 
         "The business needs to generate an objective, multi-factor customer health risk score for all active accounts so that customer success managers can identify and prioritize vulnerable accounts for outreach before their renewal deadline.", 
         "SH-03, SH-02", 
         "HMW-10", 
         "Must", 
         "100% of active accounts scored into High, Medium, and Low risk tiers based on defined usage, ticket, satisfaction, and renewal criteria."),
        
        ("BR-11", 
         "The business needs to export filtered lists of at-risk customer accounts and segment metrics to standard spreadsheet format so that customer success managers and analysts can execute offline account intervention workflows.", 
         "SH-03, SH-08", 
         "HMW-02, HMW-10", 
         "Could", 
         "Successful CSV data export containing filtered account records, risk categories, and renewal dates generated in under 3 seconds."),
        
        ("BR-12", 
         "The business needs to trigger automated real-time email or SMS alert notifications to account managers when an enterprise client enters high-risk status so that emergency customer retention workflows can begin immediately.", 
         "SH-03", 
         "HMW-01, HMW-07, HMW-10", 
         "Won't", 
         "Formally documented as out of scope for Phase 1 to maintain focus on core analytics; scheduled for future system phase.")
    ]
    t_br = doc.add_table(rows=len(br_data) + 1, cols=6)
    format_table(t_br, br_headers, br_data, [Inches(0.7), Inches(2.3), Inches(0.9), Inches(0.7), Inches(0.6), Inches(1.5)])

    # 3. MoSCoW Summary
    add_h1("3. MoSCoW Prioritization Summary")
    p_mos = doc.add_paragraph()
    p_mos.add_run("The 12 business requirements are distributed across the MoSCoW hierarchy to balance delivery feasibility with stakeholder value:\n"
                  "• Must Have (Core Capabilities): 7 requirements (58.3%) — BR-01, BR-02, BR-03, BR-06, BR-07, BR-09, BR-10.\n"
                  "• Should Have (High Value): 3 requirements (25.0%) — BR-04, BR-05, BR-08.\n"
                  "• Could Have (Desirable): 1 requirement (8.3%) — BR-11.\n"
                  "• Won't Have (Deferred to Future Phases): 1 requirement (8.3%) — BR-12.")

    # 4. Data Needs per Requirement
    add_h1("4. Data Entities Required per Business Requirement")
    p_data_intro = doc.add_paragraph()
    p_data_intro.add_run("To ensure feasibility, each business requirement is mapped to the explicit business data elements it consumes:")

    data_headers = ["BR ID", "Requirement Summary", "Required Business Data Elements"]
    data_reqs = [
        ("BR-01", "Overall Churn & Retention Rates", "Account identifiers, subscription start/end dates, renewal status, cancellation effective date."),
        ("BR-02", "Segmented Retention Analysis", "Customer firmographics (industry, company size, geographic region), subscription plan tier, status."),
        ("BR-03", "Cohort Survival Analysis", "Account activation/signup date (cohort month), active subscription status across 24 monthly periods."),
        ("BR-04", "Tenure & Customer Lifetime Value", "Account active tenure duration (months), monthly recurring price tier, historical billing/payment totals."),
        ("BR-05", "Acquisition Channel Performance", "Customer acquisition source attribution (Direct, Referral, Organic, Paid Ads, Partners), churn flag, tenure."),
        ("BR-06", "Support SLA Impact Analysis", "Support ticket records, ticket creation timestamp, resolution elapsed hours, customer ID, churn status."),
        ("BR-07", "Usage Telemetry & Drop-Off", "Monthly login days, core feature usage interaction counts, activity month, customer ID, churn date."),
        ("BR-08", "Satisfaction vs. Renewal Analysis", "Customer survey submissions, CSAT rating (1-5), Net Promoter Score (0-10), survey date, renewal outcome."),
        ("BR-09", "MRR Churn & Net Revenue Retention", "Monthly recurring subscription fees, upgrade/expansion revenue, contraction revenue, churn dates."),
        ("BR-10", "At-Risk Account Health Scoring", "30-day login frequency change, pending open support tickets, latest CSAT score, days remaining until contract expiry."),
        ("BR-11", "Data Export for Operational Outreach", "Account metadata, contact details, computed risk scores, account owner, contract renewal date."),
        ("BR-12", "Automated Real-Time Alerts", "Account manager contact endpoints (email/SMS), event webhooks, real-time risk threshold triggers (deferred).")
    ]
    t_data = doc.add_table(rows=len(data_reqs) + 1, cols=3)
    format_table(t_data, data_headers, data_reqs, [Inches(0.8), Inches(2.2), Inches(3.7)])

    # 5. Gap Analysis
    add_h1("5. Bidirectional Traceability & Gap Analysis")
    p_gap = doc.add_paragraph()
    p_gap.add_run("A rigorous bidirectional audit was conducted between the 10 How Might We (HMW) design statements and the 12 Business Requirements (BRs) to verify completeness and identify potential requirements gaps:")
    
    gap_headers = ["Analysis Dimension", "Audit Question", "Audit Finding & Verification"]
    gap_data = [
        ("HMW-to-BR Traceability", "Are there any orphan HMW statements without a corresponding Business Requirement?", 
         "Zero Gaps (100% Covered). Every HMW (HMW-01 through HMW-10) generates at least one formal BR."),
        ("BR-to-HMW Traceability", "Are there any extraneous Business Requirements that do not trace back to an authorized HMW?", 
         "Zero Gaps (100% Traceable). Every BR (BR-01 through BR-12) originates directly from an established HMW source."),
        ("Scope Boundary Audit", "Are out-of-scope or deferred capabilities appropriately categorized?", 
         "Compliant. Automated alerting (BR-12) is explicitly designated as Won't Have, preventing scope creep."),
        ("Verification Testability", "Can every requirement be validated through an objective, measurable test?", 
         "100% Testable. Every BR possesses a concrete verification criterion based on mathematical parity or empirical evaluation.")
    ]
    t_gap = doc.add_table(rows=len(gap_data) + 1, cols=3)
    format_table(t_gap, gap_headers, gap_data, [Inches(1.5), Inches(2.2), Inches(3.0)])

    out_path = os.path.join("docs", "03_Needs_Analysis.docx")
    doc.save(out_path)
    print(f"Successfully generated {out_path}")

if __name__ == "__main__":
    build_step3_document()
