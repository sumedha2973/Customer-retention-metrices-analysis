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

def build_step2_document():
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
    run_title = p_title.add_run("HOW MIGHT WE (HMW) STATEMENTS")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Customer Retention Metrics Analysis — Design Thinking Facilitation")
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
        ("Deliverable", "Step 2: How Might We (HMW) Statements"),
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
                if j == 0 and ("HMW-" in str(val) or "SH-" in str(val) or "OQ-" in str(val)):
                    r.font.bold = True
                    r.font.color.rgb = STEEL

    # 1. Stakeholder Pain Points
    add_h1("1. Stakeholder Pain Points Analysis")
    p_intro = doc.add_paragraph()
    p_intro.add_run("Following the design-thinking framework and building directly upon Step 1 (Project Understanding), each stakeholder's role is examined to isolate the critical frictions, operational blindspots, and systemic pain points they experience in the current state.")

    pain_headers = ["SH ID", "Stakeholder & Role", "Key Pain Points in Current State"]
    pain_data = [
        ("SH-01", "Chief Executive Officer\n(Executive Sponsor)", 
         "1. Cannot assess overall customer retention health without waiting 3-4 days for manual monthly reports.\n2. Lacks clarity on whether high customer churn is secretly diluting business growth ('leaky bucket' syndrome)."),
        ("SH-02", "Head of Customer Success\n(Primary Business Owner)", 
         "1. Has no standardized, repeatable churn definition across customer segments and regions.\n2. Cannot prioritize customer success team bandwidth because at-risk accounts are identified only after cancellation."),
        ("SH-03", "Customer Success Managers\n(Daily Operational Users)", 
         "1. Waste hours cross-referencing disconnected billing and ticketing spreadsheets prior to client renewal calls.\n2. Lack early behavioral warning indicators when an assigned client begins disengaging weeks before renewal."),
        ("SH-04", "Head of Sales & Marketing\n(Acquisition Owner)", 
         "1. Evaluates marketing channels purely on raw acquisition volume rather than long-term customer retention.\n2. Cannot tell which campaigns bring high-churn, short-lived accounts versus high-value loyal customers."),
        ("SH-05", "Product Manager\n(Product Owner)", 
         "1. Has no empirical data linking specific feature adoption or login inactivity to subsequent customer churn.\n2. Lacks visibility into how feature usage trajectories diverge between retained and churned accounts."),
        ("SH-06", "Support Team Lead\n(Service Owner)", 
         "1. Cannot prove to executive leadership how helpdesk resolution delays and ticket spikes impact customer cancellations.\n2. Lacks a unified interface connecting repeated support complaints directly to renewal outcomes."),
        ("SH-07", "Finance Manager\n(Financial Controller)", 
         "1. Obscured revenue impact: churn is reported only as account headcount rather than Monthly Recurring Revenue (MRR) lost.\n2. Lacks visibility into Net Revenue Retention (NRR) and cannot produce reliable cash-flow forecasts."),
        ("SH-08", "Data / BI Analyst\n(Technical Custodian)", 
         "1. Spends 3 to 4 working days every month manually extracting, cleaning, and VLOOKUP-merging fragile Excel dumps.\n2. Frequently faces conflicting stakeholder requests due to differing manual calculation methods across teams."),
        ("SH-09", "Ms. Anjali Pandey (Mentor)\n(Academic Guide)", 
         "1. Needs assurance that student project follows rigorous analytical, statistical, and software engineering methodologies.\n2. Requires transparent end-to-end traceability from operational questions to data models and test cases."),
        ("SH-10", "Sumedha (Student)\n(Developer & Analyst)", 
         "1. Must deliver a complete, professional full-stack data analytics application within semester time constraints.\n2. Needs a clear, unambiguous specification baseline so development and data generation remain strictly on track.")
    ]
    t_pain = doc.add_table(rows=len(pain_data) + 1, cols=3)
    format_table(t_pain, pain_headers, pain_data, [Inches(0.8), Inches(2.2), Inches(3.5)])

    # 2. How Might We Framework
    add_h1("2. How Might We (HMW) Design Thinking Statements")
    p_hmw_desc = doc.add_paragraph()
    p_hmw_desc.add_run("In design thinking, 'How Might We' (HMW) statements reframe identified pain points into forward-looking, constructive problem-framing opportunities. They strike an intentional balance: broad enough to invite creative data analytical solutions, yet specific enough to guide direct application functionality without prematurely dictating technical implementation details (avoiding anti-patterns like 'How might we build a dashboard?').")

    # 3. Master HMW Table
    add_h1("3. HMW Traceability Matrix")
    p_matrix_desc = doc.add_paragraph()
    p_matrix_desc.add_run("The following table establishes direct alignment between the HMW statements, their primary stakeholder champions, the underlying operational pain points, and the corresponding Operational Questions (OQs) formulated in Step 1.")

    hmw_headers = ["HMW ID", "How Might We ...?", "Stakeholder", "Pain Point Addressed", "Linked OQ"]
    hmw_data = [
        ("HMW-01", 
         "How might we provide leadership with an immediate, consistent view of overall retention and churn rates to eliminate reporting delays and conflicting numbers?", 
         "SH-01, SH-02, SH-08", 
         "Manual reporting delay (3-4 days) and conflicting churn formulas across teams.", 
         "OQ-01"),
        
        ("HMW-02", 
         "How might we help customer success managers uncover retention variations across plans, industries, sizes, and regions to deploy targeted segment interventions?", 
         "SH-02, SH-03", 
         "No segment visibility; unable to detect which client profiles or tiers struggle to retain.", 
         "OQ-02"),
        
        ("HMW-03", 
         "How might we enable product and success teams to track customer survival rates across onboarding cohorts over time to assess long-term customer longevity?", 
         "SH-02, SH-05", 
         "No cohort tracking; unable to tell if newer customer batches retain better than older ones.", 
         "OQ-03"),
        
        ("HMW-04", 
         "How might we quantify average customer tenure and Customer Lifetime Value (CLV) across segments so leadership can allocate retention investments where ROI is highest?", 
         "SH-01, SH-07", 
         "Customer value is treated equally without understanding lifetime financial worth per tier.", 
         "OQ-04"),
        
        ("HMW-05", 
         "How might we reveal which acquisition channels attract customers who stay the longest so marketing can shift spending toward sustainable, long-term growth?", 
         "SH-04", 
         "Marketing optimizes for raw acquisition volume rather than post-sale customer retention.", 
         "OQ-05"),
        
        ("HMW-06", 
         "How might we help support leads demonstrate the statistical impact of support ticket surges and resolution delays on customer cancellations?", 
         "SH-06", 
         "Support SLAs and repeated helpdesk tickets are not linked with empirical evidence to churn.", 
         "OQ-06"),
        
        ("HMW-07", 
         "How might we detect customer activity drop-offs and feature disengagement before customers make the final decision to cancel their subscriptions?", 
         "SH-05, SH-03", 
         "Telemetry usage data is unlinked to attrition; silent churn happens without prior notice.", 
         "OQ-07"),
        
        ("HMW-08", 
         "How might we connect customer satisfaction ratings (CSAT & NPS) directly to renewal outcomes to evaluate how sentiment predicts account retention?", 
         "SH-02, SH-03", 
         "Survey scores remain in isolated spreadsheets unconnected to customer renewal outcomes.", 
         "OQ-08"),
        
        ("HMW-09", 
         "How might we quantify the monthly recurring revenue (MRR) lost to churn and calculate Net Revenue Retention (NRR) to reflect true cash-flow stability?", 
         "SH-07, SH-01", 
         "Churn is tracked only as account headcount, hiding substantial monthly revenue losses.", 
         "OQ-09"),
        
        ("HMW-10", 
         "How might we equip customer success managers with a rule-based early-warning risk score to proactively intervene with vulnerable accounts before renewal deadlines pass?", 
         "SH-03, SH-02", 
         "Retention actions occur reactively after accounts have already churned or contracts lapsed.", 
         "OQ-10")
    ]
    t_hmw = doc.add_table(rows=len(hmw_data) + 1, cols=5)
    format_table(t_hmw, hmw_headers, hmw_data, [Inches(0.8), Inches(2.2), Inches(1.1), Inches(1.6), Inches(0.8)])

    # 4. Top 5 Priorities
    add_h1("4. Priority HMW Statements (Top 5 Selection)")
    p_prio_intro = doc.add_paragraph()
    p_prio_intro.add_run("To maximize business impact within the technical scope of our project, five core HMW statements are designated as PRIORITY. These prioritize proactive risk mitigation, executive clarity, and cohort-driven insights:")

    prio_headers = ["Priority Rank", "HMW ID", "HMW Statement Summary", "One-Line Strategic Justification"]
    prio_data = [
        ("PRIORITY 1", "HMW-10", "Automated Early-Warning Risk Scoring", 
         "Directly drives proactive intervention by converting customer success from reactive churn clean-up to timely, pre-renewal account rescues."),
        ("PRIORITY 2", "HMW-01", "Unified Baseline Retention & Churn Rates", 
         "Establishes the indispensable single source of truth for retention KPIs, eliminating multi-day manual reporting delays and conflicting formulas."),
        ("PRIORITY 3", "HMW-03", "Cohort Retention Tracking Over Time", 
         "Reveals true customer lifecycle decay curves across onboarding classes, showing whether product and onboarding changes improve retention."),
        ("PRIORITY 4", "HMW-09", "MRR Churn & Net Revenue Retention (NRR)", 
         "Translates raw customer counts into recurring financial impact, giving executive and finance leadership revenue-weighted business reality."),
        ("PRIORITY 5", "HMW-07", "Behavioral Disengagement & Usage Decline", 
         "Unlocks the earliest leading indicator of customer dissatisfaction—digital inactivity—weeks before survey complaints or formal cancellations occur.")
    ]
    t_prio = doc.add_table(rows=len(prio_data) + 1, cols=4)
    format_table(t_prio, prio_headers, prio_data, [Inches(1.0), Inches(0.8), Inches(2.0), Inches(2.7)])

    # 5. Coverage Check
    add_h1("5. Traceability & Answerability Check")
    p_cov_intro = doc.add_paragraph()
    p_cov_intro.add_run("Result: 100% Coverage Verified. ").bold = True
    p_cov_intro.add_run("Every single Operational Question (OQ-01 through OQ-10) established in Step 1 maps to at least one dedicated How Might We (HMW) statement. This guarantees that design thinking remains rigorously tethered to the foundational business inquiries.")

    cov_headers = ["OQ ID", "Operational Question Description", "Primary HMW", "Coverage Status"]
    cov_data = [
        ("OQ-01", "Overall monthly and quarterly churn and retention rates", "HMW-01", "Fully Covered"),
        ("OQ-02", "Retention breakdown across customer segments and regions", "HMW-02", "Fully Covered"),
        ("OQ-03", "Cohort retention curves after 1, 3, 6, and 12 months", "HMW-03", "Fully Covered"),
        ("OQ-04", "Average customer lifetime and Customer Lifetime Value (CLV)", "HMW-04", "Fully Covered"),
        ("OQ-05", "Acquisition channel longevity and retention comparison", "HMW-05", "Fully Covered"),
        ("OQ-06", "Support ticket volume and resolution SLA impact on churn", "HMW-06", "Fully Covered"),
        ("OQ-07", "Product usage and feature adoption trends prior to churn", "HMW-07", "Fully Covered"),
        ("OQ-08", "Customer satisfaction scores (CSAT, NPS) correlation with renewals", "HMW-08", "Fully Covered"),
        ("OQ-09", "Monthly recurring revenue (MRR) churn and Net Revenue Retention", "HMW-09", "Fully Covered"),
        ("OQ-10", "Identification and ranking of currently at-risk active customers", "HMW-10", "Fully Covered")
    ]
    t_cov = doc.add_table(rows=len(cov_data) + 1, cols=4)
    format_table(t_cov, cov_headers, cov_data, [Inches(0.8), Inches(3.2), Inches(1.2), Inches(1.3)])

    out_path = os.path.join("docs", "02_HMW_Statements.docx")
    doc.save(out_path)
    print(f"Successfully generated {out_path}")

if __name__ == "__main__":
    build_step2_document()
