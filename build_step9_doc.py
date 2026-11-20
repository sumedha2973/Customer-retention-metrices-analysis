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

def build_step9_doc():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(5)

    NAVY = RGBColor(0x1B, 0x36, 0x5D)
    STEEL = RGBColor(0x2B, 0x4C, 0x7E)

    # Document Header
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run("ENTITY RELATIONSHIP (ER) DIAGRAM SPECIFICATION")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Customer Retention Metrics Analysis — Conceptual, Logical & Physical ER Model")
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(12)
    run_sub.font.color.rgb = STEEL

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(13)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(12.5)
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
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = STEEL
        return p

    def add_code_box(text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        cell.width = Inches(6.5)
        set_cell_background(cell, "F4F6F9")
        set_cell_margins(cell, 60, 60, 90, 90)
        set_table_borders(tbl, color="CBD5E1", sz="4")
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(text)
        r.font.name = 'Consolas'
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def format_table(table, headers, data, widths, align_center_cols=[]):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table, color="B0C4DE", sz="4")
        hdr_row = table.rows[0]
        for j, h in enumerate(headers):
            cell = hdr_row.cells[j]
            cell.width = widths[j]
            set_cell_background(cell, "1B365D")
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if j in align_center_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            r.font.size = Pt(8.5)
        
        for i, row_data in enumerate(data):
            row = table.rows[i + 1]
            bg_color = "F7F9FC" if i % 2 == 1 else "FFFFFF"
            for j, val in enumerate(row_data):
                cell = row.cells[j]
                cell.width = widths[j]
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, 50, 50, 80, 80)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                if j in align_center_cols:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(str(val))
                r.font.size = Pt(8)
                if j == 0 and ("REL-" in str(val) or "RULE-" in str(val) or "DP-" in str(val)):
                    r.font.bold = True
                    r.font.color.rgb = STEEL

    # Metadata Table
    meta_headers = ["Attribute", "Specification Details"]
    meta_data = [
        ("Project Name", "Customer Retention Metrics Analysis (CRMA)"),
        ("Model Type", "Physical Relational Entity Relationship Diagram (Crow's Foot 3NF)"),
        ("Storage Implementation", "SQLite 3.42+ (data/app.db)"),
        ("Student Author", "Sumedha (Roll No: 2503201450146, MCA 2nd Year, ABES Engineering College)"),
        ("Academic Mentor", "Ms. Anjali Pandey, Assistant Professor, ABES Engineering College"),
        ("Entity Scope", "5 Tables: 2 Master Dimensions + 3 Core Fact Tables"),
        ("Deliverables", "docs/09_ER_Diagram.docx and embedded high-resolution docs/09_ER_Diagram.png"),
        ("Version / Date", "v1.0 | 20 November 2026")
    ]
    t_meta = doc.add_table(rows=len(meta_data) + 1, cols=2)
    format_table(t_meta, meta_headers, meta_data, [Inches(2.2), Inches(4.3)], align_center_cols=[0])

    # 1. Graphical ER Diagram Image
    add_h1("1. Visual Entity Relationship Diagram")
    p_img_intro = doc.add_paragraph()
    p_img_intro.add_run("The high-resolution diagram below illustrates the physical database design, table schemas, primary/foreign keys, and relational cardinalities implemented in SQLite:")

    img_path = os.path.join("docs", "09_ER_Diagram.png")
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Inches(6.5))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(10)
        r_cap = p_cap.add_run("Figure 9.1: Physical Entity Relationship Diagram for Customer Retention Metrics Analysis (3NF Star-Schema)")
        r_cap.font.size = Pt(8.5)
        r_cap.font.italic = True
        r_cap.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    # 2. Mermaid erDiagram Code
    add_h1("2. Formal Mermaid erDiagram Specification")
    p_mm_intro = doc.add_paragraph()
    p_mm_intro.add_run("The exact Mermaid modeling code representing every table entity, data attribute, key constraint, and active verb-labeled relationship is provided below (renderable directly in Mermaid Live):")

    mermaid_code = (
        "erDiagram\n"
        "    dim_customers ||--o{ fact_subscriptions : \"holds (1:N)\"\n"
        "    dim_plans ||--o{ fact_subscriptions : \"governs (1:N)\"\n"
        "    dim_customers ||--o{ fact_support_tickets : \"logs (1:N)\"\n"
        "    dim_customers ||--o{ fact_monthly_usage : \"generates (1:N)\"\n\n"
        "    dim_customers {\n"
        "        TEXT customer_id PK \"CUST-XXXXX\"\n"
        "        TEXT company_name \"Corporate name\"\n"
        "        TEXT industry \"EdTech, FinTech, Healthcare, etc.\"\n"
        "        TEXT company_size \"Micro, Small, Medium\"\n"
        "        TEXT region \"North, South, East, West\"\n"
        "        TEXT acquisition_channel \"Organic, Referral, Paid, etc.\"\n"
        "        DATE signup_date \"Activation date\"\n"
        "        TEXT signup_cohort_month \"YYYY-MM cohort\"\n"
        "    }\n\n"
        "    dim_plans {\n"
        "        TEXT plan_id PK \"PLAN-01..03\"\n"
        "        TEXT plan_name UK \"Starter, Pro, Enterprise\"\n"
        "        INTEGER plan_tier_level \"1, 2, 3\"\n"
        "        REAL monthly_base_price \"INR list price\"\n"
        "        INTEGER max_user_seats \"User allowance\"\n"
        "        TEXT sla_support_level \"Standard, Priority, Dedicated\"\n"
        "    }\n\n"
        "    fact_subscriptions {\n"
        "        TEXT subscription_id PK \"SUB-XXXXXX\"\n"
        "        TEXT customer_id FK \"References dim_customers\"\n"
        "        TEXT plan_id FK \"References dim_plans\"\n"
        "        DATE start_date \"Term start date\"\n"
        "        DATE end_date \"Term end date\"\n"
        "        TEXT billing_cadence \"Monthly or Annual\"\n"
        "        REAL monthly_fee_inr \"Contracted MRR\"\n"
        "        TEXT subscription_status \"Active, Churned, Renewed\"\n"
        "        INTEGER is_churned \"0 or 1\"\n"
        "        DATE churn_date \"Null if active\"\n"
        "        INTEGER tenure_months \"Active lifespan\"\n"
        "    }\n\n"
        "    fact_support_tickets {\n"
        "        TEXT ticket_id PK \"TCK-XXXXXX\"\n"
        "        TEXT customer_id FK \"References dim_customers\"\n"
        "        DATETIME created_at \"Ticket opened\"\n"
        "        DATETIME resolved_at \"Ticket closed\"\n"
        "        REAL resolution_hours \"SLA turnaround hours\"\n"
        "        TEXT priority \"Low, Medium, High, Critical\"\n"
        "        TEXT category \"Billing, Bug, Feature, Tech\"\n"
        "        INTEGER ticket_csat \"1 to 5 rating\"\n"
        "    }\n\n"
        "    fact_monthly_usage {\n"
        "        TEXT usage_id PK \"USG-XXXXXX\"\n"
        "        TEXT customer_id FK \"References dim_customers\"\n"
        "        TEXT activity_month \"YYYY-MM\"\n"
        "        INTEGER login_days \"0 to 31\"\n"
        "        INTEGER core_feature_clicks \"Engagement count\"\n"
        "        INTEGER active_users_count \"Monthly active seats\"\n"
        "        INTEGER csat_score \"1 to 5 survey\"\n"
        "        INTEGER nps_score \"0 to 10 survey\"\n"
        "    }"
    )
    add_code_box(mermaid_code)

    # 3. Plain English Narrative
    add_h1("3. Relational Cardinality & Plain-English Business Narratives")
    p_rel_intro = doc.add_paragraph()
    p_rel_intro.add_run("Each relationship in the ER model represents an unambiguous business relationship governed by enterprise business rules:")

    rel_headers = ["Relationship ID", "Parent -> Child", "Cardinality", "Verb Label", "Plain-English Enterprise Description"]
    rel_data = [
        ("REL-01", "dim_customers -> fact_subscriptions", "1 : N (1 to 0..*)", "holds", 
         "One customer account holds zero, one, or many historical subscription contract terms over time (allowing upgrades, renewals, and cancellations)."),
        ("REL-02", "dim_plans -> fact_subscriptions", "1 : N (1 to 0..*)", "governs", 
         "One catalog subscription plan tier governs zero, one, or many active customer subscription contracts."),
        ("REL-03", "dim_customers -> fact_support_tickets", "1 : N (1 to 0..*)", "logs", 
         "One customer account logs zero, one, or many support helpdesk tickets over its relationship lifecycle."),
        ("REL-04", "dim_customers -> fact_monthly_usage", "1 : N (1 to 0..*)", "generates", 
         "One customer account generates zero, one, or many monthly telemetry and survey records (constrained to exactly one per calendar month).")
    ]
    t_rel = doc.add_table(rows=len(rel_data) + 1, cols=5)
    format_table(t_rel, rel_headers, rel_data, [Inches(0.9), Inches(1.8), Inches(1.1), Inches(0.8), Inches(2.9)], align_center_cols=[0, 2, 3])

    # 4. Business Rules Enforced
    add_h1("4. Business Rules Enforced by the Relational Architecture")
    p_rules_intro = doc.add_paragraph()
    p_rules_intro.add_run("The database design directly hardcodes the business rules established in Steps 1, 5, and 8 at the schema layer:")

    br_enforce_headers = ["Rule ID", "Governing Business Rule", "Technical Schema Enforcement Mechanism"]
    br_enforce_data = [
        ("RULE-01", "Standard Churn & Grace Period Lifecycle", 
         "Contracts track start_date, end_date, and is_churned. Accounts failing to renew within 30 days of end_date transition to subscription_status = 'Churned' with an explicit churn_date populated."),
        ("RULE-02", "Referential Integrity & Commercial Auditability", 
         "Deleting an account with active or historical financial contracts is blocked via ON DELETE RESTRICT on REL-01, preserving financial ledgers for statutory audit."),
        ("RULE-03", "Support Ticket SLA Attribution", 
         "Tickets record created_at and resolved_at timestamps, with resolution_hours constrained to >= 0.0 and priority restricted to ('Low', 'Medium', 'High', 'Critical')."),
        ("RULE-04", "Periodic Telemetry Non-Duplication", 
         "The composite constraint UNIQUE(customer_id, activity_month) in fact_monthly_usage physically prevents double-counting monthly logins or multiple survey entries for the same billing cycle."),
        ("RULE-05", "Customer Lifetime Value & Revenue Integrity", 
         "monthly_fee_inr and monthly_base_price enforce CHECK(>= 0.0), ensuring negative pricing anomalies cannot distort CLV or MRR waterfall calculations.")
    ]
    t_br_enf = doc.add_table(rows=len(br_enforce_data) + 1, cols=3)
    format_table(t_br_enf, br_enforce_headers, br_enforce_data, [Inches(1.0), Inches(2.2), Inches(3.3)], align_center_cols=[0])

    # 5. Design Problems Identified & Resolved
    add_h1("5. Design Problems Identified, Evaluated, and Resolved")
    p_dp_intro = doc.add_paragraph()
    p_dp_intro.add_run("During the structural modeling phase, four classical relational design challenges were analyzed and systematically resolved to guarantee high performance and 3NF normalization:")

    dp_headers = ["Problem ID", "Potential Design Anti-Pattern", "Architectural Risk / Consequence", "Implemented Architectural Fix"]
    dp_data = [
        ("DP-01", "Many-to-Many Relationship between Customers and Plans", 
         "Customers switch, upgrade, or downgrade plans over time. Direct M:N modeling loses contract start/end dates and billing rates.", 
         "Resolved: Introduced fact_subscriptions as an associative contract fact table, capturing historical plan transitions, specific contracted MRR, and renewal flags."),
        ("DP-02", "Granularity Mismatch (Event Logs vs. Monthly Analysis)", 
         "Support tickets occur continuously, whereas churn is analyzed monthly. Joining raw tickets directly to subscriptions creates Cartesian product row explosions.", 
         "Resolved: Preserved fact_support_tickets at atomic ticket grain, while the analytical view (vw_customer_retention_analytics) pre-aggregates total tickets and average SLA turnaround."),
        ("DP-03", "Survey Feedback Separation & Bias Mitigation", 
         "Conflating transactional post-ticket feedback with overall quarterly sentiment distorts NPS analysis.", 
         "Resolved: Segregated ticket_csat (transactional support rating in fact_support_tickets) from longitudinal csat_score and nps_score (recorded in fact_monthly_usage)."),
        ("DP-04", "Orphan Records & Cascade Deletion Risks", 
         "Careless CASCADE deletes on customer accounts could inadvertently wipe audit-mandated financial revenue ledgers.", 
         "Resolved: Fine-grained referential actions: ON DELETE RESTRICT on financial contracts (fact_subscriptions), but ON DELETE CASCADE on operational logs (tickets and usage).")
    ]
    t_dp = doc.add_table(rows=len(dp_data) + 1, cols=4)
    format_table(t_dp, dp_headers, dp_data, [Inches(0.8), Inches(1.8), Inches(1.9), Inches(2.0)], align_center_cols=[0])

    os.makedirs("docs", exist_ok=True)
    out_path = os.path.join("docs", "09_ER_Diagram.docx")
    doc.save(out_path)
    print(f"Successfully generated {out_path}")

if __name__ == "__main__":
    build_step9_doc()
