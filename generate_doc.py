import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

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

def create_document():
    doc = docx.Document()
    
    # Page setup - 1 inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Base styling
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Palette
    NAVY = RGBColor(0x1B, 0x36, 0x5D)
    STEEL = RGBColor(0x2B, 0x4C, 0x7E)
    CHARCOAL = RGBColor(0x33, 0x33, 0x33)

    # Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run("PROJECT UNDERSTANDING")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = NAVY

    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Customer Retention Metrics Analysis")
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
        ("Deliverable", "Step 1: Project Understanding"),
        ("Version / date", "v1.0 | 18 November 2026")
    ]
    t_meta = doc.add_table(rows=len(meta_data), cols=2)
    t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_meta, color="CCCCCC", sz="4")
    col_widths = [Inches(1.8), Inches(4.7)]
    for i, (k, v) in enumerate(meta_data):
        row = t_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_margins(c0, 70, 70, 100, 100)
        set_cell_margins(c1, 70, 70, 100, 100)
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
            set_cell_margins(cell, 90, 90, 110, 110)
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
                set_cell_margins(cell, 80, 80, 110, 110)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                if j in align_right_cols:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r = p.add_run(str(val))
                r.font.size = Pt(9)
                if j == 0 and ("SH-" in str(val) or "OQ-" in str(val) or "OBJ-" in str(val) or "A-" in str(val) or "C-" in str(val) or "SC-" in str(val)):
                    r.font.bold = True
                    r.font.color.rgb = STEEL

    # 1. Project description
    add_h1("1. Project description")
    p = doc.add_paragraph()
    p.add_run("Customer Retention Metrics Analysis is a Python and Streamlit data analytics application that helps a company understand how many of its customers stay, how many leave, and why. The app reads customer, subscription, billing, support and usage data from a SQLite database, calculates standard retention metrics, runs simple statistical tests, and shows the results as interactive Plotly charts on a clean dashboard.")
    
    p2 = doc.add_paragraph()
    p2.add_run("In plain words: instead of someone opening five spreadsheets once a quarter and guessing, a manager opens one web page and sees the answer to questions such as “How many customers did we lose last month?”, “Which type of customer stays longest?” and “Who might leave next?”.")

    p3 = doc.add_paragraph()
    p3.add_run("For this project, an assumed organisation is used:\n")
    p3.add_run("• Organisation (assumed): ").bold = True
    p3.add_run("NoidaSoft Technologies Pvt. Ltd., a fictitious IT company in Noida with about 500 employees.\n")
    p3.add_run("• Business: ").bold = True
    p3.add_run("It sells subscription-based software products (for example, a billing tool and a helpdesk tool) and managed IT support plans to about 5,000 small and medium business customers across India.\n")
    p3.add_run("• Revenue model: ").bold = True
    p3.add_run("Customers pay monthly or annually and must renew. A customer who does not renew is a lost customer (churn).\n")
    p3.add_run("• Data: ").bold = True
    p3.add_run("All data is synthetic (computer-generated). No real person, company or personal data is used anywhere.")

    p4 = doc.add_paragraph()
    p4.add_run("The project will be built in 18 steps, from understanding the problem (this step) through requirements, database design, dataset generation, building and testing, deployment, the project document and the final presentation.")

    # 2. Background
    add_h1("2. Background: what happens today and why it is a problem")
    add_h2("2.1 What happens today")
    p_b1 = doc.add_paragraph()
    p_b1.add_run("• Customer data is spread across separate sources: the CRM export (customer details), the billing system (subscriptions and payments), the helpdesk tool (support tickets), the product logs (logins and feature use) and the survey tool (CSAT and NPS).\n"
                 "• Once a month, one analyst downloads these files into Excel, joins them manually and prepares a retention report for the leadership meeting.\n"
                 "• The report usually shows only one or two numbers (total customers and customers lost). It does not compare segments, cohorts or reasons.\n"
                 "• Customer success managers learn that a customer is leaving only when the renewal date has already passed.")

    add_h2("2.2 Why this is a problem")
    prob_headers = ["#", "Problem today", "Effect on the business"]
    prob_data = [
        ("1", "Manual, slow reporting (3 to 4 working days each month)", "Information is already outdated when managers see it."),
        ("2", "Different people calculate churn in different ways", "Numbers do not match between teams and trust in the data is low."),
        ("3", "No cohort or segment view", "The company cannot tell whether new customers are staying longer than old ones, or which plan or region is weak."),
        ("4", "Support, usage and satisfaction data are never linked to churn", "Likely causes of churn stay as opinions, not evidence."),
        ("5", "No early-warning view of at-risk customers", "Retention actions happen after the customer has left."),
        ("6", "Revenue impact of churn is not tracked clearly", "Leadership cannot see how much monthly revenue is being lost.")
    ]
    t_prob = doc.add_table(rows=len(prob_data) + 1, cols=3)
    format_table(t_prob, prob_headers, prob_data, [Inches(0.5), Inches(2.8), Inches(3.2)])

    p_b2 = doc.add_paragraph()
    p_b2.paragraph_format.space_before = Pt(6)
    p_b2.add_run("Industry experience widely shows that keeping an existing customer costs much less than winning a new one, so even a small improvement in retention has a large effect on revenue. This is why the company wants a reliable, repeatable way to measure retention.")

    # 3. Problem statement
    add_h1("3. Problem statement")
    p_ps = doc.add_paragraph()
    r_ps = p_ps.add_run("NoidaSoft Technologies cannot quickly and consistently measure how many customers it retains, which segments and behaviours drive churn, or which customers are at risk, because its customer, billing, support and usage data are scattered and analysed manually in spreadsheets.")
    r_ps.bold = True

    # 4. Intended audience / users
    add_h1("4. Intended audience / users")
    p_aud = doc.add_paragraph()
    p_aud.add_run("The app is a view-and-explore dashboard. Users do not enter data; they open the app, apply filters and read the charts.")
    
    aud_headers = ["User group", "Who they are", "How often they open the app", "What they do in the app"]
    aud_data = [
        ("Leadership", "CEO, COO, Head of Customer Success", "Monthly (before the review meeting)", "Read headline KPIs: retention rate, churn rate, revenue churn."),
        ("Customer Success team", "Customer success managers (about 15)", "Weekly", "Check at-risk customers, segment performance and cohort trends."),
        ("Product team", "Product managers", "Monthly", "Compare usage of retained and churned customers; spot weak features."),
        ("Support team leads", "Helpdesk managers", "Monthly", "See how tickets and resolution time relate to churn."),
        ("Data / BI analyst", "The analyst who builds today’s report", "Weekly", "Explore data, view statistical test results, export tables.")
    ]
    t_aud = doc.add_table(rows=len(aud_data) + 1, cols=4)
    format_table(t_aud, aud_headers, aud_data, [Inches(1.3), Inches(1.8), Inches(1.4), Inches(2.0)])

    # 5. Stakeholders
    add_h1("5. Stakeholders")
    sh_headers = ["SH ID", "Stakeholder", "Role", "Interest", "Influence (H/M/L)", "Decision they make with the app"]
    sh_data = [
        ("SH-01", "Chief Executive Officer", "Executive sponsor", "Healthy recurring revenue and growth", "H", "Approves company retention targets and budget."),
        ("SH-02", "Head of Customer Success", "Primary business owner", "Lower churn and higher renewals", "H", "Chooses which segments and customers to focus on each month."),
        ("SH-03", "Customer Success Managers", "Daily users", "Know which accounts need attention", "M", "Decide whom to call or visit before renewal."),
        ("SH-04", "Head of Sales / Marketing", "Acquisition owner", "Bring in customers who stay", "M", "Shift spending to channels with better retention."),
        ("SH-05", "Product Manager", "Product owner", "Features that keep customers engaged", "M", "Prioritise features and fixes based on usage gaps."),
        ("SH-06", "Support Team Lead", "Service owner", "Faster, better support", "M", "Adjust support staffing and response targets."),
        ("SH-07", "Finance Manager", "Revenue reporting", "Accurate revenue churn and forecasts", "M", "Update revenue forecast using churn trends."),
        ("SH-08", "Data / BI Analyst", "Builder and maintainer", "Less manual work, consistent numbers", "L", "Choose metric definitions and refresh the data."),
        ("SH-09", "Ms. Anjali Pandey (Mentor)", "Academic guide", "Correct method and quality of the project", "H", "Reviews and approves each project deliverable."),
        ("SH-10", "Sumedha (Student)", "Developer and analyst", "Complete a working, well-documented project", "H", "Makes design and build decisions within the project.")
    ]
    t_sh = doc.add_table(rows=len(sh_data) + 1, cols=6)
    format_table(t_sh, sh_headers, sh_data, [Inches(0.7), Inches(1.3), Inches(1.1), Inches(1.3), Inches(0.8), Inches(1.3)])
    p_sh_note = doc.add_paragraph()
    p_sh_note.paragraph_format.space_before = Pt(4)
    r_shn = p_sh_note.add_run("Note: SH-09 and SH-10 are academic stakeholders; SH-01 to SH-08 are the business stakeholders of the assumed company.")
    r_shn.font.italic = True
    r_shn.font.size = Pt(9)

    # 6. Operational questions
    add_h1("6. Operational questions the app must answer")
    p_oq = doc.add_paragraph()
    p_oq.add_run("Each question is written so that it can be answered with a number, a table or a chart from the data. The last two columns form the answerability check.")

    oq_headers = ["OQ ID", "Operational question", "Data needed", "Answered by (metric / chart / test)"]
    oq_data = [
        ("OQ-01", "What are the monthly and quarterly customer churn rate and retention rate overall?", "Customers; subscription start and end dates; status", "KPI cards; line chart of churn and retention rate by month"),
        ("OQ-02", "How does retention differ by segment (plan type, industry, company size, region)?", "Customer attributes; subscription status", "Grouped bar chart; segment table; chi-square test"),
        ("OQ-03", "How many customers of each signup-month cohort are still active after 1, 3, 6 and 12 months?", "Signup date; subscription end date", "Cohort retention heatmap"),
        ("OQ-04", "What is the average customer lifetime and customer lifetime value (CLV) by segment?", "Tenure in months; monthly fee; payments", "KPI table; bar chart of average tenure and CLV"),
        ("OQ-05", "Which acquisition channels bring the customers who stay longest?", "Acquisition channel; tenure; churn flag", "Bar chart; retention by channel; chi-square test"),
        ("OQ-06", "Do customers with more support tickets or slower resolution churn more?", "Tickets per customer; resolution hours; churn flag", "Box plot; correlation; Mann-Whitney or t-test"),
        ("OQ-07", "Do churned customers use the product less than retained customers?", "Monthly logins; feature-use count; churn flag", "Line chart of usage before churn; t-test"),
        ("OQ-08", "How do satisfaction scores (CSAT, NPS) relate to renewal?", "Survey scores; renewal outcome", "Bar chart by score band; correlation"),
        ("OQ-09", "How much monthly recurring revenue (MRR) is lost to churn, and what is net revenue retention?", "Monthly fee; start and end dates; plan changes", "MRR waterfall or line chart; NRR KPI card"),
        ("OQ-10", "Which active customers are currently at risk of leaving?", "Recent usage drop; open tickets; low survey score; renewal date", "Rule-based risk score; ranked at-risk table")
    ]
    t_oq = doc.add_table(rows=len(oq_data) + 1, cols=4)
    format_table(t_oq, oq_headers, oq_data, [Inches(0.8), Inches(2.2), Inches(1.8), Inches(1.7)])

    # 7. Objectives (SMART)
    add_h1("7. Objectives (SMART)")
    obj_headers = ["ID", "Objective", "Specific", "Measurable", "Achievable / Relevant", "Time-bound"]
    obj_data = [
        ("OBJ-1", "Create a clean synthetic database", "SQLite database with customers, subscriptions, payments, tickets, usage and surveys", "At least 5,000 customers and 24 months of history; 6 tables with keys and no orphan records", "Generated with Python; no real data needed", "By Step 10"),
        ("OBJ-2", "Calculate the core retention KPIs", "Churn rate, retention rate, cohort retention, tenure, CLV, MRR churn, NRR", "At least 8 KPIs; each matches an independent SQL/pandas recalculation exactly", "Standard formulas; pandas and SQL", "By Step 15"),
        ("OBJ-3", "Visualise retention insights", "Interactive Plotly charts and cohort heatmap with filters", "At least 10 charts covering all 10 OQs; filters for segment and date range", "Plotly and Streamlit are free", "By Step 15"),
        ("OBJ-4", "Identify churn drivers statistically", "Tests linking support, usage, satisfaction and channel to churn", "At least 3 tests (chi-square, t-test or Mann-Whitney, correlation) reported with p-value and plain-English meaning", "Uses SciPy; suitable for MCA level", "By Step 15"),
        ("OBJ-5", "Ensure quality through testing", "Pytest tests for calculations and data checks, plus manual UI test cases", "At least 25 test cases; all critical tests pass; at least 70% coverage of the calculation module", "Pytest is free and simple", "By Step 15"),
        ("OBJ-6", "Deploy and document", "Public app on Streamlit Community Cloud, project document and slides", "App loads in under 5 seconds; opens without errors; documents follow the ABES templates", "Free hosting; templates provided", "Deploy by Step 16; docs by Step 18")
    ]
    t_obj = doc.add_table(rows=len(obj_data) + 1, cols=6)
    format_table(t_obj, obj_headers, obj_data, [Inches(0.7), Inches(1.3), Inches(1.3), Inches(1.5), Inches(1.0), Inches(0.7)])

    # 8. Scope
    add_h1("8. Scope")
    scope_headers = ["In scope", "Out of scope"]
    scope_data = [
        ("Synthetic data of a subscription business (24 months, about 5,000 customers)", "Real customer data or any personal data"),
        ("SQLite database design and ER diagram", "Connecting to live CRM, billing or helpdesk systems"),
        ("Data cleaning and exploratory data analysis (EDA) in Python", "Real-time or automatic daily data refresh"),
        ("Retention KPIs: churn, retention, cohort, tenure, CLV, MRR churn, NRR", "Machine learning churn prediction models (a simple rule-based risk score is used instead)"),
        ("Statistical tests: chi-square, t-test or Mann-Whitney, correlation", "A/B testing of retention campaigns"),
        ("Interactive Plotly charts and cohort heatmap", "Sending emails or alerts to customers or staff"),
        ("Streamlit app with Tailwind-styled interface and filters", "User login, roles and permissions"),
        ("CSV export of at-risk customers and summary tables", "Mobile app version"),
        ("Pytest tests and manual test cases", "Load or security penetration testing"),
        ("Deployment on Streamlit Community Cloud", "Paid cloud services or custom domain"),
        ("Project Document (ABES Word template) and ABES slide deck", "Cost-benefit and marketing plans for the assumed company")
    ]
    t_scope = doc.add_table(rows=len(scope_data) + 1, cols=2)
    format_table(t_scope, scope_headers, scope_data, [Inches(3.25), Inches(3.25)])

    # 9. Assumptions and constraints
    add_h1("9. Assumptions and constraints")
    add_h2("9.1 Assumptions")
    assump_headers = ["ID", "Assumption"]
    assump_data = [
        ("A-01", "The organisation NoidaSoft Technologies Pvt. Ltd. is fictitious; it is used only to give the project a realistic business setting."),
        ("A-02", "All data is synthetic and generated with Python using fixed random seeds, so results can be repeated."),
        ("A-03", "The company is subscription-based; customers pay monthly or annually and renew or leave."),
        ("A-04", "Churn definition: a customer has churned in a month if their last subscription ended in that month and was not renewed within 30 days."),
        ("A-05", "The data covers 24 months (January 2024 to December 2025) and is a static snapshot, not a live feed."),
        ("A-06", "All amounts are in Indian Rupees (INR); there is no currency conversion."),
        ("A-07", "A customer is a company account, not an individual person; names are fictitious or coded (for example, CUST-00123)."),
        ("A-08", "The generated data will contain realistic patterns (for example, more tickets and lower usage before churn) so that the analysis has something to find; the patterns are built in deliberately and are not real findings."),
        ("A-09", "The app is read-only and is used by a small number of people at a time."),
        ("A-10", "A simple rule-based risk score is accepted by the mentor in place of a machine learning model.")
    ]
    t_assump = doc.add_table(rows=len(assump_data) + 1, cols=2)
    format_table(t_assump, assump_headers, assump_data, [Inches(0.8), Inches(5.7)])

    add_h2("9.2 Constraints")
    const_headers = ["ID", "Constraint"]
    const_data = [
        ("C-01", "Single student working alone; the project must be finished within the ABES semester schedule set by the mentor."),
        ("C-02", "Free tools only: Python, SQLite, Streamlit, Plotly, SciPy, pytest and Streamlit Community Cloud."),
        ("C-03", "Streamlit Community Cloud has limited memory and CPU, so the database must stay small (under about 50 MB) and queries must be efficient."),
        ("C-04", "SQLite is a single-file database; it suits this project but is not designed for many simultaneous writers."),
        ("C-05", "Streamlit has limited support for Tailwind CSS, so Tailwind styling will be added through custom HTML and CSS blocks inside the app rather than through a full front-end framework."),
        ("C-06", "No code is written before Step 10; Steps 1 to 9 produce documents and designs only."),
        ("C-07", "Final documents must follow the ABES Word template and slide deck format.")
    ]
    t_const = doc.add_table(rows=len(const_data) + 1, cols=2)
    format_table(t_const, const_headers, const_data, [Inches(0.8), Inches(5.7)])

    # 10. Success criteria
    add_h1("10. Success criteria (measurable)")
    sc_headers = ["ID", "Success criterion", "Target", "How it will be checked"]
    sc_data = [
        ("SC-01", "All operational questions answered", "10 of 10 OQs answered by a chart, KPI or test", "Walk through each OQ in the running app"),
        ("SC-02", "KPI accuracy", "100% match with independent SQL/pandas recalculation for all KPIs", "Pytest comparison tests"),
        ("SC-03", "Data quality", "0 duplicate IDs, 0 orphan foreign keys, 0 negative amounts in the dataset", "Data validation script and tests"),
        ("SC-04", "Dataset size", "At least 5,000 customers and 24 months of history", "Row counts from the database"),
        ("SC-05", "Statistical analysis", "At least 3 tests with p-value, effect description and plain-English conclusion", "Review of the statistics page"),
        ("SC-06", "Test results", "At least 25 test cases; all critical cases pass; at least 70% coverage of the calculation module", "Pytest and coverage report"),
        ("SC-07", "Performance", "Dashboard loads in under 5 seconds and filter changes respond in under 3 seconds", "Manual timing on the deployed app"),
        ("SC-08", "Deployment", "Public Streamlit URL works with no errors for a fresh visitor", "Open the link in a new browser"),
        ("SC-09", "Usability", "A new user finds any KPI within 3 clicks; at least 4 of 5 classmates rate the app clear (4 or above out of 5)", "Short feedback form with classmates"),
        ("SC-10", "Documentation", "Project Document and slide deck approved by the mentor", "Mentor review sign-off")
    ]
    t_sc = doc.add_table(rows=len(sc_data) + 1, cols=4)
    format_table(t_sc, sc_headers, sc_data, [Inches(0.8), Inches(2.0), Inches(2.0), Inches(1.7)])

    # 11. Check
    add_h1("11. Check: can every OQ be answered from data?")
    p_chk = doc.add_paragraph()
    p_chk.add_run("Result: Yes. ").bold = True
    p_chk.add_run("All 10 operational questions can be answered from data that the synthetic dataset will contain. The table below shows the planned data sources and which OQs depend on each. These are planned sources only; the exact tables and columns will be finalised in Step 8 (Database Design).")

    chk_headers = ["Planned data source", "Key fields (planned)", "Supports"]
    chk_data = [
        ("Customers", "Customer ID, signup date, plan type, industry, company size, region, acquisition channel", "OQ-02, OQ-03, OQ-05"),
        ("Subscriptions", "Subscription ID, customer ID, start date, end date, monthly fee, status, renewal flag", "OQ-01, OQ-03, OQ-04, OQ-09, OQ-10"),
        ("Payments", "Payment ID, customer ID, date, amount", "OQ-04, OQ-09"),
        ("Support tickets", "Ticket ID, customer ID, opened date, resolution hours, category", "OQ-06, OQ-10"),
        ("Product usage", "Customer ID, month, logins, feature-use count", "OQ-07, OQ-10"),
        ("Surveys", "Survey ID, customer ID, date, CSAT, NPS", "OQ-08, OQ-10")
    ]
    t_chk = doc.add_table(rows=len(chk_data) + 1, cols=3)
    format_table(t_chk, chk_headers, chk_data, [Inches(1.5), Inches(3.5), Inches(1.5)])

    p_watch = doc.add_paragraph()
    p_watch.paragraph_format.space_before = Pt(8)
    p_watch.add_run("Two points to watch in later steps:\n").bold = True
    p_watch.add_run("1. OQ-10 (at-risk customers): ").bold = True
    p_watch.add_run("This depends on a rule-based score that combines several fields. The rules and thresholds must be defined clearly in the BRD (Step 5) so the answer is repeatable and not a guess.\n")
    p_watch.add_run("2. OQ-04 (CLV) and OQ-09 (NRR): ").bold = True
    p_watch.add_run("These need agreed formulas. The formulas will be fixed in the BRD and PRD to avoid the “different people, different numbers” problem described in Section 2.")

    p_oq_m = doc.add_paragraph()
    p_oq_m.paragraph_format.space_before = Pt(6)
    p_oq_m.add_run("Open questions for the mentor (optional; an assumption is used if no reply is received):\n").bold = True
    p_oq_m.add_run("• OQ-M1: ").bold = True
    p_oq_m.add_run("Is a rule-based risk score acceptable in place of a prediction model? (Assumed yes: A-10.)\n")
    p_oq_m.add_run("• OQ-M2: ").bold = True
    p_oq_m.add_run("Is the assumed 24-month period and about 5,000 customers acceptable? (Assumed yes: A-05.)")

    os.makedirs("docs", exist_ok=True)
    out_path = os.path.join("docs", "01_Project_Understanding.docx")
    doc.save(out_path)
    print(f"Successfully generated {out_path}")

    # Also save as 01_Customer_Retention_Project_Understanding.docx
    alt_path = os.path.join("docs", "01_Customer_Retention_Project_Understanding.docx")
    doc.save(alt_path)
    print(f"Successfully generated {alt_path}")

if __name__ == "__main__":
    create_document()
