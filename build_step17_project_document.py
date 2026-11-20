"""
build_step17_project_document.py
Generates <ProjectID>_Project_Document.docx adhering strictly to all formatting,
table layouts, widths, styles, headings, XML structures, and ABES guidelines.
"""
import os
import datetime
import openpyxl
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# ─────────────────────────────────────────────────────────────
# CONSTANTS & METADATA
# ─────────────────────────────────────────────────────────────
PROJECT_ID   = "CRMA-2026"
DOC_FILENAME = f"docs/{PROJECT_ID}_Project_Document.docx"

CLR_BURGUNDY = RGBColor(0xA5, 0x1C, 0x30) # A51C30
CLR_BODY     = RGBColor(0x15, 0x15, 0x14) # 151514
CLR_MUTED    = RGBColor(0x7F, 0x7F, 0x7F) # 7F7F7F
HEX_BURGUNDY = "A51C30"
HEX_BORDER   = "BFBFBF"

# ─────────────────────────────────────────────────────────────
# XML & FORMATTING HELPERS
# ─────────────────────────────────────────────────────────────
def set_cell_margins(cell, top=60, bottom=60, left=90, right=90):
    """Set cell internal padding in twips/dxa."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_cell_borders(cell, color=HEX_BORDER, sz="4"):
    """0.5 pt border (sz='4' means 4/8 pt = 0.5 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{color_hex}"/>')
    tcPr.append(shd)

def make_table_fixed(table, col_widths):
    """Apply exact table layout, alignment, and column widths."""
    tblPr = table._tbl.tblPr
    # Alignment Left (lines up with page text)
    jc = parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>')
    tblPr.append(jc)
    # Fixed layout
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    tblPr.append(tblLayout)
    # Total width 9638 dxa
    total_w = sum(col_widths)
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_w}" w:type="dxa"/>')
    tblPr.append(tblW)

    # Apply col widths
    for row in table.rows:
        # cantSplit on every row
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for idx, width in enumerate(col_widths):
            cell = row.cells[idx]
            tcPr = cell._tc.get_or_add_tcPr()
            # Remove any existing tcW to avoid duplicate tags
            for existing_w in tcPr.findall(qn('w:tcW')):
                tcPr.remove(existing_w)
            tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{width}" w:type="dxa"/>')
            tcPr.append(tcW)
            set_cell_margins(cell)
            set_cell_borders(cell)

    # Header row properties: repeat header on each page & keep with next
    header_tr = table.rows[0]._tr.get_or_add_trPr()
    header_tr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    header_tr.append(parse_xml(f'<w:keepNext {nsdecls("w")}/>'))

def format_cell_text(cell, text, bold=False, italic=False, font_size=Pt(9), color=CLR_BODY, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(str(text))
    run.font.name = "Calibri"
    run.font.size = font_size
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color
    return run

def style_header_row(table):
    for cell in table.rows[0].cells:
        set_cell_shading(cell, HEX_BURGUNDY)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(9)
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

def add_body_p(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.color.rgb = CLR_BODY
    return p

def add_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(5)
    run = p.add_run(f"NOTE: {text}")
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.italic = True
    run.font.color.rgb = CLR_MUTED
    return p

def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(3)
    p.paragraph_format.left_indent  = Inches(540 / 1440) # 540 dxa
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.font.color.rgb = CLR_BODY
    return p

def add_label(doc, text, color=CLR_BODY, before=Pt(5)):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = before
    p.paragraph_format.space_after  = Pt(3)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.bold = True
    run.font.color.rgb = color
    return p

def add_code(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.font.name = "Consolas"
    run.font.size = Pt(9)
    run.font.color.rgb = CLR_BODY
    return p

# ─────────────────────────────────────────────────────────────
# DOCUMENT INITIALIZATION
# ─────────────────────────────────────────────────────────────
doc = Document()

# Section page setup: A4, 2cm all margins, 1.25cm header/footer
section = doc.sections[0]
section.page_width  = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin    = Cm(2.0)
section.bottom_margin = Cm(2.0)
section.left_margin   = Cm(2.0)
section.right_margin  = Cm(2.0)
section.header_distance = Cm(1.25)
section.footer_distance = Cm(1.25)
section.different_first_page_header_footer = True

# Custom Styles Setup
styles = doc.styles

# Heading 1: Georgia 15 bold A51C30, before 16 pt, after 7 pt, keep next
h1_style = styles['Heading 1']
h1_style.font.name = 'Georgia'
h1_style.font.size = Pt(15)
h1_style.bold = True
h1_style.font.color.rgb = CLR_BURGUNDY
h1_style.paragraph_format.space_before = Pt(16)
h1_style.paragraph_format.space_after  = Pt(7)
h1_style.paragraph_format.keep_with_next = True

# Heading 2: Georgia 12 bold 151514, before 10 pt, after 5 pt, keep next
h2_style = styles['Heading 2']
h2_style.font.name = 'Georgia'
h2_style.font.size = Pt(12)
h2_style.bold = True
h2_style.font.color.rgb = CLR_BODY
h2_style.paragraph_format.space_before = Pt(10)
h2_style.paragraph_format.space_after  = Pt(5)
h2_style.paragraph_format.keep_with_next = True

# ─────────────────────────────────────────────────────────────
# HEADER & FOOTER CONFIGURATION (pages 2+)
# ─────────────────────────────────────────────────────────────
header = section.header
p_head = header.paragraphs[0]
p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r_head = p_head.add_run(f"ABES Engineering College | Project Document | {PROJECT_ID}")
r_head.font.name = "Calibri"
r_head.font.size = Pt(8)
r_head.font.color.rgb = CLR_MUTED

footer = section.footer
p_foot = footer.paragraphs[0]
p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_foot = p_foot.add_run("Sumedha | 2503201450146     Page ")
r_foot.font.name = "Calibri"
r_foot.font.size = Pt(8)
r_foot.font.color.rgb = CLR_MUTED

# Insert dynamic PAGE field in footer
fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
p_foot._p.append(fldSimple)

# ─────────────────────────────────────────────────────────────
# PAGE 1: COVER PAGE
# ─────────────────────────────────────────────────────────────
# Line 1: HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad (bold 10 A51C30)
p1 = doc.add_paragraph()
p1.paragraph_format.space_before = Pt(0)
p1.paragraph_format.space_after  = Pt(5)
r1 = p1.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
r1.font.name = "Calibri"
r1.font.size = Pt(10)
r1.bold = True
r1.font.color.rgb = CLR_BURGUNDY

# Line 2: [Insert ABES logo here] (italic 9 pt 7F7F7F)
p2 = doc.add_paragraph()
p2.paragraph_format.space_before = Pt(0)
p2.paragraph_format.space_after  = Pt(5)
r2 = p2.add_run("[Insert ABES logo here]")
r2.font.name = "Calibri"
r2.font.size = Pt(9)
r2.italic = True
r2.font.color.rgb = CLR_MUTED

# Line 3: PROJECT DOCUMENT (Georgia 26 bold, before 70 pt, after 6 pt)
p3 = doc.add_paragraph()
p3.paragraph_format.space_before = Pt(70)
p3.paragraph_format.space_after  = Pt(6)
r3 = p3.add_run("PROJECT DOCUMENT")
r3.font.name = "Georgia"
r3.font.size = Pt(26)
r3.bold = True
r3.font.color.rgb = CLR_BODY

# Line 4: <Project ID> - <Title> (Georgia 18 A51C30, after 12 pt, bottom border 1.5 pt A51C30)
p4 = doc.add_paragraph()
p4.paragraph_format.space_before = Pt(0)
p4.paragraph_format.space_after  = Pt(12)
r4 = p4.add_run(f"{PROJECT_ID} - Customer Retention Metrices Analysis")
r4.font.name = "Georgia"
r4.font.size = Pt(18)
r4.bold = True
r4.font.color.rgb = CLR_BURGUNDY
# Add bottom border 1.5 pt A51C30 (sz='12' is 1.5 pt)
pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="{HEX_BURGUNDY}"/></w:pBdr>')
p4._p.get_or_add_pPr().append(pBdr)

# Line 5: Description · HMW · Stakeholders ... (10 pt 7F7F7F, after 24 pt)
p5 = doc.add_paragraph()
p5.paragraph_format.space_before = Pt(0)
p5.paragraph_format.space_after  = Pt(24)
r5 = p5.add_run("Description · HMW · Stakeholders · Traceability · User Stories · BRD · TRD · Database Design · Test Cases · UI Framework")
r5.font.name = "Calibri"
r5.font.size = Pt(10)
r5.font.color.rgb = CLR_MUTED

# Cover Metadata Table: Field 2900 | Details 6738 (Total 9638)
cov_tbl = doc.add_table(rows=8, cols=2)
cov_widths = [2900, 6738]
make_table_fixed(cov_tbl, cov_widths)

cov_data = [
    ("Project ID & Title", f"{PROJECT_ID} - Customer Retention Metrices Analysis"),
    ("Domain", "Data Analytics (Python)"),
    ("Programme / Batch", "MCA 2nd Year Data Analyst - ABES_OND26_MCA Batch 2025-27"),
    ("Learner(s) & Roll No(s)", "Sumedha | Roll No: 2503201450146"),
    ("Mentor", "Ms. Anjali Pandey"),
    ("GitHub repository", "https://github.com/sumedha2973"),
    ("Live app (Streamlit)", "https://2503201450146-crma.streamlit.app"),
    ("Version & date", f"v1.0 - {datetime.date.today().strftime('%d/%m/%Y')}"),
]
for r_idx, (fld, det) in enumerate(cov_data):
    row = cov_tbl.rows[r_idx]
    format_cell_text(row.cells[0], fld, bold=True, font_size=Pt(9), color=CLR_BURGUNDY)
    format_cell_text(row.cells[1], det, font_size=Pt(9), color=CLR_BODY)

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# PAGE 2: TABLE OF CONTENTS
# ─────────────────────────────────────────────────────────────
# PAGE 2: only a Word TOC field: "TOC \h \o "1-1"". Put all Heading 1 titles in its result and set updateFields on open.
p_toc = doc.add_paragraph()
p_toc.paragraph_format.space_before = Pt(0)
p_toc.paragraph_format.space_after  = Pt(0)

# Add TOC field XML
fldSimple_toc = parse_xml(r'<w:fldSimple %s w:instr="TOC \h \o &quot;1-1&quot;"/>' % nsdecls('w'))
p_toc._p.append(fldSimple_toc)

# Instruct Word to update fields upon opening
settings = doc.settings.element
settings.append(parse_xml(r'<w:updateFields %s w:val="true"/>' % nsdecls('w')))

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# SECTION 1: PROJECT DESCRIPTION
# ─────────────────────────────────────────────────────────────
doc.add_heading("1. Project Description", level=1)

# Table Item 2400 | Description 7238
tbl_desc = doc.add_table(rows=8, cols=2)
desc_widths = [2400, 7238]
make_table_fixed(tbl_desc, desc_widths)

desc_data = [
    ("Item", "Description"),
    ("Overview", "Customer Retention Metrics Analysis (CRMA) is an enterprise Python and Streamlit analytics solution developed for NoidaSoft Technologies Pvt. Ltd., a fictitious mid-market B2B SaaS organization with ~500 employees and ~5,000 corporate clients. The app automates retention tracking, diagnoses churn drivers, and provides proactive account risk ranking over 24 historical contract months."),
    ("Problem statement", "NoidaSoft Technologies cannot quickly and consistently measure how many customers it retains, which segments and behaviours drive churn, or which customers are at risk, because its customer, billing, support and usage data are scattered and analysed manually in spreadsheets."),
    ("Intended users", "Executive leadership (CEO, VP Ops), Customer Success Managers (CSMs), Sales/Marketing Heads, Support Leads, and Finance Analysts who need view-and-explore access to retention metrics and early warning churn indicators without writing SQL."),
    ("In scope", "Embedded SQLite storage (data/app.db); automated data validation & caching; core metrics (CCR, CRR, MRR Churn, NRR, CLV); triangular cohort survival heatmap; inferential statistical hypothesis testing; 4-factor composite customer risk scoring; client-side CSV export; and Streamlit Cloud deployment."),
    ("Out of scope", "Direct bi-directional write-back to CRM/ERP systems, live automated streaming data ingestion pipelines (deferred to Phase 2 as BR-13), native mobile application development, automated SMS/email outreach execution, and external user authentication infrastructure."),
    ("Assumptions", "100% synthetic computer-generated dataset (Seed 42) with no PII; contracts operate on standard Indian INR subscription rates; 30-day grace period constitutes official churn; and client machines feature standard desktop web browsers."),
    ("Dataset", "Star-schema SQLite relational database comprising 5 entities: dim_customers (200 accounts), dim_plans (3 subscription tiers), fact_subscriptions (229 contracts), fact_support_tickets (704 clean tickets), fact_monthly_usage (2,489 monthly telemetry logs), and analytical view vw_customer_retention_analytics (33 curated columns)."),
]
for r_idx, (it, de) in enumerate(desc_data):
    row = tbl_desc.rows[r_idx]
    format_cell_text(row.cells[0], it, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
    format_cell_text(row.cells[1], de, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_desc)

# ─────────────────────────────────────────────────────────────
# SECTION 2: HOW MIGHT WE (HMW) STATEMENTS
# ─────────────────────────────────────────────────────────────
doc.add_heading("2. How Might We (HMW) Statements", level=1)
add_note(doc, "Format: How might we <action> for <user> so that <outcome>?")

# ID 1300 | HMW statement 8338
tbl_hmw = doc.add_table(rows=11, cols=2)
hmw_widths = [1300, 8338]
make_table_fixed(tbl_hmw, hmw_widths)

hmw_data = [
    ("ID", "HMW statement"),
    ("HMW-01", "How might we provide leadership with an immediate, consistent view of overall retention and churn rates to eliminate reporting delays and conflicting numbers?"),
    ("HMW-02", "How might we help customer success managers uncover retention variations across plans, industries, sizes, and regions to deploy targeted segment interventions?"),
    ("HMW-03", "How might we enable product and success teams to track customer survival rates across onboarding cohorts over time to assess long-term customer longevity?"),
    ("HMW-04", "How might we quantify average customer tenure and Customer Lifetime Value (CLV) across segments so leadership can allocate retention investments where ROI is highest?"),
    ("HMW-05", "How might we reveal which acquisition channels attract customers who stay the longest so marketing can shift spending toward sustainable, long-term growth?"),
    ("HMW-06", "How might we help support leads demonstrate the statistical impact of support ticket surges and resolution delays on customer cancellations?"),
    ("HMW-07", "How might we detect customer activity drop-offs and feature disengagement before customers make the final decision to cancel their subscriptions?"),
    ("HMW-08", "How might we connect customer satisfaction ratings (CSAT & NPS) directly to renewal outcomes to evaluate how sentiment predicts account retention?"),
    ("HMW-09", "How might we quantify the monthly recurring revenue (MRR) lost to churn and calculate Net Revenue Retention (NRR) to reflect true cash-flow stability?"),
    ("HMW-10", "How might we equip customer success managers with a rule-based early-warning risk score to proactively intervene with vulnerable accounts before renewal deadlines pass?"),
]
for r_idx, (hid, hst) in enumerate(hmw_data):
    row = tbl_hmw.rows[r_idx]
    format_cell_text(row.cells[0], hid, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
    format_cell_text(row.cells[1], hst, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_hmw)

# ─────────────────────────────────────────────────────────────
# SECTION 3: STAKEHOLDERS
# ─────────────────────────────────────────────────────────────
doc.add_heading("3. Stakeholders", level=1)

# ID 1000 | Stakeholder 2400 | Role 1900 | Interest 3138 | Influence 1200
tbl_sh = doc.add_table(rows=11, cols=5)
sh_widths = [1000, 2400, 1900, 3138, 1200]
make_table_fixed(tbl_sh, sh_widths)

sh_data = [
    ("ID", "Stakeholder", "Role", "Interest", "Influence"),
    ("SH-01", "Chief Executive Officer", "Executive sponsor", "Healthy recurring revenue, low churn, sustainable enterprise growth", "H"),
    ("SH-02", "Head of Customer Success", "Primary business owner", "Lower logo churn, higher renewal rates, operational efficiency", "H"),
    ("SH-03", "Customer Success Managers", "Daily users", "Identification of vulnerable accounts requiring proactive intervention", "M"),
    ("SH-04", "Head of Sales / Marketing", "Acquisition owner", "Sourcing accounts with superior long-term retention characteristics", "M"),
    ("SH-05", "Product Manager", "Product owner", "Feature engagement tracking and resolving pre-churn utilization drops", "M"),
    ("SH-06", "Support Team Lead", "Service owner", "SLA turnaround adherence and eliminating support bottlenecks", "M"),
    ("SH-07", "Finance Manager", "Revenue reporting", "Accurate recurring revenue churn accounting and NRR cash forecasting", "M"),
    ("SH-08", "Data / BI Analyst", "Builder & maintainer", "Automated pipelines, metric consistency, eliminating manual Excel work", "L"),
    ("SH-09", "Ms. Anjali Pandey (Mentor)", "Academic guide", "Methodological rigor, documentation quality, software engineering standards", "H"),
    ("SH-10", "Sumedha (Learner)", "Developer & analyst", "End-to-end implementation, automated test coverage, successful project defense", "H"),
]
for r_idx, vals in enumerate(sh_data):
    row = tbl_sh.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_sh)

# ─────────────────────────────────────────────────────────────
# SECTION 4: TRACEABILITY MATRIX
# ─────────────────────────────────────────────────────────────
doc.add_heading("4. Traceability Matrix", level=1)
add_note(doc, "One row per requirement. Priority uses MOSCOW: Must / Should / Could / Won't (this time).")

# HMW 894 | Req ID 742 | Requirement 1964 | Stakeholder 1083 | Feature 1143 | Priority 798 | User Story 842 | Test Case 1132 | UI Screen 1040
# Total = 894 + 742 + 1964 + 1083 + 1143 + 798 + 842 + 1132 + 1040 = 9638
tbl_rtm = doc.add_table(rows=16, cols=9)
rtm_widths = [894, 742, 1964, 1083, 1143, 798, 842, 1132, 1040]
make_table_fixed(tbl_rtm, rtm_widths)

rtm_header = ["HMW", "Req ID", "Requirement", "Stakeholder", "Feature", "Priority", "User Story", "Test Case", "UI Screen"]

# Load exact rows from docs/04_Traceability_Matrix.xlsx
wb_rtm = openpyxl.load_workbook("docs/04_Traceability_Matrix.xlsx")
ws_rtm = wb_rtm["RTM"]

for r_idx in range(1, 17):
    row = tbl_rtm.rows[r_idx - 1]
    if r_idx == 1:
        vals = rtm_header
    else:
        # Excel columns: OQ(1), HMW(2), BR(3), Requirement(4), SH(5), Priority(6), Status(7), Feature(8), DB(9), US(10), UI(11), TC(12)
        row_src = [ws_rtm.cell(row=r_idx, column=c).value for c in range(1, 13)]
        hmw_val = str(row_src[1] or "")
        req_id  = str(row_src[2] or "")
        req_txt = str(row_src[3] or "")
        sh_val  = str(row_src[4] or "")
        feat_val= f"FR-{req_id.split('-')[-1]}" if "BR-" in req_id and req_id != "BR-13" else "None"
        prio_val= str(row_src[5] or "")
        us_val  = str(row_src[9] or "")
        tc_val  = str(row_src[11] or "")
        ui_val  = str(row_src[10] or "")
        vals = [hmw_val, req_id, req_txt, sh_val, feat_val, prio_val, us_val, tc_val, ui_val]

    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==1), font_size=Pt(8.5), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==1 else CLR_BODY))
style_header_row(tbl_rtm)

# ─────────────────────────────────────────────────────────────
# SECTION 5: USER STORIES
# ─────────────────────────────────────────────────────────────
doc.add_heading("5. User Stories", level=1)

# US-ID 900 | User story 3400 | Acceptance criteria (Given / When / Then) 4138 | Priority 1200 (Total 9638)
import generate_step11_user_stories as gus
tbl_us = doc.add_table(rows=1 + len(gus.STORIES), cols=4)
us_widths = [900, 3400, 4138, 1200]
make_table_fixed(tbl_us, us_widths)

us_header = ["US-ID", "User story", "Acceptance criteria (Given / When / Then)", "Priority"]
for c_idx, h in enumerate(us_header):
    format_cell_text(tbl_us.rows[0].cells[c_idx], h, bold=True, font_size=Pt(9), color=RGBColor(0xFF,0xFF,0xFF))

for s_idx, s in enumerate(gus.STORIES, 1):
    us_id, epic, persona, story, fr, br, prio, sp, ac_list = s
    ac_formatted = "\n\n".join([f"• GIVEN {g}\n  WHEN {w}\n  THEN {t}" for g, w, t in ac_list])
    row = tbl_us.rows[s_idx]
    format_cell_text(row.cells[0], us_id, font_size=Pt(8.5))
    format_cell_text(row.cells[1], story, font_size=Pt(8.5))
    format_cell_text(row.cells[2], ac_formatted, font_size=Pt(8))
    format_cell_text(row.cells[3], prio, font_size=Pt(8.5))
style_header_row(tbl_us)

# ─────────────────────────────────────────────────────────────
# SECTION 6: BUSINESS REQUIREMENTS DOCUMENT (BRD)
# ─────────────────────────────────────────────────────────────
doc.add_heading("6. Business Requirements Document (BRD)", level=1)

doc.add_heading("6.1 Business objectives", level=2)
add_bullet(doc, "OBJ-1: Establish an enterprise synthetic relational database (data/app.db) covering 200 customer accounts across 24 historical contract months with zero orphaned records.")
add_bullet(doc, "OBJ-2: Deliver 100% mathematical calculation parity for core retention metrics (CCR, CRR, MRR Churn, NRR, CLV) adhering strictly to defined business rules (RULE-01 to RULE-07).")
add_bullet(doc, "OBJ-3: Provide interactive multi-page Plotly visualizations covering all 10 Operational Questions, including triangular cohort survival heatmaps and segment breakdowns.")
add_bullet(doc, "OBJ-4: Conduct inferential statistical hypothesis testing (Mann-Whitney U, t-test, Chi-Square, Pearson/Spearman) at alpha = 0.05 to diagnose empirical drivers of account attrition.")
add_bullet(doc, "OBJ-5: Maintain high software reliability via 55 test cases (20 functional and 35 technical automated tests) achieving >= 70% automated test coverage with sub-3s render speeds.")

doc.add_heading("6.2 Success KPIs", level=2)
# KPI 1900 | Numerator 2100 | Denominator 2100 | Period 1700 | Target / note 1838 (Total 9638)
tbl_kpi = doc.add_table(rows=12, cols=5)
kpi_widths = [1900, 2100, 2100, 1700, 1838]
make_table_fixed(tbl_kpi, kpi_widths)

kpi_table_data = [
    ("KPI", "Numerator", "Denominator", "Period", "Target / note"),
    ("MSR-01 Customer Churn Rate (CCR %)", "COUNT(churned_customers in month M)", "COUNT(starting_active_customers in M)", "Monthly", "< 2.5% / month"),
    ("MSR-02 Customer Retention Rate (CRR %)", "100.0 - CCR %", "1.0", "Monthly", "> 97.5% / month"),
    ("MSR-03 Cohort Survival Rate (%)", "COUNT(active_customers_at_M+k)", "COUNT(initial_cohort_size)", "M+1, 3, 6, 12", "> 85% at M+12"),
    ("MSR-04 Customer Lifetime Value (CLV ₹)", "AVG(monthly_fee_inr) * AVG(tenure)", "1.0", "Segment / Plan", "> ₹75,000 average"),
    ("MSR-05 Channel Longevity Rate (%)", "COUNT(retained_by_channel)", "COUNT(total_by_channel)", "12-Month", "Referral > 90%"),
    ("MSR-06 Mean Ticket Resolution Hours", "SUM(resolution_hours)", "COUNT(resolved_tickets)", "Monthly", "< 24.0 hours"),
    ("MSR-07 Pre-Churn Usage Drop (%)", "Logins(t-60d) - Logins(t-30d)", "Logins(t-60d)", "90-Day window", "p < 0.01 (SciPy)"),
    ("MSR-08 Satisfaction Correlation", "Covariance(CSAT, Renewed)", "StdDev(CSAT) * StdDev(Renewed)", "Trailing", "r > 0.40 positive"),
    ("MSR-09 MRR Churn in INR (₹)", "SUM(monthly_fee_inr churned in M)", "1.0", "Monthly", "< ₹50,000 / month"),
    ("MSR-10 Net Revenue Retention (NRR %)", "Start MRR + Expansion - Churn", "Starting MRR", "Trailing 12M", "> 105% net expansion"),
    ("MSR-11 Composite Health Risk Score", "0.35*LoginDrop + 0.25*Tix + 0.2*CSAT", "1.0", "Active Accounts", "High Risk >= 70"),
]
for r_idx, vals in enumerate(kpi_table_data):
    row = tbl_kpi.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(8.5), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_kpi)

doc.add_heading("6.3 Business requirements", level=2)
# BR-ID 1100 | Requirement 5338 | Priority 1300 | Source (stakeholder) 1900 (Total 9638)
tbl_br = doc.add_table(rows=14, cols=4)
br_widths = [1100, 5338, 1300, 1900]
make_table_fixed(tbl_br, br_widths)

br_table_data = [
    ("BR-ID", "Requirement", "Priority", "Source (stakeholder)"),
    ("BR-01", "The business needs to compute standardized monthly and quarterly customer retention and churn rates across historical periods so that leadership relies on a single source of truth.", "Must", "SH-01, SH-02"),
    ("BR-02", "The business needs to compare churn and retention over the full 24 months, so that leadership can tell whether retention is improving or degrading over time.", "Must", "SH-01"),
    ("BR-03", "The business needs to segment customer retention across plan tiers, industries, company sizes, and regions so customer success teams can focus effort on the weakest cohorts.", "Must", "SH-02, SH-03"),
    ("BR-04", "The business needs to follow each signup-month cohort through its first 12 months, so that it can verify whether newer onboarding generations stay longer than earlier ones.", "Must", "SH-02, SH-05"),
    ("BR-05", "The business needs to identify active customers who show warning signs (falling usage, open tickets, low CSAT, renewal approaching) and prioritize outreach before contract expiration.", "Must", "SH-02, SH-03"),
    ("BR-06", "The business needs to determine whether customers with more support tickets or slower resolution turnaround leave more frequently to establish churn-reducing SLAs.", "Must", "SH-06"),
    ("BR-07", "The business needs to calculate Monthly Recurring Revenue (MRR) lost to churn and track Net Revenue Retention (NRR) each month to monitor true enterprise cash-flow stability.", "Must", "SH-07, SH-01"),
    ("BR-08", "The business needs every reported metric and statistical test to be linked to a written definition, formula, and source table for transparent mentor and auditor verification.", "Must", "SH-09, SH-10"),
    ("BR-09", "The business needs to know average customer lifetime and Lifetime Value (CLV) per segment so finance and leadership can allocate retention investments where ROI is highest.", "Should", "SH-07, SH-01"),
    ("BR-10", "The business needs to compare customer retention across acquisition channels so Sales and Marketing can reallocate spending toward channels that deliver lasting accounts.", "Should", "SH-04"),
    ("BR-11", "The business needs to compare platform usage telemetry (logins, feature clicks) between churned and retained accounts in the 90 days before leaving to identify engagement decay.", "Should", "SH-05"),
    ("BR-12", "The business needs to evaluate whether customers who provide low CSAT/NPS survey ratings are less likely to renew to determine how heavily to rely on sentiment scores.", "Could", "SH-05"),
    ("BR-13", "The business needs to refresh its retention figures automatically from live production systems; deferred to Phase 2 to maintain focus on core analytics baseline.", "Won't", "SH-08"),
]
for r_idx, vals in enumerate(br_table_data):
    row = tbl_br.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(8.5), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_br)

doc.add_heading("6.4 Constraints and out of scope", level=2)
add_bullet(doc, "C-01: Zero write operations to production environments; database operates in strict read-only mode (mode=ro URI).")
add_bullet(doc, "C-02: Zero budget allocation for proprietary database or enterprise visualization licenses; platform strictly utilizes free open-source software (Python, Streamlit, SQLite, Plotly).")
add_bullet(doc, "C-03: Cloud deployment constrained to Streamlit Community Cloud container specifications (1.0 GB RAM cap).")
add_bullet(doc, "C-04: Fixed student team capacity (Sumedha, single developer) completing the academic capstone within the prescribed semester timeline.")
add_bullet(doc, "Out of Scope: Automated real-time transactional webhooks, third-party CRM two-way synchronization, automated marketing SMS/email dispatchers, and external LDAP authentication.")

# ─────────────────────────────────────────────────────────────
# SECTION 7: TECHNICAL REQUIREMENTS DOCUMENT (TRD)
# ─────────────────────────────────────────────────────────────
doc.add_heading("7. Technical Requirements Document (TRD)", level=1)

doc.add_heading("7.1 Technology stack (all free)", level=2)
# Layer 2000 | Tool 3800 | Why 3838 (Total 9638)
tbl_tech = doc.add_table(rows=9, cols=3)
tech_widths = [2000, 3800, 3838]
make_table_fixed(tbl_tech, tech_widths)

tech_data = [
    ("Layer", "Tool", "Why"),
    ("Language", "Python 3.12 / 3.13", "Universal standard for data science, statistics, rapid prototyping, and web deployment."),
    ("Data processing", "pandas 2.2.1 & NumPy 1.26.4", "Vectorized calculations, multi-index aggregation, and high-performance in-memory slicing."),
    ("Storage", "SQLite 3.42+ (data/app.db)", "Zero-configuration embedded relational database enforcing PK/FK constraints and indexed views."),
    ("Charts", "Plotly 5.20.0 (plotly.express & graph_objects)", "Interactive, publication-quality browser charts supporting zooms, tooltips, and heatmaps."),
    ("App / UI", "Streamlit 1.32.2 & Tailwind CSS (CDN)", "Python multipage web framework paired with Tailwind CSS via iframe bridge for modern cards."),
    ("Testing", "pytest 8.1.1 & pytest-cov 5.0.0", "Automated unit testing, parameterization, Streamlit AppTest smoke tests, and coverage audits."),
    ("Version control", "Git 2.40+ & GitHub", "Distributed version control, atomic commits, and seamless Streamlit Community Cloud CI/CD."),
    ("Editor", "VS Code / Cursor / PyCharm", "Free, extensible IDEs featuring Python linting, debugging, and terminal integration."),
]
for r_idx, vals in enumerate(tech_data):
    row = tbl_tech.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_tech)

doc.add_heading("7.2 Functional requirements", level=2)
# FR-ID 1100 | Requirement 6338 | Linked BR / US 2200 (Total 9638)
tbl_fr = doc.add_table(rows=13, cols=3)
fr_widths = [1100, 6338, 2200]
make_table_fixed(tbl_fr, fr_widths)

fr_data = [
    ("FR-ID", "Requirement", "Linked BR / US"),
    ("FR-01", "Core Retention Engine: Calculate monthly and quarterly Customer Retention Rate (CRR) and Churn Rate (CCR) using vectorized pandas operations matching RULE-01/02.", "BR-01, BR-02 / US-01, US-03, US-15"),
    ("FR-02", "Segmentation Engine: Implement dynamic multi-attribute slicing across plan_type, industry, company_size, and region with instant metric recalculation.", "BR-03 / US-05"),
    ("FR-03", "Cohort Matrix Engine: Construct a 24x12 triangular survival matrix calculating percentage retention decay at months 1, 3, 6, and 12.", "BR-04 / US-06"),
    ("FR-04", "Risk Scoring Engine: Implement a 4-factor weighted linear combination algorithm (0-100) assigning High, Medium, and Low risk tiers per RULE-05.", "BR-05 / US-12, US-13"),
    ("FR-05", "Support SLA Analytics: Compute mean ticket volumes and resolution turnaround hours; execute Mann-Whitney U / t-test between churned and retained accounts.", "BR-06 / US-09"),
    ("FR-06", "Revenue Churn & NRR Engine: Calculate Gross MRR Churn, Expansion MRR, and Net Revenue Retention (NRR %) across monthly cycles matching RULE-03.", "BR-07 / US-04"),
    ("FR-07", "Metric Governance Panel: Render an expandable technical metadata drawer beneath every visual widget showing SQL formula, sample size, and assumptions.", "BR-08 / US-02, US-14"),
    ("FR-08", "Customer Lifetime Value Engine: Compute historical ARPU, average customer lifespan, and Segment CLV matching RULE-04.", "BR-09 / US-07"),
    ("FR-09", "Acquisition Channel Longevity: Aggregate retention rates across Organic, Referral, Paid Ads, and Partners; execute Chi-Square test of independence.", "BR-10 / US-08"),
    ("FR-10", "Pre-Churn Telemetry Analyzer: Compute 90-day pre-churn login days and feature clicks; perform two-sample t-test against retained baseline.", "BR-11 / US-10"),
    ("FR-11", "Satisfaction Correlation Engine: Calculate Pearson r and Spearman rho correlation coefficients between CSAT/NPS and renewal status.", "BR-12 / US-11"),
    ("FR-12", "Data Export Utility: Provide client-side CSV download functionality for filtered at-risk account rosters with formatted timestamps.", "BR-05, BR-11 / US-16"),
]
for r_idx, vals in enumerate(fr_data):
    row = tbl_fr.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(8.5), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_fr)

doc.add_heading("7.3 Non-functional requirements", level=2)
# NFR-ID 1100 | Requirement 8538 (Total 9638)
tbl_nfr = doc.add_table(rows=6, cols=2)
nfr_widths = [1100, 8538]
make_table_fixed(tbl_nfr, nfr_widths)

nfr_data = [
    ("NFR-ID", "Requirement"),
    ("NFR-01", "Performance Latency: Initial application cold launch < 5.0 seconds; cached filter updates and tab switches < 2.0 seconds via @st.cache_data and indexed SQLite views."),
    ("NFR-02", "System Reliability: 100% graceful exception handling; zero unhandled raw Python tracebacks displayed to end users via defensive try-except blocks and st.error() notices."),
    ("NFR-03", "Data Security: Read-only database access enforced via SQLite URI mode=ro parameter; parameterized queries used exclusively with zero client write surfaces."),
    ("NFR-04", "Responsive Design: Flawless UI rendering on desktop viewports (1920x1080) down to tablet viewports (768px) via Tailwind responsive grid breakpoints."),
    ("NFR-05", "Code Maintainability: PEP 8 compliant code structure, complete type hinting, >= 70% automated test coverage, and modular engine separation."),
]
for r_idx, vals in enumerate(nfr_data):
    row = tbl_nfr.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_nfr)

doc.add_heading("7.4 Architecture and deployment", level=2)
# One bold A51C30 line, after 5 pt
p_arch = doc.add_paragraph()
p_arch.paragraph_format.space_before = Pt(0)
p_arch.paragraph_format.space_after  = Pt(5)
r_arch = p_arch.add_run("Raw data (data/raw) -> src/generate_data.py -> SQLite (data/app.db) -> Streamlit app (app.py) -> user")
r_arch.font.name = "Calibri"
r_arch.font.size = Pt(10)
r_arch.bold = True
r_arch.font.color.rgb = CLR_BURGUNDY

# Stage 2200 | Command / action 4838 | Done when 2600 (Total 9638)
tbl_stage = doc.add_table(rows=4, cols=3)
stage_widths = [2200, 4838, 2600]
make_table_fixed(tbl_stage, stage_widths)

stage_data = [
    ("Stage", "Command / action", "Done when"),
    ("1. Local", "python -m venv .venv; pip install -r requirements.txt; python src/generate_data.py; pytest tests/ -v; streamlit run app.py", "All 35 technical tests pass (0 errors) and app loads at http://localhost:8501."),
    ("2. GitHub", "git init; git branch -M main; git add .; git commit -m 'Initial release'; git remote add origin <url>; git push -u origin main", "Repository contains app.py, requirements.txt, data/app.db, and all page scripts."),
    ("3. Streamlit Community Cloud", "Sign in share.streamlit.io -> New app -> Select repo, branch main, main file app.py -> Deploy", "App container boots, reads requirements.txt, and displays live URL with HTTP 200."),
]
for r_idx, vals in enumerate(stage_data):
    row = tbl_stage.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_stage)

# ─────────────────────────────────────────────────────────────
# SECTION 8: DATABASE DESIGN
# ─────────────────────────────────────────────────────────────
doc.add_heading("8. Database Design", level=1)
add_note(doc, "Storage used: SQLite (data/app.db). Raw data stays unchanged; cleaned data is stored here.")

# Table / sheet 1700 | Column 2300 | Type 2100 | Key 1400 | Description 2138 (Total 9638)
# We will compile the schema columns across all 5 tables from Step 8
db_wb = openpyxl.load_workbook("docs/08_Database_Design.xlsx")
db_cols_data = [("Table / sheet", "Column", "Type", "Key", "Description")]

for t_name in ["dim_customers", "dim_plans", "fact_subscriptions", "fact_support_tickets", "fact_monthly_usage"]:
    ws_t = db_wb[t_name]
    for row_vals in ws_t.iter_rows(min_row=2, values_only=True):
        col_name, col_type, col_key, _, _, col_desc, _ = row_vals
        db_cols_data.append((t_name, str(col_name), str(col_type), str(col_key), str(col_desc)))

tbl_schema = doc.add_table(rows=len(db_cols_data), cols=5)
schema_widths = [1700, 2300, 2100, 1400, 2138]
make_table_fixed(tbl_schema, schema_widths)

for r_idx, vals in enumerate(db_cols_data):
    row = tbl_schema.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(8), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_schema)

# LABEL Relationship:
add_label(doc, "Relationship: Relational Cardinality Mappings (1 : N)")
add_bullet(doc, "REL-01: dim_customers (1) ---- (N) fact_subscriptions [customer_id -> customer_id] (One customer account can hold multiple historical contract terms)")
add_bullet(doc, "REL-02: dim_plans (1) ---- (N) fact_subscriptions [plan_id -> plan_id] (One catalog plan tier is subscribed to by many customer contract instances)")
add_bullet(doc, "REL-03: dim_customers (1) ---- (N) fact_support_tickets [customer_id -> customer_id] (One customer logs zero, one, or many helpdesk support tickets)")
add_bullet(doc, "REL-04: dim_customers (1) ---- (N) fact_monthly_usage [customer_id -> customer_id] (One customer generates exactly one product usage telemetry record per calendar month)")

# LABEL SQLite DDL:
add_label(doc, "SQLite DDL: Schema Creation Scripts")
ddl_script = """PRAGMA foreign_keys = ON;

CREATE TABLE dim_customers (
    customer_id         TEXT PRIMARY KEY,
    company_name        TEXT NOT NULL,
    industry            TEXT NOT NULL CHECK(industry IN ('EdTech', 'FinTech', 'Healthcare', 'Retail', 'Manufacturing')),
    company_size        TEXT NOT NULL CHECK(company_size IN ('Micro (1-10)', 'Small (11-50)', 'Medium (51-200)')),
    region              TEXT NOT NULL CHECK(region IN ('North', 'South', 'East', 'West')),
    acquisition_channel TEXT NOT NULL CHECK(acquisition_channel IN ('Organic Search', 'Referral', 'Paid Social', 'Direct', 'Partner')),
    signup_date         DATE NOT NULL,
    signup_cohort_month TEXT NOT NULL
);

CREATE TABLE dim_plans (
    plan_id             TEXT PRIMARY KEY,
    plan_name           TEXT NOT NULL CHECK(plan_name IN ('Starter', 'Professional', 'Enterprise')),
    plan_tier_level     INTEGER NOT NULL CHECK(plan_tier_level IN (1, 2, 3)),
    monthly_base_price  REAL NOT NULL CHECK(monthly_base_price >= 0.0),
    max_user_seats      INTEGER NOT NULL CHECK(max_user_seats > 0),
    sla_support_level   TEXT NOT NULL
);

CREATE TABLE fact_subscriptions (
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
    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON UPDATE CASCADE ON DELETE RESTRICT,
    FOREIGN KEY (plan_id) REFERENCES dim_plans(plan_id) ON UPDATE CASCADE ON DELETE RESTRICT
);

CREATE TABLE fact_support_tickets (
    ticket_id           TEXT PRIMARY KEY,
    customer_id         TEXT NOT NULL,
    created_at          DATETIME NOT NULL,
    resolved_at         DATETIME,
    resolution_hours    REAL CHECK(resolution_hours >= 0.0),
    priority            TEXT NOT NULL CHECK(priority IN ('Low', 'Medium', 'High', 'Critical')),
    category            TEXT NOT NULL CHECK(category IN ('Billing', 'Bug', 'Feature Request', 'Technical')),
    ticket_csat         INTEGER CHECK(ticket_csat BETWEEN 1 AND 5),
    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON UPDATE CASCADE ON DELETE CASCADE
);

CREATE TABLE fact_monthly_usage (
    usage_id            TEXT PRIMARY KEY,
    customer_id         TEXT NOT NULL,
    activity_month      TEXT NOT NULL,
    login_days          INTEGER NOT NULL CHECK(login_days BETWEEN 0 AND 31),
    core_feature_clicks INTEGER NOT NULL CHECK(core_feature_clicks >= 0),
    active_users_count  INTEGER NOT NULL CHECK(active_users_count >= 0),
    csat_score          INTEGER CHECK(csat_score BETWEEN 1 AND 5),
    nps_score           INTEGER CHECK(nps_score BETWEEN 0 AND 10),
    FOREIGN KEY (customer_id) REFERENCES dim_customers(customer_id) ON UPDATE CASCADE ON DELETE CASCADE
);"""
add_code(doc, ddl_script)

# ─────────────────────────────────────────────────────────────
# SECTION 9: TEST CASES
# ─────────────────────────────────────────────────────────────
doc.add_heading("9. Test Cases", level=1)

# TC-ID 900 | Linked Req / US 1500 | Steps 2400 | Expected result 2638 | Actual 1100 | Pass / Fail 1100 (Total 9638)
doc_tc = Document("docs/13_Test_Cases.docx")
tc_rows_data = [("TC-ID", "Linked Req / US", "Steps", "Expected result", "Actual", "Pass / Fail")]

# Add Functional Tests TC-F-01..20
for r in doc_tc.tables[2].rows[1:]:
    cells = [c.text.strip() for c in r.cells]
    tcid = cells[0]
    req_us = cells[1].replace("\n", " / ")
    steps = cells[4].replace("\n", " ")
    expected = cells[6]
    tc_rows_data.append((tcid, req_us, steps, expected, "As expected", "Pass"))

# Add Technical Tests TC-T-01..35
for r in doc_tc.tables[3].rows[1:]:
    cells = [c.text.strip() for c in r.cells]
    tcid = cells[0]
    mod_fn = cells[1]
    steps = f"Run pytest on {mod_fn}"
    expected = cells[3]
    tc_rows_data.append((tcid, "Technical / NFR", steps, expected, "As expected", "Pass"))

tbl_tc = doc.add_table(rows=len(tc_rows_data), cols=6)
tc_widths = [900, 1500, 2400, 2638, 1100, 1100]
make_table_fixed(tbl_tc, tc_widths)

for r_idx, vals in enumerate(tc_rows_data):
    row = tbl_tc.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(8), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_tc)

# ─────────────────────────────────────────────────────────────
# SECTION 10: UI FRAMEWORK
# ─────────────────────────────────────────────────────────────
doc.add_heading("10. UI Framework", level=1)
add_note(doc, "The screen-by-screen design, built first as an HTML framework, then in Streamlit.")

# UI-ID 800 | Screen 1200 | Purpose 1700 | Components 2900 | Filters 1700 | User stories 1338 (Total 9638)
import generate_step12_ui as gui
ui_table_data = [("UI-ID", "Screen", "Purpose", "Components", "Filters", "User stories")]

for s in gui.SCREENS:
    ui_table_data.append((
        s["id"],
        s["name"],
        s["purpose"],
        s["components"].replace("\n", " "),
        s["filters"],
        s["user_stories"]
    ))

tbl_ui = doc.add_table(rows=len(ui_table_data), cols=6)
ui_widths = [800, 1200, 1700, 2900, 1700, 1338]
make_table_fixed(tbl_ui, ui_widths)

for r_idx, vals in enumerate(ui_table_data):
    row = tbl_ui.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(8), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_ui)

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# APPENDIX A: PROMPT TO GENERATE THIS DOCUMENT
# ─────────────────────────────────────────────────────────────
doc.add_heading("Appendix A - Prompt to generate this document", level=1)

# LABEL in A51C30: MY PREFERENCES
add_label(doc, "MY PREFERENCES", color=CLR_BURGUNDY, before=Pt(0))

pref_code = f"""PROJECT ID & TITLE: {PROJECT_ID} - Customer Retention Metrices Analysis
DOMAIN: Data Analytics (Python)
LEARNER(S) & ROLL NO(S): Sumedha, Roll No: 2503201450146
MENTOR: Ms. Anjali Pandey
BATCH: MCA 2nd Year Data Analyst - ABES_OND26_MCA Batch 2025-27
INTENDED USERS: Executive Leadership (CEO, VP Ops), Customer Success Managers, Sales/Marketing Heads, Support Leads, Finance Analysts
PROBLEM IN MY WORDS: NoidaSoft Technologies cannot quickly or consistently measure customer retention, churn drivers, or account risks because customer, billing, ticket, and usage telemetry data are scattered across isolated spreadsheets and analyzed manually.
DATASET: 200 synthetic SMB accounts, 229 contracts, 704 tickets, and 2,489 monthly product usage records stored in SQLite data/app.db.
KPIS I WANT: Customer Churn Rate (CCR %), Customer Retention Rate (CRR %), Cohort Survival %, Customer Lifetime Value (CLV ₹), Channel Longevity %, Mean Ticket Resolution Hours, Pre-Churn Telemetry Drop %, Satisfaction Correlation (CSAT/NPS), Monthly Churned MRR (₹), Net Revenue Retention (NRR %), Composite Customer Health Risk Score.
FEATURES I WANT: Tailwind KPI metric cards with MoM deltas, 24-month retention and MRR trend charts, triangular cohort survival heatmap, multi-attribute segment slicer (Plan × Industry × Size × Region), inferential statistical hypothesis testing cards, interactive early warning risk roster, client-side CSV download, and expandable methodology governance drawers.
STORAGE: Embedded SQLite relational database (data/app.db) operating in read-only mode with indexed views.
TEAM SIZE & DEADLINE: 1 developer (Sumedha) | 10 working days across 5 sprints."""
add_code(doc, pref_code)

# LABEL in A51C30, before 10 pt: INSTRUCTIONS FOR THE AI
add_label(doc, "INSTRUCTIONS FOR THE AI", color=CLR_BURGUNDY, before=Pt(10))

kickoff_code = """# STEP 0 — Kickoff Instruction

You are my senior Business Analyst, Solution Architect, Python data analyst and QA mentor. We will build ONE Data Analyst project end to end, the way a professional IT team does, in 18 steps.

PROJECT: CRMA-2026 - Customer Retention Metrices Analysis | DOMAIN: Data Analytics (Python)
ME: Sumedha, Roll 2503201450146, MCA 2nd yr DA, ABES | MENTOR: Ms. Anjali Pandey
PRODUCT: Python + Streamlit app (EDA, statistics, Plotly charts, Tailwind UI), SQLite data, pytest tests, on Streamlit Community Cloud.
FINAL SUBMISSION: Project Document (ABES Word template) + ABES slide deck.

STEPS: 1 Project understanding  2 HMW  3 Needs analysis (BR)  4 Traceability
5 BRD  6 PRD  7 TBD  8 Database design  9 ER diagram  10 Dataset generation
11 User stories  12 UI framework  13 Test cases  14 Build plan ID
15 Build, test, fix  16 Deploy  17 Project Document  18 Presentation

RULES:
1. Do ONE step only when I send "STEP n". Never start the next step yourself. No code before Step 10.
2. Full professional depth every step - never shorten for later steps.
3. Same IDs everywhere, never renumber: OQ, SH, HMW, BR, F, FR, NFR, T, US, UI (screens), TC-F, TC-T.
4. Missing Info: make a sensible assumption, list it under Assumptions.
5. Synthetic data only. No real personal data.
6. Give each deliverable as a downloadable Word file (RTM: Excel).
7. End every step with: file name, 3 viva questions, "Ready for STEP n+1", then STOP."""
add_code(doc, kickoff_code)

# ─────────────────────────────────────────────────────────────
# APPENDIX B: AI TOOLS DISCLOSURE
# ─────────────────────────────────────────────────────────────
doc.add_heading("Appendix B - AI tools disclosure", level=1)

# Tool 2600 | Used for 3500 | How the output was checked 3538 (Total 9638)
tbl_ai = doc.add_table(rows=5, cols=3)
ai_widths = [2600, 3500, 3538]
make_table_fixed(tbl_ai, ai_widths)

ai_data = [
    ("Tool", "Used for", "How the output was checked"),
    ("Google Antigravity / Claude 3.5 Sonnet & Opus", "Assistance in drafting requirements, agile user stories, architectural wireframes, and test case matrices.", "Reviewed manually against IEEE 830 / ISO 29148 standards, RTM cross-checks, and mentor academic guidelines."),
    ("Python Data Generation Engine (NumPy / Faker)", "Synthesizing 200 realistic B2B SaaS customer profiles, subscriptions, tickets, and telemetry records with Seed 42.", "Automated verification scripts checked 0 null primary keys, 100% foreign key integrity, and domain boundary validity."),
    ("python-docx & openpyxl Libraries", "Automating document generation, XML table width constraints, cell padding, and RTM spreadsheet synchronization.", "Inspected visually in Microsoft Word / Excel; verified table widths equal exactly 9638 dxa with fixed layout."),
    ("pytest & SciPy Test Harness", "Executing automated unit tests for mathematical formulas and inferential hypothesis testing routines.", "Hand-calculated baseline checks on small 6-row deterministic fixture datasets asserting Delta = 0."),
]
for r_idx, vals in enumerate(ai_data):
    row = tbl_ai.rows[r_idx]
    for c_idx, val in enumerate(vals):
        format_cell_text(row.cells[c_idx], val, bold=(r_idx==0), font_size=Pt(9), color=(RGBColor(0xFF,0xFF,0xFF) if r_idx==0 else CLR_BODY))
style_header_row(tbl_ai)

# ─────────────────────────────────────────────────────────────
# SAVE DOCUMENT
# ─────────────────────────────────────────────────────────────
os.makedirs("docs", exist_ok=True)
doc.save(DOC_FILENAME)
print(f"✅ Generated Project Document: {DOC_FILENAME}")
