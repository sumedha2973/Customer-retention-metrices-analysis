"""
Step 11 — User Stories Document Generator
Customer Retention Metrics Analysis (CRMA)
Sumedha | Roll 2503201450146 | MCA 2nd Yr | ABES
"""
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os, datetime

# ── Colour palette (matches earlier steps) ──
NAVY   = RGBColor(0x1B, 0x36, 0x5D)
STEEL  = RGBColor(0x2B, 0x4C, 0x7E)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY = "F2F4F7"
HEADER_BG  = "1B365D"
ALT_ROW    = "EDF1F7"
ACCENT_BG  = "E8EDF5"

doc = Document()

# ── Page setup ──
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.0)
    section.right_margin  = Cm(2.0)

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)
style.paragraph_format.space_after = Pt(4)
style.paragraph_format.line_spacing = 1.15

# ── Helper functions ──
def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_text(cell, text, bold=False, font_size=Pt(9), color=None, alignment=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if alignment:
        p.alignment = alignment
    run = p.add_run(str(text))
    run.font.name = 'Calibri'
    run.font.size = font_size
    run.bold = bold
    if color:
        run.font.color.rgb = color

def add_header_row(table, texts, font_size=Pt(8.5)):
    for i, txt in enumerate(texts):
        cell = table.rows[0].cells[i]
        set_cell_shading(cell, HEADER_BG)
        set_cell_text(cell, txt, bold=True, font_size=font_size, color=WHITE, alignment=WD_ALIGN_PARAGRAPH.CENTER)

def add_data_row(table, row_idx, values, font_size=Pt(8.5)):
    row = table.rows[row_idx]
    bg = ALT_ROW if row_idx % 2 == 0 else "FFFFFF"
    for i, val in enumerate(values):
        cell = row.cells[i]
        set_cell_shading(cell, bg)
        set_cell_text(cell, val, font_size=font_size)

def add_heading(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = NAVY
        run.font.name = 'Calibri'
    return h

def add_body(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(4)
    return p

def add_info_box(title, content):
    """Create a coloured info box using a 1x1 table"""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, ACCENT_BG)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(title + "\n")
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(10)
    run.font.color.rgb = NAVY
    run2 = p.add_run(content)
    run2.font.name = 'Calibri'
    run2.font.size = Pt(9)
    # Set table width
    tbl_element = tbl._tbl
    tbl_element.attrib[qn('w:w')] = str(Cm(16).emu)
    doc.add_paragraph()

# ── ALL USER STORIES DATA ──
# Epics
EPICS = {
    "E-01": ("Data Pipeline & Quality", "Data ingestion, cleaning, validation, and caching pipeline"),
    "E-02": ("Executive Dashboard", "Top-level KPI overview for leadership decision-making"),
    "E-03": ("Retention Analysis", "Cohort analysis, segmentation, and trend exploration"),
    "E-04": ("Churn Drivers & Insights", "Statistical analysis of factors driving churn"),
    "E-05": ("Quality & Governance", "Data provenance, metric definitions, and export capabilities"),
}

# Personas (from stakeholders)
PERSONAS = {
    "VP Operations":    "SH-01",
    "CS Manager":       "SH-02",
    "CS Executive":     "SH-03",
    "Sales Head":       "SH-04",
    "Product Manager":  "SH-05",
    "Support Lead":     "SH-06",
    "Finance Analyst":  "SH-07",
    "Data Analyst":     "SH-08",
    "Mentor":           "SH-09",
    "Examiner":         "SH-10",
}

# User stories: (US-ID, Epic, Persona, Story, Feature/FR, BR, Priority, SP, AC_list)
STORIES = [
    # ── EPIC E-01: Data Pipeline & Quality ──
    ("US-01", "E-01", "Data Analyst",
     "As a Data Analyst, I want the app to load all five database tables into memory with schema validation on startup, so that I am certain every downstream metric rests on clean, verified data.",
     "FR-01, FR-02", "BR-01, BR-08", "Must", 3,
     [
         ("Given the app starts up",
          "When the data loader reads data/app.db",
          "Then all 5 tables load successfully, PK uniqueness is asserted, FK integrity is 100%, and domain ranges (CSAT 1-5, login 0-31, fee ≥ 0) pass without error."),
         ("Given a schema assertion fails",
          "When the validation detects a violation",
          "Then a structured st.error() message names the exact check, table, and column that failed — no raw traceback is shown."),
     ]),
    ("US-02", "E-01", "Mentor",
     "As a Mentor reviewing the project, I want every computed metric to display its formula, data source, sample size, and assumptions in an expandable drawer, so that I can verify analytical rigour without reading source code.",
     "FR-07", "BR-08", "Must", 2,
     [
         ("Given any KPI card or chart is visible",
          "When the user clicks the 'ℹ Methodology' expander beneath it",
          "Then a panel shows the SQL/pandas formula, the source table(s), the row count used, and any filtering assumptions."),
     ]),

    # ── EPIC E-02: Executive Dashboard ──
    ("US-03", "E-02", "VP Operations",
     "As the VP of Operations, I want a single executive overview page showing the current month's Churn Rate, Retention Rate, MRR Lost, NRR %, and Active Customers as KPI cards with month-over-month deltas, so that I can assess company health in under 10 seconds.",
     "FR-01, FR-06", "BR-01, BR-02, BR-07", "Must", 5,
     [
         ("Given the Executive Overview page loads",
          "When the latest month's data is available",
          "Then 5 KPI cards (CCR %, CRR %, MRR Lost ₹, NRR %, Active Count) display with green/red delta badges comparing to the previous month."),
         ("Given CCR for the latest month is 2.8%",
          "When the previous month's CCR was 3.1%",
          "Then the delta badge shows '▼ 0.3 pp' in green (improvement)."),
     ]),
    ("US-04", "E-02", "Finance Analyst",
     "As a Finance Analyst, I want a 24-month MRR trend line chart on the overview page showing gross MRR, churned MRR, and net MRR, so that I can see the revenue impact of churn over time and forecast future losses.",
     "FR-06", "BR-07", "Must", 3,
     [
         ("Given the Executive Overview page is open",
          "When I look at the MRR trend chart",
          "Then a Plotly line chart shows three series (Gross MRR, Churned MRR, Net MRR) across all 24 months with ₹ INR formatting on the Y-axis."),
         ("Given I hover over any data point",
          "When the tooltip appears",
          "Then it shows the exact month, metric name, and INR value rounded to two decimals."),
     ]),
    ("US-15", "E-02", "VP Operations",
     "As the VP of Operations, I want a churn vs. retention rate trend chart on the overview page showing 24 months of monthly CCR and CRR side by side, so that I can spot seasonal patterns and long-term direction at a glance.",
     "FR-01", "BR-01, BR-02", "Must", 3,
     [
         ("Given the Executive Overview page is open",
          "When I view the retention trend chart",
          "Then a dual-axis Plotly chart displays CCR % and CRR % for each of the 24 months with clearly labelled axes."),
     ]),

    # ── EPIC E-03: Retention Analysis ──
    ("US-05", "E-03", "CS Manager",
     "As a Customer Success Manager, I want to filter all metrics by plan type, industry, company size, and region using sidebar multi-select dropdowns, so that I can compare retention across any combination of segments instantly.",
     "FR-02", "BR-03", "Must", 5,
     [
         ("Given the Segment Comparison page is open",
          "When I select 'Starter' in the Plan Type filter and 'FinTech' in the Industry filter",
          "Then every KPI card, chart, and table on the page recalculates to show only Starter-plan FinTech customers, and the filter summary bar shows the active selections."),
         ("Given all filters are cleared",
          "When I click 'Reset Filters'",
          "Then the page reverts to showing all 200 customers with no active filter."),
     ]),
    ("US-06", "E-03", "CS Manager",
     "As a Customer Success Manager, I want a cohort survival heatmap showing what percentage of each signup-month cohort is still active at months 1, 3, 6, and 12, so that I can tell whether newer cohorts retain better than older ones.",
     "FR-03", "BR-04", "Must", 5,
     [
         ("Given the Cohort Analysis page loads",
          "When the survival matrix is computed",
          "Then a triangular heatmap displays signup cohorts (rows, YYYY-MM) versus lifecycle months (columns, 1/3/6/12) with colour intensity proportional to survival percentage (green = high, red = low)."),
         ("Given I hover over the cell for cohort 2024-03 at month 6",
          "When the tooltip appears",
          "Then it shows 'Cohort 2024-03 | Month 6 | 82.5% survived | 33 of 40 customers'."),
     ]),
    ("US-07", "E-03", "Finance Analyst",
     "As a Finance Analyst, I want a CLV comparison bar chart grouped by plan tier and industry, so that I can identify which customer segments generate the most lifetime value.",
     "FR-08", "BR-09", "Should", 3,
     [
         ("Given the Segment Comparison page is open",
          "When the CLV chart section is visible",
          "Then a grouped Plotly bar chart shows average CLV (₹ INR) for each plan tier (Starter, Professional, Enterprise) broken down by industry."),
         ("Given Enterprise FinTech has the highest CLV",
          "When I read the chart",
          "Then the Enterprise FinTech bar is visually the tallest and displays its exact INR value on hover."),
     ]),
    ("US-08", "E-03", "Sales Head",
     "As the Sales Head, I want a channel retention comparison chart showing 12-month retention rate by acquisition channel with a statistical significance indicator, so that I can redirect marketing budget to channels that bring lasting customers.",
     "FR-09", "BR-10", "Should", 3,
     [
         ("Given the Segment Comparison page is open",
          "When the channel retention chart loads",
          "Then a horizontal bar chart displays 12-month retention rate for each of the 5 acquisition channels (Organic, Referral, Paid Social, Direct, Partner)."),
         ("Given the Chi-Square test p-value < 0.05",
          "When the chart renders",
          "Then a green badge '✓ Statistically Significant (p = 0.03)' appears below the chart; otherwise a grey badge '✗ Not Significant (p = 0.42)' appears."),
     ]),

    # ── EPIC E-04: Churn Drivers & Insights ──
    ("US-09", "E-04", "Support Lead",
     "As the Support Lead, I want a support-impact analysis showing whether churned customers had significantly more tickets and slower resolution times than retained customers, so that I can prioritise SLA improvements that reduce churn.",
     "FR-05", "BR-06", "Must", 5,
     [
         ("Given the Churn Drivers page is open",
          "When the support analysis section loads",
          "Then side-by-side box plots compare ticket volume and resolution hours between churned and retained groups."),
         ("Given the Mann-Whitney U test has been computed",
          "When the result is displayed",
          "Then a summary card shows the U-statistic, p-value, effect size (r), and a plain-English interpretation (e.g., 'Churned customers had significantly more tickets, p = 0.001')."),
     ]),
    ("US-10", "E-04", "Product Manager",
     "As a Product Manager, I want a 90-day pre-churn usage decline chart comparing login activity and feature engagement of churned customers against retained ones, so that I can identify where engagement drops before cancellation.",
     "FR-10", "BR-11", "Should", 5,
     [
         ("Given the Churn Drivers page is open",
          "When the pre-churn telemetry section loads",
          "Then a dual-line Plotly chart shows average login_days and core_feature_clicks for churned vs. retained customers across the 3-month window before churn/end date."),
         ("Given the two-sample t-test results are computed",
          "When displayed beneath the chart",
          "Then it shows the t-statistic, p-value, 95% CI for the difference, and states whether the decline is statistically significant at α = 0.05."),
     ]),
    ("US-11", "E-04", "Product Manager",
     "As a Product Manager, I want a satisfaction-renewal correlation panel showing Pearson and Spearman correlations between CSAT/NPS scores and renewal status, so that I can decide how heavily to invest in improving satisfaction scores.",
     "FR-11", "BR-12", "Could", 3,
     [
         ("Given the Churn Drivers page is open",
          "When the satisfaction correlation section loads",
          "Then a scatter plot with regression line shows latest CSAT score (x) vs. renewal flag (y), and a results card shows Pearson r, Spearman ρ, p-values, and sample size."),
         ("Given more than 60% of CSAT values are null",
          "When the analysis runs",
          "Then a warning badge '⚠ Limited sample: N = {count} of 200 customers have survey data' appears above the chart."),
     ]),

    # ── EPIC E-05: Quality & Governance ──
    ("US-12", "E-05", "CS Executive",
     "As a Customer Success Executive, I want an at-risk customer roster page showing all active accounts sorted by health risk score (highest first) with risk tier badges, so that I can prioritise my daily outreach to the most endangered accounts.",
     "FR-04", "BR-05", "Must", 5,
     [
         ("Given the Risk Roster page loads",
          "When active customer data is available",
          "Then a sortable table shows columns: Customer ID, Company, Plan, Industry, Risk Score (0-100), Risk Tier (High/Medium/Low), Tenure, Next Renewal, and each row has a coloured badge (Red = High ≥ 70, Amber = Medium 40-69, Green = Low < 40)."),
         ("Given I click the 'Risk Score' column header",
          "When the sort toggles",
          "Then rows reorder from highest to lowest (or vice versa)."),
     ]),
    ("US-13", "E-05", "CS Executive",
     "As a Customer Success Executive, I want to filter the risk roster by risk tier (High / Medium / Low) and by plan type, so that I can focus my intervention on high-risk Enterprise customers who represent the most revenue.",
     "FR-04, FR-02", "BR-05, BR-03", "Must", 2,
     [
         ("Given the Risk Roster page is open",
          "When I select 'High' in the Risk Tier filter and 'Enterprise' in the Plan filter",
          "Then only customers with risk_score ≥ 70 on the Enterprise plan appear in the table, and the count badge updates (e.g., 'Showing 4 of 200')."),
     ]),
    ("US-14", "E-05", "Examiner",
     "As an Examiner evaluating the project, I want a Methodology & Data Dictionary page listing every metric's definition, formula, statistical test, significance threshold, and data lineage, so that I can verify the analytical approach without reviewing code.",
     "FR-07", "BR-08", "Must", 3,
     [
         ("Given the Methodology page loads",
          "When I scroll through the content",
          "Then I see: (a) a glossary of all 11 KPI measures with formulas matching Step 8 KPIs_Measures, (b) statistical test descriptions for each OQ, (c) the data dictionary for all 5 tables, and (d) a note on data-quality issues and how they were handled."),
     ]),
    ("US-16", "E-05", "Data Analyst",
     "As a Data Analyst, I want a CSV download button on the risk roster page that exports the currently filtered at-risk accounts with all visible columns and a timestamp, so that I can share the list with the CS team via email.",
     "FR-12", "BR-05, BR-11", "Should", 2,
     [
         ("Given the Risk Roster page shows a filtered list of customers",
          "When I click the '⬇ Download CSV' button",
          "Then the browser downloads a file named 'at_risk_customers_YYYYMMDD_HHMMSS.csv' containing exactly the rows and columns currently visible in the table."),
     ]),
]

# Sprint Plan
SPRINT_PLAN = [
    ("Sprint 1\n(Day 1-2)", "E-01: Data Pipeline",
     "US-01, US-02", "5 SP",
     "Data loader, schema validation, caching, methodology expander component — foundation for all pages."),
    ("Sprint 2\n(Day 3-4)", "E-02: Executive Dashboard",
     "US-03, US-04, US-15", "11 SP",
     "KPI cards with deltas, MRR trend chart, churn/retention trend chart — the VP's landing page."),
    ("Sprint 3\n(Day 5-7)", "E-03: Retention Analysis",
     "US-05, US-06, US-07, US-08", "16 SP",
     "Sidebar filters, cohort heatmap, CLV comparison, channel retention — deep-dive segmentation pages."),
    ("Sprint 4\n(Day 8-9)", "E-04: Churn Drivers",
     "US-09, US-10, US-11", "13 SP",
     "Support impact analysis, pre-churn telemetry, satisfaction correlation — statistical insights."),
    ("Sprint 5\n(Day 10)", "E-05: Quality & Governance",
     "US-12, US-13, US-14, US-16", "12 SP",
     "Risk roster with filters, methodology page, CSV export — governance and actionability."),
]


# ════════════════════════════════════════════════════
# BUILD THE DOCUMENT
# ════════════════════════════════════════════════════

# ── Title Page ──
doc.add_paragraph()
doc.add_paragraph()
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run("Customer Retention Metrics Analysis")
run.bold = True
run.font.size = Pt(26)
run.font.color.rgb = NAVY
run.font.name = 'Calibri'

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run2 = subtitle.add_run("Step 11 — User Stories & Sprint Plan")
run2.font.size = Pt(16)
run2.font.color.rgb = STEEL
run2.font.name = 'Calibri'

doc.add_paragraph()

# Document control table
info_data = [
    ("Project", "Customer Retention Metrics Analysis (CRMA)"),
    ("Author", "Sumedha | Roll 2503201450146 | MCA 2nd Yr, ABES"),
    ("Mentor", "Ms. Anjali Pandey"),
    ("Version", "1.0"),
    ("Date", datetime.date.today().strftime("%d %B %Y")),
    ("Status", "Baselined"),
    ("Input Documents", "01 Project Understanding, 03 Needs Analysis (BRs),\n05 BRD, 07 TRD (FRs), 08 Database Design"),
]
info_tbl = doc.add_table(rows=len(info_data), cols=2)
info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(info_data):
    set_cell_shading(info_tbl.rows[i].cells[0], ACCENT_BG)
    set_cell_text(info_tbl.rows[i].cells[0], k, bold=True, font_size=Pt(10), color=NAVY)
    set_cell_text(info_tbl.rows[i].cells[1], v, font_size=Pt(10))

doc.add_page_break()

# ── Table of Contents placeholder ──
add_heading("Table of Contents", level=1)
toc_items = [
    "1. Introduction & Scope",
    "2. Personas",
    "3. Epic Structure",
    "4. User Stories — Full Catalogue",
    "   4.1 Epic E-01: Data Pipeline & Quality",
    "   4.2 Epic E-02: Executive Dashboard",
    "   4.3 Epic E-03: Retention Analysis",
    "   4.4 Epic E-04: Churn Drivers & Insights",
    "   4.5 Epic E-05: Quality & Governance",
    "5. User Story Summary Table",
    "6. Sprint Plan",
    "7. Story-to-Requirement Traceability",
    "8. Coverage Analysis",
    "9. Assumptions",
]
for item in toc_items:
    p = doc.add_paragraph(item)
    p.paragraph_format.space_after = Pt(1)

doc.add_page_break()

# ═══════════ 1. INTRODUCTION ═══════════
add_heading("1. Introduction & Scope", level=1)
add_body(
    "This document defines the Agile User Stories for the Customer Retention Metrics Analysis (CRMA) "
    "application. Each story is derived from the Business Requirements (BRs) established in Step 3, "
    "the Functional Requirements (FRs) specified in Step 7 (TRD), and the database design from Step 8. "
    "Since Step 6 (PRD) was skipped, features are mapped directly from BRs and FRs."
)
add_body(
    "The stories follow the standard Agile format: 'As a <persona>, I want <goal>, so that <benefit>.' "
    "Each story includes Given/When/Then acceptance criteria, is sized in story points (1-8 scale using "
    "modified Fibonacci), and is grouped into one of five Epics. All stories are buildable in under "
    "half a day (≤ 5 story points)."
)

add_info_box("📋 Scope Summary",
    f"• Total User Stories: {len(STORIES)}\n"
    f"• Epics: {len(EPICS)}\n"
    f"• Total Story Points: {sum(s[7] for s in STORIES)} SP\n"
    f"• Priority Split: {sum(1 for s in STORIES if s[6]=='Must')} Must / "
    f"{sum(1 for s in STORIES if s[6]=='Should')} Should / "
    f"{sum(1 for s in STORIES if s[6]=='Could')} Could\n"
    f"• Sprints: 5 (10 working days)\n"
    f"• OQ Coverage: OQ-01 through OQ-10 (100%)")

# ═══════════ 2. PERSONAS ═══════════
add_heading("2. Personas", level=1)
add_body(
    "The following personas are drawn from the stakeholder register (Step 1, SH-01 to SH-10). "
    "Each persona represents a distinct user role who interacts with the CRMA application."
)

persona_tbl = doc.add_table(rows=1 + len(PERSONAS), cols=4)
persona_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(persona_tbl, ["Persona", "SH ID", "Primary Goal", "Frequency of Use"])

persona_details = [
    ("VP Operations", "SH-01", "Monitor overall retention health and strategic KPIs", "Daily — morning check"),
    ("CS Manager", "SH-02", "Deep-dive into segments and cohorts for team strategy", "Daily"),
    ("CS Executive", "SH-03", "Identify and act on at-risk accounts for outreach", "Multiple times daily"),
    ("Sales Head", "SH-04", "Evaluate acquisition channel quality for budget allocation", "Weekly"),
    ("Product Manager", "SH-05", "Understand usage-churn relationship for roadmap decisions", "Weekly"),
    ("Support Lead", "SH-06", "Assess support-churn correlation for SLA improvement", "Weekly"),
    ("Finance Analyst", "SH-07", "Track MRR impact and CLV for financial forecasting", "Weekly/Monthly"),
    ("Data Analyst", "SH-08", "Validate data pipeline and export reports", "Daily"),
    ("Mentor", "SH-09", "Verify analytical rigour and methodology", "Per review session"),
    ("Examiner", "SH-10", "Evaluate project completeness and correctness", "During evaluation"),
]
for i, (persona, sh, goal, freq) in enumerate(persona_details, 1):
    add_data_row(persona_tbl, i, [persona, sh, goal, freq])

# ═══════════ 3. EPIC STRUCTURE ═══════════
doc.add_page_break()
add_heading("3. Epic Structure", level=1)
add_body(
    "Stories are grouped into five Epics, following the natural data-to-insight flow of the application:"
)

epic_tbl = doc.add_table(rows=1 + len(EPICS), cols=4)
epic_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(epic_tbl, ["Epic ID", "Epic Name", "Description", "Stories"])

epic_story_map = {}
for s in STORIES:
    epic = s[1]
    epic_story_map.setdefault(epic, []).append(s[0])

for i, (eid, (ename, edesc)) in enumerate(EPICS.items(), 1):
    stories_in = ", ".join(epic_story_map.get(eid, []))
    add_data_row(epic_tbl, i, [eid, ename, edesc, stories_in])

# ═══════════ 4. USER STORIES — FULL CATALOGUE ═══════════
doc.add_page_break()
add_heading("4. User Stories — Full Catalogue", level=1)

current_epic = None
for s in STORIES:
    us_id, epic_id, persona, story_text, fr_ids, br_ids, priority, sp, ac_list = s

    # Epic sub-heading
    if epic_id != current_epic:
        current_epic = epic_id
        epic_name = EPICS[epic_id][0]
        add_heading(f"4.{int(epic_id.split('-')[1])} Epic {epic_id}: {epic_name}", level=2)

    # Story card as a table
    add_heading(f"{us_id}", level=3)

    card_tbl = doc.add_table(rows=7, cols=2)
    card_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    card_data = [
        ("User Story ID", us_id),
        ("Epic", f"{epic_id} — {EPICS[epic_id][0]}"),
        ("Persona", f"{persona} ({PERSONAS.get(persona, '')})"),
        ("User Story", story_text),
        ("Linked FR / BR", f"FR: {fr_ids}  |  BR: {br_ids}"),
        ("Priority", priority),
        ("Story Points", str(sp)),
    ]
    for i, (k, v) in enumerate(card_data):
        set_cell_shading(card_tbl.rows[i].cells[0], ACCENT_BG)
        set_cell_text(card_tbl.rows[i].cells[0], k, bold=True, font_size=Pt(9), color=NAVY)
        set_cell_text(card_tbl.rows[i].cells[1], v, font_size=Pt(9))

    # Acceptance criteria
    ac_heading = doc.add_paragraph()
    ac_heading.paragraph_format.space_before = Pt(6)
    run = ac_heading.add_run("Acceptance Criteria:")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = NAVY
    run.font.name = 'Calibri'

    for ac_idx, (given, when, then) in enumerate(ac_list, 1):
        ac_tbl = doc.add_table(rows=3, cols=2)
        ac_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        # AC number header
        ac_items = [
            ("GIVEN", given),
            ("WHEN", when),
            ("THEN", then),
        ]
        for j, (kw, val) in enumerate(ac_items):
            set_cell_shading(ac_tbl.rows[j].cells[0], "D6E4F0" if kw == "GIVEN" else ("E8F5E9" if kw == "WHEN" else "FFF3E0"))
            set_cell_text(ac_tbl.rows[j].cells[0], kw, bold=True, font_size=Pt(8), color=NAVY)
            set_cell_text(ac_tbl.rows[j].cells[1], val, font_size=Pt(8.5))
        if ac_idx < len(ac_list):
            doc.add_paragraph()  # space between ACs

    doc.add_paragraph()  # space between stories

# ═══════════ 5. SUMMARY TABLE ═══════════
doc.add_page_break()
add_heading("5. User Story Summary Table", level=1)
add_body("The following table provides a consolidated view of all 16 user stories for quick reference:")

sum_tbl = doc.add_table(rows=1 + len(STORIES), cols=7)
sum_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(sum_tbl, ["US ID", "Persona", "Epic", "FR", "BR", "Priority", "SP"])

for i, s in enumerate(STORIES, 1):
    us_id, epic_id, persona, _, fr_ids, br_ids, priority, sp, _ = s
    add_data_row(sum_tbl, i, [us_id, persona, epic_id, fr_ids, br_ids, priority, str(sp)])

# Totals row
total_row = sum_tbl.add_row()
for cell in total_row.cells:
    set_cell_shading(cell, HEADER_BG)
set_cell_text(total_row.cells[0], "TOTAL", bold=True, font_size=Pt(8.5), color=WHITE)
set_cell_text(total_row.cells[5], f"{len(STORIES)} stories", bold=True, font_size=Pt(8.5), color=WHITE)
set_cell_text(total_row.cells[6], f"{sum(s[7] for s in STORIES)} SP", bold=True, font_size=Pt(8.5), color=WHITE)

# ═══════════ 6. SPRINT PLAN ═══════════
doc.add_page_break()
add_heading("6. Sprint Plan", level=1)
add_body(
    "The 16 stories are distributed across 5 two-day sprints (10 working days total). "
    "Each sprint delivers a usable vertical slice of the application. "
    "Velocity assumption: ~12 story points per sprint (one developer)."
)

sprint_tbl = doc.add_table(rows=1 + len(SPRINT_PLAN), cols=5)
sprint_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(sprint_tbl, ["Sprint / Days", "Epic Focus", "Stories", "Total SP", "Deliverable"])

for i, (sprint, focus, stories, sp, deliverable) in enumerate(SPRINT_PLAN, 1):
    add_data_row(sprint_tbl, i, [sprint, focus, stories, sp, deliverable])

# Sprint burndown note
add_info_box("📊 Velocity & Capacity",
    "• Team size: 1 developer (Sumedha)\n"
    "• Sprint length: 2 days\n"
    "• Average velocity: ~11.4 SP/sprint\n"
    f"• Total capacity: 10 days | {sum(s[7] for s in STORIES)} story points\n"
    "• Buffer: Sprint 5 has lighter load (12 SP) for integration testing and polish")

# ═══════════ 7. TRACEABILITY ═══════════
doc.add_page_break()
add_heading("7. Story-to-Requirement Traceability", level=1)
add_body(
    "The following matrix shows the bidirectional mapping between User Stories, "
    "Functional Requirements (FRs), and Business Requirements (BRs)."
)

trace_tbl = doc.add_table(rows=1 + len(STORIES), cols=5)
trace_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(trace_tbl, ["US ID", "FR(s)", "BR(s)", "OQ(s) Answered", "Epic"])

# Map stories to OQs via FR -> OQ mapping
FR_OQ_MAP = {
    "FR-01": "OQ-01", "FR-02": "OQ-02", "FR-03": "OQ-03", "FR-04": "OQ-10",
    "FR-05": "OQ-06", "FR-06": "OQ-09", "FR-07": "All", "FR-08": "OQ-04",
    "FR-09": "OQ-05", "FR-10": "OQ-07", "FR-11": "OQ-08", "FR-12": "OQ-10",
}

for i, s in enumerate(STORIES, 1):
    us_id, epic_id, _, _, fr_ids, br_ids, _, _, _ = s
    # Derive OQs from FRs
    oqs = set()
    for fr in fr_ids.replace(" ", "").split(","):
        oq = FR_OQ_MAP.get(fr, "")
        if oq:
            oqs.add(oq)
    oq_str = ", ".join(sorted(oqs)) if oqs else "—"
    add_data_row(trace_tbl, i, [us_id, fr_ids, br_ids, oq_str, epic_id])

# ═══════════ 8. COVERAGE ANALYSIS ═══════════
doc.add_page_break()
add_heading("8. Coverage Analysis", level=1)

# OQ coverage
add_heading("8.1 OQ Coverage", level=2)
oq_coverage = {
    "OQ-01": ["US-03", "US-15"],
    "OQ-02": ["US-05"],
    "OQ-03": ["US-06"],
    "OQ-04": ["US-07"],
    "OQ-05": ["US-08"],
    "OQ-06": ["US-09"],
    "OQ-07": ["US-10"],
    "OQ-08": ["US-11"],
    "OQ-09": ["US-03", "US-04"],
    "OQ-10": ["US-12", "US-13", "US-16"],
}

oq_tbl = doc.add_table(rows=1 + len(oq_coverage), cols=3)
oq_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(oq_tbl, ["OQ ID", "Covering User Stories", "Status"])

for i, (oq, us_list) in enumerate(oq_coverage.items(), 1):
    add_data_row(oq_tbl, i, [oq, ", ".join(us_list), "✅ Covered"])

add_body("")
add_info_box("✅ Coverage Result", "All 10 Operational Questions (OQ-01 to OQ-10) are covered by at least one User Story.\nAll 12 FRs (FR-01 to FR-12) are addressed.\nAll Must-priority BRs have corresponding Must-priority User Stories.")

# BR coverage
add_heading("8.2 BR Coverage", level=2)
br_coverage = {
    "BR-01": ["US-01", "US-03", "US-15"], "BR-02": ["US-03", "US-15"],
    "BR-03": ["US-05", "US-13"], "BR-04": ["US-06"],
    "BR-05": ["US-12", "US-13", "US-16"], "BR-06": ["US-09"],
    "BR-07": ["US-03", "US-04"], "BR-08": ["US-02", "US-14"],
    "BR-09": ["US-07"], "BR-10": ["US-08"],
    "BR-11": ["US-10", "US-16"], "BR-12": ["US-11"],
    "BR-13": ["Deferred (Won't priority)"],
}

br_tbl = doc.add_table(rows=1 + len(br_coverage), cols=3)
br_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(br_tbl, ["BR ID", "Covering User Stories", "Status"])

for i, (br, us_list) in enumerate(br_coverage.items(), 1):
    status = "⏳ Deferred" if "Deferred" in us_list[0] else "✅ Covered"
    add_data_row(br_tbl, i, [br, ", ".join(us_list), status])

# ═══════════ 9. ASSUMPTIONS ═══════════
doc.add_page_break()
add_heading("9. Assumptions", level=1)

assumptions = [
    "A-11: Since Step 6 (PRD) was skipped, Feature IDs are not assigned. User Stories map directly to FRs from the TRD (Step 7) and BRs from the Needs Analysis (Step 3).",
    "A-12: Story points are estimated for a single MCA student developer working approximately 6 productive hours per day.",
    "A-13: All user stories target the web-based Streamlit application; no mobile or native desktop interface is in scope.",
    "A-14: 'Under half a day' means ≤ 5 story points per story, achievable within 3-4 hours of focused development.",
    "A-15: The 5-sprint plan assumes sequential development; no parallel development streams.",
    "A-16: US-02 (methodology expander) and US-14 (methodology page) are complementary — US-02 is the per-widget inline drawer, US-14 is the full standalone reference page.",
    "A-17: BR-13 (live data refresh) is deferred as Won't priority per Step 3; no User Story is allocated for it in this phase.",
]

for a in assumptions:
    p = doc.add_paragraph(a, style='List Bullet')
    p.paragraph_format.space_after = Pt(3)

# ═══════════ SAVE ═══════════
output_path = r"c:\3rd sem\HCL-anti\docs\11_User_Stories.docx"
os.makedirs(os.path.dirname(output_path), exist_ok=True)
doc.save(output_path)
print(f"✅ Saved: {output_path}")
print(f"   Stories: {len(STORIES)}")
print(f"   Epics: {len(EPICS)}")
print(f"   Total SP: {sum(s[7] for s in STORIES)}")
