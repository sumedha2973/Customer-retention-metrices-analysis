"""
Step 14 — BUILD PLAN Document Generator (Word Copy)
Customer Retention Metrics Analysis (CRMA)
Sumedha | Roll 2503201450146 | MCA 2nd Yr | ABES
"""
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os, datetime

NAVY       = RGBColor(0x1B, 0x36, 0x5D)
STEEL      = RGBColor(0x2B, 0x4C, 0x7E)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
HEADER_BG  = "1B365D"
ALT_ROW    = "EDF1F7"
ACCENT_BG  = "E8EDF5"
CODE_BG    = "F4F6F9"

doc = Document()

for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.0)
    section.right_margin  = Cm(2.0)

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10)
style.paragraph_format.space_after = Pt(3)
style.paragraph_format.line_spacing = 1.15

def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_text(cell, text, bold=False, font_size=Pt(8.5), color=None, alignment=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
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

def add_data_row(table, row_idx, values, font_size=Pt(8)):
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
    p.paragraph_format.space_after = Pt(3)
    return p

def add_info_box(title, content):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, ACCENT_BG)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(3)
    run = p.add_run(title + "\n")
    run.bold = True; run.font.name = 'Calibri'; run.font.size = Pt(9.5); run.font.color.rgb = NAVY
    run2 = p.add_run(content)
    run2.font.name = 'Calibri'; run2.font.size = Pt(8.5)
    doc.add_paragraph()

def add_code_box(code_text, title=None):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, CODE_BG)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    if title:
        rt = p.add_run(title + "\n")
        rt.bold = True; rt.font.name = 'Consolas'; rt.font.size = Pt(8.5); rt.font.color.rgb = NAVY
    run = p.add_run(code_text)
    run.font.name = 'Consolas'; run.font.size = Pt(7.5)
    doc.add_paragraph()

# ─────────────────────────────────────────────────────────────
# Title Page
# ─────────────────────────────────────────────────────────────
doc.add_paragraph()
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run("Customer Retention Metrics Analysis")
run.bold = True; run.font.size = Pt(24); run.font.color.rgb = NAVY; run.font.name = 'Calibri'

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run2 = subtitle.add_run("Step 14 — Master Build Plan & Execution Blueprint")
run2.font.size = Pt(15); run2.font.color.rgb = STEEL; run2.font.name = 'Calibri'

doc.add_paragraph()

info_data = [
    ("Project", "Customer Retention Metrics Analysis (CRMA)"),
    ("Author", "Sumedha | Roll 2503201450146 | MCA 2nd Yr, ABES"),
    ("Mentor", "Ms. Anjali Pandey"),
    ("Version", "1.0"),
    ("Date", datetime.date.today().strftime("%d %B %Y")),
    ("Status", "Baselined for Step 15 Execution"),
    ("Scope", "15 Atomic Build Tasks (BUILD-01..15) + DoD + 10 Viva Notes"),
    ("Inputs", "01 Project Understanding, 03 Needs Analysis, 05 BRD, 07 TRD, 08 DB Design, 11 User Stories, 12 UI Framework, 13 Test Cases"),
]
info_tbl = doc.add_table(rows=len(info_data), cols=2)
info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(info_data):
    set_cell_shading(info_tbl.rows[i].cells[0], ACCENT_BG)
    set_cell_text(info_tbl.rows[i].cells[0], k, bold=True, font_size=Pt(9.5), color=NAVY)
    set_cell_text(info_tbl.rows[i].cells[1], v, font_size=Pt(9))

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 1. Project Summary & Links
# ─────────────────────────────────────────────────────────────
add_heading("1. Project Summary & Deliverable Registry", level=1)
add_body(
    "The Customer Retention Metrics Analysis (CRMA) platform is an enterprise analytics application developed for "
    "NoidaSoft Technologies Pvt. Ltd., an Indian mid-market B2B SaaS enterprise. The solution tracks, diagnoses, and "
    "forecasts customer retention, subscription churn, and recurring revenue decay over a 24-month horizon (January 2024 to December 2025). "
    "Operating over an embedded SQLite database (`data/app.db`) comprising 200 customer accounts, 229 contract lifecycles, "
    "704 customer service tickets, and 2,489 monthly product telemetry records, the application empowers executive leadership and "
    "customer success teams with real-time operational insights. Features include interactive cohort survival heatmaps, multi-attribute "
    "firmographic slicing, inferential statistical hypothesis testing (Mann-Whitney U, Student's t-test, Chi-Square, Pearson/Spearman correlations), "
    "an automated 4-factor customer health risk scoring engine (RULE-05), and comprehensive analytical governance."
)

doc_links = [
    ("docs/01_Project_Understanding.docx", "Problem statement, 10 Stakeholders (SH-01..10), 10 Operational Questions (OQ-01..10), SMART objectives"),
    ("docs/02_HMW_Statements.docx", "Design thinking synthesis, stakeholder pain points, 10 How Might We statements (HMW-01..10)"),
    ("docs/03_Needs_Analysis.docx", "12 Business Requirements (BR-01..12) with MoSCoW prioritization, measurement criteria, and gap analysis"),
    ("docs/04_Traceability_Matrix.xlsx", "Complete bidirectional RTM (v2.0 Final) tracing OQ -> HMW -> BR -> DB -> US -> UI -> TC across 15 rows"),
    ("docs/05_BRD.docx", "Formal Business Requirements Document, RACI matrix, 7 business calculation rules (RULE-01..07)"),
    ("docs/07_TRD.docx", "Technical Architecture, 12 Functional Requirements (FR-01..12), 5 NFRs, Streamlit caching & deployment"),
    ("docs/08_Database_Design.xlsx", "Relational schema (5 tables, 1 analytical view, 33 columns), DDL, and 11 KPI definitions"),
    ("docs/09_ER_Diagram.docx / .png", "Relational Entity-Relationship diagram, cardinality definitions, and schema normalization narrative"),
    ("data/crma_data.xlsx / app.db", "Verified synthetic dataset (Seed 42) and data quality audit report"),
    ("docs/11_User_Stories.docx", "16 Agile User Stories (US-01..16) across 5 Epics with Given/When/Then acceptance criteria and 5-sprint plan"),
    ("docs/12_UI_Framework.docx", "Multipage UI layout, ASCII wireframes for 8 screens (UI-01..08), Tailwind design tokens"),
    ("docs/13_Test_Cases.docx", "Test specification with 20 Functional Tests (TC-F-01..20) and 35 Automated Technical Tests (TC-T-01..35)"),
]

tbl_docs = doc.add_table(rows=1 + len(doc_links), cols=2)
tbl_docs.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_docs, ["Document Deliverable", "Contents & Purpose"], font_size=Pt(8.5))
for i, row in enumerate(doc_links, 1):
    add_data_row(tbl_docs, i, row, font_size=Pt(8))

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 2. Folder Structure
# ─────────────────────────────────────────────────────────────
add_heading("2. Final Target Folder Structure", level=1)
add_body("The project adheres strictly to the modular architectural layout specified in TRD Section 2:")

fs_code = """c:\\3rd sem\\HCL-anti\\
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
├── BUILD_PLAN.md                 # Markdown execution plan for AI/developer
├── README.md                     # Comprehensive project documentation & local deployment runbook
└── requirements.txt              # Pinned deterministic dependencies"""
add_code_box(fs_code, title="Directory Layout")

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 3. Build Tasks (BUILD-01 to BUILD-15)
# ─────────────────────────────────────────────────────────────
add_heading("3. Master Build Tasks (BUILD-01 to BUILD-15)", level=1)
add_body(
    "Step 15 executes these 15 discrete tasks sequentially. Each task specifies its operational goal, "
    "target files, linked user stories/features, exact implementation directives, verification criteria, "
    "and mandatory test assertions."
)

tasks = [
    ("BUILD-01", "Setup, Environment & Dependencies", "requirements.txt, .streamlit/config.toml", "Setup, NFR-05",
     "Pin requirements.txt strictly; configure .streamlit/config.toml (primaryColor=#2B4C7E).", "pip list / imports", "TC-T-01"),
    ("BUILD-02", "Data Layer & Validation Cache", "src/data_loader.py", "US-01, FR-01..02, NFR-01",
     "Implement SQLite read-only URI connection, @st.cache_data(ttl=3600), categorical downcasting, and validation assertions.", "load_analytical_view() runs < 200ms", "TC-T-01..14"),
    ("BUILD-03", "Metrics & Cohort Engine", "src/metrics.py", "US-03, US-04, US-06, FR-01,03,06",
     "Implement vectorized functions for CCR %, CRR %, MRR movements, CLV, and triangular cohort survival matrix with zero guards.", "Formulas verify on known data", "TC-T-15..19, 26..27"),
    ("BUILD-04", "Statistical Tests & Risk Scoring", "src/stats.py, src/risk_scoring.py", "US-08..12, FR-04,05,09..11",
     "Wrap SciPy tests (Mann-Whitney U, t-test, Chi-square, correlation); compute RULE-05 4-factor risk score [0,100].", "p-values in [0,1]; risk tiers High/Med/Low", "TC-T-20..25"),
    ("BUILD-05", "UI Components & Global Filters", "src/ui.py, src/components.py, src/filters.py", "US-02, US-05, FR-02,07",
     "Implement methodology drawer, Plotly chart wrapper, st.session_state filter manager with reset button.", "Session state persists; reset works", "TC-F-07..08, TC-T-34"),
    ("BUILD-06", "UI-01: Executive Overview", "app.py", "US-03, US-04, US-15, FR-01,06, UI-01",
     "Build landing page: 5 Tailwind KPI cards with deltas, 24m CCR/CRR dual line chart, 24m MRR trend chart, methodology drawers.", "app.py renders in < 2s", "TC-F-01..04, TC-T-28,30,31,33"),
    ("BUILD-07", "UI-02: Cohort Analysis", "pages/1_Cohorts.py", "US-06, FR-03, UI-02",
     "Build triangular survival heatmap (Plotly) + multi-line survival decay curves + summary callout.", "Heatmap cells match retention", "TC-F-05..06, TC-T-32"),
    ("BUILD-08", "UI-03: Segment Comparison", "pages/2_Segments.py", "US-05, US-07, US-08, FR-02,08,09, UI-03",
     "Multi-attribute retention bar charts, CLV grouped bars by plan x industry, channel longevity with Chi-square badge.", "Slicing works; empty alert works", "TC-F-07..10, 19, TC-T-32"),
    ("BUILD-09", "UI-04: Churn Drivers", "pages/3_Drivers.py", "US-09..11, FR-05,10,11, UI-04",
     "Support SLA box plots + MW-U card, 90-day pre-churn telemetry drop + t-test card, CSAT correlation scatter + sample warning.", "All 3 stat cards render valid text", "TC-F-11..13, TC-T-32"),
    ("BUILD-10", "UI-05: Risk Roster & Export", "pages/4_Risk_Roster.py", "US-12..13, US-16, FR-04,12, UI-05",
     "Risk count cards (High/Med/Low), sortable roster table with colored badges, min-score slider, and client-side CSV download.", "CSV exports filtered table", "TC-F-14..16, TC-T-32"),
    ("BUILD-11", "UI-06: Data Quality & EDA", "pages/5_Data_Quality.py", "US-01, UI-06",
     "Dataset metrics cards, column missing-value heatmap (DQ-01), feature distribution histograms, and structured DQ audit table.", "DQ-01..03 clearly audited", "TC-T-10, TC-T-32"),
    ("BUILD-12", "UI-07 & 08: Methodology & About", "pages/6_Methodology.py, pages/7_About.py", "US-02, US-14, FR-07, UI-07..08",
     "Render 11 KPI formal definitions (MSR-01..11), statistical specs, 5-table data dictionary, and ABES student/mentor attribution.", "MSR-01..11 match Step 8", "TC-F-17..18, TC-T-32"),
    ("BUILD-13", "Test Suite Verification", "tests/test_*.py", "NFR-01, NFR-02, NFR-05",
     "Execute full pytest suite across data, analysis, and AppTest smoke layers; verify >= 70% coverage.", "pytest returns 35 passed", "All 35 TC-T tests"),
    ("BUILD-14", "README & Documentation", "README.md, .gitignore", "Handover, Governance",
     "Author comprehensive enterprise README.md with architecture diagram, local setup commands, and viva guide.", "README renders cleanly", "File check"),
    ("BUILD-15", "Cloud Deployment Packaging", "Git root", "Cloud Deployment",
     "Prepare repository for Streamlit Community Cloud; verify relative paths and clean git tracking.", "Cloud URL live HTTP 200", "Live cloud probe"),
]

tbl_tasks = doc.add_table(rows=1 + len(tasks), cols=7)
tbl_tasks.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_tasks, ["Task ID", "Goal", "Target Files", "Requirements", "Exact Directives", "Verification", "Must-Pass Tests"], font_size=Pt(7.5))
for i, row in enumerate(tasks, 1):
    add_data_row(tbl_tasks, i, row, font_size=Pt(7))

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 4. Definition of Done & Commands
# ─────────────────────────────────────────────────────────────
add_heading("4. Definition of Done (DoD) & Execution Commands", level=1)
add_body(
    "The project is formally declared complete and ready for final academic evaluation when: "
    "(1) 100% of Must-Have Business Requirements (BR-01..08) and Functional Requirements (FR-01..12) are operational; "
    "(2) All 35 automated technical tests pass with zero failures; (3) All 20 manual functional tests pass; "
    "(4) SQLite database satisfies all relational and domain constraints; (5) UI renders with consistent enterprise styling; "
    "(6) Zero raw tracebacks appear under empty selections; (7) Initial load time is under 3 seconds; and "
    "(8) Comprehensive README and live cloud deployment are accessible."
)

cmd_code = """# Local Environment Setup & Installation
python -m venv venv
venv\\Scripts\\activate
pip install --upgrade pip
pip install -r requirements.txt

# Run Local Streamlit Application
streamlit run app.py

# Execute Full Automated Pytest Test Suite
pytest tests/ -v

# Run Test Coverage Audit
pytest --cov=src tests/ --cov-report=term-missing"""
add_code_box(cmd_code, title="Standard Execution Commands")

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 5. Viva Preparation: 10 Critical Questions & Answers
# ─────────────────────────────────────────────────────────────
add_heading("5. Viva Preparation: 10 Critical Questions & Answers", level=1)

viva_qa = [
    ("Q1: Business difference between Customer Churn Rate (CCR) and Net Revenue Retention (NRR)?",
     "CCR measures account/logo loss percentage regardless of contract size. NRR measures net retained recurring revenue from existing cohorts, factoring in account expansion (upgrades) and contraction (downgrades) alongside churn."),
    ("Q2: Why use embedded SQLite rather than in-memory pandas or flat CSV files?",
     "SQLite enforces relational schema constraints (PK uniqueness, FK referential integrity, domain CHECKs) at the storage level, supports SQL pre-aggregation via indexed views, and minimizes cloud memory footprint with zero credential needs."),
    ("Q3: Why choose the Mann-Whitney U test for support ticket analysis (OQ-06)?",
     "Support ticket volumes and turnaround hours are highly skewed count/duration distributions violating Gaussian normality. Mann-Whitney U is a robust non-parametric test comparing stochastic dominance without normality assumptions."),
    ("Q4: How does the composite customer health risk score work (RULE-05)?",
     "Applies a 4-factor weighted linear model: 35% login drop (trailing 30d vs prior 60d), 25% open tickets, 20% low CSAT, and 20% renewal proximity (<= 60 days). Clamped to [0, 100], scores >= 70 alert High Risk."),
    ("Q5: How does @st.cache_data improve performance and prevent stale data?",
     "Serializes function outputs in memory across user filter reruns, preventing repeated SQLite disk I/O and achieving sub-second renders. Stale data is prevented by setting ttl=3600 (1 hour) and using immutable read-only DB connections."),
    ("Q6: Why inject intentional data quality issues (DQ-01, DQ-02, DQ-03)?",
     "Real-world SaaS data is dirty. Injecting 65.8% survey nulls mirrors realistic survey response rates; duplicate tickets test deduplication; and 5 extreme turnaround outliers (>200h) test pipeline resistance against skewed distributions."),
    ("Q7: Why load Tailwind CSS via st.components.v1.html() rather than plain Streamlit markdown?",
     "Standard Streamlit metric widgets offer limited styling control. Tailwind via an iframe bridge enables custom responsive enterprise cards with drop shadows, custom gradients, rounded pills, and dynamic semantic delta badges."),
    ("Q8: How does the application prevent SQL injection vulnerabilities?",
     "Database connections are opened read-only (mode=ro URI). All data extraction runs through pre-defined view selections or parameterized queries in src/data_loader.py; zero dynamic string concatenation is passed to SQLite."),
    ("Q9: How do you interpret a triangular cohort survival heatmap?",
     "Rows represent signup cohorts (YYYY-MM); columns represent lifecycle milestones (M1, M3, M6, M12). The triangular shape reflects cohort age; reading a row shows cohort decay, while reading a column compares cohort quality over time."),
    ("Q10: What is the significance of the Gherkin Given/When/Then syntax in your test cases?",
     "Provides an unambiguous contract between business requirements and QA testing. 'Given' sets preconditions, 'When' defines the user action, and 'Then' asserts expected behavior, mapping directly into pytest assertions."),
]

tbl_viva = doc.add_table(rows=1 + len(viva_qa), cols=2)
tbl_viva.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_viva, ["Viva Question", "Authoritative Answer"], font_size=Pt(8.5))
for i, row in enumerate(viva_qa, 1):
    add_data_row(tbl_viva, i, row, font_size=Pt(8))

# Save
output_path = r"c:\3rd sem\HCL-anti\docs\14_BUILD_PLAN.docx"
doc.save(output_path)
print(f"✅ Saved: {output_path}")
print(f"   Tasks: {len(tasks)}")
print(f"   Viva Q&A: {len(viva_qa)}")
