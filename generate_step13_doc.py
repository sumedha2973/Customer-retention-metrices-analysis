"""
Step 13 — Test Cases Document Generator
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
run2 = subtitle.add_run("Step 13 — Test Cases (Functional & Technical)")
run2.font.size = Pt(15); run2.font.color.rgb = STEEL; run2.font.name = 'Calibri'

doc.add_paragraph()

info_data = [
    ("Project", "Customer Retention Metrics Analysis (CRMA)"),
    ("Author", "Sumedha | Roll 2503201450146 | MCA 2nd Yr, ABES"),
    ("Mentor", "Ms. Anjali Pandey"),
    ("Version", "1.0"),
    ("Date", datetime.date.today().strftime("%d %B %Y")),
    ("Status", "Baselined"),
    ("Test Scope", "20 Functional Tests (TC-F-01..20) + 35 Technical Tests (TC-T-01..35)"),
    ("Input Documents", "03 Needs Analysis (BRs), 07 TRD (FRs), 11 User Stories (ACs), 12 UI Framework"),
]
info_tbl = doc.add_table(rows=len(info_data), cols=2)
info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(info_data):
    set_cell_shading(info_tbl.rows[i].cells[0], ACCENT_BG)
    set_cell_text(info_tbl.rows[i].cells[0], k, bold=True, font_size=Pt(9.5), color=NAVY)
    set_cell_text(info_tbl.rows[i].cells[1], v, font_size=Pt(9))

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 1. Test Strategy & Overview
# ─────────────────────────────────────────────────────────────
add_heading("1. Test Strategy & Quality Assurance Overview", level=1)
add_body(
    "This document establishes the comprehensive quality assurance and testing suite for the CRMA platform. "
    "The testing strategy covers two primary paradigms: (A) Functional Testing (TC-F) driven directly by "
    "the Gherkin Acceptance Criteria (Given/When/Then) defined across all 16 User Stories, ensuring complete user-facing "
    "validation of filters, KPI cards, Plotly charts, responsive layouts, data exports, and error boundary fallbacks; "
    "and (B) Technical Automated Unit & Smoke Testing (TC-T) implemented via pytest across tests/test_data.py, "
    "tests/test_analysis.py, and tests/test_app.py."
)

add_info_box("📊 Test Suite Metrics & Verification Targets",
    "• Total Test Cases: 55 (20 Functional + 35 Technical)\n"
    "• Functional Tests: TC-F-01 to TC-F-20 (covering filters, edge cases, zero results, downloads, responsive UI)\n"
    "• Technical Tests: TC-T-01 to TC-T-35 (data integrity, statistical rigor, KPI formulas, AppTest smoke tests)\n"
    "• Automated Test Files: tests/test_data.py (14 tests), tests/test_analysis.py (13 tests), tests/test_app.py (8 tests)\n"
    "• Requirement Coverage: 100% of Business Requirements (BR-01..12), User Stories (US-01..16), and Operational Questions (OQ-01..10)\n"
    "• Execution Baseline Status: 'Not run' (Functional manual tests) / 'Baselined & Ready for Phase 15 Execution' (Pytest suite)")

# ─────────────────────────────────────────────────────────────
# 2. Section A: Functional Test Cases (TC-F-01 to TC-F-20)
# ─────────────────────────────────────────────────────────────
add_heading("2. Section A: Functional Test Cases (TC-F-01 to TC-F-20)", level=1)
add_body(
    "Functional test cases validate application behaviors against stakeholder requirements. "
    "Each test prescribes concrete preconditions, user steps, input data, expected outcomes, priority, and initial status."
)

tc_f_data = [
    ("TC-F-01", "BR-01, BR-02\nUS-03, US-15", "UI-01", "App running; data/app.db loaded",
     "1. Navigate to Executive Overview.\n2. Observe 5 top KPI cards.",
     "Latest active month (2025-12)",
     "Cards display CCR %, CRR %, MRR Lost (₹), NRR %, and Active count with MoM delta badges.", "Must", "Not run"),
    ("TC-F-02", "BR-01, BR-02\nUS-03", "UI-01", "Overview page open",
     "1. Change Month selector in sidebar to '2024-06'.\n2. Verify KPI recalculation.",
     "Month = '2024-06'",
     "All 5 KPI cards update immediately to reflect 2024-06 metrics; deltas calculate vs 2024-05.", "Must", "Not run"),
    ("TC-F-03", "BR-01, BR-02\nUS-15", "UI-01", "Overview page open",
     "1. Inspect 24-month retention trend chart.\n2. Hover over month 2025-03.",
     "Chart hover on 2025-03",
     "Plotly dual-axis line shows CCR % and CRR %; tooltip reveals exact rates summing to 100%.", "Must", "Not run"),
    ("TC-F-04", "BR-07\nUS-04", "UI-01", "Overview page open",
     "1. Scroll to MRR Trend chart.\n2. Toggle series visibility in legend.",
     "Click 'Churned MRR' legend item",
     "Gross, Churned, and Net MRR lines render in ₹ INR; toggling legend filters series cleanly.", "Must", "Not run"),
    ("TC-F-05", "BR-04\nUS-06", "UI-02", "Cohort Analysis page open",
     "1. Navigate to Cohorts.\n2. Inspect triangular survival heatmap.",
     "24 signup cohorts × 4 lifecycle intervals (M1, M3, M6, M12)",
     "Heatmap renders triangular grid; values decrease from left to right; cells color-coded.", "Must", "Not run"),
    ("TC-F-06", "BR-04\nUS-06", "UI-02", "Cohorts page open",
     "1. Hover over cell (Cohort 2024-01, Month 12).\n2. Verify tooltip data.",
     "Cell 2024-01 @ M12",
     "Tooltip shows exact cohort size, active count at M12, and retention %.", "Should", "Not run"),
    ("TC-F-07", "BR-03\nUS-05", "UI-03", "Segment Comparison page open",
     "1. Select 'Starter' in Plan filter.\n2. Select 'FinTech' in Industry filter.",
     "Plan='Starter', Industry='FinTech'",
     "All visual widgets recalculate to isolate Starter FinTech segment; filter badge shows active criteria.", "Must", "Not run"),
    ("TC-F-08", "BR-03\nUS-05", "UI-03", "Filtered segment comparison open",
     "1. Click 'Reset Filters' button in sidebar.",
     "Click reset button",
     "All dropdowns revert to 'All'; charts and KPI cards restore to the full 200 customer baseline.", "Must", "Not run"),
    ("TC-F-09", "BR-09\nUS-07", "UI-03", "Segment Comparison page open",
     "1. Scroll to Customer Lifetime Value chart.\n2. Compare Enterprise vs Starter bars.",
     "Segment CLV grouped bar chart",
     "CLV grouped bar chart displays mean CLV in ₹; Enterprise tier shows significantly higher CLV.", "Should", "Not run"),
    ("TC-F-10", "BR-10\nUS-08", "UI-03", "Segment Comparison page open",
     "1. Scroll to Channel Longevity section.\n2. Verify ranking and significance badge.",
     "5 channels (Referral, Organic, Paid Social, Direct, Partner)",
     "Horizontal bar chart ranks channels by 12m retention; Chi-Square significance badge displays p-value.", "Should", "Not run"),
    ("TC-F-11", "BR-06\nUS-09", "UI-04", "Churn Drivers page open",
     "1. Navigate to Churn Drivers.\n2. Inspect Support SLA box plots.",
     "Support ticket counts & resolution hours for churned vs retained",
     "Side-by-side box plots reveal distributions; Mann-Whitney U test card shows statistic, p-value, and text.", "Must", "Not run"),
    ("TC-F-12", "BR-11\nUS-10", "UI-04", "Churn Drivers page open",
     "1. Scroll to Pre-Churn Telemetry section.\n2. Observe 90-day trajectory lines.",
     "Login days & feature clicks across t-3, t-2, t-1 months",
     "Dual-line chart displays noticeable usage drop for churned accounts; t-test card reports p < 0.05.", "Should", "Not run"),
    ("TC-F-13", "BR-12\nUS-11", "UI-04", "Churn Drivers page open",
     "1. Scroll to Satisfaction Correlation section.\n2. Inspect scatter plot.",
     "CSAT/NPS scores vs renewal status flag",
     "Scatter plot with trendline renders; Pearson r and Spearman rho cards appear; warning badge indicates sample size.", "Could", "Not run"),
    ("TC-F-14", "BR-05\nUS-12", "UI-05", "Risk Roster page open",
     "1. Navigate to Risk Roster.\n2. Inspect summary counts and table sorting.",
     "Active customer accounts with health_risk_score",
     "Risk cards show counts (High/Medium/Low); roster table is sorted descending by risk score.", "Must", "Not run"),
    ("TC-F-15", "BR-05\nUS-13", "UI-05", "Risk Roster page open",
     "1. Filter Risk Tier = 'High'.\n2. Filter Plan = 'Enterprise'.",
     "Tier='High', Plan='Enterprise'",
     "Table displays only High-Risk Enterprise customers with red badges; row count updates correctly.", "Must", "Not run"),
    ("TC-F-16", "BR-05, BR-11\nUS-16", "UI-05", "Filtered Risk Roster open",
     "1. Click 'Download CSV' button in top toolbar.\n2. Open downloaded file.",
     "Filtered roster subset",
     "Browser downloads 'at_risk_customers_YYYYMMDD_HHMMSS.csv'; columns match table; values match active filters.", "Should", "Not run"),
    ("TC-F-17", "BR-08\nUS-02", "UI-01..05", "Any analytical chart card",
     "1. Click 'ℹ Methodology' expander beneath chart.\n2. Review documentation.",
     "Methodology expander toggle",
     "Drawer expands showing computational SQL/pandas formula, source tables, row counts, and assumptions.", "Must", "Not run"),
    ("TC-F-18", "BR-08\nUS-14", "UI-08", "Methodology page open",
     "1. Navigate to Methodology & Data Dictionary.\n2. Verify 11 measures and table schemas.",
     "Complete governance reference catalog",
     "Comprehensive catalog renders all 11 KPIs (MSR-01..11), statistical test formulas, and 5-table data dictionary.", "Must", "Not run"),
    ("TC-F-19", "BR-03\nUS-05", "UI-03", "Segment Comparison open",
     "1. Select an impossible filter combo (e.g. Micro + Enterprise + North where count=0).",
     "Zero-record filter selection",
     "App handles empty subset gracefully; renders informative notice 'No matching customer records found'; zero crash.", "Must", "Not run"),
    ("TC-F-20", "NFR-04\nAll US", "All screens", "Desktop browser (1920x1080)",
     "1. Emulate tablet viewport (768px).\n2. Emulate mobile viewport (375px).",
     "Chrome DevTools responsive emulation",
     "Layout shifts to 2-column then 1-column; KPI cards stack vertically; charts remain legible; horizontal scroll disabled.", "Should", "Not run"),
]

tbl_f = doc.add_table(rows=1 + len(tc_f_data), cols=9)
tbl_f.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_f, ["ID", "Linked Req/US", "UI", "Pre-condition", "Steps", "Test Data", "Expected Result", "Priority", "Status"], font_size=Pt(8))

for i, row in enumerate(tc_f_data, 1):
    add_data_row(tbl_f, i, row, font_size=Pt(7.5))

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 3. Section B: Technical Automated Test Cases (TC-T-01 to TC-T-35)
# ─────────────────────────────────────────────────────────────
add_heading("3. Section B: Technical Automated Test Cases (TC-T-01 to TC-T-35)", level=1)
add_body(
    "Technical test cases represent automated unit, integration, and smoke assertions implemented in Python/pytest. "
    "They rigorously test database schema integrity, domain bounds, formula accuracy on controlled datasets, "
    "scientific computing packages, and Streamlit execution speed."
)

tc_t_data = [
    # Data Layer (tests/test_data.py)
    ("TC-T-01", "tests/test_data.py::test_TC_T_01_db_file_exists", "SQLite DB file exists and opens in read-only mode", "data/app.db connects cleanly; sqlite_version() returns string"),
    ("TC-T-02", "tests/test_data.py::test_TC_T_02_all_tables_exist", "Presence of 5 relational tables in sqlite_master", "dim_customers, dim_plans, fact_subscriptions, fact_support_tickets, fact_monthly_usage all exist"),
    ("TC-T-03", "tests/test_data.py::test_TC_T_03_row_counts", "Row counts match design specifications", "dim_customers (200), dim_plans (3), fact_subscriptions (229), tickets (704), usage (2489) within design limits"),
    ("TC-T-04", "tests/test_data.py::test_TC_T_04_view_columns", "Analytical view vw_customer_retention_analytics structure", "View exists and exposes >= 30 analytical columns (actual: 33)"),
    ("TC-T-05", "tests/test_data.py::test_TC_T_05_pk_uniqueness", "Primary Key uniqueness across all tables", "customer_id, plan_id, subscription_id, ticket_id, usage_id have 0 duplicates"),
    ("TC-T-06", "tests/test_data.py::test_TC_T_06_foreign_key_integrity", "Foreign key referential integrity", "Zero orphan rows across fact_subscriptions, fact_support_tickets, fact_monthly_usage"),
    ("TC-T-07", "tests/test_data.py::test_TC_T_07_domain_ranges", "Domain bounds for financial, tenure, usage, & survey fields", "monthly_fee >= 0, login_days in [0,31], CSAT in [1,5], NPS in [0,10], tenure >= 1, resolution_hours >= 0"),
    ("TC-T-08", "tests/test_data.py::test_TC_T_08_required_columns", "Mandatory schema column coverage", "All required enterprise fields present in every table DataFrame"),
    ("TC-T-09", "tests/test_data.py::test_TC_T_09_mandatory_non_null", "Mandatory non-nullable columns check", "0 null values across critical identifiers, contract dates, and status codes"),
    ("TC-T-10", "tests/test_data.py::test_TC_T_10_no_duplicate_tickets", "DQ-02 ticket deduplication resolution", "Database contains 704 unique ticket IDs; duplicate CSV rows cleaned upon ingestion"),
    ("TC-T-11", "tests/test_data.py::test_TC_T_11_exactly_three_plans", "Subscription plan catalog integrity", "dim_plans contains exactly Starter, Professional, and Enterprise tiers"),
    ("TC-T-12", "tests/test_data.py::test_TC_T_12_is_churned_binary", "Binary domain check for churn status flag", "fact_subscriptions.is_churned contains only values in {0, 1}"),
    ("TC-T-13", "tests/test_data.py::test_TC_T_13_risk_score_range", "Health risk composite score bounds", "health_risk_score in vw_customer_retention_analytics strictly within [0.0, 100.0]"),
    ("TC-T-14", "tests/test_data.py::test_TC_T_14_risk_tier_values", "Risk tier categorical values domain", "risk_tier contains only 'High', 'Medium', or 'Low'"),
    
    # Analysis & Math Layer (tests/test_analysis.py)
    ("TC-T-15", "tests/test_analysis.py::test_TC_T_15_churn_rate_formula", "Customer Churn Rate (CCR %) formula on known dataset", "Exact match to manual calculation (50.0% on 3 churned / 6 active)"),
    ("TC-T-16", "tests/test_analysis.py::test_TC_T_16_retention_rate_formula", "Customer Retention Rate identity check", "CRR % strictly equals (100.0 - CCR %) across boundaries [0, 100]"),
    ("TC-T-17", "tests/test_analysis.py::test_TC_T_17_mrr_churn_formula", "Monthly Recurring Revenue churn sum formula", "Sum of churned subscription fees exactly equals INR 42,500.0 on test fixture"),
    ("TC-T-18", "tests/test_analysis.py::test_TC_T_18_clv_formula", "Customer Lifetime Value calculation accuracy", "CLV = AVG(MRR) * AVG(Tenure) verified against hand-calculated test fixture"),
    ("TC-T-19", "tests/test_analysis.py::test_TC_T_19_cohort_survival", "Cohort survival decay computation", "Survival percentage at Month 6 matches 5/6 = 83.33% on known cohort data"),
    ("TC-T-20", "tests/test_analysis.py::test_TC_T_20_risk_score_weights", "Risk score composite weights & clamping (RULE-05)", "Weights (35/25/20/20) correctly applied; clamped strictly to [0, 100]"),
    ("TC-T-21", "tests/test_analysis.py::test_TC_T_21_correlation_positive", "Pearson correlation mathematical sanity", "r is bounded in [-1.0, 1.0]; positive correlation between CSAT and renewal"),
    ("TC-T-22", "tests/test_analysis.py::test_TC_T_22_mannwhitney_runs", "Mann-Whitney U statistical execution via SciPy", "Returns non-negative U-stat and p-value in [0.0, 1.0] for ticket distributions"),
    ("TC-T-23", "tests/test_analysis.py::test_TC_T_23_ttest_runs", "Two-sample Student's t-test execution", "Calculates valid t-statistic and significant p-value on diverging usage means"),
    ("TC-T-24", "tests/test_analysis.py::test_TC_T_24_chisquare_runs", "Chi-Square test of independence on channels", "Executes on 2x5 contingency matrix; DOF = 4; p-value bounded in [0.0, 1.0]"),
    ("TC-T-25", "tests/test_analysis.py::test_TC_T_25_spearman_runs", "Spearman rank correlation coefficient", "Returns valid monotonic rho in [-1.0, 1.0] and p-value for survey rankings"),
    ("TC-T-26", "tests/test_analysis.py::test_TC_T_26_churn_rate_empty_guard", "Zero-division error handling on empty data slice", "Returns 0.0 without raising ZeroDivisionError when denominator is zero"),
    ("TC-T-27", "tests/test_analysis.py::test_TC_T_27_nrr_above_100", "Net Revenue Retention formula logic", "NRR exceeds 100% when expansion revenue outpaces contraction & churn"),
    
    # App & Smoke Layer (tests/test_app.py)
    ("TC-T-28", "tests/test_app.py::test_TC_T_28_app_py_exists", "Root application entrypoint existence", "app.py exists and file size > 100 bytes"),
    ("TC-T-29", "tests/test_app.py::test_TC_T_29_page_files_exist", "Presence of 7 multipage script files in pages/", "All 7 page files present in pages/ directory"),
    ("TC-T-30", "tests/test_app.py::test_TC_T_30_app_starts", "Streamlit AppTest initialization smoke test", "App starts and executes without unhandled Python exceptions"),
    ("TC-T-31", "tests/test_app.py::test_TC_T_31_overview_has_kpis", "Executive overview KPI rendering verification", "AppTest detects metrics or HTML card components on landing page"),
    ("TC-T-32", "tests/test_app.py::test_TC_T_32_pages_render", "Multipage AppTest rendering sweep", "Each of the 7 subpages renders successfully without exceptions"),
    ("TC-T-33", "tests/test_app.py::test_TC_T_33_load_time_under_3s", "Initial render latency benchmark (NFR-01)", "Total app startup and initial execution completes in under 3.0 seconds"),
    ("TC-T-34", "tests/test_app.py::test_TC_T_34_sidebar_filters_present", "Sidebar filter controls rendering", "Sidebar contains functional selectbox/multiselect filter widgets"),
    ("TC-T-35", "tests/test_app.py::test_TC_T_35_no_raw_tracebacks", "Global error boundary validation (NFR-02)", "Zero raw unhandled exception tracebacks exposed across all pages"),
]

tbl_t = doc.add_table(rows=1 + len(tc_t_data), cols=4)
tbl_t.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_t, ["Test ID", "Test Function / Module", "Check Description", "Expected Result"], font_size=Pt(8))

for i, row in enumerate(tc_t_data, 1):
    add_data_row(tbl_t, i, row, font_size=Pt(7.5))

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 4. Final Traceability & Requirement Mapping
# ─────────────────────────────────────────────────────────────
add_heading("4. Final Traceability Matrix & Test Coverage Analysis", level=1)
add_body(
    "Every Business Requirement (BR-01 through BR-12) is mapped to at least one functional test and at least "
    "one technical automated test. BR-13 (live database refresh) is explicitly documented as Deferred (Phase 2)."
)

req_cov = [
    ("BR-01", "Official Churn & Retention Definition", "TC-F-01, TC-F-02, TC-F-03", "TC-T-15, TC-T-16, TC-T-26", "✅ Fully Covered"),
    ("BR-02", "24-Month Trend Analysis", "TC-F-01, TC-F-03", "TC-T-03, TC-T-15, TC-T-16", "✅ Fully Covered"),
    ("BR-03", "Multidimensional Customer Slicing", "TC-F-07, TC-F-08, TC-F-19", "TC-T-08, TC-T-11, TC-T-34", "✅ Fully Covered"),
    ("BR-04", "12-Month Cohort Survival Matrix", "TC-F-05, TC-F-06", "TC-T-08, TC-T-19", "✅ Fully Covered"),
    ("BR-05", "Health Risk Scoring & At-Risk Roster", "TC-F-14, TC-F-15, TC-F-16", "TC-T-13, TC-T-14, TC-T-20", "✅ Fully Covered"),
    ("BR-06", "Support SLA & Resolution Impact", "TC-F-11", "TC-T-07, TC-T-10, TC-T-22", "✅ Fully Covered"),
    ("BR-07", "MRR Churn & Revenue Retention", "TC-F-01, TC-F-04", "TC-T-07, TC-T-17, TC-T-27", "✅ Fully Covered"),
    ("BR-08", "Transparent Metric Governance", "TC-F-17, TC-F-18", "TC-T-04, TC-T-08", "✅ Fully Covered"),
    ("BR-09", "Customer Lifetime Value (CLV)", "TC-F-09", "TC-T-07, TC-T-18", "✅ Fully Covered"),
    ("BR-10", "Channel Longevity Analysis", "TC-F-10", "TC-T-08, TC-T-24", "✅ Fully Covered"),
    ("BR-11", "Pre-Churn Usage Telemetry", "TC-F-12, TC-F-16", "TC-T-07, TC-T-23", "✅ Fully Covered"),
    ("BR-12", "Survey Satisfaction Correlation", "TC-F-13", "TC-T-07, TC-T-21, TC-T-25", "✅ Fully Covered"),
    ("BR-13", "Live Automated Data Refresh", "Deferred (Phase 2)", "Deferred (Phase 2)", "⏳ Deferred (Won't)"),
]

tbl_cov = doc.add_table(rows=1 + len(req_cov), cols=5)
tbl_cov.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(tbl_cov, ["BR ID", "Requirement Description", "Functional Tests (TC-F)", "Technical Tests (TC-T)", "Coverage Status"], font_size=Pt(8.5))

for i, row in enumerate(req_cov, 1):
    add_data_row(tbl_cov, i, row, font_size=Pt(8))

doc.add_page_break()

# ─────────────────────────────────────────────────────────────
# 5. Test Execution Guidelines
# ─────────────────────────────────────────────────────────────
add_heading("5. Test Execution Guidelines & Tooling", level=1)
add_body(
    "To execute automated technical tests, pytest is run against the test modules. "
    "In Step 13, data layer tests (TC-T-01..14) and mathematical formula tests (TC-T-15..27) are fully passing. "
    "Streamlit AppTest smoke tests (TC-T-28..35) will execute once page scripts are constructed in Step 15."
)

add_info_box("💻 Pytest Execution Commands",
    "1. Run data layer integrity tests:\n"
    "   python -m pytest tests/test_data.py -v\n\n"
    "2. Run analytical formulas & statistical tests:\n"
    "   python -m pytest tests/test_analysis.py -v\n\n"
    "3. Run entire test suite (Step 15):\n"
    "   python -m pytest tests/ -v\n\n"
    "4. Generate test coverage report:\n"
    "   python -m pytest --cov=src tests/ --cov-report=term-missing")

# Save
output_path = r"c:\3rd sem\HCL-anti\docs\13_Test_Cases.docx"
doc.save(output_path)
print(f"✅ Saved: {output_path}")
print(f"   Functional Tests: {len(tc_f_data)}")
print(f"   Technical Tests: {len(tc_t_data)}")
