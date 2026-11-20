"""
Step 12 — UI Framework Document Generator
Customer Retention Metrics Analysis (CRMA)
Sumedha | Roll 2503201450146 | MCA 2nd Yr | ABES
"""
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os, datetime

# ── Colour palette ──
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
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, ACCENT_BG)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(title + "\n")
    run.bold = True; run.font.name = 'Calibri'; run.font.size = Pt(10); run.font.color.rgb = NAVY
    run2 = p.add_run(content)
    run2.font.name = 'Calibri'; run2.font.size = Pt(9)
    doc.add_paragraph()

def add_code_box(code_text, title=None):
    """Monospaced code in a shaded 1×1 table cell."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, CODE_BG)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    if title:
        run_t = p.add_run(title + "\n")
        run_t.bold = True; run_t.font.name = 'Consolas'; run_t.font.size = Pt(8); run_t.font.color.rgb = NAVY
    run = p.add_run(code_text)
    run.font.name = 'Consolas'; run.font.size = Pt(7.5)
    doc.add_paragraph()


# ════════════════════════════════════════════════════════
# SCREEN DEFINITIONS
# ════════════════════════════════════════════════════════
SCREENS = [
    {
        "id": "UI-01", "name": "Executive Overview",
        "purpose": "Top-level KPI dashboard for leadership — assess company retention health in under 10 seconds.",
        "components": "5 Tailwind KPI cards (CCR%, CRR%, MRR Lost ₹, NRR%, Active Count) with MoM deltas;\n24-month dual-axis CCR/CRR trend line chart;\n24-month MRR waterfall/line chart (Gross, Churned, Net);\nSnapshot summary sentence.",
        "filters": "Date range selector (month picker); Compare toggle (MoM / QoQ)",
        "user_stories": "US-03, US-04, US-15",
        "oqs": "OQ-01, OQ-09",
        "wireframe": """
┌─────────────────────────────────────────────────────────────────────┐
│  [🏠 CRMA]  Executive Overview | Cohorts | Segments | Drivers |   │
│             Risk Roster | Data Quality | Methodology | About      │
├──────────┬──────────────────────────────────────────────────────────┤
│ SIDEBAR  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐  │
│          │  │ CCR %    │ │ CRR %    │ │ MRR Lost │ │ NRR %    │  │
│ 📅 Month │  │ 2.3%     │ │ 97.7%    │ │ ₹41,200  │ │ 107.2%   │  │
│ [Jan 25] │  │ ▼0.2pp 🟢│ │ ▲0.2pp 🟢│ │ ▼₹3.1K 🟢│ │ ▲1.1pp 🟢│  │
│          │  └──────────┘ └──────────┘ └──────────┘ └──────────┘  │
│ 📊 View  │  ┌──────────────────────────────────────────────────┐  │
│ ○ MoM    │  │  Churn Rate vs Retention Rate — 24-Month Trend  │  │
│ ○ QoQ    │  │  [Plotly dual-axis line: CCR% left, CRR% right] │  │
│          │  │  x: months  y1: CCR%  y2: CRR%                  │  │
│          │  └──────────────────────────────────────────────────┘  │
│          │  ┌──────────────────────────────────────────────────┐  │
│          │  │  Monthly Recurring Revenue Trend                 │  │
│          │  │  [Plotly line: Gross MRR, Churned MRR, Net MRR]  │  │
│          │  │  x: months  y: ₹ INR                             │  │
│          │  └──────────────────────────────────────────────────┘  │
│          │  ℹ Methodology (expander)                             │
└──────────┴──────────────────────────────────────────────────────────┘"""
    },
    {
        "id": "UI-02", "name": "Cohort Analysis",
        "purpose": "Track signup-month cohorts through their lifecycle to see whether newer customers stay longer than earlier ones.",
        "components": "Cohort survival heatmap (triangular matrix, Plotly heatmap);\nCohort decay line chart (Plotly multi-line);\nCohort summary stats table;\nMethodology expander.",
        "filters": "Cohort range (start/end month); Plan type filter",
        "user_stories": "US-06",
        "oqs": "OQ-03",
        "wireframe": """
┌──────────┬──────────────────────────────────────────────────────────┐
│ SIDEBAR  │  Cohort Survival Heatmap                                │
│          │  ┌──────────────────────────────────────────────────┐   │
│ 📅 Cohort│  │       M1    M3    M6    M12                      │   │
│ range    │  │ 24-01 95%   88%   82%   74%  ████████████████    │   │
│ [24-01 to│  │ 24-02 97%   91%   85%   78%  ████████████████    │   │
│  25-06]  │  │ 24-03 93%   86%   79%   —    ████████████████    │   │
│          │  │ ...                                               │   │
│ 🏷 Plan  │  └──────────────────────────────────────────────────┘   │
│ [All]    │                                                         │
│          │  Cohort Decay Curves                                    │
│          │  ┌──────────────────────────────────────────────────┐   │
│          │  │  [Plotly multi-line: one line per cohort]         │   │
│          │  │  x: lifecycle month  y: survival %               │   │
│          │  └──────────────────────────────────────────────────┘   │
│          │  ℹ Methodology (expander)                              │
└──────────┴──────────────────────────────────────────────────────────┘"""
    },
    {
        "id": "UI-03", "name": "Segment Comparison",
        "purpose": "Compare retention, CLV, and channel effectiveness across customer segments to direct effort to the weakest areas.",
        "components": "Multi-select segment filters (plan, industry, size, region, channel);\nSegment retention bar chart (grouped);\nCLV comparison grouped bar chart;\nChannel retention horizontal bar + Chi-Square badge;\nSegment summary table;\nMethodology expander.",
        "filters": "Plan type (multi); Industry (multi); Company size (multi); Region (multi); Acquisition channel (multi); Reset button",
        "user_stories": "US-05, US-07, US-08",
        "oqs": "OQ-02, OQ-04, OQ-05",
        "wireframe": """
┌──────────┬──────────────────────────────────────────────────────────┐
│ SIDEBAR  │  ┌────────────────────────┐ ┌────────────────────────┐  │
│          │  │ Segment Retention Rate │ │ Avg CLV by Plan × Ind. │  │
│ 🏷 Plan  │  │ [Plotly grouped bar]   │ │ [Plotly grouped bar]   │  │
│ ☑ All    │  │ x: segments y: CRR%   │ │ x: plan  fill: indust. │  │
│          │  └────────────────────────┘ └────────────────────────┘  │
│ 🏭 Ind.  │                                                        │
│ ☑ All    │  Channel Retention (12-month)                           │
│          │  ┌──────────────────────────────────────────────────┐   │
│ 📏 Size  │  │ Referral    ████████████████████ 89%              │   │
│ ☑ All    │  │ Partner     ██████████████████   84%              │   │
│          │  │ Organic     █████████████████    82%              │   │
│ 🌍 Region│  │ Direct      ███████████████      75%              │   │
│ ☑ All    │  │ Paid Social ████████████         69%              │   │
│          │  │ ✓ Statistically Significant (χ² p = 0.02)        │   │
│ 📡 Chan. │  └──────────────────────────────────────────────────┘   │
│ ☑ All    │                                                        │
│          │  [🔄 Reset Filters]                                     │
│          │  ℹ Methodology (expander)                              │
└──────────┴──────────────────────────────────────────────────────────┘"""
    },
    {
        "id": "UI-04", "name": "Churn Drivers",
        "purpose": "Statistical analysis of factors driving churn — support tickets, usage decline, and satisfaction scores.",
        "components": "Support impact: side-by-side box plots (ticket volume, resolution hours) for churned vs. retained + Mann-Whitney U card;\nPre-churn telemetry: dual-line chart (login days, feature clicks 90d window) + t-test card;\nSatisfaction correlation: scatter + regression + Pearson/Spearman card;\nMethodology expander.",
        "filters": "Analysis period (date range); Segment filter (optional cross-filter)",
        "user_stories": "US-09, US-10, US-11",
        "oqs": "OQ-06, OQ-07, OQ-08",
        "wireframe": """
┌──────────┬──────────────────────────────────────────────────────────┐
│ SIDEBAR  │  Support Impact on Churn                                │
│          │  ┌────────────────────────┐ ┌────────────────────────┐  │
│ 📅 Period│  │ Ticket Volume          │ │ Resolution Hours       │  │
│ [Full]   │  │ [Box plot: Churned vs  │ │ [Box plot: Churned vs  │  │
│          │  │  Retained]             │ │  Retained]             │  │
│ 🏷 Seg.  │  └────────────────────────┘ └────────────────────────┘  │
│ [All]    │  ┌──────────────────────────────────────────────────┐   │
│          │  │ Mann-Whitney U = 1842 | p = 0.001 | r = 0.38    │   │
│          │  │ ✓ Churned customers have significantly more      │   │
│          │  │   tickets (median 6 vs 3, p < 0.001)             │   │
│          │  └──────────────────────────────────────────────────┘   │
│          │                                                         │
│          │  Pre-Churn Usage Decline (90-day window)                │
│          │  ┌──────────────────────────────────────────────────┐   │
│          │  │ [Dual-line: login_days + feature_clicks           │   │
│          │  │  for churned vs retained, months -3 to 0]        │   │
│          │  │  + t-test result card beneath                    │   │
│          │  └──────────────────────────────────────────────────┘   │
│          │                                                         │
│          │  Satisfaction–Renewal Correlation                       │
│          │  ┌──────────────────────────────────────────────────┐   │
│          │  │ [Scatter: CSAT (x) vs is_churned (y) + reg line] │   │
│          │  │ Pearson r = 0.52 | Spearman ρ = 0.48 | p < 0.01│   │
│          │  │ ⚠ Limited sample: N = 68 of 200 have survey data│   │
│          │  └──────────────────────────────────────────────────┘   │
│          │  ℹ Methodology (expander)                              │
└──────────┴──────────────────────────────────────────────────────────┘"""
    },
    {
        "id": "UI-05", "name": "Risk Roster",
        "purpose": "Actionable at-risk customer list sorted by health risk score for daily CS outreach prioritisation.",
        "components": "Risk summary KPI cards (High/Medium/Low counts);\nSortable, filterable data table with coloured risk badges;\nCSV download button;\nRisk score breakdown tooltip;\nMethodology expander.",
        "filters": "Risk tier (High / Medium / Low); Plan type; Industry; Min risk score slider",
        "user_stories": "US-12, US-13, US-16",
        "oqs": "OQ-10",
        "wireframe": """
┌──────────┬──────────────────────────────────────────────────────────┐
│ SIDEBAR  │  ┌──────────┐ ┌──────────┐ ┌──────────┐               │
│          │  │ 🔴 High  │ │ 🟡 Medium│ │ 🟢 Low   │               │
│ ⚠ Risk   │  │   12     │ │   34     │ │   154    │               │
│ ☑ High   │  └──────────┘ └──────────┘ └──────────┘               │
│ ☑ Medium │                                                        │
│ ☐ Low    │  At-Risk Customer Roster  [⬇ Download CSV]             │
│          │  ┌─────┬──────────┬──────┬────┬─────┬──────┬────────┐  │
│ 🏷 Plan  │  │ ID  │ Company  │ Plan │Ind.│Score│ Tier │Renewal │  │
│ [All]    │  ├─────┼──────────┼──────┼────┼─────┼──────┼────────┤  │
│          │  │C-042│ Shankar  │ Ent  │Fin │ 87  │🔴High│ 15 Jan │  │
│ 🏭 Ind.  │  │C-118│ Himani   │ Pro  │EdT │ 74  │🔴High│ 03 Feb │  │
│ [All]    │  │C-005│ Chaitra  │ Str  │HC  │ 52  │🟡Med │ 28 Feb │  │
│          │  │ ... │  ...     │ ...  │... │ ... │ ...  │  ...   │  │
│ 🎚 Min   │  └─────┴──────────┴──────┴────┴─────┴──────┴────────┘  │
│ Score    │  Showing 46 of 200 active customers                    │
│ [40───]  │  ℹ Methodology (expander)                              │
└──────────┴──────────────────────────────────────────────────────────┘"""
    },
    {
        "id": "UI-06", "name": "Data Quality & EDA",
        "purpose": "Transparency on the data pipeline — show raw data profiles, cleaning steps, distributions, and quality metrics.",
        "components": "Dataset overview cards (table counts, date range, completeness %);\nNull-value heatmap per column;\nDistribution histograms for key numeric columns;\nOutlier detection summary;\nDuplicate removal log;\nData lineage diagram.",
        "filters": "Table selector (dropdown); Column selector (multi-select)",
        "user_stories": "US-01 (data loading/validation)",
        "oqs": "Cross-cutting (data quality underpins all OQs)",
        "wireframe": """
┌──────────┬──────────────────────────────────────────────────────────┐
│ SIDEBAR  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐  │
│          │  │ Tables   │ │ Total    │ │ Complete │ │ Date     │  │
│ 📋 Table │  │   5      │ │ 3,625    │ │  94.2%   │ │ Jan 24 - │  │
│ [All]    │  │          │ │ rows     │ │          │ │ Dec 25   │  │
│          │  └──────────┘ └──────────┘ └──────────┘ └──────────┘  │
│ 📊 Col.  │                                                        │
│ [Select] │  Null Value Heatmap                                     │
│          │  ┌──────────────────────────────────────────────────┐   │
│          │  │ [Plotly heatmap: columns vs tables, shade=null%] │   │
│          │  └──────────────────────────────────────────────────┘   │
│          │                                                         │
│          │  Distributions           │  DQ Issues Found            │
│          │  ┌─────────────────────┐ │ ┌────────────────────────┐  │
│          │  │ [Histograms for     │ │ │ DQ-01: 65.8% null CSAT │  │
│          │  │  tenure, fees, etc.]│ │ │ DQ-02: 3 dup tickets   │  │
│          │  └─────────────────────┘ │ │ DQ-03: 5 outlier hours │  │
│          │                          │ └────────────────────────┘  │
└──────────┴──────────────────────────────────────────────────────────┘"""
    },
    {
        "id": "UI-07", "name": "About",
        "purpose": "Project metadata, team information, tech stack, and academic context.",
        "components": "Project info card (title, domain, purpose);\nTeam card (Sumedha, mentor, institution);\nTech stack badges;\nData sources and refresh note;\nGitHub repo link.",
        "filters": "None",
        "user_stories": "Cross-cutting",
        "oqs": "N/A (informational)",
        "wireframe": """
┌──────────────────────────────────────────────────────────────────────┐
│  Customer Retention Metrics Analysis                                │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │ 📋 Project Information                                       │   │
│  │ Title: Customer Retention Metrics Analysis (CRMA)            │   │
│  │ Domain: B2B SaaS — Customer Success & Retention              │   │
│  │ Company: NoidaSoft Technologies Pvt. Ltd. (fictitious)       │   │
│  └──────────────────────────────────────────────────────────────┘   │
│  ┌──────────────────────┐  ┌───────────────────────────────────┐   │
│  │ 👤 Team              │  │ 🛠 Tech Stack                     │   │
│  │ Sumedha              │  │ Python 3.12 | Streamlit 1.32      │   │
│  │ Roll: 2503201450146  │  │ pandas | NumPy | SciPy            │   │
│  │ MCA 2nd Yr, ABES     │  │ Plotly | Tailwind CSS             │   │
│  │ Mentor: Ms. Pandey   │  │ SQLite | pytest                   │   │
│  └──────────────────────┘  └───────────────────────────────────┘   │
│  Data: Synthetic (seed=42) | 200 customers | 24 months              │
└──────────────────────────────────────────────────────────────────────┘"""
    },
    {
        "id": "UI-08", "name": "Methodology & Data Dictionary",
        "purpose": "Full reference page for metric definitions, statistical tests, data dictionary, and analytical assumptions — for mentor/examiner verification.",
        "components": "11 KPI/measure definition cards with formulas;\nStatistical test reference table (test, hypothesis, α, assumptions);\nComplete data dictionary (all 5 tables + view);\nData quality handling notes;\nMethodology expander on every other page links here.",
        "filters": "Section jump links (anchor navigation)",
        "user_stories": "US-02, US-14",
        "oqs": "All (cross-cutting governance)",
        "wireframe": """
┌──────────────────────────────────────────────────────────────────────┐
│  Methodology & Data Dictionary                                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │ § Metric Definitions                                         │   │
│  │ ┌────────┬────────────────────────┬──────────────────────┐   │   │
│  │ │ MSR-01 │ Customer Churn Rate    │ (churned / active)×100│   │   │
│  │ │ MSR-02 │ Customer Retention Rate│ 100 - CCR%            │   │   │
│  │ │ ...    │ ...                    │ ...                   │   │   │
│  │ └────────┴────────────────────────┴──────────────────────┘   │   │
│  └──────────────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │ § Statistical Tests                                          │   │
│  │ OQ-06: Mann-Whitney U | H₀: ticket dist. equal | α = 0.05  │   │
│  │ OQ-07: Two-sample t   | H₀: login means equal  | α = 0.05  │   │
│  │ OQ-08: Pearson/Spearman | H₀: ρ = 0 | α = 0.05             │   │
│  └──────────────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │ § Data Dictionary (5 tables + 1 view, 41 columns)            │   │
│  │ [Expandable section per table with column details]           │   │
│  └──────────────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────────────┘"""
    },
]

# Chart-OQ mapping
CHART_OQ = [
    ("OQ-01", "CCR% / CRR% 24-month trend", "Dual-axis line chart", "Line charts best show temporal trends and direction; dual-axis avoids scale distortion between complementary metrics."),
    ("OQ-02", "Segment retention comparison", "Grouped bar chart", "Grouped bars enable side-by-side comparison of categorical segments; colour coding distinguishes plan/industry/size/region."),
    ("OQ-03", "Cohort survival matrix", "Heatmap (triangular)", "Heatmaps encode three dimensions (cohort × lifecycle month × survival%) in a compact visual; colour intensity conveys magnitude intuitively."),
    ("OQ-04", "CLV by plan × industry", "Grouped bar chart", "Bar charts effectively compare a single continuous metric (CLV ₹) across categorical groups with clear magnitude differences."),
    ("OQ-05", "Channel retention rates", "Horizontal bar chart", "Horizontal bars accommodate long channel labels and naturally rank-order from best to worst retention."),
    ("OQ-06", "Support tickets: churned vs retained", "Side-by-side box plots", "Box plots reveal distribution shape, median, quartiles, and outliers — critical for statistical comparison of two groups."),
    ("OQ-07", "Pre-churn 90-day usage decline", "Multi-line chart", "Line charts show the temporal decay of login/feature engagement leading up to churn; two groups (churned vs retained) as separate series."),
    ("OQ-08", "CSAT/NPS vs renewal correlation", "Scatter + regression line", "Scatter plots with regression lines visually communicate correlation strength and direction; outliers are immediately visible."),
    ("OQ-09", "MRR trend (gross/churned/net)", "Multi-line chart", "Line charts show MRR movement over time; three series (gross, churned, net) reveal the gap between revenue earned and lost."),
    ("OQ-10", "Risk score distribution", "Risk roster table + KPI cards", "A sortable data table is the most actionable format for CS executives who need to identify specific at-risk accounts by name."),
]

# Reusable components
COMPONENTS = [
    ("KPI Metric Card", "render_kpi_card(title, value, delta, status)",
     "bg-white rounded-xl shadow-sm border border-slate-200 p-4\ntext-slate-500 text-xs uppercase tracking-wider\ntext-3xl font-bold text-slate-800\ntext-sm (green-600 / red-500 for delta)",
     "Displays a headline metric with MoM delta badge. Status-aware colouring (green = good, red = bad). Responsive: full-width on mobile, 4-col grid on desktop."),
    ("Chart Card", "render_chart_card(title, fig, methodology_text)",
     "bg-white rounded-lg shadow-sm border border-slate-100 p-5\ntext-lg font-semibold text-slate-700 mb-3\n[Plotly fig rendered via st.plotly_chart]\ndetails > summary: text-xs text-indigo-600 cursor-pointer",
     "Wraps a Plotly figure with a title and a collapsible methodology expander beneath it."),
    ("Stat Result Card", "render_stat_card(test_name, stat, p_value, interpretation)",
     "bg-gradient-to-r from-blue-50 to-indigo-50 rounded-lg p-4 border-l-4 border-indigo-500\ntext-sm font-mono text-slate-600\ntext-base font-medium (green-700 / amber-600)",
     "Shows a statistical test result with the test statistic, p-value, significance badge, and plain-English interpretation."),
    ("Risk Badge", "render_risk_badge(tier)",
     "inline-flex px-2 py-0.5 rounded-full text-xs font-semibold\nbg-red-100 text-red-700 (High)\nbg-amber-100 text-amber-700 (Medium)\nbg-green-100 text-green-700 (Low)",
     "Semantic colour badge for risk tier — inline in table cells."),
    ("Section Header", "render_section_header(title, subtitle)",
     "text-2xl font-bold text-slate-800 border-b-2 border-indigo-500 pb-2\ntext-sm text-slate-500 mt-1",
     "Page section header with accent underline. Used at the top of each analytical section."),
    ("Filter Summary Bar", "render_filter_bar(active_filters)",
     "bg-indigo-50 rounded-lg px-4 py-2 flex flex-wrap gap-2\nspan: bg-indigo-100 text-indigo-800 text-xs px-2 py-1 rounded-full",
     "Shows active filter selections as pill badges below the sidebar. Includes a clear-all button."),
]

# Theme colours
THEME = {
    "primaryColor":      "#2B4C7E",
    "backgroundColor":   "#FFFFFF",
    "secondaryBackgroundColor": "#F0F2F6",
    "textColor":         "#1B365D",
    "font":              "sans serif",
}


# ════════════════════════════════════════════════════════
# BUILD DOCUMENT
# ════════════════════════════════════════════════════════

# ── Title Page ──
doc.add_paragraph()
doc.add_paragraph()
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run("Customer Retention Metrics Analysis")
run.bold = True; run.font.size = Pt(26); run.font.color.rgb = NAVY; run.font.name = 'Calibri'

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run2 = subtitle.add_run("Step 12 — UI Framework & Wireframes")
run2.font.size = Pt(16); run2.font.color.rgb = STEEL; run2.font.name = 'Calibri'

doc.add_paragraph()

info_data = [
    ("Project", "Customer Retention Metrics Analysis (CRMA)"),
    ("Author", "Sumedha | Roll 2503201450146 | MCA 2nd Yr, ABES"),
    ("Mentor", "Ms. Anjali Pandey"),
    ("Version", "1.0"),
    ("Date", datetime.date.today().strftime("%d %B %Y")),
    ("Status", "Baselined"),
    ("Input Documents", "01 Project Understanding, 07 TRD (FRs/NFRs),\n08 Database Design, 11 User Stories"),
]
info_tbl = doc.add_table(rows=len(info_data), cols=2)
info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(info_data):
    set_cell_shading(info_tbl.rows[i].cells[0], ACCENT_BG)
    set_cell_text(info_tbl.rows[i].cells[0], k, bold=True, font_size=Pt(10), color=NAVY)
    set_cell_text(info_tbl.rows[i].cells[1], v, font_size=Pt(10))

doc.add_page_break()

# ── TOC ──
add_heading("Table of Contents", level=1)
toc_items = [
    "1. Screen Registry",
    "2. Screen Wireframes (UI-01 to UI-08)",
    "3. Global Layout & Navigation",
    "4. Colour Theme & Typography",
    "5. Reusable Tailwind Components",
    "6. Best Chart for Each OQ",
    "7. .streamlit/config.toml",
    "8. Starter src/ui.py Code",
    "9. RTM Update Summary",
    "10. Assumptions",
]
for item in toc_items:
    p = doc.add_paragraph(item)
    p.paragraph_format.space_after = Pt(1)

doc.add_page_break()

# ═══════════ 1. SCREEN REGISTRY ═══════════
add_heading("1. Screen Registry", level=1)
add_body("The CRMA application comprises 8 screens (pages), each mapped to specific User Stories and Operational Questions:")

reg_tbl = doc.add_table(rows=1 + len(SCREENS), cols=6)
reg_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(reg_tbl, ["UI ID", "Screen Name", "Purpose (brief)", "Key Components", "User Stories", "OQs"])

for i, s in enumerate(SCREENS, 1):
    purpose_brief = s["purpose"].split("—")[0].strip() if "—" in s["purpose"] else s["purpose"][:60] + "..."
    add_data_row(reg_tbl, i, [
        s["id"], s["name"], purpose_brief,
        s["components"].split(";")[0].strip() + " ...",
        s["user_stories"], s["oqs"]
    ])

# ═══════════ 2. SCREEN WIREFRAMES ═══════════
doc.add_page_break()
add_heading("2. Screen Wireframes (UI-01 to UI-08)", level=1)
add_body(
    "Each wireframe uses ASCII art to show the layout of sidebar filters, KPI cards, charts, tables, "
    "and interactive components. The wireframes are designed for desktop (1920×1080) with responsive "
    "breakpoints for tablet (768px)."
)

for s in SCREENS:
    doc.add_page_break()
    add_heading(f"{s['id']} — {s['name']}", level=2)

    # Purpose
    add_body(f"Purpose: {s['purpose']}")

    # Details table
    det_tbl = doc.add_table(rows=4, cols=2)
    det_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    det_data = [
        ("Components", s["components"]),
        ("Sidebar Filters", s["filters"]),
        ("User Stories", s["user_stories"]),
        ("OQs Answered", s["oqs"]),
    ]
    for j, (k, v) in enumerate(det_data):
        set_cell_shading(det_tbl.rows[j].cells[0], ACCENT_BG)
        set_cell_text(det_tbl.rows[j].cells[0], k, bold=True, font_size=Pt(9), color=NAVY)
        set_cell_text(det_tbl.rows[j].cells[1], v, font_size=Pt(9))

    doc.add_paragraph()

    # Wireframe
    add_heading("Wireframe", level=3)
    add_code_box(s["wireframe"].strip(), title=f"{s['id']} — {s['name']}")


# ═══════════ 3. GLOBAL LAYOUT ═══════════
doc.add_page_break()
add_heading("3. Global Layout & Navigation", level=1)

add_heading("3.1 Navigation Architecture", level=2)
add_body(
    "The app uses Streamlit's native multipage architecture (pages/ directory). "
    "Navigation appears as a horizontal tab bar at the top, styled with Tailwind CSS "
    "in a custom header component rendered via st.components.v1.html()."
)

nav_items = [
    "🏠 Executive Overview (app.py — landing page)",
    "📊 Cohort Analysis (pages/1_Cohorts.py)",
    "🔀 Segment Comparison (pages/2_Segments.py)",
    "🔍 Churn Drivers (pages/3_Drivers.py)",
    "⚠️ Risk Roster (pages/4_Risk_Roster.py)",
    "🧹 Data Quality (pages/5_Data_Quality.py)",
    "📖 Methodology (pages/6_Methodology.py)",
    "ℹ️ About (pages/7_About.py)",
]
for item in nav_items:
    p = doc.add_paragraph(item, style='List Bullet')
    p.paragraph_format.space_after = Pt(2)

add_heading("3.2 Shared State Management", level=2)
add_body(
    "Global filters (date range, plan type, industry, company size, region) are stored in "
    "st.session_state and persist across page navigations. Each page reads from session state "
    "and applies the active filters to its data before computing metrics."
)

add_code_box("""# Shared filter state pattern (src/filters.py)
def init_filters():
    defaults = {
        'plan_filter': [],
        'industry_filter': [],
        'size_filter': [],
        'region_filter': [],
        'channel_filter': [],
        'date_range': ('2024-01', '2025-12'),
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def render_sidebar_filters(df):
    with st.sidebar:
        st.session_state.plan_filter = st.multiselect(
            '🏷 Plan Type', df['plan_name'].unique(),
            default=st.session_state.plan_filter)
        # ... repeat for industry, size, region, channel
        if st.button('🔄 Reset Filters'):
            for k in defaults: st.session_state[k] = defaults[k]
            st.rerun()""", title="src/filters.py — Shared Filter Pattern")

add_heading("3.3 Responsive Design", level=2)
add_body(
    "Tailwind's responsive utilities ensure the layout adapts across screen sizes:"
)
resp_tbl = doc.add_table(rows=1 + 3, cols=3)
resp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(resp_tbl, ["Breakpoint", "Width", "Layout Behaviour"])
add_data_row(resp_tbl, 1, ["Mobile", "< 768px", "Single column; KPI cards stack vertically; charts full-width"])
add_data_row(resp_tbl, 2, ["Tablet", "768px – 1024px", "2-column grid for KPI cards; charts full-width"])
add_data_row(resp_tbl, 3, ["Desktop", "> 1024px", "4-column KPI grid; side-by-side charts; full sidebar"])


# ═══════════ 4. COLOUR THEME ═══════════
doc.add_page_break()
add_heading("4. Colour Theme & Typography", level=1)

add_heading("4.1 Colour Palette", level=2)
color_tbl = doc.add_table(rows=1 + 7, cols=4)
color_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(color_tbl, ["Role", "Colour Name", "Hex Code", "Usage"])
colors = [
    ("Primary", "Navy", "#1B365D", "Headers, headings, primary text"),
    ("Secondary", "Steel Blue", "#2B4C7E", "Accents, links, active states"),
    ("Positive", "Emerald", "#059669", "Retention up, good delta, low risk"),
    ("Negative", "Rose", "#E11D48", "Churn up, bad delta, high risk"),
    ("Warning", "Amber", "#D97706", "Medium risk, caution badges"),
    ("Background", "White", "#FFFFFF", "Page background, card background"),
    ("Surface", "Slate 50", "#F8FAFC", "Sidebar, secondary background, alt rows"),
]
for i, (role, name, hex_code, usage) in enumerate(colors, 1):
    add_data_row(color_tbl, i, [role, name, hex_code, usage])

add_heading("4.2 Typography", level=2)
typo_tbl = doc.add_table(rows=1 + 4, cols=3)
typo_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(typo_tbl, ["Element", "Font", "Style"])
typo_data = [
    ("Page Title", "Inter / system sans-serif", "text-2xl font-bold text-slate-800"),
    ("Section Header", "Inter / system sans-serif", "text-lg font-semibold text-slate-700"),
    ("Body Text", "Inter / system sans-serif", "text-sm text-slate-600 leading-relaxed"),
    ("Code / Formulas", "JetBrains Mono / Consolas", "text-xs font-mono text-slate-500"),
]
for i, (el, font, sty) in enumerate(typo_data, 1):
    add_data_row(typo_tbl, i, [el, font, sty])


# ═══════════ 5. REUSABLE COMPONENTS ═══════════
doc.add_page_break()
add_heading("5. Reusable Tailwind Components", level=1)
add_body(
    "The following components are defined in src/components.py and reused across all pages. "
    "Each is rendered as an HTML block inside st.components.v1.html() with the Tailwind CDN."
)

for comp_name, func_sig, tw_classes, description in COMPONENTS:
    add_heading(comp_name, level=3)
    comp_tbl = doc.add_table(rows=4, cols=2)
    comp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_data = [
        ("Function", func_sig),
        ("Tailwind Classes", tw_classes),
        ("Description", description),
        ("Used In", "All pages" if "Card" in comp_name or "Header" in comp_name else "Specific pages"),
    ]
    for j, (k, v) in enumerate(comp_data):
        set_cell_shading(comp_tbl.rows[j].cells[0], ACCENT_BG)
        set_cell_text(comp_tbl.rows[j].cells[0], k, bold=True, font_size=Pt(9), color=NAVY)
        set_cell_text(comp_tbl.rows[j].cells[1], v, font_size=Pt(9))
    doc.add_paragraph()


# ═══════════ 6. BEST CHART FOR EACH OQ ═══════════
doc.add_page_break()
add_heading("6. Best Chart for Each OQ", level=1)
add_body(
    "Each Operational Question is paired with the optimal Plotly chart type, chosen for clarity, "
    "statistical appropriateness, and the target persona's decision-making context."
)

chart_tbl = doc.add_table(rows=1 + len(CHART_OQ), cols=4)
chart_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(chart_tbl, ["OQ", "Visual", "Chart Type", "Why This Chart"])

for i, (oq, visual, chart_type, reason) in enumerate(CHART_OQ, 1):
    add_data_row(chart_tbl, i, [oq, visual, chart_type, reason])


# ═══════════ 7. CONFIG.TOML ═══════════
doc.add_page_break()
add_heading("7. .streamlit/config.toml", level=1)
add_body("The Streamlit theme configuration file harmonises the native canvas with the Tailwind-styled components:")

config_toml = """[theme]
primaryColor = "#2B4C7E"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#1B365D"
font = "sans serif"

[server]
headless = true
port = 8501

[browser]
gatherUsageStats = false"""

add_code_box(config_toml, title=".streamlit/config.toml")


# ═══════════ 8. STARTER src/ui.py ═══════════
doc.add_page_break()
add_heading("8. Starter src/ui.py", level=1)
add_body(
    "The starter UI module provides the Tailwind-powered header and KPI card components. "
    "These are rendered via st.components.v1.html() with the Tailwind CDN loaded in-frame."
)

ui_code = r'''"""
src/ui.py — Tailwind-styled UI components for CRMA
Renders HTML via st.components.v1.html() with Tailwind CDN.
"""
import streamlit.components.v1 as components


def render_header(title: str = "Customer Retention Metrics Analysis",
                  subtitle: str = "NoidaSoft Technologies Pvt. Ltd."):
    """Render the app header with Tailwind styling."""
    html = f"""
    <script src="https://cdn.tailwindcss.com"></script>
    <div class="bg-gradient-to-r from-[#1B365D] to-[#2B4C7E]
                rounded-xl shadow-lg px-6 py-4 mb-4">
        <h1 class="text-2xl font-bold text-white tracking-tight">
            {title}
        </h1>
        <p class="text-sm text-blue-200 mt-1">{subtitle}</p>
    </div>
    """
    components.html(html, height=90)


def render_kpi_card(title: str, value: str, delta: str = "",
                    status: str = "neutral"):
    """Render a single KPI metric card with delta badge.

    Args:
        title:  Metric label (e.g., "Churn Rate")
        value:  Formatted value (e.g., "2.3%")
        delta:  Change text (e.g., "▼ 0.2 pp")
        status: "positive" (green), "negative" (red), or "neutral" (grey)
    """
    color_map = {
        "positive": ("text-green-600", "bg-green-50"),
        "negative": ("text-red-500",   "bg-red-50"),
        "neutral":  ("text-slate-500", "bg-slate-50"),
    }
    text_color, badge_bg = color_map.get(status, color_map["neutral"])

    delta_html = ""
    if delta:
        delta_html = f"""
        <span class="inline-flex items-center px-2 py-0.5
                      rounded-full text-xs font-medium
                      {badge_bg} {text_color}">
            {delta}
        </span>
        """

    html = f"""
    <script src="https://cdn.tailwindcss.com"></script>
    <div class="bg-white rounded-xl shadow-sm border border-slate-200
                p-4 hover:shadow-md transition-shadow duration-200">
        <p class="text-xs font-medium text-slate-500 uppercase
                  tracking-wider mb-1">{title}</p>
        <div class="flex items-end gap-2">
            <p class="text-3xl font-bold text-slate-800">{value}</p>
            {delta_html}
        </div>
    </div>
    """
    components.html(html, height=110)


def render_kpi_row(metrics: list[dict]):
    """Render a row of KPI cards in a responsive grid.

    Args:
        metrics: List of dicts with keys: title, value, delta, status
    """
    cards_html = ""
    for m in metrics:
        text_color, badge_bg = {
            "positive": ("text-green-600", "bg-green-50"),
            "negative": ("text-red-500",   "bg-red-50"),
            "neutral":  ("text-slate-500", "bg-slate-50"),
        }.get(m.get("status", "neutral"), ("text-slate-500", "bg-slate-50"))

        delta_html = ""
        if m.get("delta"):
            delta_html = f"""
            <span class="inline-flex items-center px-2 py-0.5
                          rounded-full text-xs font-medium
                          {badge_bg} {text_color}">
                {m['delta']}
            </span>
            """

        cards_html += f"""
        <div class="bg-white rounded-xl shadow-sm border border-slate-200
                    p-4 hover:shadow-md transition-shadow duration-200">
            <p class="text-xs font-medium text-slate-500 uppercase
                      tracking-wider mb-1">{m['title']}</p>
            <div class="flex items-end gap-2">
                <p class="text-3xl font-bold text-slate-800">{m['value']}</p>
                {delta_html}
            </div>
        </div>
        """

    html = f"""
    <script src="https://cdn.tailwindcss.com"></script>
    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-{min(len(metrics), 4)}
                gap-4 mb-4">
        {cards_html}
    </div>
    """
    height = 130 if len(metrics) <= 4 else 250
    components.html(html, height=height)
'''

add_code_box(ui_code, title="src/ui.py — Full Starter Code")


# ═══════════ 9. RTM UPDATE SUMMARY ═══════════
doc.add_page_break()
add_heading("9. RTM Update Summary", level=1)
add_body(
    "The Traceability Matrix (04_Traceability_Matrix.xlsx) has been updated. "
    "The 'UI Screen' column is now 100% populated across all 15 RTM rows."
)

rtm_tbl = doc.add_table(rows=1 + 9, cols=3)
rtm_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_header_row(rtm_tbl, ["UI Screen", "Screen Name", "RTM Rows Linked"])
rtm_data = [
    ("UI-01", "Executive Overview", "Rows 2, 3, 10, 11, 14 (OQ-01, OQ-09)"),
    ("UI-02", "Cohort Analysis", "Row 5 (OQ-03)"),
    ("UI-03", "Segment Comparison", "Rows 4, 7, 12 (OQ-02, OQ-04, OQ-05)"),
    ("UI-04", "Churn Drivers", "Rows 8, 9, 13 (OQ-06, OQ-07, OQ-08)"),
    ("UI-05", "Risk Roster", "Row 6 (OQ-10)"),
    ("UI-06", "Data Quality & EDA", "Cross-cutting (data validation)"),
    ("UI-07", "About", "Informational (not linked to specific BR)"),
    ("UI-08", "Methodology & Data Dictionary", "Row 16 (All OQs, BR-08)"),
    ("Deferred", "N/A (BR-13 Won't)", "Row 15"),
]
for i, (ui, name, rows) in enumerate(rtm_data, 1):
    add_data_row(rtm_tbl, i, [ui, name, rows])

add_info_box("✅ RTM Column Status After Step 12",
    "• OQ: 100% populated (Step 1)\n"
    "• HMW: 100% populated (Step 2)\n"
    "• BR: 100% populated (Step 3)\n"
    "• DB table.column: 100% populated (Step 8)\n"
    "• User Story: 100% populated (Step 11)\n"
    "• UI Screen: 100% populated (Step 12) ✅\n"
    "• Feature: Skipped (Step 6 PRD not done)\n"
    "• Test Case: To be filled in Step 13")


# ═══════════ 10. ASSUMPTIONS ═══════════
doc.add_page_break()
add_heading("10. Assumptions", level=1)
assumptions = [
    "A-18: Tailwind CSS is loaded from CDN (https://cdn.tailwindcss.com) inside st.components.v1.html() iframes. No local Tailwind build pipeline is needed.",
    "A-19: The multipage app uses Streamlit's native pages/ directory convention — no custom routing library.",
    "A-20: UI-06 (Data Quality) and UI-07 (About) are informational pages not directly linked to specific BRs in the RTM, but support the overall analytical narrative.",
    "A-21: Responsive design targets desktop-first with tablet support via Tailwind breakpoints. No dedicated mobile-app view is designed.",
    "A-22: The Inter font is loaded from Google Fonts CDN inside Tailwind components. The native Streamlit canvas uses the default sans-serif font configured in config.toml.",
    "A-23: Chart interactivity (hover, zoom, pan) uses Plotly's built-in toolbar and is not customised beyond tooltip formatting.",
    "A-24: The 'ℹ Methodology' expander on each page uses st.expander() (native Streamlit), not a Tailwind component, to avoid iframe nesting complexity.",
]
for a in assumptions:
    p = doc.add_paragraph(a, style='List Bullet')
    p.paragraph_format.space_after = Pt(3)


# ═══════════ SAVE ═══════════
output_path = r"c:\3rd sem\HCL-anti\docs\12_UI_Framework.docx"
doc.save(output_path)
print(f"✅ Saved: {output_path}")
print(f"   Screens: {len(SCREENS)}")
print(f"   Components: {len(COMPONENTS)}")
print(f"   Chart-OQ mappings: {len(CHART_OQ)}")
