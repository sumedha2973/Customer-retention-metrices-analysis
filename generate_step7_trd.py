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

def build_step7_trd():
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
    run_title = p_title.add_run("TECHNICAL REQUIREMENTS DOCUMENT (TRD)")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Customer Retention Metrics Analysis — Technical Architecture & Engineering Specifications")
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
        set_cell_margins(cell, 70, 70, 100, 100)
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
                if j == 0 and ("FR-" in str(val) or "OQ-" in str(val) or "TR-" in str(val) or "NFR-" in str(val)):
                    r.font.bold = True
                    r.font.color.rgb = STEEL

    # Metadata Table
    meta_headers = ["Engineering Attribute", "Technical Specification"]
    meta_data = [
        ("System Name", "Customer Retention Metrics Analysis (CRMA)"),
        ("Architecture Style", "Modular Multi-Page Data Analytics Web Application"),
        ("Target Runtime Environment", "Python 3.12 LTS (64-bit)"),
        ("Core Application Framework", "Streamlit 1.32.2 (Multi-page configuration)"),
        ("Database Storage Engine", "SQLite3 (Embedded relational storage at data/app.db)"),
        ("Analytical & Statistical Stack", "pandas 2.2.1, NumPy 1.26.4, SciPy 1.12.0, statsmodels 0.14.1"),
        ("Visualization Engine", "Plotly 5.20.0 (Graph Objects & Plotly Express)"),
        ("Front-End UI Enhancements", "Tailwind CSS (CDN-injected via st.components.v1.html iframe)"),
        ("Testing & Quality Framework", "pytest 8.1.1, pytest-cov, Streamlit AppTest native testing"),
        ("Hosting & CI/CD Target", "Streamlit Community Cloud connected to GitHub Repository")
    ]
    t_meta = doc.add_table(rows=len(meta_data) + 1, cols=2)
    format_table(t_meta, meta_headers, meta_data, [Inches(2.2), Inches(4.3)], align_center_cols=[0])

    # 1. System Architecture
    add_h1("1. System Architecture & End-to-End Pipeline")
    p_arch = doc.add_paragraph()
    p_arch.add_run("The CRMA platform is architected as an uncoupled, layered data analytics pipeline designed to enforce separation of concerns between raw storage, in-memory caching, vectorized mathematical modeling, statistical inferential testing, and UI rendering.")

    p_pipe_desc = doc.add_paragraph()
    p_pipe_desc.add_run("The data flow traverses five distinct architectural stages: Data Layer (SQLite) -> Ingestion & Caching Layer (st.cache_data) -> Analytical & Statistical Processing Core -> Presentation Layer (Plotly & Tailwind Components) -> Deployment & QA Target.")

    arch_code = (
        "Flowchart Architecture Representation:\n"
        "[SQLite Database (data/app.db)]\n"
        "      │ (Indexed SQL Queries)\n"
        "      ▼\n"
        "[Ingestion & Caching Layer (src/data_loader.py)]\n"
        "      │ ├── @st.cache_data (TTL=3600s, max_entries=10)\n"
        "      │ └── Schema validation & null-handling\n"
        "      ▼\n"
        "[Analytical Core Modules (src/)]\n"
        "      │ ├── src/metrics.py       --> CRR, CCR, MRR Churn, NRR, CLV\n"
        "      │ ├── src/stats.py         --> Chi-Square, Two-Sample t-Test, Mann-Whitney U, Correlation\n"
        "      │ └── src/risk_scoring.py  --> Multi-factor health score (0-100)\n"
        "      ▼\n"
        "[Presentation & Dashboard Layer (pages/)]\n"
        "      │ ├── src/visualizations.py --> Plotly interactive charts & cohort heatmaps\n"
        "      │ ├── src/components.py     --> Tailwind CSS KPI cards & headers via st.components.v1.html\n"
        "      │ └── pages/1..5            --> Streamlit Multipage views\n"
        "      ▼\n"
        "[Continuous Testing & Cloud Delivery]\n"
        "        ├── pytest + Streamlit AppTest (tests/)\n"
        "        └── Streamlit Community Cloud (GitHub CI/CD)"
    )
    add_code_box(arch_code)

    # 2. Folder Structure
    add_h1("2. Standard Repository & Modular Directory Structure")
    p_folder = doc.add_paragraph()
    p_folder.add_run("The codebase strictly follows industry best practices for modular Python applications, ensuring complete isolation of presentation logic from data computation:")

    tree_text = (
        "CRMA-Project/\n"
        "├── app.py                      # Application router, global layout, navigation landing\n"
        "├── pages/                      # Multi-page analytical views\n"
        "│   ├── 1_Executive_Overview.py # Headline KPIs, retention/churn monthly trajectories (OQ-01, OQ-09)\n"
        "│   ├── 2_Cohort_Analysis.py    # 12-month cohort survival triangular heatmap (OQ-03)\n"
        "│   ├── 3_Segment_DeepDive.py   # Multi-dimensional breakdowns, channels, CLV (OQ-02, OQ-04, OQ-05)\n"
        "│   ├── 4_Churn_Drivers.py      # Statistical hypothesis testing: tickets, usage, CSAT (OQ-06, OQ-07, OQ-08)\n"
        "│   └── 5_At_Risk_Accounts.py   # Rule-based early warning risk scores & CSV export (OQ-10)\n"
        "├── src/                        # Reusable core analytical modules\n"
        "│   ├── __init__.py\n"
        "│   ├── data_loader.py          # SQLite connection, query caching, schema validation\n"
        "│   ├── metrics.py              # Vectorized math: CRR, CCR, MRR churn, NRR, CLV\n"
        "│   ├── stats.py                # SciPy/statsmodels hypothesis testing engine\n"
        "│   ├── risk_scoring.py         # Multi-factor customer health score algorithm\n"
        "│   ├── visualizations.py       # Plotly chart factories (line, bar, heatmap, waterfall)\n"
        "│   └── components.py           # Tailwind CSS components injected via st.components.v1.html\n"
        "├── data/                       # Local data storage and generation scripts\n"
        "│   ├── app.db                  # Embedded SQLite database (~5,000 customers, 24 months)\n"
        "│   └── generate_data.py        # Deterministic synthetic data generation engine (Step 10)\n"
        "├── tests/                      # Automated test suite\n"
        "│   ├── __init__.py\n"
        "│   ├── test_metrics.py         # Mathematical parity tests asserting Delta = 0 vs SQL\n"
        "│   ├── test_data_integrity.py  # Zero nulls, key constraints, valid ranges\n"
        "│   ├── test_stats.py           # Statistical outputs and p-value validation\n"
        "│   └── test_app_ui.py          # Streamlit AppTest native page rendering checks\n"
        "├── docs/                       # Project lifecycle specifications (Steps 1-18)\n"
        "├── .streamlit/\n"
        "│   └── config.toml             # Custom theme settings (palette, typography, layout)\n"
        "├── requirements.txt            # Pinned package dependencies\n"
        "└── README.md                   # Installation, testing, and deployment runbook"
    )
    add_code_box(tree_text)

    # 3. Functional Requirements
    add_h1("3. Functional Requirements Specification (FR-01 to FR-12)")
    p_fr_intro = doc.add_paragraph()
    p_fr_intro.add_run("The functional requirements map business capabilities (BRs) directly into technical modules within the codebase:")

    fr_headers = ["FR ID", "Technical Requirement Description", "Target Module", "Linked BR", "Linked OQ"]
    fr_data = [
        ("FR-01", "Core Retention Engine: Calculate monthly and quarterly Customer Retention Rate (CRR) and Churn Rate (CCR) using vectorized pandas operations matching RULE-01/02.", "src/metrics.py", "BR-01, BR-02", "OQ-01"),
        ("FR-02", "Segmentation Engine: Implement dynamic multi-attribute slicing across plan_type, industry, company_size, and region with instant metric recalculation.", "src/metrics.py", "BR-03", "OQ-02"),
        ("FR-03", "Cohort Matrix Engine: Construct a 24x12 triangular survival matrix calculating percentage retention decay at months 1, 3, 6, and 12.", "src/metrics.py", "BR-04", "OQ-03"),
        ("FR-04", "Risk Scoring Engine: Implement a 4-factor weighted linear combination algorithm (0-100) assigning High, Medium, and Low risk tiers per RULE-05.", "src/risk_scoring.py", "BR-05", "OQ-10"),
        ("FR-05", "Support SLA Analytics: Compute mean ticket volumes and resolution turnaround hours; execute Mann-Whitney U / t-test between churned and retained accounts.", "src/stats.py", "BR-06", "OQ-06"),
        ("FR-06", "Revenue Churn & NRR Engine: Calculate Gross MRR Churn, Expansion MRR, and Net Revenue Retention (NRR %) across monthly cycles matching RULE-03.", "src/metrics.py", "BR-07", "OQ-09"),
        ("FR-07", "Metric Governance Panel: Render an expandable technical metadata drawer beneath every visual widget showing SQL formula, sample size, and assumptions.", "src/components.py", "BR-08", "All (OQ-01..10)"),
        ("FR-08", "Customer Lifetime Value Engine: Compute historical ARPU, average customer lifespan, and Segment CLV matching RULE-04.", "src/metrics.py", "BR-09", "OQ-04"),
        ("FR-09", "Acquisition Channel Longevity: Aggregate retention rates across Organic, Referral, Paid Ads, and Partners; execute Chi-Square test of independence.", "src/stats.py", "BR-10", "OQ-05"),
        ("FR-10", "Pre-Churn Telemetry Analyzer: Compute 90-day pre-churn login days and feature clicks; perform two-sample t-test against retained baseline.", "src/stats.py", "BR-11", "OQ-07"),
        ("FR-11", "Satisfaction Correlation Engine: Calculate Pearson r and Spearman rho correlation coefficients between CSAT/NPS and renewal status.", "src/stats.py", "BR-12", "OQ-08"),
        ("FR-12", "Data Export Utility: Provide client-side CSV download functionality for filtered at-risk account rosters with formatted timestamps.", "src/data_loader.py", "BR-05, BR-11", "OQ-10")
    ]
    t_fr = doc.add_table(rows=len(fr_data) + 1, cols=5)
    format_table(t_fr, fr_headers, fr_data, [Inches(0.7), Inches(2.5), Inches(1.3), Inches(1.1), Inches(0.9)], align_center_cols=[0, 3, 4])

    # 4. Data Layer Architecture
    add_h1("4. Data Layer Architecture: Loading, Caching, and Validation")
    add_h2("4.1 Database Access & Connection Management")
    p_dl = doc.add_paragraph()
    p_dl.add_run("Data is stored in an embedded SQLite database (`data/app.db`). To prevent file-locking contentions and ensure zero security vulnerabilities:\n"
                 "• Connections are established using Python's native `sqlite3` module in read-only URI mode: `file:data/app.db?mode=ro`.\n"
                 "• Database tables include explicit primary keys and B-Tree indexes on foreign keys (`idx_sub_customer`, `idx_ticket_customer`, `idx_usage_customer`).")

    add_h2("4.2 Streamlit In-Memory Caching Strategy")
    p_cache = doc.add_paragraph()
    p_cache.add_run("To achieve sub-second page transitions and respect Streamlit Community Cloud memory quotas (~1 GB RAM):\n"
                    "• The primary data extraction function in `src/data_loader.py` is decorated with `@st.cache_data(ttl=3600, show_spinner=False)`.\n"
                    "• In-memory dataframes are downcast to memory-efficient types (e.g., `category` for low-cardinality strings like `plan_type` and `industry`, `int32` for IDs, `float32` for fees).\n"
                    "• Filter operations leverage pandas vectorized boolean indexing rather than repeated SQL query dispatches.")

    add_h2("4.3 Schema Validation & Data Integrity Rules")
    p_val = doc.add_paragraph()
    p_val.add_run("Upon initial cache load, `src/data_loader.py` executes strict automated sanity assertions:\n"
                  "1. Primary Key Uniqueness: Asserts `df['customer_id'].is_unique == True`.\n"
                  "2. Foreign Key Integrity: Asserts 100% of foreign keys in child tables resolve to valid parent IDs.\n"
                  "3. Domain Range Checks: Asserts `monthly_fee >= 0`, `csat_score.between(1, 5)`, `login_days.between(0, 31)`.\n"
                  "4. Non-Null Checks: Mandatory attributes (IDs, dates, status flags) must possess zero null values.")

    # 5. Statistical Methods for Each OQ
    add_h1("5. Statistical Methodology Specification (OQ-01 through OQ-10)")
    p_stats_intro = doc.add_paragraph()
    p_stats_intro.add_run("In strict adherence to MCA-level analytical rigor, every operational question is paired with a formal mathematical or statistical technique implemented in `src/stats.py` and `src/metrics.py`:")

    stats_headers = ["OQ ID", "Analytical Investigation", "Statistical / Mathematical Technique", "Python Implementation & SciPy Module", "Statistical Output"]
    stats_data = [
        ("OQ-01", "Overall Churn & Retention Rates", "Vectorized Ratio Aggregation & Trailing 3-Month Simple Moving Average (SMA-3)", "pandas .rolling(window=3).mean()", "Monthly CCR %, CRR %, and 3-month moving trend line"),
        ("OQ-02", "Segment Retention Differences", "Pearson Chi-Square Contingency Test of Independence with Cramer's V effect size", "scipy.stats.chi2_contingency", "Chi-Square statistic, p-value (<0.05), Cramer's V, degrees of freedom"),
        ("OQ-03", "Cohort Retention Curves", "Longitudinal Survival Analysis & Retention Decay Probability", "pandas .pivot_table() cohort triangular matrix", "Percentage of starting cohort surviving at M+1, M+3, M+6, M+12"),
        ("OQ-04", "Customer Lifetime Value (CLV)", "Parametric ARPU Lifespan Model & Kruskal-Wallis H-Test across tiers", "scipy.stats.kruskal", "Mean/median CLV (INR), H-statistic, p-value across plan tiers"),
        ("OQ-05", "Acquisition Channel Longevity", "Chi-Square Test of Independence across acquisition categories", "scipy.stats.chi2_contingency", "Chi-Square statistic, p-value, bar chart with 95% confidence intervals"),
        ("OQ-06", "Support SLAs & Ticket Volume", "Independent Two-Sample Welch's t-test (or Mann-Whitney U for skewed data)", "scipy.stats.mannwhitneyu / ttest_ind", "t/U statistic, p-value (<0.05), median resolution turnaround difference"),
        ("OQ-07", "Usage Telemetry Drop-off", "Independent Two-Sample Student's t-Test on 90-day pre-churn active days", "scipy.stats.ttest_ind", "t-statistic, p-value (<0.01), percentage drop in active logins"),
        ("OQ-08", "Satisfaction vs. Renewal", "Bivariate Pearson (r) & Spearman Rank (rho) Correlation Analysis", "scipy.stats.pearsonr / spearmanr", "Correlation coefficients (r, rho), p-values, renewal probability by tier"),
        ("OQ-09", "MRR Churn & NRR Dynamics", "Financial Waterfall Accounting Model & Net Revenue Retention Ratio", "Custom vectorized pandas accounting logic", "Starting MRR, Expansion MRR, Churned MRR, and NRR % (Benchmark >100%)"),
        ("OQ-10", "At-Risk Account Identification", "Multi-Criteria Decision Analysis (MCDA) with Min-Max Feature Normalization", "Custom weighted linear scoring algorithm", "Risk Score (0-100), risk tier distribution, ranked at-risk data table")
    ]
    t_stats = doc.add_table(rows=len(stats_data) + 1, cols=5)
    format_table(t_stats, stats_headers, stats_data, [Inches(0.7), Inches(1.8), Inches(1.8), Inches(1.3), Inches(0.9)], align_center_cols=[0])

    # 6. UI Architecture & Styling
    add_h1("6. Front-End UI Architecture: Tailwind CSS & Streamlit Design")
    add_h2("6.1 Tailwind CSS Integration via Streamlit Components")
    p_tw = doc.add_paragraph()
    p_tw.add_run("Standard Streamlit KPI metrics provide limited styling flexibility. To achieve a clean, modern enterprise interface, custom HTML cards styled with Tailwind CSS utility classes are rendered using Streamlit's official iframe bridge:\n"
                 "• `src/components.py` encapsulates a reusable `render_kpi_card(title, value, delta, status)` function.\n"
                 "• The component dispatches HTML containing Tailwind CDN scripts (`https://cdn.tailwindcss.com`) inside `streamlit.components.v1.html(..., height=110)`.\n"
                 "• Cards feature subtle drop shadows (`shadow-sm`), rounded borders (`rounded-xl`), slate headers (`text-slate-500`), and semantic status badges (green for positive retention, red for elevated churn).")

    add_h2("6.2 Global Theme Configuration (.streamlit/config.toml)")
    p_theme = doc.add_paragraph()
    p_theme.add_run("The native Streamlit canvas is styled to harmonize with the corporate palette defined in Steps 1–5:")
    toml_code = (
        "[theme]\n"
        "primaryColor = '#1B365D'          # Deep Corporate Navy\n"
        "backgroundColor = '#F8FAFC'       # Light Slate Canvas\n"
        "secondaryBackgroundColor = '#FFFFFF' # Pure White Container Cards\n"
        "textColor = '#1E293B'              # Charcoal High-Contrast Body Text\n"
        "font = 'sans serif'\n\n"
        "[server]\n"
        "headless = true\n"
        "enableCORS = false\n"
        "maxUploadSize = 50"
    )
    add_code_box(toml_code)

    # 7. Non-Functional Requirements
    add_h1("7. Non-Functional Requirements (Performance, Security, Reliability)")
    nfr_headers = ["NFR ID", "Quality Dimension", "Target Technical SLA", "Architectural Enforcement"]
    nfr_data = [
        ("NFR-01", "Performance Latency", "Initial app launch < 5.0 seconds; filter updates < 2.0 seconds.", "@st.cache_data for dataset caching; indexed SQLite foreign keys; pre-aggregated cohort tables."),
        ("NFR-02", "System Reliability", "100% graceful error handling; zero raw unhandled Python tracebacks.", "defensive try-except wrappers on all analytical functions displaying structured st.error() notices."),
        ("NFR-03", "Data Security", "Read-only database access; zero SQL injection vulnerability surface.", "SQLite opened with mode=ro URI parameter; parameterized queries used exclusively; zero client writes."),
        ("NFR-04", "Responsive Design", "Flawless UI rendering on desktop (1920x1080) down to tablet (768px).", "Tailwind flexbox/grid responsive breakpoints (grid-cols-1 md:grid-cols-2 lg:grid-cols-4)."),
        ("NFR-05", "Code Maintainability", "PEP 8 compliance, full type hinting, >= 70% automated test coverage.", "Ruff/Flake8 linting, mypy typing checks, and pytest test suite enforced in local build plan.")
    ]
    t_nfr = doc.add_table(rows=len(nfr_data) + 1, cols=4)
    format_table(t_nfr, nfr_headers, nfr_data, [Inches(0.8), Inches(1.4), Inches(2.1), Inches(2.2)], align_center_cols=[0])

    # 8. Pinned Requirements.txt
    add_h1("8. Software Dependencies & Pinned requirements.txt")
    p_req = doc.add_paragraph()
    p_req.add_run("To ensure absolute build determinism across local workstations and Streamlit Community Cloud, all package versions are strictly pinned:")

    req_text = (
        "# Core Web Framework\n"
        "streamlit==1.32.2\n\n"
        "# Data Manipulation & Computational Analytics\n"
        "pandas==2.2.1\n"
        "numpy==1.26.4\n\n"
        "# Statistical Modeling & Inferential Testing\n"
        "scipy==1.12.0\n"
        "statsmodels==0.14.1\n\n"
        "# Interactive Visualizations\n"
        "plotly==5.20.0\n\n"
        "# Quality Assurance & Automated Testing\n"
        "pytest==8.1.1\n"
        "pytest-cov==5.0.0\n\n"
        "# Reporting & Office Document Generation\n"
        "openpyxl==3.1.5\n"
        "python-docx==1.2.0"
    )
    add_code_box(req_text)

    # 9. Deployment Architecture
    add_h1("9. Cloud Deployment & Operational Runbook")
    p_dep = doc.add_paragraph()
    p_dep.add_run("The application is architected for zero-configuration, continuous deployment on Streamlit Community Cloud:\n"
                  "• Continuous Delivery Pipeline: Pushing commits to the `main` branch of the GitHub repository triggers an automated build and container restart on Streamlit Cloud.\n"
                  "• Environment Configuration: The runtime automatically detects Python 3.12 and installs dependencies strictly from `requirements.txt`.\n"
                  "• Storage Persistence: The embedded database `data/app.db` is bundled directly within the repository, ensuring zero external cloud database credential requirements.\n"
                  "• Health Probing: Automated uptime checks verify HTTP 200 responses and render latency upon container restart.")

    # 10. Technical Risks
    add_h1("10. Technical Risk Assessment & Mitigation Matrix")
    tr_headers = ["Risk ID", "Technical Risk Description", "Likelihood", "Impact", "Architectural Mitigation Strategy"]
    tr_data = [
        ("TR-01", "Tailwind CDN iframe cropping or height mismatch in Streamlit.", "Medium", "Medium", 
         "Standardize iframe height parameters in st.components.v1.html to fixed pixel values (e.g., height=110px) with overflow-hidden."),
        ("TR-02", "Streamlit Community Cloud 1 GB memory limit exceeded during cohort pivot.", "Low", "High", 
         "Downcast numeric datatypes to int32/float32; pre-calculate cohort survival matrices rather than recalculating on every user interaction."),
        ("TR-03", "Plotly rendering latency with large scatter or time-series plots.", "Low", "Medium", 
         "Pre-aggregate raw daily transaction data into monthly metrics; cap displayed points at monthly summary levels."),
        ("TR-04", "SQLite database file corruption during local test execution.", "Low", "High", 
         "Enforce read-only URI connection strings in application code; maintain reproducible synthetic generation script (data/generate_data.py).")
    ]
    t_tr = doc.add_table(rows=len(tr_data) + 1, cols=5)
    format_table(t_tr, tr_headers, tr_data, [Inches(0.8), Inches(2.2), Inches(0.9), Inches(0.8), Inches(1.8)], align_center_cols=[0, 2, 3])

    os.makedirs("docs", exist_ok=True)
    out_path = os.path.join("docs", "07_TRD.docx")
    doc.save(out_path)
    print(f"Successfully generated {out_path}")

if __name__ == "__main__":
    build_step7_trd()
