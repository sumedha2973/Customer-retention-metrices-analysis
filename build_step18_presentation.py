"""
build_step18_presentation.py
Generates the official 11-slide PowerPoint presentation docs/CRMA-2026_Presentation.pptx
in strict compliance with all ABES guidelines, coordinates, formatting, and notes.
"""
import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ─────────────────────────────────────────────────────────────
# CONSTANTS & METADATA
# ─────────────────────────────────────────────────────────────
PROJECT_ID = "CRMA-2026"
TITLE = "Customer Retention Metrices Analysis"
LEARNER = "Sumedha"
ROLL_NO = "2503201450146"
MENTOR = "Ms. Anjali Pandey"
DATE_STR = "November 2026"
APP_URL = "https://2503201450146-crma.streamlit.app"
GITHUB_URL = "https://github.com/sumedha2973"
PPTX_PATH = f"docs/{PROJECT_ID}_Presentation.pptx"

# COLOUR PALETTE
C_RED   = RGBColor(0xA5, 0x1C, 0x30) # A51C30
C_BLACK = RGBColor(0x15, 0x15, 0x14) # 151514
C_DARK  = RGBColor(0x3A, 0x3A, 0x38) # 3A3A38
C_GREY  = RGBColor(0x5E, 0x5E, 0x5B) # 5E5E5B
C_CARD  = RGBColor(0xF4, 0xF1, 0xF1) # F4F1F1
C_PINK  = RGBColor(0xD9, 0xB5, 0xBC) # D9B5BC
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_BORDER= RGBColor(0xD9, 0xD9, 0xD9)
C_NUM_GREY = RGBColor(0x8A, 0x8A, 0x86)

# ─────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────
def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    return prs

def set_zero_insets(tf):
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)

def add_dark_slide_base(prs, slide):
    # Background 151514
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = C_BLACK

    # [ABES logo] (.5,.35,1.6,.5) italic 10 pink, middle, default insets
    tb_logo = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_logo = tb_logo.text_frame
    tf_logo.word_wrap = True
    tf_logo.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_logo.paragraphs[0]
    p.text = "[ABES logo]"
    p.font.name = "Calibri"
    p.font.size = Pt(10)
    p.font.italic = True
    p.font.color.rgb = C_PINK

    # Header right (.2.2,.35,7.3,.5) 11 pink, right, middle, default insets
    tb_hdr = slide.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf_hdr = tb_hdr.text_frame
    tf_hdr.word_wrap = True
    tf_hdr.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_hdr.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    p.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.color.rgb = C_PINK

def add_content_slide_base(prs, slide, title_text, slide_num):
    # White background by default
    # Title (.5,.3,9,.75) Cambria 28 bold black, centred, middle, default insets
    tb_t = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    tf_t.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_t.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = title_text
    p.font.name = "Cambria"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = C_BLACK

    # Footer (.5,5.2,7.5,.3) 9 grey, middle, default insets
    tb_f = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
    tf_f = tb_f.text_frame
    tf_f.word_wrap = True
    tf_f.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_f.paragraphs[0]
    p.text = f"HCLTech Career Shaper ABES Engineering College | {PROJECT_ID} - {TITLE}"
    p.font.name = "Calibri"
    p.font.size = Pt(9)
    p.font.color.rgb = C_GREY

    # Slide number (9,5.2,.5,.3) 9, colour 8A8A86, right, default insets
    tb_n = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
    tf_n = tb_n.text_frame
    tf_n.word_wrap = True
    tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_n.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    p.text = str(slide_num)
    p.font.name = "Calibri"
    p.font.size = Pt(9)
    p.font.color.rgb = C_NUM_GREY

def draw_card(slide, x, y, w, h):
    # rounded rectangle, fill and line F4F1F1
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_CARD
    shape.line.color.rgb = C_CARD
    return shape

def draw_badge(slide, x, y, text):
    # circle .42 red; t centred, 12 bold white
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_RED
    shape.line.color.rgb = C_RED
    tf = shape.text_frame
    set_zero_insets(tf)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = str(text)
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    return shape

def add_badge_card(slide, x, y, w, h, badge_txt, head_txt, body_txt):
    draw_card(slide, x, y, w, h)
    draw_badge(slide, x + 0.2, y + 0.2, badge_txt)
    
    # HEAD (x+.75, y+.17, w-.9, .45) 15 bold black, middle
    tb_h = slide.shapes.add_textbox(Inches(x + 0.75), Inches(y + 0.17), Inches(w - 0.9), Inches(0.45))
    tf_h = tb_h.text_frame
    set_zero_insets(tf_h)
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_h.paragraphs[0]
    p.text = head_txt
    p.font.name = "Calibri"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = C_BLACK

    # BODY (x+.2, y+.72, w-.4, h-.85) 13 dark, top
    tb_b = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_b = tb_b.text_frame
    set_zero_insets(tf_b)
    tf_b.vertical_anchor = MSO_ANCHOR.TOP
    p = tf_b.paragraphs[0]
    p.text = body_txt
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.color.rgb = C_DARK

def add_plain_card(slide, x, y, w, h, head_txt, body_txt):
    draw_card(slide, x, y, w, h)
    # HEAD (x+.2, y+.17, w-.35, .45) 15 bold black, middle
    tb_h = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.17), Inches(w - 0.35), Inches(0.45))
    tf_h = tb_h.text_frame
    set_zero_insets(tf_h)
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_h.paragraphs[0]
    p.text = head_txt
    p.font.name = "Calibri"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = C_BLACK

    # BODY (x+.2, y+.72, w-.4, h-.85) 13 dark, top
    tb_b = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_b = tb_b.text_frame
    set_zero_insets(tf_b)
    tf_b.vertical_anchor = MSO_ANCHOR.TOP
    p = tf_b.paragraphs[0]
    p.text = body_txt
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.color.rgb = C_DARK

def format_table_cell(cell, text, is_header=False):
    cell.margin_left = Inches(0.1)
    cell.margin_right = Inches(0.1)
    cell.margin_top = Inches(0.05)
    cell.margin_bottom = Inches(0.05)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_RED if is_header else C_WHITE
    tf = cell.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.bold = is_header
    p.font.color.rgb = C_WHITE if is_header else C_BLACK

def set_speaker_note(slide, note_text):
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = note_text

# ─────────────────────────────────────────────────────────────
# BUILD THE 11 SLIDES
# ─────────────────────────────────────────────────────────────
def build_presentation():
    prs = create_deck()
    blank_layout = prs.slide_layouts[6]

    # ═════════════════════════════════════════════════════════
    # SLIDE 1: DARK (Title Slide)
    # ═════════════════════════════════════════════════════════
    s1 = prs.slides.add_slide(blank_layout)
    add_dark_slide_base(prs, s1)
    
    # Title (.5,1.55,9,1.5) Cambria 36 bold white, left, bottom, default insets
    tb_title = s1.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    tf_title.vertical_anchor = MSO_ANCHOR.BOTTOM
    p = tf_title.paragraphs[0]
    p.text = f"{PROJECT_ID} - {TITLE}"
    p.font.name = "Cambria"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = C_WHITE

    # Lines (.5,3.2,9,1.7) 15 F4F1F1, left, top, 4 pt before each line, default insets
    tb_lines = s1.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_lines = tb_lines.text_frame
    tf_lines.word_wrap = True
    tf_lines.vertical_anchor = MSO_ANCHOR.TOP

    lines_content = [
        (f"{LEARNER} | Roll No {ROLL_NO}", C_CARD),
        ("MCA 2nd Year Data Analyst", C_CARD),
        (f"Mentor: {MENTOR}", C_CARD),
        (DATE_STR, C_PINK)
    ]
    for idx, (txt, clr) in enumerate(lines_content):
        p = tf_lines.paragraphs[0] if idx == 0 else tf_lines.add_paragraph()
        p.space_before = Pt(4)
        p.text = txt
        p.font.name = "Calibri"
        p.font.size = Pt(15)
        p.font.color.rgb = clr

    set_speaker_note(s1, "Good morning esteemed evaluators and mentor. I am Sumedha, presenting project CRMA-2026: Customer Retention Metrics Analysis. This project delivers an automated analytics platform for B2B SaaS retention and churn diagnostics.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 2: CONTENT "Problem and How Might We"
    # ═════════════════════════════════════════════════════════
    s2 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s2, "Problem and How Might We", 2)

    # BADGE CARD (.5,1.25,4.3,3.7) P, Problem statement: 2-3 sentences
    prob_body = (
        "NoidaSoft Technologies cannot measure customer retention or churn drivers consistently. "
        "Customer, billing, ticket, and usage telemetry data remain isolated across manual spreadsheets. "
        "Account managers discover churned clients only after renewal deadlines pass."
    )
    add_badge_card(s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem statement", prob_body)

    # BADGE CARD (5.2,1.25,4.3,1.75) 1, HMW-01: statement
    hmw1_body = "Provide leadership with an immediate, consistent view of overall retention and churn rates to eliminate reporting delays and conflicting numbers across teams."
    add_badge_card(s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01", hmw1_body)

    # BADGE CARD (5.2,3.2,4.3,1.75) 2, HMW-02: statement
    hmw2_body = "Help customer success managers uncover retention variations across plans, industries, sizes, and regions to deploy targeted segment interventions."
    add_badge_card(s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02", hmw2_body)

    set_speaker_note(s2, "Our core problem is manual, fragmented spreadsheet reporting that causes 4-day delays and reactive churn discovery. Our primary HMWs focus on establishing a single source of truth for leadership and empowering CSMs with granular segment visibility.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 3: CONTENT "Stakeholders and users"
    # ═════════════════════════════════════════════════════════
    s3 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s3, "Stakeholders and users", 3)

    # BADGE CARDS y 1.35, w 2.85, h 3.5, x .5 / 3.55 / 6.6, badges 1 / 2 / 3
    b1_txt = "Head of Customer Success (SH-02).\nNeeds retention visibility by segment to prioritize monthly renewals.\n\nInfluence: High"
    add_badge_card(s3, 0.5, 1.35, 2.85, 3.5, "1", "Primary user", b1_txt)

    b2_txt = "Customer Success Managers (SH-03).\nNeed early risk scores and at-risk lists to contact clients before expiration.\n\nInfluence: Medium"
    add_badge_card(s3, 3.55, 1.35, 2.85, 3.5, "2", "Secondary user", b2_txt)

    b3_txt = "Faculty Mentor (SH-09).\nEvaluates analytical rigor, data architecture, and software test coverage.\n\nInfluence: High"
    add_badge_card(s3, 6.6, 1.35, 2.85, 3.5, "3", "Mentor / Evaluator", b3_txt)

    set_speaker_note(s3, "We mapped ten enterprise stakeholders. The Head of CS is our primary business user directing retention strategy; daily CSMs are secondary users acting on risk rosters; and our mentor guides software quality and academic rigor.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 4: CONTENT "Requirements and traceability"
    # ═════════════════════════════════════════════════════════
    s4 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s4, "Requirements and traceability", 4)

    # TABLE at (.5,1.3), column widths 1, 3.4, 1, 1.2, 1.2, 1.2 (Total 9.0 in)
    rows_data = [
        ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"],
        ["HMW-01", "Official 24-month churn and retention calculation engine", "Must", "US-03", "TC-F-01", "UI-01"],
        ["HMW-02", "Multi-attribute segment filtering (Plan, Industry, Size)", "Must", "US-05", "TC-F-07", "UI-03"],
        ["HMW-03", "12-month cohort survival decay matrix & heatmap", "Must", "US-06", "TC-F-05", "UI-02"],
        ["HMW-10", "4-factor composite customer health risk scoring roster", "Must", "US-12", "TC-F-14", "UI-05"],
    ]
    t_shape = s4.shapes.add_table(len(rows_data), 6, Inches(0.5), Inches(1.3), Inches(9.0), Inches(2.5))
    table = t_shape.table
    col_w = [1.0, 3.4, 1.0, 1.2, 1.2, 1.2]
    for idx, w in enumerate(col_w):
        table.columns[idx].width = Inches(w)

    for r_idx, r_data in enumerate(rows_data):
        for c_idx, val in enumerate(r_data):
            format_table_cell(table.cell(r_idx, c_idx), val, is_header=(r_idx==0))

    # Note .3 below table (.5, y9, .35) italic 12 grey
    tb_note = s4.shapes.add_textbox(Inches(0.5), Inches(4.3), Inches(9.0), Inches(0.35))
    tf_note = tb_note.text_frame
    set_zero_insets(tf_note)
    p = tf_note.paragraphs[0]
    p.text = "Full matrix: Project Document, section 4"
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.italic = True
    p.font.color.rgb = C_GREY

    set_speaker_note(s4, "Every requirement originates from an approved HMW and traces bidirectionally to an Agile User Story, automated test case, and Streamlit screen. This complete thread eliminates scope creep and guarantees accountability.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 5: CONTENT "Data and preparation"
    # ═════════════════════════════════════════════════════════
    s5 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s5, "Data and preparation", 5)

    # Four stats, x .5 / 2.8 / 5.1 / 7.4, w 2.1
    # value y 1.3 h .8 Cambria 30 bold red, middle; label y 2.1 h .35, 12 grey, middle
    stats_data = [
        ("2,489", "telemetry rows"),
        ("33", "analytical columns"),
        ("24 Mo", "contract timeline"),
        ("3 Issues", "synthetic DQ fixes")
    ]
    for idx, (val, lbl) in enumerate(stats_data):
        x_pos = 0.5 + idx * 2.3
        tb_val = s5.shapes.add_textbox(Inches(x_pos), Inches(1.3), Inches(2.1), Inches(0.8))
        tf_val = tb_val.text_frame
        set_zero_insets(tf_val)
        tf_val.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_val.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = val
        p.font.name = "Cambria"
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = C_RED

        tb_lbl = s5.shapes.add_textbox(Inches(x_pos), Inches(2.1), Inches(2.1), Inches(0.35))
        tf_lbl = tb_lbl.text_frame
        set_zero_insets(tf_lbl)
        tf_lbl.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_lbl.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = lbl
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.color.rgb = C_GREY

    # PLAIN CARD (.5,2.7,4.3,2.25) Source
    src_body = (
        "Dataset: Fully synthetic B2B SaaS relational database (Seed 42) stored in SQLite (data/app.db).\n"
        "Representation: One row per customer contract term (229 contracts across 200 Indian corporate accounts) "
        "and one row per monthly customer usage log (2,489 records)."
    )
    add_plain_card(s5, 0.5, 2.7, 4.3, 2.25, "Source", src_body)

    # PLAIN CARD (5.2,2.7,4.3,2.25) Cleaning
    clean_body = (
        "1. Deduplication (DQ-02): Resolved 3 duplicate ticket IDs via SQLite PK constraints.\n"
        "2. Missingness (DQ-01): Managed 65.8% survey nulls via survey-conditional slicing.\n"
        "3. Outliers (DQ-03): Mitigated 5 extreme resolution turnaround spikes (>200h) using non-parametric statistics."
    )
    add_plain_card(s5, 5.2, 2.7, 4.3, 2.25, "Cleaning", clean_body)

    set_speaker_note(s5, "We generated a 24-month relational database in SQLite with 2,489 monthly telemetry records. We injected three realistic data quality issues — duplicate tickets, missing survey scores, and SLA turnaround outliers — and built cleaning pipelines to handle each cleanly.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 6: CONTENT "Solution and technology"
    # ═════════════════════════════════════════════════════════
    s6 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s6, "Solution and technology", 6)

    # Five steps, x .5 / 2.35 / 4.2 / 6.05 / 7.9, y 1.5, w 1.6, h 1.2
    # rounded, radius .1, fill F4F1F1; text centred, middle, 12 bold black, two lines
    # Dashboard step: red fill, white text
    steps_info = [
        ("Raw data\n(CSV / xlsx)", False),
        ("Clean & KPIs\n(Python, pandas)", False),
        ("Store\n(SQLite app.db)", False),
        ("Dashboard\n(Streamlit)", True),
        ("Deploy\n(GitHub -> cloud)", False),
    ]
    step_xs = [0.5, 2.35, 4.2, 6.05, 7.9]
    for idx, (st_text, is_dash) in enumerate(steps_info):
        x = step_xs[idx]
        sh = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.5), Inches(1.6), Inches(1.2))
        sh.fill.solid()
        sh.fill.fore_color.rgb = C_RED if is_dash else C_CARD
        sh.line.color.rgb = C_RED if is_dash else C_CARD
        tf = sh.text_frame
        set_zero_insets(tf)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = st_text
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE if is_dash else C_BLACK

    # ">" between steps, x 2.1 / 3.95 / 5.8 / 7.65, y 1.5, w .25, h 1.2: 18 bold red, centred
    arrow_xs = [2.1, 3.95, 5.8, 7.65]
    for ax in arrow_xs:
        tb_a = s6.shapes.add_textbox(Inches(ax), Inches(1.5), Inches(0.25), Inches(1.2))
        tf_a = tb_a.text_frame
        set_zero_insets(tf_a)
        tf_a.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_a.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = ">"
        p.font.name = "Calibri"
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = C_RED

    # (.5,3.15,9,.5) 13 dark, middle: bold "Tools:" + my tools separated by " · "
    tb_tools = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    tf_tools = tb_tools.text_frame
    set_zero_insets(tf_tools)
    tf_tools.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_tools.paragraphs[0]
    r_b = p.add_run()
    r_b.text = "Tools: "
    r_b.font.bold = True
    r_b.font.color.rgb = C_DARK
    r_t = p.add_run()
    r_t.text = "Python 3.12 · Streamlit 1.32 · SQLite 3 · pandas 2.2 · SciPy 1.12 · Plotly 5.20 · Tailwind CSS · pytest 8.1"
    r_t.font.color.rgb = C_DARK
    p.font.name = "Calibri"
    p.font.size = Pt(13)

    # (.5,3.75,9,.5) 13 italic grey, middle: Run locally...
    tb_cmd = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    tf_cmd = tb_cmd.text_frame
    set_zero_insets(tf_cmd)
    tf_cmd.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_cmd.paragraphs[0]
    p.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.italic = True
    p.font.color.rgb = C_GREY

    set_speaker_note(s6, "Our architecture follows a clean 5-stage pipeline: raw data is cleaned into SQLite, processed through vectorized pandas and SciPy engines, and rendered via a high-performance Streamlit dashboard deployed on the cloud.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 7: CONTENT "Dashboard"
    # ═════════════════════════════════════════════════════════
    s7 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s7, "Dashboard", 7)

    # Frame (.5,1.25,5.5,3.7): fill F4F1F1, dashed 1 pt BFBFBF
    sh_frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    sh_frame.fill.solid()
    sh_frame.fill.fore_color.rgb = C_CARD
    sh_frame.line.color.rgb = RGBColor(0xBF, 0xBF, 0xBF)
    sh_frame.line.width = Pt(1)
    tf_fr = sh_frame.text_frame
    set_zero_insets(tf_fr)
    tf_fr.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_fr.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "[Paste a dashboard screenshot here]"
    p.font.name = "Calibri"
    p.font.size = Pt(14)
    p.font.italic = True
    p.font.color.rgb = C_GREY

    # BADGE x 6.7, y 1.3 / 2.2 / 3.1 / 4.0, numbers 1-4
    # Beside each, (7.25,y,2.25,.42), middle, 12 bold red "UI-01" + short black screen purpose
    screens_list = [
        ("1", 1.3, "UI-01", "Executive KPI Cockpit"),
        ("2", 2.2, "UI-02", "Cohort Survival Heatmap"),
        ("3", 3.1, "UI-03", "Segment & CLV Deep-Dive"),
        ("4", 4.0, "UI-05", "Proactive Health Risk Roster"),
    ]
    for num_t, y_val, ui_code, purpose in screens_list:
        draw_badge(s7, 6.7, y_val, num_t)
        tb_ui = s7.shapes.add_textbox(Inches(7.25), Inches(y_val), Inches(2.25), Inches(0.42))
        tf_ui = tb_ui.text_frame
        set_zero_insets(tf_ui)
        tf_ui.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_ui.paragraphs[0]
        r1 = p.add_run()
        r1.text = f"{ui_code} "
        r1.font.bold = True
        r1.font.color.rgb = C_RED
        r2 = p.add_run()
        r2.text = purpose
        r2.font.bold = False
        r2.font.color.rgb = C_BLACK
        p.font.name = "Calibri"
        p.font.size = Pt(12)

    set_speaker_note(s7, "Here is the interactive application interface. It showcases 5 headline Tailwind KPI cards, 24-month retention trends, and instant sidebar multi-attribute filtering. Users can seamlessly drill into cohorts, segment CLVs, and actionable risk rosters.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 8: CONTENT "Key findings"
    # ═════════════════════════════════════════════════════════
    s8 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s8, "Key findings", 8)

    # Three CARDS y 1.3, w 2.85, h 3.6, x .5 / 3.55 / 6.6
    # Inside, at x+.2, w 2.45:
    # number y 1.5 h 1 Cambria 40 bold red, middle;
    # meaning y 2.6 h .9, 14 bold black, top;
    # Evidence: KPI or test y 3.6 h 1.1, 12 grey, top.
    findings_data = [
        ("6.1 vs 3.4", "Tickets trigger churn", "Mann-Whitney U p=0.001;\nchurned clients had 2x more support tickets."),
        ("23% vs 10%", "Starter churns faster", "Starter tier had 23% churn vs 9.8% for Enterprise (CLV INR 27K vs 179K)."),
        ("25% vs 16%", "Paid Social is weakest", "Chi-Square p=0.024;\nPartner & Referral channels stay significantly longer.")
    ]
    card_xs = [0.5, 3.55, 6.6]
    for idx, (num_val, mean_val, evid_val) in enumerate(findings_data):
        x = card_xs[idx]
        draw_card(s8, x, 1.3, 2.85, 3.6)

        tb_n = s8.shapes.add_textbox(Inches(x + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        tf_n = tb_n.text_frame
        set_zero_insets(tf_n)
        tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf_n.paragraphs[0]
        p.text = num_val
        p.font.name = "Cambria"
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = C_RED

        tb_m = s8.shapes.add_textbox(Inches(x + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        tf_m = tb_m.text_frame
        set_zero_insets(tf_m)
        tf_m.vertical_anchor = MSO_ANCHOR.TOP
        p = tf_m.paragraphs[0]
        p.text = mean_val
        p.font.name = "Calibri"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = C_BLACK

        tb_e = s8.shapes.add_textbox(Inches(x + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        tf_e = tb_e.text_frame
        set_zero_insets(tf_e)
        tf_e.vertical_anchor = MSO_ANCHOR.TOP
        p = tf_e.paragraphs[0]
        p.text = f"Evidence: {evid_val}"
        p.font.name = "Calibri"
        p.font.size = Pt(12)
        p.font.color.rgb = C_GREY

    set_speaker_note(s8, "Our analysis revealed three major empirical findings: churned customers log nearly twice as many support tickets; Starter tier accounts churn at over double the rate of Enterprise accounts; and Paid Social brings the highest-churn clients.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 9: CONTENT "Testing and validation"
    # ═════════════════════════════════════════════════════════
    s9 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s9, "Testing and validation", 9)

    # Passed / total (.5,1.35,2.6,1) Cambria 44 bold red, middle
    tb_score = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    tf_score = tb_score.text_frame
    set_zero_insets(tf_score)
    tf_score.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_score.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "55 / 55"
    p.font.name = "Cambria"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = C_RED

    # (.5,2.4,2.6,.9) 13 grey, middle: test cases passed / pytest: n / manual: n
    tb_sub = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    tf_sub = tb_sub.text_frame
    set_zero_insets(tf_sub)
    tf_sub.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_sub.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "test cases passed\npytest: 35 / manual: 20"
    p.font.name = "Calibri"
    p.font.size = Pt(13)
    p.font.color.rgb = C_GREY

    # TABLE at (3.4,1.35), column widths 1.1, 3.9, 1.1 (Total 6.1 in)
    test_rows = [
        ["Test", "Checks", "Result"],
        ["TC-T-05..06", "Primary key uniqueness & 100% foreign key referential integrity", "Pass"],
        ["TC-T-15..20", "Mathematical formula parity on deterministic fixture datasets", "Pass"],
        ["TC-F-14..16", "Health risk scoring tier categorization & CSV roster download", "Pass"],
    ]
    t_test_shape = s9.shapes.add_table(len(test_rows), 3, Inches(3.4), Inches(1.35), Inches(6.1), Inches(2.5))
    t_test = t_test_shape.table
    test_widths = [1.1, 3.9, 1.1]
    for idx, w in enumerate(test_widths):
        t_test.columns[idx].width = Inches(w)

    for r_idx, r_data in enumerate(test_rows):
        for c_idx, val in enumerate(r_data):
            format_table_cell(t_test.cell(r_idx, c_idx), val, is_header=(r_idx==0))

    set_speaker_note(s9, "Quality was verified through 55 comprehensive test cases. All 35 automated technical pytest checks passed cleanly with zero formula discrepancies, and 20 manual end-to-end functional test cases verified the complete user journey.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 10: CONTENT "Recommendation, limitations and next steps"
    # ═════════════════════════════════════════════════════════
    s10 = prs.slides.add_slide(blank_layout)
    add_content_slide_base(prs, s10, "Recommendation, limitations and next steps", 10)

    # BADGE CARDS y 1.25, w 2.85, h 3.7, x .5 / 3.58 / 6.65
    # R Recommendation; L Limitations; N Next steps
    rec_body = (
        "Reallocate CS effort toward accounts with >= 4 tickets and login drops > 25%.\n\n"
        "Offer dedicated onboarding to Starter tier clients to reduce first-year attrition."
    )
    add_badge_card(s10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendation", rec_body)

    lim_body = (
        "Synthetic dataset (Seed 42) reflects simulated SMB patterns rather than live telemetry.\n\n"
        "CSAT ratings had 65.8% missingness reflecting realistic survey response friction."
    )
    add_badge_card(s10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations", lim_body)

    nxt_body = (
        "Implement BR-13 automated live database synchronization in Phase 2.\n\n"
        "Integrate automated webhook alerts to trigger email/SMS notifications when risk >= 70."
    )
    add_badge_card(s10, 6.65, 1.25, 2.85, 3.7, "N", "Next steps", nxt_body)

    set_speaker_note(s10, "Our key recommendation is deploying proactive outreach when support tickets exceed four. An honest limitation is that our data is synthetic with 65% survey non-response. Next steps will focus on live ERP sync and automated alerting.")

    # ═════════════════════════════════════════════════════════
    # SLIDE 11: DARK (Closing Slide)
    # ═════════════════════════════════════════════════════════
    s11 = prs.slides.add_slide(blank_layout)
    add_dark_slide_base(prs, s11)

    # Title (.5,1.55,9,1.5) Cambria 36 bold white, left, bottom, default insets
    tb_t11 = s11.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t11 = tb_t11.text_frame
    tf_t11.word_wrap = True
    tf_t11.vertical_anchor = MSO_ANCHOR.BOTTOM
    p = tf_t11.paragraphs[0]
    p.text = "Live demo and Q&A"
    p.font.name = "Cambria"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = C_WHITE

    # Lines (.5,3.2,9,1.7) 15 F4F1F1, left, top, 4 pt before each line, default insets
    tb_l11 = s11.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_l11 = tb_l11.text_frame
    tf_l11.word_wrap = True
    tf_l11.vertical_anchor = MSO_ANCHOR.TOP

    lines11_content = [
        (f"App: {APP_URL}", C_CARD),
        (f"Code: {GITHUB_URL}", C_CARD),
        ("Thank you", C_PINK)
    ]
    for idx, (txt, clr) in enumerate(lines11_content):
        p = tf_l11.paragraphs[0] if idx == 0 else tf_l11.add_paragraph()
        p.space_before = Pt(4)
        p.text = txt
        p.font.name = "Calibri"
        p.font.size = Pt(15)
        p.font.color.rgb = clr

    set_speaker_note(s11, "I will now open the live Streamlit application for the interactive demonstration. Thank you for your time and guidance, and I welcome any questions.")

    # Save
    os.makedirs(os.path.dirname(PPTX_PATH), exist_ok=True)
    prs.save(PPTX_PATH)
    print(f"✅ Presentation successfully saved: {PPTX_PATH}")

if __name__ == "__main__":
    build_presentation()
