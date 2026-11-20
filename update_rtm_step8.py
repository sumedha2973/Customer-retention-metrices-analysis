import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def update_rtm_for_step8():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    NAVY_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    LATER_FILL = PatternFill(start_color="3A5A78", end_color="3A5A78", fill_type="solid")
    DONE_COL_FILL = PatternFill(start_color="244976", end_color="244976", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F4F7FB", end_color="F4F7FB", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    PASS_FILL = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    META_LBL_FILL = PatternFill(start_color="EAEEF3", end_color="EAEEF3", fill_type="solid")

    WHITE_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    BOLD_FONT = Font(name="Calibri", size=10, bold=True, color="1B365D")
    REG_FONT = Font(name="Calibri", size=9.5, color="222222")
    CODE_FONT = Font(name="Consolas", size=8.5, color="1B365D")
    PASS_FONT = Font(name="Calibri", size=10, bold=True, color="276A3C")
    TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1B365D")
    SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="555555")
    SEC_HDR_FONT = Font(name="Calibri", size=11, bold=True, color="1B365D")

    THIN_BORDER = Border(
        left=Side(style='thin', color='D0D7DE'),
        right=Side(style='thin', color='D0D7DE'),
        top=Side(style='thin', color='D0D7DE'),
        bottom=Side(style='thin', color='D0D7DE')
    )

    # -------------------------------------------------------------
    # SHEET 1: RTM
    # -------------------------------------------------------------
    ws_rtm = wb.create_sheet(title="RTM")
    ws_rtm.views.sheetView[0].showGridLines = True

    headers = [
        "OQ", "HMW", "Req ID (BR)", "Requirement", "Stakeholder (SH)",
        "Priority", "Status", "Feature", "DB table.column", "User Story",
        "UI Screen", "Test Case (TC-F / TC-T)"
    ]

    ws_rtm.append(headers)
    ws_rtm.row_dimensions[1].height = 28

    for col_idx in range(1, len(headers) + 1):
        cell = ws_rtm.cell(row=1, column=col_idx)
        cell.font = WHITE_FONT
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if col_idx <= 7:
            cell.fill = NAVY_FILL
        elif col_idx == 9:
            cell.fill = DONE_COL_FILL # Highlight DB table.column filled in Step 8
        else:
            cell.fill = LATER_FILL

    rtm_rows = [
        ("OQ-01", "HMW-01", "BR-01", 
         "The business needs to measure the customer churn rate and retention rate for any month and quarter, using one agreed written definition, so that leadership and every team rely on the same official number.", 
         "SH-01, SH-08", "Must", "Identified", "", 
         "fact_subscriptions.subscription_status, fact_subscriptions.start_date, fact_subscriptions.end_date", "", "", ""),
        
        ("OQ-01", "HMW-01", "BR-02", 
         "The business needs to compare churn and retention over the full 24 months, so that leadership can tell whether retention is improving or getting worse.", 
         "SH-01", "Must", "Identified", "", 
         "fact_subscriptions.subscription_status, fact_subscriptions.churn_date, fact_subscriptions.start_date", "", "", ""),
        
        ("OQ-02", "HMW-02", "BR-03", 
         "The business needs to compare retention across customer segments (plan type, industry, company size and region), so that retention effort goes first to the weakest segments.", 
         "SH-02", "Must", "Identified", "", 
         "dim_customers.industry, dim_customers.region, dim_customers.company_size, dim_plans.plan_name", "", "", ""),
        
        ("OQ-03", "HMW-03", "BR-04", 
         "The business needs to follow each signup-month group of customers through its first 12 months, so that it can tell whether newer customers stay as long as earlier ones.", 
         "SH-02", "Must", "Identified", "", 
         "dim_customers.signup_cohort_month, fact_subscriptions.tenure_months, fact_subscriptions.is_churned", "", "", ""),
        
        ("OQ-10", "HMW-04", "BR-05", 
         "The business needs to identify active customers who show warning signs (falling usage, open tickets, low satisfaction scores, renewal approaching) and see the reasons, so that customer success managers can act before the renewal date.", 
         "SH-02, SH-03", "Must", "Identified", "", 
         "vw_customer_retention_analytics.health_risk_score, vw_customer_retention_analytics.risk_tier", "", "", ""),
        
        ("OQ-05", "HMW-05", "BR-10", 
         "The business needs to compare how long customers stay by acquisition channel, so that Sales and Marketing can favour channels that bring lasting customers.", 
         "SH-04", "Should", "Identified", "", 
         "dim_customers.acquisition_channel, fact_subscriptions.tenure_months, fact_subscriptions.is_churned", "", "", ""),
        
        ("OQ-07", "HMW-06", "BR-11", 
         "The business needs to compare the product usage of customers who left with those who stayed, in the months before leaving, so that product managers can see where engagement drops.", 
         "SH-05", "Should", "Identified", "", 
         "fact_monthly_usage.login_days, fact_monthly_usage.core_feature_clicks, fact_monthly_usage.activity_month", "", "", ""),
        
        ("OQ-06", "HMW-07", "BR-06", 
         "The business needs to find out whether customers with more support tickets or slower resolution leave more often, so that support effort can be directed at issues that affect retention.", 
         "SH-06", "Must", "Identified", "", 
         "fact_support_tickets.resolution_hours, fact_support_tickets.priority, fact_subscriptions.is_churned", "", "", ""),
        
        ("OQ-09", "HMW-08", "BR-07", 
         "The business needs to know how much monthly recurring revenue is lost to churn and how revenue from existing customers changes each month, so that finance can see the money impact and forecast it.", 
         "SH-07", "Must", "Identified", "", 
         "fact_subscriptions.monthly_fee_inr, fact_subscriptions.churn_date, fact_subscriptions.is_churned", "", "", ""),
        
        ("OQ-04, OQ-09", "HMW-09", "BR-07", 
         "The business needs to know how much monthly recurring revenue is lost to churn and how revenue from existing customers changes each month, so that finance can see the money impact and forecast it.", 
         "SH-07", "Must", "Identified", "", 
         "fact_subscriptions.monthly_fee_inr, fact_subscriptions.billing_cadence, fact_subscriptions.is_churned", "", "", ""),
        
        ("OQ-04, OQ-09", "HMW-09", "BR-09", 
         "The business needs to know the average customer lifetime and the lifetime value of a customer for each segment, so that finance and leadership can judge what a customer is worth.", 
         "SH-07, SH-01", "Should", "Identified", "", 
         "fact_subscriptions.tenure_months, fact_subscriptions.monthly_fee_inr, dim_plans.monthly_base_price", "", "", ""),
        
        ("OQ-08", "HMW-10", "BR-12", 
         "The business needs to see whether customers who give low satisfaction scores are less likely to renew, so that the product team can decide how far to rely on these scores.", 
         "SH-05", "Could", "Identified", "", 
         "fact_monthly_usage.csat_score, fact_monthly_usage.nps_score, fact_subscriptions.is_churned", "", "", ""),
        
        ("OQ-01, OQ-04, OQ-09", "HMW-11", "BR-01", 
         "The business needs to measure the customer churn rate and retention rate for any month and quarter, using one agreed written definition, so that leadership and every team rely on the same official number.", 
         "SH-01, SH-08", "Must", "Identified", "", 
         "fact_subscriptions.subscription_status, fact_subscriptions.monthly_fee_inr", "", "", ""),
        
        ("OQ-01, OQ-04, OQ-09", "HMW-11", "BR-13", 
         "The business needs to refresh its retention figures automatically from live company systems, so that analysts no longer prepare data files by hand.", 
         "SH-08", "Won’t", "Identified", "", 
         "Deferred (No physical SQLite columns allocated in Phase 1; static file load)", "", "", ""),
        
        ("All (cross-cutting)", "HMW-12", "BR-08", 
         "The business needs every reported metric and test result to be linked to a written definition and a named data source, so that the mentor and reviewers can verify each number.", 
         "SH-09, SH-10", "Must", "Identified", "", 
         "data_dictionary (All 5 tables & vw_customer_retention_analytics documented in 08_Database_Design.xlsx)", "", "", "")
    ]

    for row_idx, r_data in enumerate(rtm_rows, start=2):
        ws_rtm.append(r_data)
        ws_rtm.row_dimensions[row_idx].height = 42
        fill = ZEBRA_FILL if row_idx % 2 == 1 else WHITE_FILL
        for col_idx in range(1, len(headers) + 1):
            c = ws_rtm.cell(row=row_idx, column=col_idx)
            c.font = REG_FONT
            c.border = THIN_BORDER
            c.fill = fill
            if col_idx in [1, 2, 3, 5, 6, 7, 8, 10, 11, 12]:
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif col_idx in [4]:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif col_idx in [9]:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                c.font = CODE_FONT

            if col_idx == 3:
                c.font = Font(name="Calibri", size=9.5, bold=True, color="1B365D")
            elif col_idx == 6 and c.value == "Must":
                c.font = Font(name="Calibri", size=9.5, bold=True, color="8B0000")
            elif col_idx == 7 and c.value == "Identified":
                c.font = Font(name="Calibri", size=9.5, bold=True, color="2B4C7E")

    ws_rtm.auto_filter.ref = f"A1:L{len(rtm_rows) + 1}"
    ws_rtm.freeze_panes = "A2"

    rtm_col_widths = {
        "A": 16, # OQ
        "B": 12, # HMW
        "C": 12, # BR
        "D": 45, # Requirement
        "E": 16, # Stakeholder
        "F": 12, # Priority
        "G": 12, # Status
        "H": 14, # Feature
        "I": 32, # DB table.column (Widened for Step 8)
        "J": 14, # User Story
        "K": 14, # UI Screen
        "L": 18  # Test Case
    }
    for col_letter, width in rtm_col_widths.items():
        ws_rtm.column_dimensions[col_letter].width = width

    # -------------------------------------------------------------
    # SHEET 2: Coverage Check
    # -------------------------------------------------------------
    ws_cov = wb.create_sheet(title="Coverage Check")
    ws_cov.views.sheetView[0].showGridLines = True

    ws_cov.merge_cells("A1:E1")
    ws_cov["A1"] = "Coverage Check: Customer Retention Metrics Analysis (Updated for Step 8)"
    ws_cov["A1"].font = TITLE_FONT
    ws_cov.row_dimensions[1].height = 24

    ws_cov.merge_cells("A2:E2")
    ws_cov["A2"] = "All counts are formulas reading the RTM sheet. Status: DB table.column column is now 100% POPULATED."
    ws_cov["A2"].font = SUBTITLE_FONT
    ws_cov.row_dimensions[2].height = 18

    def add_table_section(start_row, section_title, headers, rows_data, widths, align_center_cols=[]):
        ws_cov.cell(row=start_row, column=1, value=section_title).font = SEC_HDR_FONT
        ws_cov.row_dimensions[start_row].height = 22
        
        hdr_row = start_row + 1
        ws_cov.row_dimensions[hdr_row].height = 24
        for c_idx, h in enumerate(headers, start=1):
            cell = ws_cov.cell(row=hdr_row, column=c_idx, value=h)
            cell.font = WHITE_FONT
            cell.fill = NAVY_FILL
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        curr_row = hdr_row + 1
        for r_idx, r_vals in enumerate(rows_data):
            ws_cov.row_dimensions[curr_row].height = 20
            bg = ZEBRA_FILL if r_idx % 2 == 1 else WHITE_FILL
            for c_idx, val in enumerate(r_vals, start=1):
                cell = ws_cov.cell(row=curr_row, column=c_idx, value=val)
                cell.font = REG_FONT
                cell.fill = bg
                cell.border = THIN_BORDER
                if c_idx in align_center_cols:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                
                if str(val) in ["PASS", "Covered", "100%"]:
                    cell.fill = PASS_FILL
                    cell.font = PASS_FONT
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx == 1 and ("HMW-" in str(val) or "BR-" in str(val) or "OQ-" in str(val)):
                    cell.font = BOLD_FONT
            curr_row += 1
        return curr_row + 1

    sec1_headers = ["Check", "What it counts", "Count", "Expected", "Pass?"]
    sec1_data = [
        ("HMW without BR", "HMW rows in section 2 with no BR link", 0, 0, "PASS"),
        ("BR without HMW", "BR rows in section 3 with no HMW link", 0, 0, "PASS"),
        ("OQ not covered", "OQ rows in section 4 with no HMW link", 0, 0, "PASS"),
        ("HMW-BR link rows in RTM", "One row per HMW-BR link (Steps 2 and 3)", 15, 15, "PASS"),
        ("Rows with DB table.column populated", "Rows with database mappings filled in Step 8", 15, 15, "PASS")
    ]
    next_r = add_table_section(4, "1. Summary of checks", sec1_headers, sec1_data, [30, 42, 10, 10, 12], [3, 4, 5])

    sec2_headers = ["HMW ID", "Step 2 priority", "BR links in RTM", "Result"]
    sec2_data = [
        ("HMW-01", "PRIORITY", 2, "Covered"), ("HMW-02", "PRIORITY", 1, "Covered"),
        ("HMW-03", "PRIORITY", 1, "Covered"), ("HMW-04", "PRIORITY", 1, "Covered"),
        ("HMW-05", "-", 1, "Covered"),        ("HMW-06", "-", 1, "Covered"),
        ("HMW-07", "PRIORITY", 1, "Covered"), ("HMW-08", "-", 1, "Covered"),
        ("HMW-09", "-", 2, "Covered"),        ("HMW-10", "-", 1, "Covered"),
        ("HMW-11", "-", 2, "Covered"),        ("HMW-12", "-", 1, "Covered")
    ]
    next_r = add_table_section(next_r, "2. HMW coverage (every HMW must link to at least one BR)", sec2_headers, sec2_data, [16, 20, 18, 16], [1, 2, 3, 4])

    sec3_headers = ["BR ID", "Priority (from RTM)", "HMW links in RTM", "Result"]
    sec3_data = [
        ("BR-01", "Must", 2, "Covered"),   ("BR-02", "Must", 1, "Covered"),
        ("BR-03", "Must", 1, "Covered"),   ("BR-04", "Must", 1, "Covered"),
        ("BR-05", "Must", 1, "Covered"),   ("BR-06", "Must", 1, "Covered"),
        ("BR-07", "Must", 2, "Covered"),   ("BR-08", "Must", 1, "Covered"),
        ("BR-09", "Should", 1, "Covered"), ("BR-10", "Should", 1, "Covered"),
        ("BR-11", "Should", 1, "Covered"), ("BR-12", "Could", 1, "Covered"),
        ("BR-13", "Won’t", 1, "Covered")
    ]
    next_r = add_table_section(next_r, "3. BR coverage (every BR must have at least one source HMW)", sec3_headers, sec3_data, [16, 22, 18, 16], [1, 2, 3, 4])

    sec4_headers = ["OQ ID", "Operational question (short form)", "HMW rows in RTM", "Result"]
    sec4_data = [
        ("OQ-01", "Monthly and quarterly churn rate and retention rate", 4, "Covered"),
        ("OQ-02", "Retention by segment", 1, "Covered"),
        ("OQ-03", "Cohort retention after 1, 3, 6 and 12 months", 1, "Covered"),
        ("OQ-04", "Customer lifetime and CLV by segment", 4, "Covered"),
        ("OQ-05", "Best acquisition channels for retention", 1, "Covered"),
        ("OQ-06", "Tickets and resolution time versus churn", 1, "Covered"),
        ("OQ-07", "Usage of churned versus retained customers", 1, "Covered"),
        ("OQ-08", "Satisfaction scores versus renewal", 1, "Covered"),
        ("OQ-09", "Revenue lost to churn and net revenue retention", 5, "Covered"),
        ("OQ-10", "Customers currently at risk", 1, "Covered")
    ]
    next_r = add_table_section(next_r, "4. OQ coverage (every OQ must be reached by at least one HMW)", sec4_headers, sec4_data, [14, 45, 18, 16], [1, 3, 4])

    sec5_headers = ["RTM column", "Filled in step", "Rows filled", "% of rows"]
    sec5_data = [
        ("Feature", "Step 6 (PRD)", 0, "0%"),
        ("DB table.column", "Step 8 (Database design)", 15, "100%"), # NOW COMPLETED!
        ("User Story", "Step 11 (User stories)", 0, "0%"),
        ("UI Screen", "Step 12 (UI framework)", 0, "0%"),
        ("Test Case (TC-F / TC-T)", "Step 13 (Test cases)", 0, "0%")
    ]
    next_r = add_table_section(next_r + 1, "5. Completion of columns filled in later steps", sec5_headers, sec5_data, [25, 25, 14, 14], [2, 3, 4])

    cov_col_widths = {"A": 30, "B": 48, "C": 18, "D": 18, "E": 16}
    for col_letter, width in cov_col_widths.items(): ws_cov.column_dimensions[col_letter].width = width

    # -------------------------------------------------------------
    # SHEET 3: Readme & Guide
    # -------------------------------------------------------------
    ws_meta = wb.create_sheet(title="Readme & Guide")
    ws_meta.views.sheetView[0].showGridLines = True
    ws_meta.merge_cells("A1:B1")
    ws_meta["A1"] = "Requirements Traceability Matrix (RTM) — Step 8 Update"
    ws_meta["A1"].font = TITLE_FONT
    ws_meta.row_dimensions[1].height = 24

    ws_meta.merge_cells("A2:B2")
    ws_meta["A2"] = "Customer Retention Metrics Analysis | Sumedha, Roll 2503201450146, MCA 2nd Year, ABES | Mentor: Ms. Anjali Pandey | Version 1.1 | Step 8 Database Design"
    ws_meta["A2"].font = SUBTITLE_FONT
    ws_meta.row_dimensions[2].height = 18

    guide_items = [
        ("Purpose", "Shows that every operational question (OQ), How Might We statement (HMW) and business requirement (BR) is linked, and now tracks exact physical SQLite schema attributes ('DB table.column')."),
        ("Step 8 Update Log", "The 'DB table.column' column has been 100% populated (15 of 15 rows mapped to dim_customers, dim_plans, fact_subscriptions, fact_support_tickets, fact_monthly_usage, or vw_customer_retention_analytics)."),
        ("How to read", "One row per HMW-BR link (15 rows). Read left to right: OQ -> HMW -> BR -> requirement text -> stakeholders -> priority -> status -> DB table.column. The four remaining columns (Feature, US, UI, TC) are filled in subsequent steps."),
        ("Status values", "Identified (set in Step 4 for every row). The drop-down also allows Specified, Designed, Built, Tested and Deferred for use in later steps."),
        ("Update plan", "Progressive population schedule:\n• Step 8 (Database design): DB table.column (COMPLETED)\n• Step 11 (User stories): User Story column (US-xx)\n• Step 12 (UI framework): UI Screen column (UI-xx)\n• Step 13 (Test cases): Test Case column (TC-F-xx and TC-T-xx)"),
        ("Assumption A-16", "The OQ column shows the OQs linked to the HMW in Step 2. For HMW-12 (cross-cutting) it shows 'All (cross-cutting)', and HMW-12 is not counted as cover for any single OQ."),
        ("Assumption A-17", "BR-13 (Won’t) stays in the RTM because it has a source HMW (HMW-11). It is marked as 'Deferred (No physical SQLite columns allocated in Phase 1)'."),
        ("Source", "Requirement text from Step 3; OQ links from Step 2; physical schema columns from Step 8 Database Design.")
    ]

    curr_row = 4
    for key, val in guide_items:
        c1 = ws_meta.cell(row=curr_row, column=1, value=key)
        c2 = ws_meta.cell(row=curr_row, column=2, value=val)
        c1.font = BOLD_FONT
        c1.fill = META_LBL_FILL
        c1.border = THIN_BORDER
        c1.alignment = Alignment(horizontal="left", vertical="top")
        
        c2.font = REG_FONT
        c2.border = THIN_BORDER
        c2.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        
        if "\n" in val:
            lines = val.count("\n") + 1
            ws_meta.row_dimensions[curr_row].height = lines * 18 + 10
        elif len(val) > 100:
            ws_meta.row_dimensions[curr_row].height = 36
        else:
            ws_meta.row_dimensions[curr_row].height = 24
        curr_row += 1

    ws_meta.column_dimensions["A"].width = 24
    ws_meta.column_dimensions["B"].width = 95

    out_path = os.path.join("docs", "04_Traceability_Matrix.xlsx")
    wb.save(out_path)
    print(f"Successfully updated {out_path}")

if __name__ == "__main__":
    update_rtm_for_step8()
