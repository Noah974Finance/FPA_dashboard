import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_range

def create_template(filename="FPA_Template_Blank.xlsx"):
    workbook = xlsxwriter.Workbook(filename)
    
    title_format = workbook.add_format({'bold': True, 'font_size': 14})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1})
    section_format = workbook.add_format({'bold': True, 'bg_color': '#D0D0D0', 'border': 1})
    
    input_format = workbook.add_format({'bg_color': '#FFF9C4', 'border': 1})
    input_pct_format = workbook.add_format({'bg_color': '#FFF9C4', 'border': 1, 'num_format': '0.0%'})
    
    calc_format = workbook.add_format({'bg_color': '#F5F5F5', 'border': 1, 'font_color': '#505050'})
    calc_pct_format = workbook.add_format({'bg_color': '#F5F5F5', 'border': 1, 'font_color': '#505050', 'num_format': '0.0%'})
    
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    weeks = [f"Wk {i}" for i in range(1, 14)]
    
    # ── 1. BvA Variance ────────────────────────────────────────────────────────
    ws1 = workbook.add_worksheet("1. BvA Variance")
    ws1.write(0, 0, "Company Name - FPA Template | FY 2024", title_format)
    ws1.set_column('A:A', 35)
    
    headers_bva = ["Metric"] + months + ["YTD Budget", "YTD Actual", "Var Abs", "Var %", "Status", "Type"]
    for col, h in enumerate(headers_bva):
        ws1.write(3, col, h, header_format)
        
    row = 4
    for section, items in [("REVENUE", ["Subscription Revenue", "Professional Services"]), 
                           ("EXPENSES", ["Salaries", "Marketing", "G&A"]), 
                           ("NET INCOME", ["Net Income"])]:
        ws1.write(row, 0, section, section_format)
        section_start_row = row + 1
        row += 1
        
        for item in items:
            # Budget Row
            ws1.write(row, 0, f"{item} - Budget")
            for c in range(1, 13): ws1.write(row, c, 100, input_format)
            ws1.write_formula(row, 13, f"=SUM({xl_range(row, 1, row, 12)})", calc_format)
            ws1.write(row, 14, "", calc_format) # No YTD actual for Budget line
            ws1.write(row, 15, "", calc_format)
            ws1.write(row, 16, "", calc_pct_format)
            ws1.write(row, 17, "", input_format)
            ws1.write(row, 18, "Budget")
            bud_row = row
            row += 1
            
            # Actual Row
            ws1.write(row, 0, f"{item} - Actual")
            for c in range(1, 13): ws1.write(row, c, 110, input_format)
            ws1.write(row, 13, "", calc_format)
            ws1.write_formula(row, 14, f"=SUM({xl_range(row, 1, row, 12)})", calc_format)
            ws1.write(row, 15, "", calc_format)
            ws1.write(row, 16, "", calc_pct_format)
            ws1.write(row, 17, "", input_format)
            ws1.write(row, 18, "Actual")
            act_row = row
            row += 1
            
            # Variance Row
            ws1.write(row, 0, f"{item} Variance")
            for c in range(1, 13):
                bud_cell = xl_rowcol_to_cell(bud_row, c)
                act_cell = xl_rowcol_to_cell(act_row, c)
                if section == "EXPENSES":
                    ws1.write_formula(row, c, f"={bud_cell}-{act_cell}", calc_format)
                else:
                    ws1.write_formula(row, c, f"={act_cell}-{bud_cell}", calc_format)
                    
            bud_ytd = xl_rowcol_to_cell(bud_row, 13)
            act_ytd = xl_rowcol_to_cell(act_row, 14)
            if section == "EXPENSES":
                ws1.write_formula(row, 15, f"={bud_ytd}-{act_ytd}", calc_format)
            else:
                ws1.write_formula(row, 15, f"={act_ytd}-{bud_ytd}", calc_format)
                
            var_abs = xl_rowcol_to_cell(row, 15)
            ws1.write_formula(row, 16, f"=IFERROR({var_abs}/{bud_ytd}, 0)", calc_pct_format)
            ws1.write(row, 17, "On Track", input_format)
            ws1.write(row, 18, "Variance")
            row += 1
            
        # Total Row
        ws1.write(row, 0, f"TOTAL {section}", section_format)
        for c in range(1, 13):
            # Sum only Budget rows
            col_letter = xl_rowcol_to_cell(0, c)[0]
            ws1.write_formula(row, c, f"=SUMIFS({col_letter}{section_start_row+1}:{col_letter}{row}, $S${section_start_row+1}:$S${row}, \"Budget\")", calc_format)
        
        ws1.write_formula(row, 13, f"=SUMIFS(N{section_start_row+1}:N{row}, $S${section_start_row+1}:$S${row}, \"Budget\")", calc_format)
        ws1.write_formula(row, 14, f"=SUMIFS(O{section_start_row+1}:O{row}, $S${section_start_row+1}:$S${row}, \"Actual\")", calc_format)
        if section == "EXPENSES":
            ws1.write_formula(row, 15, f"={xl_rowcol_to_cell(row, 13)}-{xl_rowcol_to_cell(row, 14)}", calc_format)
        else:
            ws1.write_formula(row, 15, f"={xl_rowcol_to_cell(row, 14)}-{xl_rowcol_to_cell(row, 13)}", calc_format)
            
        ws1.write_formula(row, 16, f"=IFERROR({xl_rowcol_to_cell(row, 15)}/{xl_rowcol_to_cell(row, 13)}, 0)", calc_pct_format)
        ws1.write(row, 17, "", input_format)
        ws1.write(row, 18, "Total")
        row += 1
        
    # ── 2. Headcount Planning ──────────────────────────────────────────────────
    ws2 = workbook.add_worksheet("2. Headcount Planning")
    ws2.write(0, 0, "Company Name - FPA Template | FY 2024", title_format)
    ws2.set_column('A:A', 30)
    
    depts = ["Sales", "Marketing", "Product/Engineering", "Customer Success", "G&A"]
    for r, dept in enumerate(depts, 4):
        ws2.write(r, 0, dept)
        ws2.write(r, 1, 10, input_format) # Start HC
        # Q1
        ws2.write(r, 2, 2, input_format) # Hires
        ws2.write(r, 3, 1, input_format) # Departs
        ws2.write_formula(r, 4, f"={xl_rowcol_to_cell(r,1)}+{xl_rowcol_to_cell(r,2)}-{xl_rowcol_to_cell(r,3)}", calc_format)
        # Q2
        ws2.write(r, 5, 2, input_format)
        ws2.write(r, 6, 1, input_format)
        ws2.write_formula(r, 7, f"={xl_rowcol_to_cell(r,4)}+{xl_rowcol_to_cell(r,5)}-{xl_rowcol_to_cell(r,6)}", calc_format)
        # Q3
        ws2.write(r, 8, 2, input_format)
        ws2.write(r, 9, 1, input_format)
        ws2.write_formula(r, 10, f"={xl_rowcol_to_cell(r,7)}+{xl_rowcol_to_cell(r,8)}-{xl_rowcol_to_cell(r,9)}", calc_format)
        # Q4
        ws2.write(r, 11, 2, input_format)
        ws2.write(r, 12, 1, input_format)
        ws2.write_formula(r, 13, f"={xl_rowcol_to_cell(r,10)}+{xl_rowcol_to_cell(r,11)}-{xl_rowcol_to_cell(r,12)}", calc_format)
        
        ws2.write(r, 14, 80000, input_format) # Avg Salary
        ws2.write(r, 15, 0.20, input_pct_format) # Benefits %
        
        # Total FTE Cost (FTE * Salary * (1+benefits)) / 4 quarters? Simplified: Average HC * Cost
        ws2.write_formula(r, 16, f"={xl_rowcol_to_cell(r,14)}*(1+{xl_rowcol_to_cell(r,15)})", calc_format)
        ws2.write_formula(r, 17, f"=AVERAGE({xl_rowcol_to_cell(r,1)},{xl_rowcol_to_cell(r,4)},{xl_rowcol_to_cell(r,7)},{xl_rowcol_to_cell(r,10)},{xl_rowcol_to_cell(r,13)})*{xl_rowcol_to_cell(r,16)}", calc_format)

    ws2.write(9, 0, "TOTAL", section_format)
    ws2.write_formula(9, 1, "=SUM(B5:B9)", calc_format)
    ws2.write_formula(9, 4, "=SUM(E5:E9)", calc_format)
    ws2.write_formula(9, 7, "=SUM(H5:H9)", calc_format)
    ws2.write_formula(9, 10, "=SUM(K5:K9)", calc_format)
    ws2.write_formula(9, 13, "=SUM(N5:N9)", calc_format)
    ws2.write_formula(9, 17, "=SUM(R5:R9)", calc_format)
    
    # FTE Summary
    for r, label in enumerate(["Beginning HC", "Total Hires", "Total Departures", "Ending HC", "Average FTE"], 13):
        ws2.write(r, 0, label)
        for c in range(1, 6): ws2.write(r, c, 50, input_format) # Keeping simple for parser
        
    # Cost Breakdown
    for r, dept in enumerate(depts, 21):
        ws2.write(r, 0, dept)
        for c in range(1, 5): ws2.write_formula(r, c, f"={xl_rowcol_to_cell(r-17, 17)}/4", calc_format)
        ws2.write_formula(r, 5, f"=SUM({xl_range(r, 1, r, 4)})", calc_format)
    ws2.write(26, 0, "TOTAL", section_format)
    for c in range(1, 6): ws2.write_formula(26, c, f"=SUM({xl_range(21, c, 25, c)})", calc_format)

    # ── 3. Revenue Forecast ────────────────────────────────────────────────────
    ws3 = workbook.add_worksheet("3. Revenue Forecast")
    ws3.write(0, 0, "Company Name - FPA Template | FY 2024", title_format)
    ws3.set_column('A:A', 25)
    
    for r, a in enumerate(["Base MRR", "Avg Growth %", "Churn %", "Expansion %"], 5):
        ws3.write(r, 0, a)
        ws3.write(r, 1, 0.05 if "%" in a else 10000, input_pct_format if "%" in a else input_format)
        
    for c, m in enumerate(months, 1): ws3.write(11, c, m, header_format)
    
    # Month 1 specific
    ws3.write(12, 0, "Beginning MRR"); ws3.write_formula(12, 1, "=B6", calc_format)
    ws3.write(13, 0, "New MRR"); ws3.write_formula(13, 1, "=B13*B7", calc_format)
    ws3.write(14, 0, "Expansion MRR"); ws3.write_formula(14, 1, "=B13*B9", calc_format)
    ws3.write(15, 0, "Churned MRR"); ws3.write_formula(15, 1, "=B13*B8", calc_format)
    ws3.write(16, 0, "Net MRR Change"); ws3.write_formula(16, 1, "=B14+B15-B16", calc_format)
    ws3.write(17, 0, "Ending MRR"); ws3.write_formula(17, 1, "=B13+B17", calc_format)
    ws3.write(18, 0, "MoM Growth %"); ws3.write_formula(18, 1, "=IFERROR(B17/B13, 0)", calc_pct_format)
    
    # Months 2-12
    for c in range(2, 13):
        prev_col = xl_rowcol_to_cell(0, c-1)[0]
        cur_col = xl_rowcol_to_cell(0, c)[0]
        ws3.write_formula(12, c, f"={prev_col}18", calc_format) # Beg = Prev End
        ws3.write(13, c, 500, input_format) # New
        ws3.write(14, c, 100, input_format) # Expansion
        ws3.write(15, c, 50, input_format)  # Churn
        ws3.write_formula(16, c, f"={cur_col}14+{cur_col}15-{cur_col}16", calc_format)
        ws3.write_formula(17, c, f"={cur_col}13+{cur_col}17", calc_format)
        ws3.write_formula(18, c, f"=IFERROR({cur_col}17/{cur_col}13, 0)", calc_pct_format)

    for r in range(12, 19):
        ws3.write_formula(r, 13, f"=SUM({xl_range(r, 1, r, 12)})", calc_format)
        
    for r, stream in enumerate(["Subscription Revenue", "Professional Services", "Other Revenue"], 22):
        ws3.write(r, 0, stream)
        for c in range(1, 13): ws3.write(r, c, 15000, input_format)
        ws3.write_formula(r, 13, f"=SUM({xl_range(r, 1, r, 12)})", calc_format)
        
    ws3.write(25, 0, "TOTAL REVENUE", section_format)
    for c in range(1, 14): ws3.write_formula(25, c, f"=SUM({xl_range(22, c, 24, c)})", calc_format)

    # ── 4. Rolling Forecast ────────────────────────────────────────────────────
    ws4 = workbook.add_worksheet("4. Rolling Forecast")
    ws4.write(0, 0, "Company Name - FPA Template | FY 2024", title_format)
    ws4.set_column('A:A', 25)
    
    for c, m in enumerate(months, 1): ws4.write(3, c, m, header_format)
    
    r = 4
    for sec, items in [("REVENUE", ["Subscription", "Services"]), ("OPERATING EXPENSES", ["Salaries", "Marketing"]), ("PROFITABILITY", ["Net Income"])]:
        ws4.write(r, 0, sec, section_format)
        r += 1
        for item in items:
            ws4.write(r, 0, item)
            for c in range(1, 13): ws4.write(r, c, 200, input_format)
            ws4.write_formula(r, 13, f"=SUM({xl_range(r, 1, r, 12)})", calc_format) # 12M Total
            ws4.write(r, 14, 2500, input_format) # Budget
            ws4.write_formula(r, 15, f"=IFERROR((N{r+1}-O{r+1})/O{r+1}, 0)", calc_pct_format)
            r += 1

    # ── 5. KPI Dashboard ───────────────────────────────────────────────────────
    ws5 = workbook.add_worksheet("5. KPI Dashboard")
    ws5.write(0, 0, "Company Name - FPA Template | FY 2024", title_format)
    ws5.set_column('A:A', 20)
    
    for r, k in enumerate(["CAC", "LTV", "Gross Margin", "Net Margin", "Burn Rate", "Runway (M)"], 5):
        ws5.write(r, 0, k)
        ws5.write(r, 1, 100, input_format)
        ws5.write(r, 2, 120, input_format)
        ws5.write(r, 3, "On Track", input_format)
        
        ws5.write(r, 6, months[r%12])
        ws5.write(r, 7, 50000, input_format)
        ws5.write(r, 8, 55000, input_format)
        if r > 5:
            ws5.write_formula(r, 9, f"=IFERROR((I{r+1}-I{r})/I{r}, 0)", calc_pct_format)
        else:
            ws5.write(r, 9, 0, calc_pct_format)
        
    for r, label in enumerate(["Cash on Hand", "Total Headcount", "ARR"], 19):
        ws5.write(r, 0, label)
        ws5.write(r, 1, 1000000, input_format)

    # ── 6. 13-Week Cash Flow ───────────────────────────────────────────────────
    ws6 = workbook.add_worksheet("6. 13-Week Cash Flow")
    ws6.write(0, 0, "Company Name - FPA Template | FY 2024", title_format)
    ws6.set_column('A:A', 30)
    
    for c, w in enumerate(weeks, 1): ws6.write(3, c, w, header_format)
    
    ws6.write(4, 0, "Opening Balance")
    ws6.write(4, 1, 500000, input_format) # Wk 1 opening
    for c in range(2, 14):
        prev_col = xl_rowcol_to_cell(0, c-1)[0]
        ws6.write_formula(4, c, f"={prev_col}20", calc_format) # Opening = Prev Ending
        
    ws6.write(7, 0, "Collections from Customers"); ws6.write_row(7, 1, [10000]*13, input_format)
    ws6.write(8, 0, "Other Income");               ws6.write_row(8, 1, [2000]*13, input_format)
    ws6.write(9, 0, "Total Inflows")
    for c in range(1, 14): ws6.write_formula(9, c, f"=SUM({xl_rowcol_to_cell(7,c)}:{xl_rowcol_to_cell(8,c)})", calc_format)
            
    outflows = ["Payroll (Bi-weekly)", "Rent & Facilities", "Marketing Spend", "Software & Tools", "Other Operating Expenses"]
    for r, label in enumerate(outflows, 12):
        ws6.write(r, 0, label)
        for c in range(1, 14): ws6.write(r, c, 3000, input_format)
            
    ws6.write(17, 0, "Total Outflows")
    for c in range(1, 14): ws6.write_formula(17, c, f"=SUM({xl_rowcol_to_cell(12,c)}:{xl_rowcol_to_cell(16,c)})", calc_format)
            
    ws6.write(18, 0, "Net Cash Flow")
    ws6.write(19, 0, "Ending Balance")
    for c in range(1, 14): 
        cur_col = xl_rowcol_to_cell(0, c)[0]
        ws6.write_formula(18, c, f"={cur_col}10-{cur_col}18", calc_format)
        ws6.write_formula(19, c, f"={cur_col}5+{cur_col}19", calc_format)

    # ── 7. Scenario Analysis ───────────────────────────────────────────────────
    ws7 = workbook.add_worksheet("7. Scenario Analysis")
    ws7.write(0, 0, "Company Name - FPA Template | FY 2024", title_format)
    ws7.set_column('A:A', 25)
    
    for r, s in enumerate(["Base Case", "Upside", "Downside"], 8):
        ws7.write(r, 0, s)
        ws7.write(r, 1, 0.10, input_pct_format)
        ws7.write(r, 2, 0.05, input_pct_format)
        ws7.write(r, 3, 0.02, input_pct_format)
        ws7.write(r, 4, f"{s} assumptions", input_format)
        
    for r, item in enumerate(["Revenue", "COGS", "Gross Profit", "R&D", "S&M", "G&A", "Operating Income"], 14):
        ws7.write(r, 0, item)
        ws7.write(r, 1, 100000, input_format) # Budget
        ws7.write_formula(r, 2, f"=B{r+1}*(1+$B$10)", calc_format) # Scenario result (using Upside row 9)
        ws7.write_formula(r, 3, f"=C{r+1}-B{r+1}", calc_format) # Var$
        ws7.write_formula(r, 4, f"=IFERROR(D{r+1}/B{r+1}, 0)", calc_pct_format) # Var%
        
    for c, s in enumerate(["Base Case", "Upside", "Downside"], 7): ws7.write(4, c, s, header_format)
    for r, m in enumerate(["Revenue", "EBITDA", "Cash Runway"], 5):
        ws7.write(r, 6, m)
        for c in range(7, 10): ws7.write(r, c, 100000, input_format)
        
    workbook.close()
    print("Template generated successfully!")

if __name__ == "__main__":
    create_template()
