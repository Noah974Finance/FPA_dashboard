"""
parser.py  –  Reads the exact 7-sheet FP&A template and returns clean dicts.

Sheet mapping:
  1. BvA Variance        →  parse_bva()
  2. Headcount Planning  →  parse_headcount()
  3. Revenue Forecast    →  parse_revenue_forecast()
  4. Rolling Forecast    →  parse_rolling_forecast()
  5. KPI Dashboard       →  parse_kpis()
  6. 13-Week Cash Flow   →  parse_cashflow()
  7. Scenario Analysis   →  parse_scenarios()
"""

from __future__ import annotations
import io
import re
import numpy as np
import pandas as pd


MONTHS = ["Jan","Feb","Mar","Apr","May","Jun",
          "Jul","Aug","Sep","Oct","Nov","Dec"]

SHEET_NAMES = {
    "bva":        "1. BvA Variance",
    "headcount":  "2. Headcount Planning",
    "revenue":    "3. Revenue Forecast",
    "rolling":    "4. Rolling Forecast",
    "kpi":        "5. KPI Dashboard",
    "cashflow":   "6. 13-Week Cash Flow",
    "scenarios":  "7. Scenario Analysis",
}


def _num(v):
    """Safe numeric cast."""
    try:
        f = float(v)
        return f if not np.isnan(f) else 0.0
    except Exception:
        return 0.0


def load_workbook(file_bytes: bytes) -> dict:
    """
    Entry point: parse all 7 sheets and return a nested dict.
    Also extracts company_name and year from the title row.
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    available = set(xl.sheet_names)

    result: dict = {"company_name": "Company", "year": "FY 2024", "sheets": {}}

    # Extract company name from the first available sheet title row
    for key, sheet in SHEET_NAMES.items():
        if sheet in available:
            raw = xl.parse(sheet, header=None)
            title = str(raw.iloc[0, 0]) if not raw.empty else ""
            m = re.search(r"-\s*(.+?)\s*\|", title)
            if m:
                result["company_name"] = m.group(1).strip()
            m2 = re.search(r"\|\s*(FY\s*\d{4}|Q\d[- ]\w+\s*\d{4})", title)
            if m2:
                result["year"] = m2.group(1).strip()
            break

    parsers = {
        "bva":       (parse_bva,       SHEET_NAMES["bva"]),
        "headcount": (parse_headcount, SHEET_NAMES["headcount"]),
        "revenue":   (parse_revenue_forecast, SHEET_NAMES["revenue"]),
        "rolling":   (parse_rolling_forecast, SHEET_NAMES["rolling"]),
        "kpi":       (parse_kpis,      SHEET_NAMES["kpi"]),
        "cashflow":  (parse_cashflow,  SHEET_NAMES["cashflow"]),
        "scenarios": (parse_scenarios, SHEET_NAMES["scenarios"]),
    }

    for key, (fn, sheet) in parsers.items():
        if sheet in available:
            try:
                result["sheets"][key] = fn(xl.parse(sheet, header=None))
            except Exception as e:
                result["sheets"][key] = {"error": str(e)}
        else:
            result["sheets"][key] = {"error": f"Sheet '{sheet}' not found"}

    return result


# ─────────────────────────────────────────────────────────────────────────────
# 1. BvA Variance
# ─────────────────────────────────────────────────────────────────────────────

def parse_bva(raw: pd.DataFrame) -> dict:
    """
    Rows follow the pattern:
      <Label> – Budget  →  12 monthly values
      <Label> – Actual  →  12 monthly values
      <Label> Variance  →  12 monthly values

    Returns:
      revenue_lines, expense_lines, net_income_lines  – each list of dicts:
        {name, budget_monthly, actual_monthly, variance_monthly,
         ytd_budget, ytd_actual, ytd_var_abs, ytd_var_pct, status, row_type}
    """
    # Row 3 = header;  cols 1-12 = Jan-Dec, col 13-17 = YTD
    data_rows = raw.iloc[4:].reset_index(drop=True)

    def _extract_row(r):
        return [_num(r.iloc[c]) for c in range(1, 13)]

    lines = []
    i = 0
    while i < len(data_rows):
        row = data_rows.iloc[i]
        label = str(row.iloc[0]).strip()

        # Section headers
        if label in ("REVENUE", "EXPENSES", "NET INCOME", "nan", ""):
            i += 1
            continue

        # Detect Budget rows (Budget / Actual / Variance triplet)
        if "Budget" in label or "budget" in label:
            clean = re.sub(r"\s*[-–]\s*Budget$", "", label, flags=re.I).strip()
            budget_vals = _extract_row(row)
            ytd_bud = _num(row.iloc[13]) if len(row) > 13 else sum(budget_vals)
            ytd_act = _num(row.iloc[14]) if len(row) > 14 else 0
            ytd_var_abs = _num(row.iloc[15]) if len(row) > 15 else 0
            ytd_var_pct = _num(row.iloc[16]) if len(row) > 16 else 0
            status = str(row.iloc[17]).strip() if len(row) > 17 else ""
            row_type = str(row.iloc[18]).strip() if len(row) > 18 else ""

            actual_vals   = [0] * 12
            variance_vals = [0] * 12

            if i + 1 < len(data_rows):
                ar = data_rows.iloc[i + 1]
                if "Actual" in str(ar.iloc[0]):
                    actual_vals = _extract_row(ar)
                    i += 1
            if i + 1 < len(data_rows):
                vr = data_rows.iloc[i + 1]
                lbl2 = str(vr.iloc[0])
                if "Variance" in lbl2 or "variance" in lbl2:
                    variance_vals = _extract_row(vr)
                    i += 1

            lines.append({
                "name":             clean,
                "budget_monthly":   budget_vals,
                "actual_monthly":   actual_vals,
                "variance_monthly": variance_vals,
                "ytd_budget":       ytd_bud,
                "ytd_actual":       ytd_act,
                "ytd_var_abs":      ytd_var_abs,
                "ytd_var_pct":      ytd_var_pct,
                "status":           status,
                "row_type":         row_type,
            })

        # Total rows (no Budget label)
        elif label.startswith("TOTAL") or label.startswith("Total"):
            budget_vals = _extract_row(row)
            ytd_bud = _num(row.iloc[13]) if len(row) > 13 else sum(budget_vals)
            ytd_act = _num(row.iloc[14]) if len(row) > 14 else 0
            ytd_var_abs = _num(row.iloc[15]) if len(row) > 15 else 0
            ytd_var_pct = _num(row.iloc[16]) if len(row) > 16 else 0
            status = str(row.iloc[17]).strip() if len(row) > 17 else ""
            row_type = str(row.iloc[18]).strip() if len(row) > 18 else "Total"

            actual_vals = [0] * 12
            if i + 1 < len(data_rows):
                ar = data_rows.iloc[i + 1]
                if "Actual" in str(ar.iloc[0]) or "actual" in str(ar.iloc[0]).lower():
                    actual_vals = _extract_row(ar)
                    i += 1

            lines.append({
                "name":             label,
                "budget_monthly":   budget_vals,
                "actual_monthly":   actual_vals,
                "variance_monthly": [a - b for a, b in zip(actual_vals, budget_vals)],
                "ytd_budget":       ytd_bud,
                "ytd_actual":       ytd_act,
                "ytd_var_abs":      ytd_var_abs,
                "ytd_var_pct":      ytd_var_pct,
                "status":           status,
                "row_type":         row_type,
            })

        i += 1

    # Split by section
    rev_lines  = [l for l in lines if l["row_type"] in ("Budget", "Total")
                  and any(k in l["name"] for k in ("Revenue","Revenue","Service","Fee","TOTAL REV"))]
    exp_lines  = [l for l in lines if l["row_type"] in ("Budget", "Total")
                  and any(k in l["name"] for k in ("Salaries","Marketing","R&D","Operations","G&A","TOTAL EXP"))]
    ni_lines   = [l for l in lines if "Net Income" in l["name"] or "NET INCOME" in l["name"]]

    # Fallback: split by order (revenue first half, then expenses)
    if not rev_lines and not exp_lines:
        total_lines = [l for l in lines if l["row_type"] == "Total"]
        rev_lines   = total_lines[:1]
        exp_lines   = total_lines[1:2]
        ni_lines    = total_lines[2:3]
        detail_lines = [l for l in lines if l["row_type"] == "Budget"]
        rev_lines    = detail_lines[:3] + rev_lines
        exp_lines    = detail_lines[3:] + exp_lines

    return {
        "revenue_lines":   rev_lines  or lines[:6],
        "expense_lines":   exp_lines  or lines[6:12],
        "net_income_lines":ni_lines   or lines[12:],
        "all_lines":       lines,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 2. Headcount Planning
# ─────────────────────────────────────────────────────────────────────────────

def parse_headcount(raw: pd.DataFrame) -> dict:
    """
    Dept table rows 4-9, summary rows 12-17, cost breakdown rows 20-26.
    """
    # Department table: row index 4..9 (0-based after iloc)
    dept_rows = raw.iloc[4:10].reset_index(drop=True)
    departments = []
    for _, r in dept_rows.iterrows():
        name = str(r.iloc[0]).strip()
        if name in ("nan", "TOTAL", "Total"):
            continue
        departments.append({
            "name":          name,
            "start_hc":      _num(r.iloc[1]),
            "q1_hires":      _num(r.iloc[2]),
            "q1_departs":    _num(r.iloc[3]),
            "q1_end":        _num(r.iloc[4]),
            "q2_hires":      _num(r.iloc[5]),
            "q2_departs":    _num(r.iloc[6]),
            "q2_end":        _num(r.iloc[7]),
            "q3_hires":      _num(r.iloc[8]),
            "q3_departs":    _num(r.iloc[9]),
            "q3_end":        _num(r.iloc[10]),
            "q4_hires":      _num(r.iloc[11]),
            "q4_departs":    _num(r.iloc[12]),
            "q4_end":        _num(r.iloc[13]),
            "avg_salary":    _num(r.iloc[14]),
            "benefits_pct":  _num(r.iloc[15]) / _num(r.iloc[14]) if _num(r.iloc[14]) else 0.25,
            "total_cost_fte":_num(r.iloc[15] if len(r) > 15 else 0) + _num(r.iloc[14]),
            "annual_cost":   _num(r.iloc[17]) if len(r) > 17 else 0,
        })

    # Totals row (row index 9)
    tot = raw.iloc[9]
    totals = {
        "start_hc":   _num(tot.iloc[1]),
        "q1_end":     _num(tot.iloc[4]),
        "q2_end":     _num(tot.iloc[7]),
        "q3_end":     _num(tot.iloc[10]),
        "q4_end":     _num(tot.iloc[13]),
        "annual_cost":_num(tot.iloc[17]) if len(tot) > 17 else 0,
    }

    # FTE summary table rows 12-17
    fte_labels = ["Beginning HC","Total Hires","Total Departures","Ending HC","Average FTE"]
    fte_data   = {}
    for idx, lbl in zip(range(13, 18), fte_labels):
        r = raw.iloc[idx]
        fte_data[lbl] = {
            "Q1": _num(r.iloc[1]),
            "Q2": _num(r.iloc[2]),
            "Q3": _num(r.iloc[3]),
            "Q4": _num(r.iloc[4]),
            "Full Year": _num(r.iloc[5]),
        }

    # Cost breakdown rows 21-26
    cost_labels = ["Sales","Marketing","Product/Engineering","Customer Success","G&A","TOTAL"]
    cost_data   = []
    for idx, lbl in zip(range(21, 27), cost_labels):
        r = raw.iloc[idx]
        cost_data.append({
            "dept":       lbl,
            "q1_cost":    _num(r.iloc[1]),
            "q2_cost":    _num(r.iloc[2]),
            "q3_cost":    _num(r.iloc[3]),
            "q4_cost":    _num(r.iloc[4]),
            "annual_cost":_num(r.iloc[5]),
        })

    return {
        "departments": departments,
        "totals":      totals,
        "fte_summary": fte_data,
        "cost_breakdown": cost_data,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 3. Revenue Forecast
# ─────────────────────────────────────────────────────────────────────────────

def parse_revenue_forecast(raw: pd.DataFrame) -> dict:
    """
    Assumptions rows 4-8, MRR waterfall rows 11-18, Revenue summary rows 21-26.
    """
    # Assumptions
    assumptions = {}
    for idx in range(5, 9):
        r = raw.iloc[idx]
        key   = str(r.iloc[0]).strip()
        value = _num(r.iloc[1])
        assumptions[key] = value

    # MRR waterfall (row 11 = header, 12-18 = data)
    mrr_labels = [
        "Beginning MRR", "New MRR", "Expansion MRR",
        "Churned MRR", "Net MRR Change", "Ending MRR", "MoM Growth %"
    ]
    mrr_data = {}
    for idx, lbl in zip(range(12, 19), mrr_labels):
        r = raw.iloc[idx]
        mrr_data[lbl] = [_num(r.iloc[c]) for c in range(1, 13)]

    # Annual totals from col 13
    mrr_annual = {}
    for idx, lbl in zip(range(12, 19), mrr_labels):
        r = raw.iloc[idx]
        mrr_annual[lbl] = _num(r.iloc[13]) if len(r) > 13 else sum(mrr_data[lbl])

    # ARR (col 14 of Ending MRR row)
    arr_row = raw.iloc[17]
    arr = _num(arr_row.iloc[14]) if len(arr_row) > 14 else 0

    # Revenue summary (rows 22-26)
    rev_streams = []
    for idx in range(22, 26):
        r = raw.iloc[idx]
        name = str(r.iloc[0]).strip()
        if name in ("nan", ""):
            continue
        monthly = [_num(r.iloc[c]) for c in range(1, 13)]
        rev_streams.append({
            "name":         name,
            "monthly":      monthly,
            "annual_total": _num(r.iloc[13]) if len(r) > 13 else sum(monthly),
            "pct_of_rev":   _num(r.iloc[14]) if len(r) > 14 else 0,
        })

    return {
        "assumptions":  assumptions,
        "mrr_waterfall":mrr_data,
        "mrr_annual":   mrr_annual,
        "arr":          arr,
        "revenue_streams": rev_streams,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 4. Rolling Forecast
# ─────────────────────────────────────────────────────────────────────────────

def parse_rolling_forecast(raw: pd.DataFrame) -> dict:
    """
    Header row 3, data rows 4-22. Columns 1-12 = 12 months, 13 = 12M Total,
    14 = Budget, 15 = Var%.
    """
    header_row = raw.iloc[3]
    # Extract month labels from cols 1-12
    month_labels = []
    for c in range(1, 13):
        val = header_row.iloc[c]
        try:
            ts = pd.to_datetime(val)
            month_labels.append(ts.strftime("%b %Y"))
        except Exception:
            month_labels.append(str(val)[:7])

    data_rows = raw.iloc[4:].reset_index(drop=True)

    def _parse_block(label_filter):
        """Return dict of {row_name: monthly_values} for rows matching section."""
        rows = {}
        in_section = False
        for _, r in data_rows.iterrows():
            lbl = str(r.iloc[0]).strip()
            if lbl == label_filter:
                in_section = True
                continue
            if in_section:
                if lbl in ("nan","") or lbl in ("REVENUE","OPERATING EXPENSES",
                                                  "PROFITABILITY","OPERATING EXPENSES"):
                    if lbl not in ("nan",""):
                        break
                    continue
                vals = [_num(r.iloc[c]) for c in range(1, 13)]
                total_12m = _num(r.iloc[13]) if len(r) > 13 else sum(vals)
                budget    = _num(r.iloc[14]) if len(r) > 14 else 0
                var_pct   = _num(r.iloc[15]) if len(r) > 15 else 0
                rows[lbl] = {
                    "monthly":   vals,
                    "total_12m": total_12m,
                    "budget":    budget,
                    "var_pct":   var_pct,
                }
        return rows

    revenue  = _parse_block("REVENUE")
    expenses = _parse_block("OPERATING EXPENSES")
    profit   = _parse_block("PROFITABILITY")

    return {
        "month_labels": month_labels,
        "revenue":      revenue,
        "expenses":     expenses,
        "profitability":profit,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 5. KPI Dashboard
# ─────────────────────────────────────────────────────────────────────────────

def parse_kpis(raw: pd.DataFrame) -> dict:
    """
    KPI table: rows 4-16 (cols 0-4: Metric, Current, Target, Status, Trend).
    Monthly revenue trend: rows 4-16 (cols 6-9: Month, MRR, Total Rev, Growth).
    """
    kpis = []
    for idx in range(5, 17):
        if idx >= len(raw):
            break
        r = raw.iloc[idx]
        metric  = str(r.iloc[0]).strip()
        current = r.iloc[1]
        target  = r.iloc[2]
        status  = str(r.iloc[3]).strip()
        trend   = str(r.iloc[4]).strip()
        if metric in ("nan", ""):
            continue
        kpis.append({
            "metric":  metric,
            "current": _num(current),
            "target":  _num(target),
            "status":  status,
            "trend":   trend,
        })

    # Monthly revenue trend (cols 6-9)
    monthly_trend = []
    for idx in range(5, 17):
        if idx >= len(raw):
            break
        r = raw.iloc[idx]
        month   = str(r.iloc[6]).strip() if len(r) > 6 else ""
        mrr     = _num(r.iloc[7]) if len(r) > 7 else 0
        total_r = _num(r.iloc[8]) if len(r) > 8 else 0
        growth  = _num(r.iloc[9]) if len(r) > 9 else 0
        if month and month != "nan":
            monthly_trend.append({
                "month": month, "mrr": mrr,
                "total_revenue": total_r, "growth": growth
            })

    # Executive summary (rows 19-22)
    summary = {}
    for idx in range(19, 23):
        if idx >= len(raw):
            break
        r = raw.iloc[idx]
        lbl = str(r.iloc[0]).strip()
        val = r.iloc[1]
        if lbl not in ("nan",""):
            summary[lbl] = val

    return {"kpis": kpis, "monthly_trend": monthly_trend, "summary": summary}


# ─────────────────────────────────────────────────────────────────────────────
# 6. 13-Week Cash Flow
# ─────────────────────────────────────────────────────────────────────────────

def parse_cashflow(raw: pd.DataFrame) -> dict:
    """
    Header: row 3 (cols 1-13 = Wk1..Wk13).
    Opening balance: row 4.
    Inflows: rows 7-9.
    Outflows: rows 12-17.
    Net cash flow: row 18.
    Ending balance: row 19.
    """
    weeks = [f"Wk {i}" for i in range(1, 14)]

    def _vals(row_idx):
        r = raw.iloc[row_idx]
        return [_num(r.iloc[c]) for c in range(1, 14)]

    opening       = _vals(4)
    collections   = _vals(7)
    other_income  = _vals(8)
    total_inflows = _vals(9)
    payroll       = _vals(12)
    rent          = _vals(13)
    marketing     = _vals(14)
    software      = _vals(15)
    other_opex    = _vals(16)
    total_outflows= _vals(17)
    net_cash      = _vals(18)
    ending        = _vals(19)

    return {
        "weeks":         weeks,
        "opening":       opening,
        "inflows": {
            "Collections from Customers": collections,
            "Other Income":               other_income,
            "Total Inflows":              total_inflows,
        },
        "outflows": {
            "Payroll (Bi-weekly)":        payroll,
            "Rent & Facilities":          rent,
            "Marketing Spend":            marketing,
            "Software & Tools":           software,
            "Other Operating Expenses":   other_opex,
            "Total Outflows":             total_outflows,
        },
        "net_cash_flow":   net_cash,
        "ending_balance":  ending,
        "total_13w":       sum(net_cash),
        "min_balance":     min(ending),
        "avg_inflow":      sum(total_inflows) / 13,
        "avg_outflow":     sum(total_outflows) / 13,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 7. Scenario Analysis
# ─────────────────────────────────────────────────────────────────────────────

def parse_scenarios(raw: pd.DataFrame) -> dict:
    """
    Scenario assumptions: rows 7-11 (Scenario, Rev Growth, OpEx Change, Churn, Description).
    Dynamic IS: rows 14-36 (Line Item, Base Budget, Scenario Result, Var$, Var%).
    Comparison summary: rows 4-12 (right side, cols 6-10).
    """
    # Scenario assumptions
    scen_rows = []
    for idx in range(8, 12):
        if idx >= len(raw):
            break
        r = raw.iloc[idx]
        name  = str(r.iloc[0]).strip()
        if name in ("nan",""):
            continue
        scen_rows.append({
            "scenario":      name,
            "rev_growth":    _num(r.iloc[1]),
            "opex_change":   _num(r.iloc[2]),
            "churn_rate":    _num(r.iloc[3]),
            "description":   str(r.iloc[4]).strip(),
        })

    # Dynamic IS (rows 14-36)
    is_rows = []
    for idx in range(14, 37):
        if idx >= len(raw):
            break
        r = raw.iloc[idx]
        item   = str(r.iloc[0]).strip()
        budget = _num(r.iloc[1])
        result = _num(r.iloc[2])
        var_d  = _num(r.iloc[3])
        var_p  = _num(r.iloc[4])
        if item in ("nan",""):
            continue
        is_rows.append({
            "item": item, "budget": budget,
            "scenario_result": result,
            "variance_abs": var_d,
            "variance_pct": var_p,
        })

    # Comparison summary (cols 6-10, rows 4-12)
    comparison_metrics = []
    col_headers = []
    header_row = raw.iloc[4]
    for c in range(6, 11):
        v = str(header_row.iloc[c]).strip()
        if v not in ("nan",""):
            col_headers.append(v)

    for idx in range(5, 13):
        if idx >= len(raw):
            break
        r = raw.iloc[idx]
        metric = str(r.iloc[6]).strip()
        if metric in ("nan",""):
            continue
        values = [_num(r.iloc[c]) for c in range(7, 11)]
        comparison_metrics.append({
            "metric": metric,
            "values": dict(zip(col_headers[1:], values)),
        })

    return {
        "scenarios":   scen_rows,
        "income_statement": is_rows,
        "comparison":  comparison_metrics,
        "col_headers": col_headers,
    }
