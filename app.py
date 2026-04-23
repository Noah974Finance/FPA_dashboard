"""
app.py  –  Enterprise FP&A Dashboard
Based 100% on the 7-sheet Excel template structure.

Run:  streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from parser import load_workbook
import auth
import storage

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FP&A Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun",
          "Jul","Aug","Sep","Oct","Nov","Dec"]

# ── Colour palette ────────────────────────────────────────────────────────────
C = {
    "blue":     "#1f77b4",
    "green":    "#2ca02c",
    "red":      "#d62728",
    "yellow":   "#ff7f0e",
    "purple":   "#9467bd",
    "cyan":     "#17becf",
    "orange":   "#ff7f0e",
}

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

  html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
  .block-container {{ padding: 1.2rem 2rem 3rem; max-width: 1600px; }}

  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
    background: var(--secondary-background-color);
    padding: 6px 8px;
    border-radius: 8px;
    border: 1px solid var(--faded-text-10);
  }}
  .stTabs [data-baseweb="tab"] {{
    border-radius: 6px; height: 38px; padding: 0 16px;
    font-weight: 600; font-size: 13px;
  }}
  .stTabs [aria-selected="true"] {{
    background: var(--background-color) !important;
    border: 1px solid var(--faded-text-20) !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
  }}

  /* KPI cards */
  .kpi-card {{
    background: var(--secondary-background-color);
    border: 1px solid var(--faded-text-10);
    border-radius: 10px;
    padding: 16px 20px;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
  }}
  .kpi-label  {{ font-size: 12px; color: var(--text-color); opacity: 0.7; text-transform: uppercase;
                  letter-spacing: .05em; margin-bottom: 6px; font-weight: 600; }}
  .kpi-value  {{ font-size: 28px; font-weight: 700; color: var(--text-color); line-height: 1.1; }}
  .kpi-delta  {{ font-size: 13px; margin-top: 6px; font-weight: 600; }}
  .kpi-g {{ color: {C['green']}; }} 
  .kpi-r {{ color: {C['red']}; }}
  .kpi-y {{ color: {C['yellow']}; }} 
  .kpi-b {{ color: {C['blue']}; }}

  /* Section header */
  .sec-hdr {{
    font-size: 15px; font-weight: 700; color: var(--text-color);
    border-bottom: 2px solid var(--faded-text-20);
    padding-bottom: 6px; margin: 20px 0 12px;
  }}

  /* Company pill */
  .co-pill {{
    display: inline-block; background: var(--secondary-background-color);
    border: 1px solid var(--faded-text-20);
    border-radius: 16px; padding: 4px 12px; font-size: 12px;
    font-weight: 600;
  }}

  /* Status badges */
  .badge-g {{ color:{C['green']}; border: 1px solid {C['green']}66; background: {C['green']}15; border-radius:6px; padding:2px 8px; font-size:11px; font-weight:600; }}
  .badge-r {{ color:{C['red']}; border: 1px solid {C['red']}66; background: {C['red']}15; border-radius:6px; padding:2px 8px; font-size:11px; font-weight:600; }}
  .badge-y {{ color:{C['yellow']}; border: 1px solid {C['yellow']}66; background: {C['yellow']}15; border-radius:6px; padding:2px 8px; font-size:11px; font-weight:600; }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Layout helpers
# ─────────────────────────────────────────────────────────────────────────────


def trim_future_zeros(vals):
    res = list(vals)
    for i in range(len(res)-1, -1, -1):
        if res[i] == 0:
            res[i] = None
        else:
            break
    return res

def chart_layout(title="", h=400):
    return dict(
        title=dict(text=title, font=dict(size=14, family="Inter"), x=0.01),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter"), height=h,
        margin=dict(l=12, r=12, t=45, b=12),
        legend=dict(orientation="h", y=-0.22, xanchor="center", x=0.5, font=dict(size=11)),
        hovermode="x unified",
    )

def kpi_card(label, value, delta=None, cls="kpi-b"):
    delta_html = f'<div class="kpi-delta {cls}">{delta}</div>' if delta else ""
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>{delta_html}</div>',
        unsafe_allow_html=True,
    )

def sec(title):
    st.markdown(f'<div class="sec-hdr">{title}</div>', unsafe_allow_html=True)

def fmt_k(v, prefix="$"):
    if abs(v) >= 1_000_000:
        return f"{prefix}{v/1_000_000:,.2f}M"
    if abs(v) >= 1_000:
        return f"{prefix}{v/1_000:,.0f}K"
    return f"{prefix}{v:,.0f}"

def fmt_pct(v, show_sign=False):
    s = "+" if show_sign and v > 0 else ""
    return f"{s}{v*100:.1f}%"

def badge(txt, kind="g"):
    return f'<span class="badge-{kind}">{txt}</span>'

# ─────────────────────────────────────────────────────────────────────────────
# Session state & Authentication
# ─────────────────────────────────────────────────────────────────────────────
import auth

if not st.session_state.get("authenticated", False):
    auth.render_auth_ui()
    st.stop()

if "companies" not in st.session_state:
    st.session_state.companies = storage.load_user_financial_data(st.session_state.user_email)
if "active" not in st.session_state:
    keys = list(st.session_state.companies.keys())
    st.session_state.active = keys[0] if keys else None

# ─────────────────────────────────────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:16px 0 8px">
      <div style="font-size:20px;font-weight:800;color:var(--text-color);">📊 FP&A Dashboard</div>
      <div style="font-size:11px;color:var(--faded-text-60);margin-top:2px;">Enterprise Financial Analytics</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📥 Import Data")
    
    try:
        with open("FPA_Template.xlsx", "rb") as file:
            st.download_button(
                label="📄 Download Blank Template",
                data=file,
                file_name="FPA_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except FileNotFoundError:
        st.warning("Blank template not found.")
        
    st.markdown("---")

    # Upload
    st.markdown("#### 📂 Upload Company File")
    st.caption("Format: `FPA_[Company Name].xlsx`  \nUse the 7-sheet template.")
    uploaded = st.file_uploader(
        "Drop Excel file here",
        type=["xlsx"],
        label_visibility="collapsed",
        key="uploader",
    )

    if uploaded is not None:
        with st.spinner("Parsing…"):
            try:
                data = load_workbook(uploaded.read())
                name = data["company_name"]
                year = data["year"]
                key = f"{name} - {year}"
                st.session_state.companies[key] = data
                st.session_state.active = key
                storage.save_financial_data(st.session_state.user_email, data)
                st.success(f"✅ Loaded **{key}**")
            except Exception as e:
                st.error(f"❌ Parse error: {e}")

    # Company selector
    companies = list(st.session_state.companies.keys())
    if companies:
        st.markdown("---")
        st.markdown("#### 🏢 Active Company")
        sel = st.selectbox(
            "Select file",
            companies,
            index=companies.index(st.session_state.active)
            if st.session_state.active in companies else 0,
            label_visibility="collapsed",
        )
        st.session_state.active = sel
        
        st.session_state.compare_mode = st.toggle("🔄 Comparison Mode")
        if st.session_state.compare_mode:
            comp_list = [c for c in companies if c != st.session_state.active]
            if comp_list:
                st.session_state.secondary = st.selectbox("Compare with:", comp_list)
            else:
                st.warning("Upload a 2nd file to enable comparison.")
                st.session_state.secondary = None
        else:
            st.session_state.secondary = None

        # Show all loaded companies
        st.markdown("**Loaded companies:**")
        for c in companies:
            active_marker = "▶ " if c == st.session_state.active else "   "
            col1, col2 = st.columns([4, 1])
            col1.markdown(
                f"<div style='color:{'var(--primary-color)' if c==st.session_state.active else 'var(--faded-text-60)'}"
                f";font-size:12px;padding:2px 0'>{active_marker}{c}</div>",
                unsafe_allow_html=True,
            )
            if col2.button("✕", key=f"del_{c}", help="Remove"):
                del st.session_state.companies[c]
                if st.session_state.active == c:
                    remaining = [x for x in companies if x != c]
                    st.session_state.active = remaining[0] if remaining else None
                st.rerun()

    st.markdown("---")
    st.caption("Upload multiple files to compare companies side-by-side.")

    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.force_logout = True
        st.session_state.companies = {}
        st.session_state.active = None
        st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# No data: welcome screen
# ─────────────────────────────────────────────────────────────────────────────
if not st.session_state.companies or st.session_state.active is None:
    st.markdown("""
    <div style="text-align:center;padding:80px 0 40px">
      <div style="font-size:56px">📊</div>
      <h1 style="font-size:2.4rem;font-weight:800;color:var(--text-color);margin:16px 0 8px">
        Enterprise FP&A Dashboard
      </h1>
      <p style="font-size:1.05rem;color:var(--faded-text-60);max-width:540px;margin:0 auto 32px">
        Upload your Excel file (using the 7-sheet FP&A template) from the sidebar
        to instantly generate your interactive financial dashboard.
      </p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    features = [
        ("📈", "BvA Variance", "Monthly Budget vs Actual with variance waterfall"),
        ("👥", "Headcount", "Department staffing plans and personnel costs"),
        ("🔮", "Revenue Forecast", "MRR waterfall, ARR, and stream breakdown"),
        ("📉", "Rolling Forecast", "12-month rolling P&L with budget comparison"),
        ("🎯", "KPI Dashboard", "Key SaaS & financial metrics with targets"),
        ("💰", "13-Week Cash Flow", "Weekly cash management and runway"),
        ("🎲", "Scenario Analysis", "Base / Optimistic / Pessimistic / Crisis"),
        ("🏢", "Multi-Company", "Upload multiple files and compare"),
    ]
    for i, (icon, title, desc) in enumerate(features):
        col = [c1, c2, c3, c4][i % 4]
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="text-align:left;margin-bottom:12px;">
              <div style="font-size:26px;margin-bottom:8px">{icon}</div>
              <div style="font-weight:700;font-size:13px;color:var(--text-color);margin-bottom:4px">{title}</div>
              <div style="font-size:11px;color:var(--faded-text-60)">{desc}</div>
            </div>""", unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# Load active company
# ─────────────────────────────────────────────────────────────────────────────
D    = st.session_state.companies[st.session_state.active]
S    = D["sheets"]
NAME = D["company_name"]
YEAR = D["year"]

# ── Page header ───────────────────────────────────────────────────────────────
col_h1, col_h2 = st.columns([3, 1])
col_h1.markdown(
    f"<h1 style='font-size:1.6rem;font-weight:800;color:var(--text-color);margin:0'>"
    f"📊 {NAME}</h1>",
    unsafe_allow_html=True,
)
col_h2.markdown(
    f"<div style='text-align:right;padding-top:6px'>"
    f"<span class='co-pill'>{YEAR}</span></div>",
    unsafe_allow_html=True,
)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📉 BvA Variance",
    "👥 Headcount",
    "🔮 Revenue Forecast",
    "📅 Rolling Forecast",
    "🎯 KPIs",
    "💰 Cash Flow",
    "🎲 Scenarios",
    "🏢 Multi-Company",
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 – Budget vs Actual Variance
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    bva = S.get("bva", {})
    if "error" in bva:
        st.error(bva["error"]); st.stop()

    all_lines = bva.get("all_lines", [])

    # Identify totals
    rev_total = next((l for l in all_lines if "TOTAL REVENUE" in l["name"] or l["name"] == "TOTAL REVENUE"), None)
    exp_total = next((l for l in all_lines if "TOTAL EXPENSES" in l["name"]), None)
    ni_total  = next((l for l in all_lines if "Net Income" in l["name"]), None)

    def _safe(l, k): return l.get(k, 0) if l else 0

    # ── KPI row ───────────────────────────────────────────────────────────────
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        ytd_rev = _safe(rev_total, "ytd_actual")
        bud_rev = _safe(rev_total, "ytd_budget")
        var_pct = (ytd_rev - bud_rev) / bud_rev if bud_rev else 0
        kpi_card("YTD Revenue", fmt_k(ytd_rev),
                 f"{'+' if var_pct>=0 else ''}{var_pct:.1%} vs Budget",
                 "kpi-g" if var_pct >= 0 else "kpi-r")
    with k2:
        ytd_exp = _safe(exp_total, "ytd_actual")
        bud_exp = _safe(exp_total, "ytd_budget")
        var_e   = (ytd_exp - bud_exp) / bud_exp if bud_exp else 0
        kpi_card("YTD Expenses", fmt_k(ytd_exp),
                 f"{'+' if var_e>=0 else ''}{var_e:.1%} vs Budget",
                 "kpi-r" if var_e > 0.05 else "kpi-g")
    with k3:
        ytd_ni = _safe(ni_total, "ytd_actual")
        bud_ni = _safe(ni_total, "ytd_budget")
        var_n  = (ytd_ni - bud_ni) / bud_ni if bud_ni else 0
        kpi_card("YTD Net Income", fmt_k(ytd_ni),
                 f"{'+' if var_n>=0 else ''}{var_n:.1%} vs Budget",
                 "kpi-g" if var_n >= 0 else "kpi-r")
    with k4:
        margin = ytd_ni / ytd_rev if ytd_rev else 0
        kpi_card("Net Margin", f"{margin:.1%}", "YTD", "kpi-b")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Revenue waterfall chart ───────────────────────────────────────────────
    sec("Revenue: Budget vs Actual")
    col_a, col_b = st.columns([3, 2])

    with col_a:
        # Line chart: total budget vs actual by month
        if rev_total:
            fig = go.Figure()
            fig.add_trace(go.Scatter(line_shape='spline', 
                x=MONTHS, y=rev_total["budget_monthly"],
                name="Budget", mode="lines+markers",
                line=dict(color=C["purple"], width=2, dash="dash"),
                marker=dict(size=5),
            ))
            fig.add_trace(go.Scatter(line_shape='spline', 
                x=MONTHS, y=trim_future_zeros(rev_total["actual_monthly"]),
                name="Actual", mode="lines+markers",
                line=dict(color=C["blue"], width=2.5),
                marker=dict(size=5),
            ))
            # Fill positive variance green, negative red
            for i in range(len(MONTHS)):
                b = rev_total["budget_monthly"][i]
                a = rev_total["actual_monthly"][i]
                color_fill = C["green"] if a >= b else C["red"]
                
            comp_key = st.session_state.get("secondary")
            if st.session_state.get("compare_mode", False) and comp_key:
                sec_bva = st.session_state.companies[comp_key]["sheets"].get("bva", {})
                sec_rev_total = next((l for l in sec_bva.get("all_lines", []) if "TOTAL REVENUE" in l["name"] or l["name"] == "TOTAL REVENUE"), None)
                if sec_rev_total:
                    fig.add_trace(go.Scatter(line_shape='spline', 
                        x=MONTHS, y=trim_future_zeros(sec_rev_total["actual_monthly"]),
                        name=f"Actual ({comp_key.split(' - ')[1] if ' - ' in comp_key else comp_key})", mode="lines+markers",
                        line=dict(color=C["yellow"], width=2.5, dash="dot"),
                        marker=dict(size=5),
                    ))
                    
            fig.update_layout(**chart_layout("Monthly Revenue: Budget vs Actual", 360))
            fig.update_layout(yaxis_tickprefix="$", yaxis_tickformat=",.0f")
            st.plotly_chart(fig, use_container_width=True, theme="streamlit")

    with col_b:
        # Monthly variance bar
        if rev_total:
            vars_m = rev_total["variance_monthly"]
            colors = [C["green"] if v >= 0 else C["red"] for v in vars_m]
            fig2 = go.Figure(go.Bar(
                x=MONTHS, y=vars_m,
                marker_color=colors,
                text=[fmt_k(v) for v in vars_m], textposition="outside",
                hovertemplate="<b>%{x}</b><br>Variance: $%{y:,.0f}<extra></extra>",
            ))
            fig2.update_layout(**chart_layout("Monthly Revenue Variance ($)", 360))
            fig2.update_layout(yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                               showlegend=False)
            st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

    # ── EBITDA waterfall bridge ───────────────────────────────────────────────
    sec("P&L Variance Bridge")
    col_c, col_d = st.columns([3, 2])

    with col_c:
        bud_e  = _safe(rev_total, "ytd_budget")
        act_e  = _safe(rev_total, "ytd_actual")
        rev_var= act_e - bud_e
        exp_b  = _safe(exp_total, "ytd_budget")
        exp_a  = _safe(exp_total, "ytd_actual")
        exp_var= -(exp_a - exp_b)  # higher costs = negative to EBITDA

        ni_b   = _safe(ni_total, "ytd_budget")
        ni_a   = _safe(ni_total, "ytd_actual")

        bridge_labels  = ["Budget Net Income", "Revenue △", "Expense △", "Actual Net Income"]
        bridge_vals    = [ni_b, rev_var, exp_var, ni_a]
        bridge_measures= ["absolute", "relative", "relative", "total"]

        fig3 = go.Figure(go.Waterfall(
            measure=bridge_measures,
            x=bridge_labels, y=bridge_vals,
            connector={"line": {"color": "rgba(128,128,128,0.5)"}},
            decreasing={"marker": {"color": C["red"]}},
            increasing={"marker": {"color": C["green"]}},
            totals={"marker": {"color": C["blue"]}},
            text=[fmt_k(v) for v in bridge_vals],
            textposition="outside",
        ))
        fig3.update_layout(**chart_layout("Net Income Bridge: Budget → Actual", 360))
        fig3.update_layout(yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                           showlegend=False)
        st.plotly_chart(fig3, use_container_width=True, theme="streamlit")

    with col_d:
        sec("YTD Variance Summary")
        # Table of all lines with YTD variance
        table_rows = []
        for l in all_lines:
            if l["row_type"] not in ("Budget", "Total"):
                continue
            var_a = l["ytd_actual"] - l["ytd_budget"]
            var_p = var_a / l["ytd_budget"] if l["ytd_budget"] else 0
            status = l.get("status", "")
            color  = "#3FB950" if "Favorable" in status else ("#F85149" if "Unfavorable" in status else "var(--faded-text-60)")
            table_rows.append({
                "Line Item":    l["name"],
                "Budget":       fmt_k(l["ytd_budget"]),
                "Actual":       fmt_k(l["ytd_actual"]),
                "Var $":        fmt_k(var_a),
                "Var %":        f"{var_p:+.1%}",
            })
        if table_rows:
            df_tbl = pd.DataFrame(table_rows)
            st.dataframe(df_tbl, use_container_width=True, height=340,
                         hide_index=True)

    # ── Expense breakdown ─────────────────────────────────────────────────────
    sec("Expense Lines: Budget vs Actual")
    exp_lines = [l for l in all_lines
                 if l["row_type"] == "Budget" and
                 any(k in l["name"] for k in ("Salaries","Marketing","R&D","Operations","G&A","Operations"))]
    if not exp_lines:
        exp_lines = [l for l in all_lines
                     if l["row_type"] == "Budget" and
                     l["name"] not in ["TOTAL REVENUE","TOTAL EXPENSES"] and
                     "Revenue" not in l["name"]]

    if exp_lines:
        fig4 = go.Figure()
        names = [l["name"] for l in exp_lines]
        buds  = [l["ytd_budget"] for l in exp_lines]
        acts  = [l["ytd_actual"] for l in exp_lines]
        fig4.add_trace(go.Bar(name="Budget", x=names, y=buds,
                              marker_color=C["purple"], opacity=0.75))
        fig4.add_trace(go.Bar(name="Actual", x=names, y=acts,
                              marker_color=C["blue"], opacity=0.85))
        fig4.update_layout(**chart_layout("YTD Expense Lines: Budget vs Actual", 340))
        fig4.update_layout(barmode="group", yaxis_tickprefix="$", yaxis_tickformat=",.0f")
        st.plotly_chart(fig4, use_container_width=True, theme="streamlit")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 – Headcount Planning
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    hc = S.get("headcount", {})
    if "error" in hc:
        st.error(hc["error"]); st.stop()

    depts = hc.get("departments", [])
    tots  = hc.get("totals", {})
    fte   = hc.get("fte_summary", {})
    costs = hc.get("cost_breakdown", [])

    # KPI row
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1: kpi_card("Starting HC",  f"{int(tots.get('start_hc',0))}", "Beginning of year", "kpi-b")
    with k2: kpi_card("Q1 End HC",    f"{int(tots.get('q1_end',0))}", "+hires -departures", "kpi-b")
    with k3: kpi_card("Q2 End HC",    f"{int(tots.get('q2_end',0))}", "+hires -departures", "kpi-b")
    with k4: kpi_card("Q4 End HC",    f"{int(tots.get('q4_end',0))}", "Year-end headcount", "kpi-g")
    with k5: kpi_card("Annual HC Cost",fmt_k(tots.get('annual_cost',0)), "All departments", "kpi-y")

    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    with col_a:
        sec("Headcount by Quarter per Department")
        if depts:
            quarters = ["Q1 End", "Q2 End", "Q3 End", "Q4 End"]
            qkeys    = ["q1_end","q2_end","q3_end","q4_end"]
            colors_d = [C["blue"],C["green"],C["purple"],C["orange"],C["cyan"]]
            fig = go.Figure()
            for i, d in enumerate(depts):
                fig.add_trace(go.Bar(
                    name=d["name"],
                    x=quarters,
                    y=[d.get(k, 0) for k in qkeys],
                    marker_color=colors_d[i % len(colors_d)],
                ))
            fig.update_layout(**chart_layout("Headcount by Department & Quarter", 380))
            fig.update_layout(barmode="stack")
            st.plotly_chart(fig, use_container_width=True, theme="streamlit")

    with col_b:
        sec("Headcount Waterfall (FTE Summary)")
        if fte:
            quarters = ["Q1","Q2","Q3","Q4"]
            begin_hc = [fte.get("Beginning HC",{}).get(q,0) for q in quarters]
            hires    = [fte.get("Total Hires",{}).get(q,0) for q in quarters]
            departures = [-fte.get("Total Departures",{}).get(q,0) for q in quarters]

            fig2 = go.Figure()
            fig2.add_trace(go.Bar(name="Begin HC", x=quarters, y=begin_hc,
                                  marker_color=C["blue"]))
            fig2.add_trace(go.Bar(name="Hires", x=quarters, y=hires,
                                  marker_color=C["green"]))
            fig2.add_trace(go.Bar(name="Departures", x=quarters, y=departures,
                                  marker_color=C["red"]))
            fig2.update_layout(**chart_layout("HC Movement by Quarter", 380))
            fig2.update_layout(barmode="relative")
            st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

    # Personnel cost donut
    col_c, col_d = st.columns([1, 2])
    with col_c:
        sec("Annual Cost per Department")
        if depts:
            names  = [d["name"] for d in depts]
            annual = [d.get("annual_cost",0) for d in depts]
            fig3 = go.Figure(go.Pie(
                labels=names, values=annual, hole=0.55,
                marker_colors=[C["blue"],C["green"],C["purple"],C["orange"],C["cyan"],C["red"],C["yellow"]],
                textinfo="percent",
                showlegend=True
            ))
            fig3.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", font=dict(),
                height=320, margin=dict(l=10,r=10,t=40,b=10),
                showlegend=True,
                title=dict(text="Cost Split", font=dict(size=13,),x=0.5),
            )
            st.plotly_chart(fig3, use_container_width=True, theme="streamlit")

    with col_d:
        sec("Personnel Cost Breakdown by Quarter")
        if costs:
            dept_costs = [c for c in costs if c["dept"] != "TOTAL"]
            fig4 = go.Figure()
            qs = ["q1_cost","q2_cost","q3_cost","q4_cost"]
            for i, dc in enumerate(dept_costs):
                fig4.add_trace(go.Bar(
                    name=dc["dept"],
                    x=["Q1","Q2","Q3","Q4"],
                    y=[dc.get(q,0) for q in qs],
                    marker_color=[C["blue"],C["green"],C["purple"],C["orange"],C["cyan"]][i % 5],
                ))
            fig4.update_layout(**chart_layout("Personnel Cost by Dept & Quarter", 320))
            fig4.update_layout(barmode="stack",
                               yaxis_tickprefix="$", yaxis_tickformat=",.0f")
            st.plotly_chart(fig4, use_container_width=True, theme="streamlit")

    # Department table
    sec("Department Detail Table")
    if depts:
        rows = []
        for d in depts:
            rows.append({
                "Department":   d["name"],
                "Start HC":     int(d["start_hc"]),
                "Q1 End":       int(d["q1_end"]),
                "Q2 End":       int(d["q2_end"]),
                "Q3 End":       int(d["q3_end"]),
                "Q4 End":       int(d["q4_end"]),
                "Avg Salary":   fmt_k(d["avg_salary"], "$"),
                "Annual Cost":  fmt_k(d["annual_cost"], "$"),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 – Revenue Forecast
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    rf = S.get("revenue", {})
    if "error" in rf:
        st.error(rf["error"]); st.stop()

    mrr     = rf.get("mrr_waterfall", {})
    streams = rf.get("revenue_streams", [])
    assum   = rf.get("assumptions", {})
    arr     = rf.get("arr", 0)

    # KPIs
    end_mrr = mrr.get("Ending MRR", [0]*12)[-1]
    new_mrr_total = sum(mrr.get("New MRR", [0]*12))
    churn_total   = abs(sum(mrr.get("Churned MRR", [0]*12)))
    exp_total_mrr = sum(mrr.get("Expansion MRR", [0]*12))

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1: kpi_card("Ending MRR",      fmt_k(end_mrr), "Dec", "kpi-b")
    with k2: kpi_card("ARR",             fmt_k(arr), "Annualised", "kpi-g")
    with k3: kpi_card("New MRR (YTD)",   fmt_k(new_mrr_total), "Total new", "kpi-g")
    with k4: kpi_card("Churned MRR",     fmt_k(churn_total), "Total lost", "kpi-r")
    with k5: kpi_card("Expansion MRR",   fmt_k(exp_total_mrr), "Upsell", "kpi-y")

    st.markdown("<br>", unsafe_allow_html=True)

    # MRR waterfall line chart
    sec("MRR Waterfall Model")
    col_a, col_b = st.columns([3, 2])

    with col_a:
        fig = go.Figure()
        colors_mrr = {
            "Beginning MRR": C["blue"],
            "New MRR":       C["green"],
            "Expansion MRR": C["cyan"],
            "Churned MRR":   C["red"],
            "Ending MRR":    C["purple"],
        }
        for lbl, clr in colors_mrr.items():
            if lbl in mrr:
                vals = mrr[lbl]
                mode = "lines+markers" if lbl in ("Beginning MRR","Ending MRR") else "bar"
                if lbl in ("Beginning MRR","Ending MRR"):
                    fig.add_trace(go.Scatter(line_shape='spline', 
                        x=MONTHS, y=vals, name=lbl,
                        mode="lines+markers",
                        line=dict(color=clr, width=2.5 if lbl=="Ending MRR" else 1.5,
                                  dash="dash" if lbl=="Beginning MRR" else "solid"),
                        marker=dict(size=5),
                    ))
        fig.update_layout(**chart_layout("Monthly MRR: Beginning → Ending", 380))
        fig.update_layout(yaxis_tickprefix="$", yaxis_tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True, theme="streamlit")

    with col_b:
        # Monthly MRR components stacked bar
        fig2 = go.Figure()
        if "New MRR" in mrr:
            fig2.add_trace(go.Bar(name="New MRR", x=MONTHS, y=mrr["New MRR"],
                                  marker_color=C["green"]))
        if "Expansion MRR" in mrr:
            fig2.add_trace(go.Bar(name="Expansion MRR", x=MONTHS, y=mrr["Expansion MRR"],
                                  marker_color=C["cyan"]))
        if "Churned MRR" in mrr:
            fig2.add_trace(go.Bar(name="Churned MRR", x=MONTHS, y=mrr["Churned MRR"],
                                  marker_color=C["red"]))
        fig2.update_layout(**chart_layout("MRR Components by Month", 380))
        fig2.update_layout(barmode="relative",
                           yaxis_tickprefix="$", yaxis_tickformat=",.0f")
        st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

    # Revenue streams
    sec("Revenue by Stream")
    col_c, col_d = st.columns([2, 1])

    with col_c:
        fig3 = go.Figure()
        stream_colors = [C["blue"], C["green"], C["yellow"]]
        for i, s in enumerate(streams):
            if s["name"] in ("TOTAL REVENUE","MoM Revenue Growth"):
                continue
            fig3.add_trace(go.Bar(
                name=s["name"], x=MONTHS, y=s["monthly"],
                marker_color=stream_colors[i % len(stream_colors)],
            ))
        total_stream = next((s for s in streams if "TOTAL" in s["name"]), None)
        if total_stream:
            fig3.add_trace(go.Scatter(line_shape='spline', 
                x=MONTHS, y=trim_future_zeros(total_stream["monthly"]),
                name="Total Revenue", mode="lines+markers",
                line=dict(color=C["orange"], width=2.5, dash="dot"),
                yaxis="y2",
            ))
        fig3.update_layout(**chart_layout("Monthly Revenue by Stream", 380))
        fig3.update_layout(
            barmode="stack",
            yaxis_tickprefix="$", yaxis_tickformat=",.0f",
            yaxis2=dict(overlaying="y", side="right", showgrid=False,
                        tickprefix="$", tickformat=",.0f"),
        )
        st.plotly_chart(fig3, use_container_width=True, theme="streamlit")

    with col_d:
        sec("Stream Mix (Annual)")
        detail_streams = [s for s in streams
                          if "TOTAL" not in s["name"] and s["annual_total"] > 0]
        if detail_streams:
            fig4 = go.Figure(go.Pie(
                labels=[s["name"] for s in detail_streams],
                values=[s["annual_total"] for s in detail_streams],
                hole=0.55,
                marker_colors=[C["blue"],C["green"],C["yellow"]],
                textinfo="label+percent",
            ))
            fig4.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", font=dict(),
                height=320, margin=dict(l=10,r=10,t=40,b=10), showlegend=False,
                title=dict(text="Annual Revenue Mix", font=dict(size=13,),x=0.5),
            )
            st.plotly_chart(fig4, use_container_width=True, theme="streamlit")

        # Assumptions table
        sec("Key Assumptions")
        if assum:
            rows = [{"Metric": k, "Value": f"{v:.2%}" if v < 1 else fmt_k(v)}
                    for k, v in assum.items()]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Rolling Forecast
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    roll = S.get("rolling", {})
    if "error" in roll:
        st.error(roll["error"]); st.stop()

    month_lbls = roll.get("month_labels", MONTHS)
    rev_r  = roll.get("revenue", {})
    exp_r  = roll.get("expenses", {})
    prof_r = roll.get("profitability", {})

    # KPIs from totals
    def _roll_kpi(section, key, field="total_12m"):
        return section.get(key, {}).get(field, 0)

    total_rev   = _roll_kpi(rev_r,  "TOTAL REVENUE")
    total_opex  = _roll_kpi(exp_r,  "TOTAL OPEX")
    ebitda_12m  = _roll_kpi(prof_r, "EBITDA")
    ebitda_margin = ebitda_12m / total_rev if total_rev else 0
    bud_rev     = _roll_kpi(rev_r,  "TOTAL REVENUE", "budget")
    rev_vs_bud  = (total_rev - bud_rev) / bud_rev if bud_rev else 0

    k1, k2, k3, k4 = st.columns(4)
    with k1: kpi_card("12M Revenue",     fmt_k(total_rev), f"{rev_vs_bud:+.1%} vs Budget",
                       "kpi-g" if rev_vs_bud >= 0 else "kpi-r")
    with k2: kpi_card("12M Total OpEx",  fmt_k(total_opex), "Rolling 12M", "kpi-b")
    with k3: kpi_card("12M EBITDA",      fmt_k(ebitda_12m), "Rolling", "kpi-g")
    with k4: kpi_card("EBITDA Margin",   f"{ebitda_margin:.1%}", "12-month avg", "kpi-b")

    st.markdown("<br>", unsafe_allow_html=True)

    # Revenue line chart with budget reference
    sec("12-Month Rolling Revenue vs Budget")
    fig = go.Figure()

    total_rev_row = rev_r.get("TOTAL REVENUE", {})
    sub_rev_row   = rev_r.get("Subscription Revenue", {})
    ps_rev_row    = rev_r.get("Professional Services", {})

    if sub_rev_row.get("monthly"):
        fig.add_trace(go.Bar(name="Subscription", x=month_lbls,
                             y=sub_rev_row["monthly"], marker_color=C["blue"]))
    if ps_rev_row.get("monthly"):
        fig.add_trace(go.Bar(name="Prof. Services", x=month_lbls,
                             y=ps_rev_row["monthly"], marker_color=C["green"]))

    other_rev = {k: v for k, v in rev_r.items()
                 if k not in ("TOTAL REVENUE","Subscription Revenue","Professional Services")}
    for k, v in other_rev.items():
        if v.get("monthly"):
            fig.add_trace(go.Bar(name=k, x=month_lbls,
                                 y=v["monthly"], marker_color=C["yellow"]))

    if total_rev_row.get("budget"):
        budget_line = [total_rev_row["budget"]/12] * len(month_lbls)
        fig.add_trace(go.Scatter(line_shape='spline', x=month_lbls, y=budget_line, name="Annualised Budget",
                                 mode="lines", line=dict(color=C["red"], width=2, dash="dot")))

    if total_rev_row.get("monthly"):
        fig.add_trace(go.Scatter(line_shape='spline', x=month_lbls, y=total_rev_row["monthly"],
                                 name="Total Revenue", mode="lines+markers",
                                 line=dict(color=C["orange"], width=2.5),
                                 marker=dict(size=5), yaxis="y2"))

    fig.update_layout(**chart_layout("Rolling 12-Month Revenue", 400))
    fig.update_layout(barmode="stack",
                      yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                      yaxis2=dict(overlaying="y", side="right", showgrid=False,
                                  tickprefix="$", tickformat=",.0f"))
    st.plotly_chart(fig, use_container_width=True, theme="streamlit")

    # EBITDA and margin
    col_a, col_b = st.columns(2)
    with col_a:
        sec("EBITDA Evolution")
        ebitda_row = prof_r.get("EBITDA", {})
        if ebitda_row.get("monthly"):
            evals = ebitda_row["monthly"]
            colors_e = [C["green"] if v >= 0 else C["red"] for v in evals]
            fig2 = go.Figure(go.Bar(
                x=month_lbls, y=evals, marker_color=colors_e,
                text=[fmt_k(v) for v in evals], textposition="outside",
            ))
            bud_e = ebitda_row.get("budget", 0)
            if bud_e:
                fig2.add_hline(y=bud_e/12, line_dash="dot",
                               line_color=C["purple"],
                               annotation_text="Budget/mo",
                               annotation_position="top right")
            fig2.update_layout(**chart_layout("Monthly EBITDA", 340))
            fig2.update_layout(showlegend=False,
                               yaxis_tickprefix="$", yaxis_tickformat=",.0f")
            st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

    with col_b:
        sec("EBITDA Margin %")
        margin_row = prof_r.get("EBITDA Margin %", {})
        rev_monthly = total_rev_row.get("monthly", [])
        ebitda_monthly = ebitda_row.get("monthly", []) if "ebitda_row" in dir() else []

        if margin_row.get("monthly"):
            margins = margin_row["monthly"]
        elif rev_monthly and ebitda_monthly:
            margins = [e/r if r else 0 for e, r in zip(ebitda_monthly, rev_monthly)]
        else:
            margins = []

        if margins:
            fig3 = go.Figure()
            fig3.add_trace(go.Scatter(line_shape='spline', 
                x=month_lbls, y=margins, name="EBITDA Margin",
                mode="lines+markers", fill="tozeroy",
                fillcolor="rgba(88, 166, 255, 0.2)",
                line=dict(color=C["blue"], width=2.5),
                marker=dict(size=5),
                hovertemplate="<b>%{x}</b><br>Margin: %{y:.1%}<extra></extra>",
            ))
            fig3.update_layout(**chart_layout("EBITDA Margin %", 340))
            fig3.update_layout(yaxis_tickformat=".0%", showlegend=False)
            st.plotly_chart(fig3, use_container_width=True, theme="streamlit")

    # Expense breakdown
    sec("Operating Expense Breakdown")
    fig4 = go.Figure()
    exp_colors = [C["blue"],C["green"],C["yellow"],C["orange"],C["cyan"],C["purple"]]
    for i, (k, v) in enumerate(exp_r.items()):
        if k == "TOTAL OPEX" or not v.get("monthly"):
            continue
        fig4.add_trace(go.Bar(name=k, x=month_lbls, y=v["monthly"],
                              marker_color=exp_colors[i % len(exp_colors)]))
    fig4.update_layout(**chart_layout("Monthly OpEx by Category", 360))
    fig4.update_layout(barmode="stack",
                       yaxis_tickprefix="$", yaxis_tickformat=",.0f")
    st.plotly_chart(fig4, use_container_width=True, theme="streamlit")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5 – KPI Dashboard
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    kpi_data = S.get("kpi", {})
    if "error" in kpi_data:
        st.error(kpi_data["error"]); st.stop()

    kpis     = kpi_data.get("kpis", [])
    trend_d  = kpi_data.get("monthly_trend", [])
    summary  = kpi_data.get("summary", {})

    # KPI grid – 4 per row
    sec("Key Performance Indicators")
    for row_start in range(0, len(kpis), 4):
        row_kpis = kpis[row_start:row_start+4]
        cols = st.columns(len(row_kpis))
        for col, k in zip(cols, row_kpis):
            with col:
                cur = k["current"]
                tgt = k["target"]
                status = k["status"]
                trend  = k["trend"]

                # Format value
                if cur > 10000:
                    val_str = fmt_k(cur)
                elif cur < 10 and cur > 0:
                    val_str = f"{cur:.3f}" if cur < 0.1 else f"{cur:.2f}"
                else:
                    val_str = f"{cur:,.0f}"

                # Colour
                cls = "kpi-g" if "On Target" in status else (
                      "kpi-r" if "Action" in status else "kpi-y")

                # Badge
                badge_cls = "g" if "On Target" in status else (
                             "r" if "Action" in status else "y")
                badge_txt = "✓ On Target" if "On Target" in status else (
                             "⚠ Monitor"  if "Monitor"  in status else "✗ Action!")

                st.markdown(f"""
                <div class="kpi-card">
                  <div class="kpi-label">{k['metric']}</div>
                  <div class="kpi-value {cls}">{val_str}</div>
                  <div class="kpi-delta" style="margin-top:6px">
                    <span class="badge-{badge_cls}">{badge_txt}</span>
                    &nbsp;<span style="color:var(--faded-text-60);font-size:11px">{trend}</span>
                  </div>
                  <div style="font-size:10px;color:var(--faded-text-60);margin-top:4px">
                    Target: {tgt:{',.0f' if tgt > 100 else '.2f'}}
                  </div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

    # Monthly trend charts
    if trend_d:
        sec("Monthly Revenue Trend")
        col_a, col_b = st.columns(2)
        months_t = [t["month"] for t in trend_d]

        with col_a:
            mrrs  = [t["mrr"] for t in trend_d]
            revs  = [t["total_revenue"] for t in trend_d]
            fig = go.Figure()
            fig.add_trace(go.Bar(name="MRR", x=months_t, y=mrrs,
                                 marker_color=C["blue"], opacity=0.8))
            fig.add_trace(go.Scatter(line_shape='spline', name="Total Revenue", x=months_t, y=revs,
                                     mode="lines+markers",
                                     line=dict(color=C["orange"], width=2.5),
                                     marker=dict(size=5), yaxis="y2"))
            fig.update_layout(**chart_layout("MRR vs Total Revenue", 360))
            fig.update_layout(yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                              yaxis2=dict(overlaying="y", side="right",
                                          showgrid=False, tickprefix="$",
                                          tickformat=",.0f"))
            st.plotly_chart(fig, use_container_width=True, theme="streamlit")

        with col_b:
            growths = [t["growth"] for t in trend_d]
            fig2 = go.Figure(go.Bar(
                x=months_t, y=growths,
                marker_color=[C["green"] if g >= 0 else C["red"] for g in growths],
                text=[f"{g:.1%}" if g else "—" for g in growths],
                textposition="outside",
                hovertemplate="<b>%{x}</b><br>Growth: %{y:.1%}<extra></extra>",
            ))
            fig2.update_layout(**chart_layout("Month-over-Month Revenue Growth", 360))
            fig2.update_layout(yaxis_tickformat=".1%", showlegend=False)
            st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

    # Executive summary
    if summary:
        sec("Executive Summary")
        for lbl, val in summary.items():
            if str(lbl) not in ("nan",""):
                if isinstance(val, (int, float)):
                    if abs(val) > 1000:
                        val_str = fmt_k(val)
                    else:
                        val_str = f"{val:,.0f}" if val > 10 else str(val)
                else:
                    val_str = str(val)
                st.markdown(f"- **{lbl}**: {val_str}")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 6 – 13-Week Cash Flow
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    cf = S.get("cashflow", {})
    if "error" in cf:
        st.error(cf["error"]); st.stop()

    weeks   = cf.get("weeks", [f"Wk {i}" for i in range(1,14)])
    ending  = cf.get("ending_balance", [0]*13)
    opening = cf.get("opening", [0]*13)
    net_cf  = cf.get("net_cash_flow", [0]*13)
    inflows = cf.get("inflows", {})
    outflows= cf.get("outflows", {})

    k1, k2, k3, k4 = st.columns(4)
    with k1: kpi_card("Opening Balance", fmt_k(opening[0] if opening else 0),
                       "Week 1", "kpi-b")
    with k2: kpi_card("Closing Balance", fmt_k(ending[-1] if ending else 0),
                       "Week 13", "kpi-g" if (ending[-1] if ending else 0) > 200000 else "kpi-r")
    with k3: kpi_card("13W Net Cash", fmt_k(cf.get("total_13w", 0)),
                       "Total", "kpi-g" if cf.get("total_13w",0) >= 0 else "kpi-r")
    with k4: kpi_card("Min Balance", fmt_k(cf.get("min_balance", 0)),
                       "Lowest week", "kpi-g" if cf.get("min_balance",0) > 200000 else "kpi-r")

    st.markdown("<br>", unsafe_allow_html=True)

    # Cash balance area chart
    sec("13-Week Cash Balance")
    MIN_THRESHOLD = 200_000
    fig = go.Figure()
    fig.add_trace(go.Scatter(line_shape='spline', 
        x=weeks, y=ending, name="Ending Cash Balance",
        mode="lines+markers", fill="tozeroy",
        fillcolor="rgba(88, 166, 255, 0.133)",
        line=dict(color=C["blue"], width=2.5),
        marker=dict(size=7,
                    color=[C["red"] if v < MIN_THRESHOLD else C["green"] for v in ending],
                    line=dict(width=2, color="rgba(0,0,0,0)")),
    ))
    fig.add_hline(y=MIN_THRESHOLD, line_dash="dot", line_color=C["red"],
                  annotation_text="Minimum Threshold ($200K)",
                  annotation_position="top right")
    fig.update_layout(**chart_layout("Weekly Cash Balance", 380))
    fig.update_layout(yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                      showlegend=False)
    st.plotly_chart(fig, use_container_width=True, theme="streamlit")

    # Weekly inflows vs outflows
    col_a, col_b = st.columns(2)
    with col_a:
        sec("Weekly Cash Inflows vs Outflows")
        tot_in  = inflows.get("Total Inflows", [0]*13)
        tot_out = outflows.get("Total Outflows", [0]*13)
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(name="Inflows", x=weeks, y=tot_in,
                              marker_color=C["green"], opacity=0.85))
        fig2.add_trace(go.Bar(name="Outflows", x=weeks,
                              y=[-v for v in tot_out],
                              marker_color=C["red"], opacity=0.85))
        fig2.add_trace(go.Scatter(line_shape='spline', name="Net Cash Flow", x=weeks, y=net_cf,
                                  mode="lines+markers",
                                  line=dict(color=C["yellow"], width=2),
                                  marker=dict(size=5)))
        fig2.update_layout(**chart_layout("Weekly Cash Flow Components", 360))
        fig2.update_layout(barmode="relative",
                           yaxis_tickprefix="$", yaxis_tickformat=",.0f")
        st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

    with col_b:
        sec("Outflow Breakdown by Category")
        outflow_items = {k: v for k, v in outflows.items() if k != "Total Outflows"}
        fig3 = go.Figure()
        out_colors = [C["blue"],C["green"],C["yellow"],C["orange"],C["cyan"]]
        for i, (cat, vals) in enumerate(outflow_items.items()):
            fig3.add_trace(go.Bar(name=cat, x=weeks, y=vals,
                                  marker_color=out_colors[i % len(out_colors)]))
        fig3.update_layout(**chart_layout("Weekly Outflow by Category", 360))
        fig3.update_layout(barmode="stack",
                           yaxis_tickprefix="$", yaxis_tickformat=",.0f")
        st.plotly_chart(fig3, use_container_width=True, theme="streamlit")

    # Cash table
    sec("13-Week Cash Flow Detail")
    table = {"Category": ["Opening Balance", "Total Inflows", "Total Outflows",
                           "Net Cash Flow", "Ending Balance"]}
    for i, wk in enumerate(weeks):
        table[wk] = [
            fmt_k(opening[i]) if i < len(opening) else "—",
            fmt_k(tot_in[i])  if i < len(tot_in)  else "—",
            fmt_k(tot_out[i]) if i < len(tot_out) else "—",
            fmt_k(net_cf[i])  if i < len(net_cf)  else "—",
            fmt_k(ending[i])  if i < len(ending)  else "—",
        ]
    st.dataframe(pd.DataFrame(table), use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 7 – Scenario Analysis
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    sc = S.get("scenarios", {})
    if "error" in sc:
        st.error(sc["error"]); st.stop()

    scenarios   = sc.get("scenarios", [])
    is_rows     = sc.get("income_statement", [])
    comparison  = sc.get("comparison", [])
    col_headers = sc.get("col_headers", [])

    k1, k2, k3, k4 = st.columns(4)
    sc_colors = [C["blue"], C["green"], C["yellow"], C["red"]]
    sc_labels = ["Base Case", "Optimistic", "Pessimistic", "Crisis"]

    # Revenue from comparison
    rev_comp = next((m for m in comparison if "Revenue" in m["metric"]), None)
    ebitda_comp = next((m for m in comparison if "EBITDA" == m["metric"]), None)
    if rev_comp:
        vals = list(rev_comp["values"].values())
        k1.metric("Base Revenue",    fmt_k(vals[0]) if len(vals) > 0 else "—")
        k2.metric("Optimistic Rev",  fmt_k(vals[1]) if len(vals) > 1 else "—",
                  delta=fmt_k(vals[1]-vals[0]) if len(vals) > 1 else "")
        k3.metric("Pessimistic Rev", fmt_k(vals[2]) if len(vals) > 2 else "—",
                  delta=fmt_k(vals[2]-vals[0]) if len(vals) > 2 else "")
        k4.metric("Crisis Rev",      fmt_k(vals[3]) if len(vals) > 3 else "—",
                  delta=fmt_k(vals[3]-vals[0]) if len(vals) > 3 else "")

    st.markdown("<br>", unsafe_allow_html=True)

    # Scenario comparison bar chart
    sec("Scenario Comparison – Key Metrics")
    col_a, col_b = st.columns(2)

    with col_a:
        if comparison and col_headers:
            metrics_to_show = ["Revenue","Gross Profit","EBITDA","Net Income"]
            filtered = [m for m in comparison
                        if m["metric"] in metrics_to_show]
            scen_names = [h for h in col_headers if h != "Metric"]

            fig = go.Figure()
            for i, scen in enumerate(scen_names[:4]):
                fig.add_trace(go.Bar(
                    name=scen,
                    x=[m["metric"] for m in filtered],
                    y=[m["values"].get(scen, 0) for m in filtered],
                    marker_color=sc_colors[i],
                ))
            fig.update_layout(**chart_layout("Revenue / Profit by Scenario", 400))
            fig.update_layout(barmode="group",
                              yaxis_tickprefix="$", yaxis_tickformat=",.0f")
            st.plotly_chart(fig, use_container_width=True, theme="streamlit")

    with col_b:
        # EBITDA margin spider / radar chart
        margin_items = [m for m in comparison if "Margin" in m["metric"]]
        
        # if we have margin items but they're all 0, try computing fallback
        if margin_items and any(v != 0 for v in margin_items[0]["values"].values()):
            margin_data = margin_items[0]["values"]
        elif ebitda_comp and rev_comp:
            rev_vals    = list(rev_comp["values"].values())
            ebitda_vals = list(ebitda_comp["values"].values())
            margins = [e/r if r else 0 for e, r in zip(ebitda_vals, rev_vals)]
            # scen_names is defined in col_a and is leaked here
            margin_data = dict(zip(scen_names[:len(margins)], margins))
        else:
            margin_data = {}

        if margin_data:
            fig2 = go.Figure(go.Bar(
                x=list(margin_data.keys()),
                y=list(margin_data.values()),
                marker_color=sc_colors[:len(margin_data)],
                text=[f"{v:.1%}" for v in margin_data.values()],
                textposition="outside",
            ))
            fig2.update_layout(**chart_layout("EBITDA Margin by Scenario", 400))
            fig2.update_layout(yaxis_tickformat=".1%", showlegend=False)
            st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

    # Dynamic Income Statement
    sec("Dynamic Income Statement (What-If Interactive)")
    
    st.markdown("### 🎛️ Interactive Simulator")
    col_s1, col_s2, col_s3 = st.columns(3)
    sim_rev = col_s1.slider("Revenue Growth (%)", -50, 100, 10, 1) / 100.0
    sim_opex = col_s2.slider("Expense Growth (%)", -50, 100, 5, 1) / 100.0
    sim_churn = col_s3.slider("Churn Rate (%)", 0, 50, 2, 1) / 100.0
    
    if is_rows:
        dynamic_is_rows = []
        for row in is_rows:
            item = row["item"].lower()
            budget = row["budget"]
            
            # Recalculate based on sliders
            if "revenue" in item or "profit" in item:
                new_res = budget * (1 + sim_rev)
            elif "cogs" in item or "r&d" in item or "s&m" in item or "g&a" in item or "expense" in item:
                new_res = budget * (1 + sim_opex)
            else:
                new_res = budget
                
            var_abs = new_res - budget
            var_pct = var_abs / budget if budget else 0
            
            dynamic_is_rows.append({
                "item": row["item"],
                "budget": budget,
                "scenario_result": new_res,
                "variance_abs": var_abs,
                "variance_pct": var_pct
            })
            
        col_c, col_d = st.columns([2, 1])
        with col_c:
            is_table = []
            for row in dynamic_is_rows:
                is_table.append({
                    "Line Item":       row["item"],
                    "Base Budget":     fmt_k(row["budget"]),
                    "Simulated Result":fmt_k(row["scenario_result"]),
                    "Variance $":      fmt_k(row["variance_abs"]),
                    "Variance %":      f"{row['variance_pct']:+.1%}" if row["variance_pct"] else "—",
                })
            df_is = pd.DataFrame(is_table)
            st.dataframe(df_is, use_container_width=True,
                         height=500, hide_index=True)

        with col_d:
            # Scenario assumptions table
            sec("Scenario Assumptions")
            if scenarios:
                rows = []
                for s in scenarios:
                    rows.append({
                        "Scenario":     s["scenario"],
                        "Rev Growth":   f"{s['rev_growth']:.0%}",
                        "OpEx Change":  f"{s['opex_change']:.0%}",
                        "Churn Rate":   f"{s['churn_rate']:.0%}",
                    })
                st.dataframe(pd.DataFrame(rows),
                             use_container_width=True, hide_index=True)

            # Scenario revenue waterfall
            rev_row = next((r for r in dynamic_is_rows if r["item"].upper() == "TOTAL REVENUE" or r["item"].upper() == "REVENUE"), None)
            if rev_row:
                bud_r    = rev_row["budget"]
                scen_res = rev_row["scenario_result"]
                fig3 = go.Figure(go.Waterfall(
                    measure=["absolute","relative","total"],
                    x=["Budget","Simulated △","Result"],
                    y=[bud_r, scen_res - bud_r, scen_res],
                    connector={"line": {"color": "rgba(128,128,128,0.5)"}},
                    decreasing={"marker":{"color":C["red"]}},
                    increasing={"marker":{"color":C["green"]}},
                    totals={"marker":{"color":C["blue"]}},
                    text=[fmt_k(v) for v in [bud_r, scen_res-bud_r, scen_res]],
                    textposition="outside",
                ))
                fig3.update_layout(**chart_layout("Simulated Revenue Bridge", 300))
                fig3.update_layout(showlegend=False,
                                   yaxis_tickprefix="$", yaxis_tickformat=",.0f")
                st.plotly_chart(fig3, use_container_width=True, theme="streamlit")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 8 – Multi-Company Comparison
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[7]:
    all_cos = st.session_state.companies
    if len(all_cos) < 2:
        st.info("Upload at least **2 company files** from the sidebar to enable comparison.")
        st.markdown("""
        Each company uses the same 7-sheet template.
        Give the file a unique company name: `FPA_CompanyName.xlsx`
        """)
    else:
        co_names = list(all_cos.keys())
        co_colors = [C["blue"],C["green"],C["yellow"],C["orange"],C["cyan"],C["purple"]]

        # Collect summary metrics per company
        def _co_metrics(name, data):
            bva  = data["sheets"].get("bva", {})
            kpid = data["sheets"].get("kpi", {})
            roll = data["sheets"].get("rolling", {})
            hc   = data["sheets"].get("headcount", {})

            all_lines = bva.get("all_lines", [])
            rev_t  = next((l for l in all_lines if "TOTAL REVENUE" in l["name"]), None)
            exp_t  = next((l for l in all_lines if "TOTAL EXPENSES" in l["name"]), None)
            ni_t   = next((l for l in all_lines if "Net Income" in l["name"]), None)

            ytd_rev = rev_t["ytd_actual"]  if rev_t else 0
            ytd_exp = exp_t["ytd_actual"]  if exp_t else 0
            ytd_ni  = ni_t["ytd_actual"]   if ni_t  else 0
            margin  = ytd_ni / ytd_rev     if ytd_rev else 0

            end_hc = hc.get("totals", {}).get("q4_end", 0)
            rev_per_hc = ytd_rev / end_hc if end_hc else 0

            # MRR from KPI
            kpis = kpid.get("kpis", [])
            mrr_kpi = next((k["current"] for k in kpis
                            if "MRR" in k["metric"] and "ARR" not in k["metric"]), 0)
            arr_kpi = next((k["current"] for k in kpis if "ARR" in k["metric"]), 0)

            return {
                "YTD Revenue":      ytd_rev,
                "YTD Expenses":     ytd_exp,
                "YTD Net Income":   ytd_ni,
                "Net Margin %":     margin,
                "Headcount (Q4)":   end_hc,
                "Rev / Headcount":  rev_per_hc,
                "MRR":              mrr_kpi,
                "ARR":              arr_kpi,
            }

        metrics_by_co = {n: _co_metrics(n, d) for n, d in all_cos.items()}
        metric_keys = list(list(metrics_by_co.values())[0].keys())

        sec("Multi-Company KPI Comparison")

        # Side-by-side KPI cards
        kpi_cols = st.columns(len(co_names))
        for col, name in zip(kpi_cols, co_names):
            m = metrics_by_co[name]
            col.markdown(f"""
            <div class="kpi-card" style="margin-bottom:8px">
              <div class="kpi-label">{name}</div>
              <div class="kpi-value" style="font-size:18px;color:var(--primary-color)">
                {fmt_k(m['YTD Revenue'])}
              </div>
              <div class="kpi-delta kpi-b">YTD Revenue</div>
            </div>""", unsafe_allow_html=True)
            col.metric("Net Margin",     f"{m['Net Margin %']:.1%}")
            col.metric("Headcount Q4",   f"{int(m['Headcount (Q4)'])}")
            col.metric("Rev/HC",         fmt_k(m["Rev / Headcount"]))

        st.markdown("<br>", unsafe_allow_html=True)

        # Grouped bar: Revenue comparison
        sec("Revenue & Profitability Comparison")
        col_a, col_b = st.columns(2)

        with col_a:
            fig = go.Figure()
            categories = ["YTD Revenue","YTD Net Income"]
            for i, name in enumerate(co_names):
                m = metrics_by_co[name]
                fig.add_trace(go.Bar(
                    name=name,
                    x=categories,
                    y=[m[c] for c in categories],
                    marker_color=co_colors[i % len(co_colors)],
                ))
            fig.update_layout(**chart_layout("Revenue & Net Income Comparison", 380))
            fig.update_layout(barmode="group",
                              yaxis_tickprefix="$", yaxis_tickformat=",.0f")
            st.plotly_chart(fig, use_container_width=True, theme="streamlit")

        with col_b:
            # Net margin bars
            fig2 = go.Figure(go.Bar(
                x=co_names,
                y=[metrics_by_co[n]["Net Margin %"] for n in co_names],
                marker_color=co_colors[:len(co_names)],
                text=[f"{metrics_by_co[n]['Net Margin %']:.1%}" for n in co_names],
                textposition="outside",
            ))
            fig2.update_layout(**chart_layout("Net Margin % by Company", 380))
            fig2.update_layout(yaxis_tickformat=".1%", showlegend=False)
            st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

        # Full comparison table
        sec("Full Metrics Table")
        rows = []
        for metric in metric_keys:
            row = {"Metric": metric}
            for name in co_names:
                val = metrics_by_co[name].get(metric, 0)
                if "%" in metric:
                    row[name] = f"{val:.1%}"
                elif isinstance(val, float) and val > 100:
                    row[name] = fmt_k(val)
                else:
                    row[name] = f"{val:,.0f}"
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="text-align:center;padding:32px 0 8px;
            color:var(--text-color);opacity:0.6;font-size:11px;border-top:1px solid var(--faded-text-20);
            margin-top:32px">
  📊 FP&A Dashboard · Enterprise Financial Analytics ·
  Active: <strong style="color:{C['blue']}">{NAME}</strong> ·
  {YEAR}
</div>
""", unsafe_allow_html=True)
