# 📊 FP&A Dashboard

Enterprise Financial Planning & Analysis dashboard built around your 7-sheet Excel template.

---

## 🚀 Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Launch the dashboard
streamlit run app.py
```

Open your browser at **http://localhost:8501**

---

## 📂 How It Works

1. **Upload your Excel file** in the sidebar (the 7-sheet template)
2. The app reads every sheet and renders interactive charts instantly
3. Modify any numbers in Excel → re-upload → dashboard updates automatically
4. Upload **multiple company files** to compare side-by-side

### File naming (optional)
`FPA_CompanyName.xlsx` → the app auto-detects "CompanyName" from the filename

---

## 📋 Excel Template – 7 Sheets

| Sheet | Tab in Dashboard | What it shows |
|-------|-----------------|---------------|
| `1. BvA Variance` | 📉 BvA Variance | Revenue & expense budget vs actual, monthly waterfall |
| `2. Headcount Planning` | 👥 Headcount | Department HC by quarter, hiring plans, personnel costs |
| `3. Revenue Forecast` | 🔮 Revenue Forecast | MRR waterfall, ARR, revenue by stream |
| `4. Rolling Forecast` | 📅 Rolling Forecast | 12-month rolling P&L vs budget |
| `5. KPI Dashboard` | 🎯 KPIs | 12 key SaaS metrics with targets and trend |
| `6. 13-Week Cash Flow` | 💰 Cash Flow | Weekly cash management with $200K threshold |
| `7. Scenario Analysis` | 🎲 Scenarios | Base / Optimistic / Pessimistic / Crisis income statement |

---

## ✏️ Modifying Your Excel

The dashboard adapts to **any numbers** you change in the template:

- Change budget figures in Sheet 1 → variance charts update
- Add/remove departments in Sheet 2 → headcount charts update
- Modify churn rate or MRR assumptions in Sheet 3 → MRR waterfall updates
- Update rolling forecast values in Sheet 4 → 12M charts update
- Change KPI targets in Sheet 5 → badge colours update
- Edit weekly cash flows in Sheet 6 → balance chart updates
- Switch scenario assumptions in Sheet 7 → IS comparison updates

---

## 🏢 Multi-Company Comparison

Upload 2+ files in the sidebar. The **Multi-Company** tab shows:
- Side-by-side KPI cards
- Revenue & net income grouped bar
- Net margin comparison
- Full metrics table

---

## 🛠 Project Structure

```
fpa_dashboard/
├── app.py           # Streamlit dashboard (all 8 tabs)
├── parser.py        # Excel reader (maps every sheet exactly)
├── requirements.txt
└── README.md
```

---

## 📦 Requirements

```
streamlit>=1.32.0
pandas>=2.0.0
numpy>=1.26.0
plotly>=5.18.0
openpyxl>=3.1.0
```
