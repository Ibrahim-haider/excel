
from __future__ import annotations

import io
from pathlib import Path
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

st.set_page_config(
    page_title="Branch Intelligence Platform",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------
# Styling
# ----------------------------
st.markdown("""
<style>
:root {
  --panel: #111827;
  --border: #243244;
  --muted: #94a3b8;
}
.block-container {padding-top: 1.2rem; padding-bottom: 2.5rem;}
[data-testid="stSidebar"] {border-right: 1px solid #243244;}
[data-testid="stMetric"] {
  background: linear-gradient(145deg, #111827, #0f172a);
  border: 1px solid #263244;
  padding: 16px;
  border-radius: 14px;
}
.hero {
  padding: 20px 22px;
  border: 1px solid #263244;
  border-radius: 16px;
  background: linear-gradient(135deg, rgba(37,99,235,.14), rgba(15,23,42,.95));
  margin-bottom: 18px;
}
.hero h1 {margin:0; font-size:2rem;}
.hero p {margin:.4rem 0 0; color:#a8b3c5;}
.badge {
  display:inline-block; padding:5px 9px; margin:3px 4px 3px 0;
  border-radius:999px; border:1px solid #334155; color:#cbd5e1;
  font-size:.78rem;
}
.section-title {font-size:1.28rem; font-weight:750; margin:8px 0 14px;}
.notice {
  border-left:4px solid #f59e0b; padding:10px 14px; background:#1f2937;
  border-radius:8px; color:#dbe4f0;
}
</style>
""", unsafe_allow_html=True)

# ----------------------------
# Authentication
# ----------------------------
USERS = {
    "admin": {"password": "admin123", "role": "Admin", "name": "Demo Administrator"},
    "manager": {"password": "manager123", "role": "Manager", "name": "Branch Manager"},
    "analyst": {"password": "analyst123", "role": "Analyst", "name": "Business Analyst"},
}

def login_screen() -> None:
    st.markdown("""
    <div class="hero">
      <h1>Branch Intelligence Platform</h1>
      <p>Automated MTD and YTD branch performance analysis.</p>
    </div>
    """, unsafe_allow_html=True)

    left, center, right = st.columns([1, 1.15, 1])
    with center:
        st.subheader("Sign in")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.button("Login", use_container_width=True, type="primary"):
            user = USERS.get(username)
            if user and user["password"] == password:
                st.session_state.authenticated = True
                st.session_state.user = username
                st.session_state.role = user["role"]
                st.session_state.name = user["name"]
                st.rerun()
            else:
                st.error("Invalid credentials.")

        with st.expander("Demo credentials"):
            st.code("Admin: admin / admin123\nManager: manager / manager123\nAnalyst: analyst / analyst123")

if not st.session_state.get("authenticated"):
    login_screen()
    st.stop()

# ----------------------------
# Helpers
# ----------------------------
REQUIRED = [
    "Branch", "Target (PKR)", "Actual Sales (PKR)",
    "Cash Sales (PKR)", "Installment Sales (PKR)",
    "Units Sold", "Avg Ticket (PKR)"
]

MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

@st.cache_data
def read_excel(source) -> pd.DataFrame:
    return pd.read_excel(source)

def money(value: float) -> str:
    value = float(value)
    if abs(value) >= 1_000_000_000:
        return f"PKR {value/1_000_000_000:.2f}B"
    if abs(value) >= 1_000_000:
        return f"PKR {value/1_000_000:.2f}M"
    if abs(value) >= 1_000:
        return f"PKR {value/1_000:.1f}K"
    return f"PKR {value:,.0f}"

def validate(df: pd.DataFrame, label: str) -> None:
    missing = [col for col in REQUIRED if col not in df.columns]
    if missing:
        st.error(f"{label} file is missing required columns: {', '.join(missing)}")
        st.stop()
    numeric = [c for c in REQUIRED if c != "Branch"]
    for col in numeric:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    if df[numeric].isna().any().any():
        st.warning(f"{label} contains invalid numeric values. Invalid entries were treated as missing.")

def build_branch_summary(df: pd.DataFrame) -> pd.DataFrame:
    out = (
        df.groupby("Branch", as_index=False)
        .agg({
            "Actual Sales (PKR)": "sum",
            "Target (PKR)": "sum",
            "Cash Sales (PKR)": "sum",
            "Installment Sales (PKR)": "sum",
            "Units Sold": "sum"
        })
    )
    out["Achievement %"] = out["Actual Sales (PKR)"] / out["Target (PKR)"].replace(0, pd.NA) * 100
    out["Contribution %"] = out["Actual Sales (PKR)"] / out["Actual Sales (PKR)"].sum() * 100
    out["Avg Ticket (PKR)"] = out["Actual Sales (PKR)"] / out["Units Sold"].replace(0, pd.NA)
    out["Variance (PKR)"] = out["Actual Sales (PKR)"] - out["Target (PKR)"]
    out["Status"] = pd.cut(
        out["Achievement %"],
        bins=[-float("inf"), 85, 100, float("inf")],
        labels=["Needs Attention", "On Track", "Strong"]
    )
    return out

def build_insights(mtd_df: pd.DataFrame, ytd_df: pd.DataFrame, branch_df: pd.DataFrame) -> list[str]:
    mtd_sales = mtd_df["Actual Sales (PKR)"].sum()
    mtd_target = mtd_df["Target (PKR)"].sum()
    ytd_sales = ytd_df["Actual Sales (PKR)"].sum()
    cash = ytd_df["Cash Sales (PKR)"].sum()
    inst = ytd_df["Installment Sales (PKR)"].sum()
    achievement = (mtd_sales / mtd_target * 100) if mtd_target else 0
    cash_share = (cash / ytd_sales * 100) if ytd_sales else 0
    inst_share = (inst / ytd_sales * 100) if ytd_sales else 0

    top = branch_df.sort_values("Actual Sales (PKR)", ascending=False).iloc[0]
    low = branch_df.sort_values("Actual Sales (PKR)", ascending=True).iloc[0]
    weak = branch_df[branch_df["Achievement %"] < 85].sort_values("Achievement %")
    over = branch_df[branch_df["Achievement %"] >= 100].sort_values("Achievement %", ascending=False)

    insights = [
        f"{top['Branch']} is the leading branch, contributing {top['Contribution %']:.1f}% of selected YTD sales.",
        f"Overall MTD target achievement is {achievement:.1f}%.",
        f"Installment sales represent {inst_share:.1f}% of YTD sales, while cash contributes {cash_share:.1f}%.",
        f"{low['Branch']} is the lowest-performing branch by selected YTD sales.",
    ]
    if len(weak):
        insights.append("Branches below the 85% performance threshold: " + ", ".join(weak["Branch"].tolist()) + ".")
    else:
        insights.append("No selected branch is below the 85% performance threshold.")
    if len(over):
        insights.append("Branches meeting or exceeding target: " + ", ".join(over["Branch"].tolist()) + ".")
    return insights

def make_excel_report(kpis: dict, branch_df: pd.DataFrame, insights: list[str]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Executive Summary"
    ws["A1"] = "Branch Intelligence Platform - Executive Summary"
    ws["A1"].font = Font(bold=True, size=16)
    row = 3
    for key, value in kpis.items():
        ws.cell(row=row, column=1, value=key).font = Font(bold=True)
        ws.cell(row=row, column=2, value=value)
        row += 1
    row += 1
    ws.cell(row=row, column=1, value="Management Insights").font = Font(bold=True, size=12)
    row += 1
    for item in insights:
        ws.cell(row=row, column=1, value=item)
        row += 1

    ws2 = wb.create_sheet("Branch Analysis")
    for c, name in enumerate(branch_df.columns, 1):
        cell = ws2.cell(row=1, column=c, value=name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    for r, record in enumerate(branch_df.to_dict("records"), 2):
        for c, name in enumerate(branch_df.columns, 1):
            value = record[name]
            if hasattr(value, "item"):
                value = value.item()
            ws2.cell(row=r, column=c, value=value)

    for wsx in [ws, ws2]:
        for col in wsx.columns:
            max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
            wsx.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

def make_pdf_report(kpis: dict, insights: list[str]) -> bytes:
    bio = io.BytesIO()
    c = canvas.Canvas(bio, pagesize=A4)
    width, height = A4
    y = height - 60
    c.setFont("Helvetica-Bold", 16)
    c.drawString(48, y, "Branch Intelligence Platform")
    y -= 24
    c.setFont("Helvetica", 10)
    c.drawString(48, y, f"Generated: {datetime.now().strftime('%d %b %Y, %I:%M %p')}")
    y -= 30

    c.setFont("Helvetica-Bold", 12)
    c.drawString(48, y, "Executive KPIs")
    y -= 20
    c.setFont("Helvetica", 10)
    for key, value in kpis.items():
        c.drawString(58, y, f"{key}: {value}")
        y -= 16

    y -= 10
    c.setFont("Helvetica-Bold", 12)
    c.drawString(48, y, "Management Conclusions")
    y -= 20
    c.setFont("Helvetica", 9)
    for item in insights:
        words = item.split()
        line = ""
        for word in words:
            test = f"{line} {word}".strip()
            if c.stringWidth(test, "Helvetica", 9) > width - 110:
                c.drawString(58, y, line)
                y -= 14
                line = word
            else:
                line = test
        if line:
            c.drawString(58, y, line)
            y -= 18
        if y < 70:
            c.showPage()
            y = height - 60
            c.setFont("Helvetica", 9)

    c.save()
    return bio.getvalue()

# ----------------------------
# Sidebar
# ----------------------------
with st.sidebar:
    st.title("📊 BIP")
    st.caption("Branch Intelligence Platform")
    st.write(f"**{st.session_state.name}**")
    st.caption(st.session_state.role)
    if st.button("Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    st.divider()
    source_mode = st.radio("Data source", ["Bundled demo data", "Upload Excel files"])

    if source_mode == "Bundled demo data":
        mtd_source = DATA_DIR / "Dummy_MTD_Sales.xlsx"
        ytd_source = DATA_DIR / "Dummy_YTD_Sales.xlsx"
        st.success("Fictional demo data loaded")
    else:
        mtd_upload = st.file_uploader("Upload MTD Excel", type=["xlsx"])
        ytd_upload = st.file_uploader("Upload YTD Excel", type=["xlsx"])
        if not mtd_upload or not ytd_upload:
            st.info("Upload both MTD and YTD files.")
            st.stop()
        mtd_source = mtd_upload
        ytd_source = ytd_upload

mtd = read_excel(mtd_source)
ytd = read_excel(ytd_source)
validate(mtd, "MTD")
validate(ytd, "YTD")

all_branches = sorted(set(mtd["Branch"]).union(set(ytd["Branch"])))
with st.sidebar:
    st.divider()
    selected_branches = st.multiselect("Branches", all_branches, default=all_branches)
    if not selected_branches:
        st.warning("Select at least one branch.")
        st.stop()

mtd_f = mtd[mtd["Branch"].isin(selected_branches)].copy()
ytd_f = ytd[ytd["Branch"].isin(selected_branches)].copy()

available_months = []
if "Month" in ytd_f.columns:
    available_months = [m for m in MONTH_ORDER if m in ytd_f["Month"].astype(str).unique()]
    with st.sidebar:
        selected_months = st.multiselect("YTD months", available_months, default=available_months)
    ytd_f = ytd_f[ytd_f["Month"].isin(selected_months)]

branch = build_branch_summary(ytd_f)

# KPI calculations
mtd_sales = mtd_f["Actual Sales (PKR)"].sum()
mtd_target = mtd_f["Target (PKR)"].sum()
ytd_sales = ytd_f["Actual Sales (PKR)"].sum()
ytd_target = ytd_f["Target (PKR)"].sum()
cash_sales = ytd_f["Cash Sales (PKR)"].sum()
inst_sales = ytd_f["Installment Sales (PKR)"].sum()
units = int(ytd_f["Units Sold"].sum())
mtd_achievement = (mtd_sales / mtd_target * 100) if mtd_target else 0
ytd_achievement = (ytd_sales / ytd_target * 100) if ytd_target else 0
cash_share = (cash_sales / ytd_sales * 100) if ytd_sales else 0
inst_share = (inst_sales / ytd_sales * 100) if ytd_sales else 0
avg_ticket = ytd_sales / units if units else 0
top = branch.sort_values("Actual Sales (PKR)", ascending=False).iloc[0]
low = branch.sort_values("Actual Sales (PKR)", ascending=True).iloc[0]

insights = build_insights(mtd_f, ytd_f, branch)

st.markdown("""
<div class="hero">
  <h1>Branch Intelligence Platform</h1>
  <p>Automated MTD and YTD performance analysis for branches, targets, cash, and installments.</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<span class="badge">MTD Analysis</span>
<span class="badge">YTD Analysis</span>
<span class="badge">Branch Ranking</span>
<span class="badge">Cash vs Installment</span>
<span class="badge">Management Conclusions</span>
<span class="badge">Excel & PDF Export</span>
""", unsafe_allow_html=True)

st.markdown('<div class="notice">Demo environment: all figures in the bundled files are fictional.</div>', unsafe_allow_html=True)

tabs = st.tabs([
    "Executive Dashboard",
    "Branch Performance",
    "Cash vs Installment",
    "Management Conclusions",
    "Reports & Data"
])

with tabs[0]:
    st.markdown('<div class="section-title">Executive overview</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("MTD Sales", money(mtd_sales))
    c2.metric("YTD Sales", money(ytd_sales))
    c3.metric("MTD Achievement", f"{mtd_achievement:.1f}%")
    c4.metric("YTD Achievement", f"{ytd_achievement:.1f}%")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Top Branch", top["Branch"])
    c6.metric("Lowest Branch", low["Branch"])
    c7.metric("YTD Units", f"{units:,}")
    c8.metric("Average Ticket", money(avg_ticket))

    left, right = st.columns([1.35, 1])
    with left:
        st.markdown('<div class="section-title">Monthly YTD sales trend</div>', unsafe_allow_html=True)
        if "Month" in ytd_f.columns and len(ytd_f):
            monthly = ytd_f.groupby("Month", as_index=False)["Actual Sales (PKR)"].sum()
            monthly["Month"] = pd.Categorical(monthly["Month"], categories=available_months, ordered=True)
            monthly = monthly.sort_values("Month")
            fig = px.line(monthly, x="Month", y="Actual Sales (PKR)", markers=True)
            fig.update_layout(yaxis_title="Sales (PKR)")
            st.plotly_chart(fig, use_container_width=True)
    with right:
        st.markdown('<div class="section-title">Payment mix</div>', unsafe_allow_html=True)
        mix = pd.DataFrame({"Payment Type": ["Cash", "Installment"], "Sales": [cash_sales, inst_sales]})
        fig = px.pie(mix, names="Payment Type", values="Sales", hole=.58)
        st.plotly_chart(fig, use_container_width=True)

with tabs[1]:
    st.markdown('<div class="section-title">Branch performance analysis</div>', unsafe_allow_html=True)
    left, right = st.columns(2)
    with left:
        ranked = branch.sort_values("Actual Sales (PKR)", ascending=True)
        fig = px.bar(ranked, x="Actual Sales (PKR)", y="Branch", orientation="h", text_auto=".2s")
        fig.update_layout(xaxis_title="YTD Sales (PKR)")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = px.bar(
            branch.sort_values("Achievement %", ascending=False),
            x="Branch", y="Achievement %", color="Status", text_auto=".1f"
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

    table = branch.copy()
    for col in ["Actual Sales (PKR)", "Target (PKR)", "Cash Sales (PKR)", "Installment Sales (PKR)", "Avg Ticket (PKR)", "Variance (PKR)"]:
        table[col] = table[col].map(lambda x: f"{x:,.0f}")
    for col in ["Achievement %", "Contribution %"]:
        table[col] = table[col].map(lambda x: f"{x:.1f}%")
    st.dataframe(table, use_container_width=True, hide_index=True)

with tabs[2]:
    st.markdown('<div class="section-title">Cash and installment analysis</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Cash Sales", money(cash_sales))
    c2.metric("Installment Sales", money(inst_sales))
    c3.metric("Cash Share", f"{cash_share:.1f}%")
    c4.metric("Installment Share", f"{inst_share:.1f}%")

    branch_mix = branch.melt(
        id_vars="Branch",
        value_vars=["Cash Sales (PKR)", "Installment Sales (PKR)"],
        var_name="Payment Type",
        value_name="Sales"
    )
    fig = px.bar(branch_mix, x="Branch", y="Sales", color="Payment Type", barmode="stack")
    fig.update_layout(xaxis_tickangle=-35)
    st.plotly_chart(fig, use_container_width=True)

with tabs[3]:
    st.markdown('<div class="section-title">Automatically generated management conclusions</div>', unsafe_allow_html=True)
    for i, item in enumerate(insights, start=1):
        st.info(f"{i}. {item}")

    st.subheader("Recommended management focus")
    weak = branch[branch["Achievement %"] < 85].sort_values("Achievement %")
    if len(weak):
        st.write("Prioritize review of:")
        for _, row in weak.iterrows():
            gap = abs(row["Variance (PKR)"])
            st.write(f"- **{row['Branch']}** — achievement {row['Achievement %']:.1f}%, target gap {money(gap)}")
    else:
        st.success("All selected branches are currently above the 85% threshold.")

with tabs[4]:
    st.markdown('<div class="section-title">Reports and underlying data</div>', unsafe_allow_html=True)

    kpis = {
        "MTD Sales": money(mtd_sales),
        "YTD Sales": money(ytd_sales),
        "MTD Achievement": f"{mtd_achievement:.1f}%",
        "YTD Achievement": f"{ytd_achievement:.1f}%",
        "Cash Share": f"{cash_share:.1f}%",
        "Installment Share": f"{inst_share:.1f}%",
        "Top Branch": str(top["Branch"]),
        "Lowest Branch": str(low["Branch"]),
        "YTD Units": f"{units:,}",
        "Average Ticket": money(avg_ticket),
    }

    excel_bytes = make_excel_report(kpis, branch, insights)
    pdf_bytes = make_pdf_report(kpis, insights)

    c1, c2 = st.columns(2)
    c1.download_button(
        "Download Excel management report",
        data=excel_bytes,
        file_name="branch_intelligence_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    c2.download_button(
        "Download PDF executive summary",
        data=pdf_bytes,
        file_name="branch_intelligence_summary.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    with st.expander("View MTD data"):
        st.dataframe(mtd_f, use_container_width=True, hide_index=True)
    with st.expander("View YTD data"):
        st.dataframe(ytd_f, use_container_width=True, hide_index=True)

st.divider()
st.caption("Milestone 1 • Presentation-ready proof of concept • Fictional data only")
