"""
HarvestBridge Automated Reporting Tool
Alitheia Capital — AI & Digital Innovation Internship Assessment
Author: Teju Adedoyin
"""

import streamlit as st
import tempfile, os
from calculator import calculate_financials
from report_generator import generate_report, fusd, fpct, fqoq
from excel_reader import read_excel
from narrative import build_narrative

st.set_page_config(page_title="HarvestBridge Reporting Tool", page_icon="🌾", layout="wide")

st.markdown("""
<style>
  .header-block {
    background:linear-gradient(135deg,#1A3C2B,#2E7D52);
    padding:1.2rem 1.5rem;border-radius:10px;margin-bottom:1.2rem;
  }
  .header-block h1{color:white;margin:0;font-size:1.5rem;}
  .header-block p{color:#B0CCBC;margin:.3rem 0 0;font-size:.85rem;}
  .sec{font-weight:600;color:#1A3C2B;border-bottom:2px solid #4CA06E;
       padding-bottom:.25rem;margin:1rem 0 .6rem;font-size:.95rem;}
  .hint{font-size:.78rem;color:#888;margin-top:-.4rem;margin-bottom:.6rem;}
  .stButton>button{background:#1A3C2B;color:white;border:none;
    border-radius:6px;padding:.55rem 1.5rem;font-weight:600;font-size:.95rem;width:100%;}
  .stButton>button:hover{background:#2E7D52;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-block">
  <h1>🌾 HarvestBridge Portfolio Reporting Tool</h1>
  <p>Alitheia Capital &nbsp;·&nbsp; Upload the Excel workbook, enter the new quarter's data, generate a report instantly.</p>
</div>
""", unsafe_allow_html=True)

# ── Step 1 — Upload Excel ─────────────────────────────────────────
st.markdown('<div class="sec">Step 1 — Upload the HarvestBridge Excel Workbook</div>',
            unsafe_allow_html=True)
uploaded = st.file_uploader("Upload HarvestBridge_Data.xlsx", type=["xlsx"])

excel_data            = None
prior_label           = None
prior_quarter         = {}
historical_quarters   = []
historical_financials = {}

if uploaded:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name
    try:
        excel_data            = read_excel(tmp_path)
        os.unlink(tmp_path)
        historical_quarters   = excel_data["quarters"]
        historical_financials = excel_data["financials"]
        prior_label           = excel_data["prior_quarter_label"]
        prior_quarter         = excel_data["prior_quarter"]
        st.success(
            f"Workbook loaded. Found **{len(historical_quarters)} quarters** "
            f"({', '.join(historical_quarters)}). "
            f"Prior quarter for QoQ: **{prior_label}**."
        )
    except Exception as e:
        st.error(f"Could not read the workbook: {e}")
else:
    st.info("Upload the Excel file above to enable automatic historical data and QoQ comparisons.")

st.markdown("---")

# ── Step 2 — New quarter inputs ───────────────────────────────────
st.markdown('<div class="sec">Step 2 - New Quarter Data</div>', unsafe_allow_html=True)

q_col, _ = st.columns([1, 4])
with q_col:
    quarter_label = st.text_input("Quarter label", value="Q2 2026",
                                   help="e.g. Q2 2026, Q3 2026")

tab_fin, tab_kpi = st.tabs(["📊 Financial Inputs (USD '000)", "📈 KPI Inputs"])

with tab_fin:
    st.markdown('<div class="sec">Revenue</div>', unsafe_allow_html=True)
    r1, r2, r3, r4 = st.columns(4)
    loan_fees = r1.number_input("Loan facilitation fees",    value=0.0, step=10.0)
    saas_rev  = r2.number_input("SaaS subscription revenue", value=0.0, step=10.0)
    insurance = r3.number_input("Insurance commissions",     value=0.0, step=10.0)
    other     = r4.number_input("Other income",              value=0.0, step=10.0)

    st.markdown('<div class="sec">Cost of Sales</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    tech    = c1.number_input("Technology & infrastructure", value=0.0, step=10.0)
    credit  = c2.number_input("Credit loss provision",       value=0.0, step=10.0)
    third_p = c3.number_input("Third-party data & APIs",     value=0.0, step=10.0)

    st.markdown('<div class="sec">Operating Expenses</div>', unsafe_allow_html=True)
    o1, o2, o3 = st.columns(3)
    sales_mkt = o1.number_input("Sales & marketing",         value=0.0, step=10.0)
    g_and_a   = o2.number_input("General & administrative",  value=0.0, step=10.0)
    rd        = o3.number_input("Research & development",    value=0.0, step=10.0)

    st.markdown('<div class="sec">Below the Line</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="hint">Expenses in brackets in the Excel are negative here. '
        'e.g. interest expense (238) → enter -238. FX gain = positive, FX loss = negative.</p>',
        unsafe_allow_html=True)
    b1, b2, b3, b4 = st.columns(4)
    dep      = b1.number_input("Depreciation & amortisation", value=0.0, step=1.0)
    interest = b2.number_input("Interest expense",            value=0.0, step=10.0)
    fx       = b3.number_input("FX gain / (loss)",            value=0.0, step=5.0)
    tax      = b4.number_input("Income tax expense",          value=0.0, step=10.0)

with tab_kpi:
    st.markdown('<div class="sec">Operations & People</div>', unsafe_allow_html=True)
    k1, k2, k3 = st.columns(3)
    cash_runway = k1.text_input("Cash runway (e.g. 22 months)", value="")
    fte         = k2.number_input("FTE headcount (total)",       value=0, step=1)
    female_fte  = k3.number_input("Female FTE headcount",        value=0, step=1)

    st.markdown('<div class="sec">Gender Lens (enter as whole numbers e.g. 40 for 40%)</div>',
                unsafe_allow_html=True)
    g1, g2, g3 = st.columns(3)
    f_board = g1.number_input("Female Board (%)",      value=0.0, step=0.5)
    f_exec  = g2.number_input("Female Executives (%)", value=0.0, step=0.5)
    f_mid   = g3.number_input("Female Mid-Mgmt (%)",   value=0.0, step=0.5)

    st.markdown('<div class="sec">Lending, Insurance & SaaS</div>', unsafe_allow_html=True)
    l1, l2, l3 = st.columns(3)
    loans    = l1.number_input("Loans disbursed (smallholder)", value=0,   step=100)
    policies = l2.number_input("Active insurance policies",     value=0,   step=100)
    claims   = l3.number_input("Claims paid (USD '000)",        value=0,   step=5)

    s1, s2 = st.columns(2)
    saas_u  = s1.number_input("Active SaaS users",             value=0, step=100)
    distrib = s2.number_input("Distributor partners (total)",   value=0, step=1)

# ── Assemble inputs & KPIs ────────────────────────────────────────
inputs = {
    "loan_fees": loan_fees, "saas_revenue": saas_rev,
    "insurance_commissions": insurance, "other_income": other,
    "tech_infrastructure": tech, "credit_loss_provision": credit,
    "third_party_data": third_p, "sales_marketing": sales_mkt,
    "general_admin": g_and_a, "rd": rd,
    "depreciation_amortisation": dep, "interest_expense": interest,
    "fx_loss": fx, "income_tax": tax,
}

# Normalise gender percentages — Excel stores as 0.4, user enters 40
def norm_pct(v):
    return v / 100 if v > 1 else v

kpis = {
    "cash_runway":             cash_runway,
    "fte_headcount":           fte,
    "female_fte":              female_fte,
    "female_board_pct":        norm_pct(f_board),
    "female_exec_pct":         norm_pct(f_exec),
    "female_midmgmt_pct":      norm_pct(f_mid),
    "loans_disbursed":         loans,
    "active_insurance_policies": policies,
    "claims_paid":             claims,
    "active_saas_users":       saas_u,
    "distributor_partners":    distrib,
}

# ── Live preview ──────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="sec">Calculated Metrics Preview</div>', unsafe_allow_html=True)

fin = calculate_financials(inputs, prior_quarter)

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Total Revenue", fusd(fin["total_revenue"]),
          f"{fqoq(fin.get('revenue_qoq'))} QoQ vs {prior_label}"
          if fin.get("revenue_qoq") else "")
m2.metric("Gross Profit",  fusd(fin["gross_profit"]),
          f"GM: {fpct(fin['gross_margin'])}")
m3.metric("EBITDA",        fusd(fin["ebitda"]),
          f"Margin: {fpct(fin['ebitda_margin'])}")
m4.metric("EBIT",          fusd(fin["ebit"]))
m5.metric("Net Income",    fusd(fin["net_income"]),
          f"{fqoq(fin.get('net_income_qoq'))} QoQ" if fin.get("net_income_qoq") else "")

if not uploaded:
    st.warning("Upload the Excel file in Step 1 to enable QoQ comparisons and the bar chart.")

# ── Generate ──────────────────────────────────────────────────────
st.markdown("---")
if st.button("⬇️  Generate PowerPoint Report"):
    if not uploaded:
        st.error("Please upload the Excel workbook first (Step 1).")
    else:
        quarterly_update = build_narrative(
            fin, kpis, quarter_label, prior_label,
            historical_financials, historical_quarters
        )

        company_info = {
            "name":    "HarvestBridge Limited",
            "country": "NG / KE",
            "sector":  "Agri-Fintech",
            "stake":   "10.6%",
            "description": (
                "HarvestBridge Limited is a Lagos and Nairobi-based agri-fintech company "
                "that provides digital working capital loans, crop insurance, and SaaS farm "
                "management tools to smallholder farmers across Nigeria and Kenya. Founded in "
                "2019, the company uses alternative data — satellite imagery, mobile money "
                "history, and input purchase records — to underwrite farmers excluded from "
                "formal banking. Its three revenue streams are loan facilitation fees, SaaS "
                "subscriptions to agri-input distributors, and embedded crop insurance commissions."
            ),
            "quarterly_update": quarterly_update,
        }

        with st.spinner("Generating report..."):
            pptx_bytes = generate_report(
                fin, kpis, quarter_label, prior_label,
                company_info, historical_financials, historical_quarters
            )
        st.success("Report ready.")
        st.download_button(
            label=f"📥  Download {quarter_label} Report (.pptx)",
            data=pptx_bytes,
            file_name=f"HarvestBridge_{quarter_label.replace(' ','_')}_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
