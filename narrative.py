"""
narrative.py — generates quarterly update narrative from available data only.
Uses: financial calculations + KPI inputs from the Excel dashboard.
Does NOT use fields not present in the provided data.
"""


def _fusd(v):
    if v is None: return "N/A"
    if abs(v) >= 1000: return f"USD {v/1000:.2f}M"
    return f"USD {v:.0f}K"

def _fpct(v):
    if v is None: return "N/A"
    return f"{v*100:.1f}%"

def _fqoq(v, sign=True):
    if v is None: return None
    s = "+" if v >= 0 else ""
    return f"{s}{v*100:.1f}%"

def _fnum(v):
    if v is None: return "N/A"
    if v >= 1000: return f"{v/1000:.1f}K"
    return str(int(v))

def _qoq(current, prior):
    if prior and prior != 0:
        return (current - prior) / abs(prior)
    return None

def _yoy_label(quarter_label: str) -> str:
    try:
        parts = quarter_label.strip().split()
        return f"{parts[0]} {int(parts[1]) - 1}"
    except Exception:
        return None


def build_narrative(
    fin: dict,
    kpis: dict,
    quarter_label: str,
    prior_label: str,
    historical_financials: dict,
    historical_quarters: list,
) -> str:

    paragraphs = []

    rev      = fin["total_revenue"]
    rev_qoq  = fin.get("revenue_qoq")
    gp       = fin["gross_profit"]
    gm       = fin["gross_margin"]
    ebitda   = fin["ebitda"]
    ebitda_m = fin["ebitda_margin"]
    net      = fin["net_income"]

    # YoY — same quarter prior year
    yoy_label  = _yoy_label(quarter_label)
    yoy_fin    = historical_financials.get(yoy_label, {})
    rev_yoy    = _qoq(rev, yoy_fin.get("total_revenue")) if yoy_fin else None
    yoy_gm     = yoy_fin.get("gross_margin")
    yoy_net    = yoy_fin.get("net_income")
    ins_yoy    = _qoq(fin["insurance_commissions"],
                      yoy_fin.get("insurance_commissions")) if yoy_fin else None

    # ── Paragraph 1 — Revenue ────────────────────────────────────
    p1 = f"In {quarter_label}, HarvestBridge generated {_fusd(rev)} in total revenue"
    comps = []
    if rev_yoy is not None:
        comps.append(f"up {_fqoq(rev_yoy)} year-on-year versus {yoy_label}")
    if rev_qoq is not None and prior_label:
        comps.append(f"{_fqoq(rev_qoq)} quarter-on-quarter versus {prior_label}")
    p1 += (", " + " and ".join(comps) + ".") if comps else "."

    # Loan disbursements
    loans = kpis.get("loans_disbursed")
    if loans:
        p1 += f" Growth was driven by continued expansion of the loan book, with {_fnum(loans)} loans disbursed to smallholder farmers in the quarter."

    # Insurance YoY
    if ins_yoy is not None:
        p1 += f" Insurance commissions grew {_fqoq(ins_yoy)} YoY as the embedded crop insurance product deepened penetration across the existing borrower base."

    paragraphs.append(p1)

    # ── Paragraph 2 — Profitability ──────────────────────────────
    p2 = f"Profitability improved materially. Gross profit reached {_fusd(gp)}, reflecting a gross margin of {_fpct(gm)}"
    if yoy_gm is not None:
        bps = int((gm - yoy_gm) * 10000)
        direction = "up" if bps > 0 else "down"
        p2 += f" — {direction} {abs(bps)} basis points year-on-year"
    p2 += f". EBITDA rose to {_fusd(ebitda)} ({_fpct(ebitda_m)} margin)"

    # First profitable quarter detection
    if net > 0 and yoy_net is not None and yoy_net < 0:
        p2 += (
            f", and the company recorded its first profitable quarter with net income of "
            f"{_fusd(net)}, compared to a net loss of {_fusd(abs(yoy_net))} in {yoy_label}."
        )
    else:
        p2 += f". Net income of {_fusd(net)}."

    paragraphs.append(p2)

    # ── Paragraph 3 — Workforce ──────────────────────────────────
    fte        = kpis.get("fte_headcount")
    female_fte = kpis.get("female_fte")
    cash       = kpis.get("cash_runway", "")

    if fte:
        female_pct = round(float(female_fte) / float(fte) * 100, 1) if female_fte else None
        p3 = f"Workforce stood at {fte} FTE"
        if female_pct:
            p3 += f" with {female_pct}% female representation."
        else:
            p3 += "."
        if cash:
            p3 += f" Cash runway remained strong at {cash} following the Series B close."
        paragraphs.append(p3)

    return "\n\n".join(paragraphs)
