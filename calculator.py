"""
calculator.py — all derived metric calculations and QoQ variances
No hardcoded quarter data. Works for any quarter.
"""


def calculate_financials(inputs: dict, prior_quarter: dict) -> dict:
    total_revenue = (
        inputs["loan_fees"] +
        inputs["saas_revenue"] +
        inputs["insurance_commissions"] +
        inputs["other_income"]
    )
    total_cos = (
        inputs["tech_infrastructure"] +
        inputs["credit_loss_provision"] +
        inputs["third_party_data"]
    )
    gross_profit  = total_revenue - total_cos
    gross_margin  = gross_profit / total_revenue if total_revenue else 0
    total_opex    = inputs["sales_marketing"] + inputs["general_admin"] + inputs["rd"]
    ebitda        = gross_profit - total_opex
    ebitda_margin = ebitda / total_revenue if total_revenue else 0
    ebit          = ebitda - inputs["depreciation_amortisation"]
    net_income = ebit + inputs["interest_expense"] + inputs["fx_loss"] + inputs["income_tax"]

    def qoq(current, prior_key):
        prior = prior_quarter.get(prior_key)
        if prior and prior != 0:
            return (current - prior) / abs(prior)
        return None

    return {
        "loan_fees": inputs["loan_fees"],
        "saas_revenue": inputs["saas_revenue"],
        "insurance_commissions": inputs["insurance_commissions"],
        "other_income": inputs["other_income"],
        "total_revenue": total_revenue,
        "tech_infrastructure": inputs["tech_infrastructure"],
        "credit_loss_provision": inputs["credit_loss_provision"],
        "third_party_data": inputs["third_party_data"],
        "total_cos": total_cos,
        "gross_profit": gross_profit,
        "gross_margin": gross_margin,
        "sales_marketing": inputs["sales_marketing"],
        "general_admin": inputs["general_admin"],
        "rd": inputs["rd"],
        "total_opex": total_opex,
        "ebitda": ebitda,
        "ebitda_margin": ebitda_margin,
        "depreciation_amortisation": inputs["depreciation_amortisation"],
        "ebit": ebit,
        "interest_expense": inputs["interest_expense"],
        "fx_loss": inputs["fx_loss"],
        "income_tax": inputs["income_tax"],
        "net_income": net_income,
        "revenue_qoq":    qoq(total_revenue, "total_revenue"),
        "ebit_qoq":       qoq(ebit,          "ebit"),
        "net_income_qoq": qoq(net_income,     "net_income"),
    }
