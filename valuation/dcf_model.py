"""
lululemon athletica: a discounted cash flow valuation from its filed financials,
computed in plain Python as the reference for the workbook.

    python valuation/dcf_model.py        print the three scenarios

Inputs are three files and nothing else: edgar_facts.json (every historical
figure, with the accession number of the 10-K it came from), market_inputs.json
(risk-free rate, equity risk premium, beta, debt, shares, each with its source)
and assumptions.json (the forecast drivers for a bear, base and bull case).
build_workbook.py writes the same calculation as Excel formulas, and
tests/test_dcf.py holds every figure in the workbook to this file.

Method
    free cash flow to the firm = EBIT x (1 - tax) + D&A - capex - increase in operating working capital
    discounted at the cost of equity (no borrowings, so it is the WACC) from the
    middle of each fiscal year back to the valuation date
    terminal value (value-driver formula) = NOPAT next year x (1 - g / RONIC) / (WACC - g)
    equity = enterprise value + cash - debt; per share over common plus exchangeable shares

Operating leases stay inside EBIT as rent (US GAAP operating lease cost), so lease
liabilities are not subtracted as debt: counting them in both places would charge
the rent twice. Share-based compensation is a real cost and is not added back.
"""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
FACTS = json.loads((HERE / "edgar_facts.json").read_text(encoding="utf-8"))
MARKET = json.loads((HERE / "market_inputs.json").read_text(encoding="utf-8"))
A = json.loads((HERE / "assumptions.json").read_text(encoding="utf-8"))
YEARS = list(FACTS["fiscal_years"])
DAYS_PER_YEAR = 365.25


def fact(key: str, year: str) -> float:
    return float(FACTS["facts"][key]["values"][year]["value"])


def history() -> dict[str, dict[str, float]]:
    out: dict[str, dict[str, float]] = {}
    for y in YEARS:
        rev = fact("revenue", y)
        nwc = (fact("receivables", y) + fact("inventory", y) + fact("prepaid_and_other_current", y)
               - fact("accounts_payable", y) - fact("accrued_liabilities", y) - fact("income_taxes_payable", y))
        out[y] = {
            "revenue": rev,
            "gross_margin": fact("gross_profit", y) / rev,
            "ebit": fact("operating_income", y),
            "ebit_margin": fact("operating_income", y) / rev,
            "other_operating_items": fact("gross_profit", y) - fact("sga", y)
            - fact("impairment_and_restructuring", y) - fact("operating_income", y),
            "tax_rate": fact("income_tax", y) / fact("pretax_income", y),
            "da_pct": fact("depreciation_and_amortization", y) / rev,
            "capex_pct": fact("capital_expenditure", y) / rev,
            "operating_nwc": nwc,
            "nwc_pct": nwc / rev,
            "sbc_pct": fact("share_based_compensation", y) / rev,
        }
    for prev, y in zip(YEARS, YEARS[1:], strict=False):
        out[y]["revenue_growth"] = out[y]["revenue"] / out[prev]["revenue"] - 1
    return out


def cost_of_capital() -> dict[str, float]:
    rf, erp, beta = (MARKET[k]["value"] for k in ("risk_free_rate", "equity_risk_premium", "beta"))
    coe = rf + beta * erp
    return {"risk_free_rate": rf, "equity_risk_premium": erp, "beta": beta, "cost_of_equity": coe, "wacc": coe}


def timing() -> list[float]:
    """Years from the valuation date to the middle of each forecast year."""
    start = date.fromisoformat(A["first_forecast_midpoint"])
    valuation = date.fromisoformat(MARKET["valuation_date"])
    first = (start - valuation).days / DAYS_PER_YEAR
    return [first + k for k in range(len(A["explicit_years"]))]


def shares() -> float:
    return float(MARKET["shares"]["common"] + MARKET["shares"]["exchangeable"])


def forecast(s: dict) -> list[dict[str, float]]:
    base = YEARS[-1]
    revenue = fact("revenue", base)
    nwc_prev = history()[base]["operating_nwc"]
    rows = []
    for k, year in enumerate(A["explicit_years"]):
        revenue = revenue * (1 + s["revenue_growth"][k])
        ebit = revenue * s["ebit_margin"][k]
        nopat = ebit * (1 - s["tax_rate"])
        da = revenue * s["da_pct"]
        capex = revenue * s["capex_pct"][k]
        nwc = revenue * s["nwc_pct"]
        fcff = nopat + da - capex - (nwc - nwc_prev)
        rows.append({"year": year, "revenue": revenue, "ebit": ebit, "nopat": nopat, "da": da, "capex": capex,
                     "nwc": nwc, "change_in_nwc": nwc - nwc_prev, "fcff": fcff})
        nwc_prev = nwc
    return rows


def value(s: dict, wacc: float | None = None, g: float | None = None) -> dict[str, float]:
    wacc = cost_of_capital()["wacc"] if wacc is None else wacc
    g = s["terminal_growth"] if g is None else g
    rows = forecast(s)
    t = timing()
    pv_fcff = sum(r["fcff"] / (1 + wacc) ** tk for r, tk in zip(rows, t, strict=True))
    last = rows[-1]
    nopat_next = last["ebit"] * (1 + g) * (1 - s["tax_rate"])
    terminal = nopat_next * (1 - g / s["terminal_return_on_new_capital"]) / (wacc - g)
    t_terminal = t[-1] + 0.5
    pv_terminal = terminal / (1 + wacc) ** t_terminal
    ev = pv_fcff + pv_terminal
    cash = fact("cash", YEARS[-1])
    equity = ev + cash - MARKET["debt"]["value"]
    return {
        "wacc": wacc, "terminal_growth": g, "pv_fcff": pv_fcff, "terminal_value": terminal,
        "pv_terminal": pv_terminal, "enterprise_value": ev, "cash": cash, "equity_value": equity,
        "value_per_share": equity / shares(), "terminal_share_of_ev": pv_terminal / ev,
        "implied_terminal_ev_ebitda": terminal / (last["ebit"] + last["da"]),
        "ev_to_fy2025_ebit": ev / fact("operating_income", YEARS[-1]),
    }


def sensitivity(scenario: str = "base") -> list[list[float]]:
    s = A["scenarios"][scenario]
    return [[value(s, w, g)["value_per_share"] for g in A["sensitivity"]["terminal_growth"]]
            for w in A["sensitivity"]["wacc"]]


def run() -> dict:
    return {
        "history": history(),
        "cost_of_capital": cost_of_capital(),
        "timing": timing(),
        "shares": shares(),
        "scenarios": {name: {"forecast": forecast(s), "value": value(s)} for name, s in A["scenarios"].items()},
        "beta_alternative": value(A["scenarios"]["base"],
                                  wacc=MARKET["risk_free_rate"]["value"]
                                  + MARKET["beta"]["alternative"]["value"] * MARKET["equity_risk_premium"]["value"]),
        "sensitivity": sensitivity(),
    }


if __name__ == "__main__":
    r = run()
    print(f"cost of equity {r['cost_of_capital']['cost_of_equity']:.2%}")
    for name, s in r["scenarios"].items():
        v = s["value"]
        print(f"{name:5}  EV ${v['enterprise_value'] / 1e9:,.2f}B  equity ${v['equity_value'] / 1e9:,.2f}B  "
              f"${v['value_per_share']:,.2f}/share  terminal {v['terminal_share_of_ev']:.0%} of EV  "
              f"implied exit {v['implied_terminal_ev_ebitda']:.1f}x EBITDA")
