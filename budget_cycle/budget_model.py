"""
The FY2027 budget cycle for Kestrel Bay, computed in pandas without the workbook.

    python budget_cycle/budget_model.py            # print the headline numbers
    python budget_cycle/budget_model.py --q1       # (re)write data/q1_*.csv, the synthetic Q1 actuals

The workbook (budget_cycle/build_workbook.py) computes every figure below with
Excel formulas; tests/test_budget_cycle.py reads its cached values and holds
them to this module. The two share inputs (data/*.csv, assumptions.json) and
nothing else.

The cycle:
    V0  top-down targets: revenue growth on the last twelve months, an EBITDA margin
    V1  bottom-up submissions: channel volumes, category prices and costs, freight,
        department opex and new hires, phased by month on 2025's seasonality
    V2  the challenge round: foodservice price on beef and seafood, a hire deferred,
        warehouse non-pay cut, technology growth capped. V2 is the approved budget.
    Three statements, monthly, for the approved budget: P&L, balance sheet with a
        revolver that keeps a minimum cash balance, and the cash flow that ties them.
    3+9 reforecast after Q1: Q1 actuals, then the remaining nine months re-driven by
        Q1's volume run-rate by channel, unit cost by category and opex by department,
        with an EBITDA bridge back to the budget.
"""

from __future__ import annotations

import calendar
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
A = json.loads((HERE / "assumptions.json").read_text(encoding="utf-8"))
MONTHS = list(range(1, 13))


def inputs() -> dict:
    base = pd.read_csv(DATA / "base_segments.csv")
    season = pd.read_csv(DATA / "seasonality.csv").set_index("category")
    opex = pd.read_csv(DATA / "opex_base.csv").set_index("dept")["amount"]
    wc = pd.read_csv(DATA / "opening_wc.csv").set_index("account")["balance"]
    return {"base": base, "season": season, "opex": opex, "wc": wc}


def build(i: dict) -> pd.DataFrame:
    """Segment x month, both versions: the workbook's Build sheet."""
    v1, v2 = A["v1_submissions"], A["v2_challenge"]
    rows = []
    for _, s in i["base"].iterrows():
        price, cost, freight = s.revenue / s.cases, s.cogs / s.cases, s.freight / s.cases
        for m in MONTHS:
            share = i["season"].loc[s.category, f"m{m:02d}"]
            cases = s.cases * (1 + v1["volume_growth"][s.channel]) * share
            live = m >= v1["price_effective_month"]
            uplift = v1["price_increase"][s.category] if live else 0.0
            extra = (v2["foodservice_extra_price"].get(s.category, 0.0)
                     if live and s.channel in v2["foodservice_channels"] else 0.0)
            unit_cost = cost * (1 + v1["cost_inflation"][s.category])
            unit_freight = freight * (1 + v1["freight_inflation"])
            rev_v1 = cases * price * (1 + uplift)
            rows.append({"month": m, "category": s.category, "channel": s.channel, "cases": cases,
                         "price_case": price * (1 + uplift) * (1 + extra), "cost_case": unit_cost,
                         "freight_case": unit_freight, "revenue_v1": rev_v1, "revenue_v2": rev_v1 * (1 + extra),
                         "cogs": cases * unit_cost, "freight": cases * unit_freight})
    return pd.DataFrame(rows)


def opex(i: dict, version: str) -> pd.DataFrame:
    v1, v2 = A["v1_submissions"], A["v2_challenge"]
    out = {}
    for dept in A["departments"]:
        growth = v1["opex_growth"][dept]
        cut = 0.0
        if version == "v2":
            growth = min(growth, v2["opex_growth_cap"].get(dept, growth))
            cut = v2["opex_base_cut"].get(dept, 0.0)
        base = i["opex"][dept] / 12 * (1 + growth) * (1 - cut)
        row = []
        for m in MONTHS:
            hires = 0.0
            for h in v1["hires"]:
                start = h["start_month"]
                if version == "v2":
                    start = v2["hire_start_month"].get(h["dept"], start)
                if h["dept"] == dept and m >= start:
                    hires += h["fte"] * h["annual_cost"] / 12
            row.append(base + hires)
        out[dept] = row
    return pd.DataFrame(out, index=MONTHS)


def consolidation(i: dict, b: pd.DataFrame) -> dict:
    base_rev = i["base"].revenue.sum()
    base_ebitda = base_rev - i["base"].cogs.sum() - i["base"].freight.sum() - i["opex"].sum()
    target_rev = base_rev * (1 + A["targets"]["revenue_growth"])
    out = {"base": {"revenue": base_rev, "ebitda": base_ebitda},
           "target": {"revenue": target_rev, "ebitda": target_rev * A["targets"]["ebitda_margin"]}}
    for v in ("v1", "v2"):
        rev = b[f"revenue_{v}"].sum()
        gm = rev - b.cogs.sum() - b.freight.sum()
        ox = opex(i, v).to_numpy().sum()
        out[v] = {"revenue": rev, "cogs": b.cogs.sum(), "freight": b.freight.sum(), "gross_margin": gm,
                  "opex": ox, "ebitda": gm - ox}
    o1, o2 = opex(i, "v1").sum(), opex(i, "v2").sum()
    out["bridge_v1_v2"] = {"price": out["v2"]["revenue"] - out["v1"]["revenue"],
                           **{f"opex {d}": float(o1[d] - o2[d]) for d in A["departments"]}}
    return out


def statements(i: dict, b: pd.DataFrame, balance_sheet: dict | None = None) -> pd.DataFrame:
    bs = A["balance_sheet"] if balance_sheet is None else balance_sheet
    ox = opex(i, "v2").sum(axis=1)
    rev = b.groupby("month").revenue_v2.sum()
    cogs = b.groupby("month").cogs.sum()
    frt = b.groupby("month").freight.sum()
    capex = pd.Series({m: float(bs["capex"].get(str(m), 0.0)) for m in MONTHS})
    days = pd.Series({m: calendar.monthrange(A["fiscal_year"], m)[1] for m in MONTHS})

    cash, rev_bal, ppe = bs["opening_cash"], bs["opening_revolver"], bs["opening_ppe"]
    ar, inv, ap = i["wc"]["Accounts receivable"], i["wc"]["Inventory"], i["wc"]["Accounts payable"]
    equity = cash + ar + inv + ppe - ap - rev_bal
    opening = {"cash": cash, "ar": ar, "inventory": inv, "ppe": ppe, "ap": ap, "revolver": rev_bal,
               "equity": equity}
    rows, cum_capex = [], 0.0
    for m in MONTHS:
        da = bs["existing_da_per_month"] + cum_capex / bs["useful_life_months"]
        ebitda = rev[m] - cogs[m] - frt[m] - ox[m]
        interest = rev_bal * bs["revolver_rate"] / 12
        ebt = ebitda - da - interest
        tax = ebt * bs["tax_rate"]
        ni = ebt - tax
        ar_m = rev[m] / days[m] * bs["dso"]
        inv_m = cogs[m] / days[m] * bs["dio"]
        ap_m = cogs[m] / days[m] * bs["dpo"]
        cfo = ni + da - (ar_m - ar) - (inv_m - inv) + (ap_m - ap)
        pre = cash + cfo - capex[m]
        draw = max(bs["minimum_cash"] - pre, -rev_bal)
        rows.append({"month": m, "revenue": rev[m], "cogs": cogs[m], "freight": frt[m],
                     "gross_margin": rev[m] - cogs[m] - frt[m], "opex": ox[m], "ebitda": ebitda, "da": da,
                     "ebit": ebitda - da, "interest": interest, "ebt": ebt, "tax": tax, "net_income": ni,
                     "capex": capex[m], "cfo": cfo, "revolver_flow": draw})
        cum_capex += capex[m]
        ppe = ppe + capex[m] - da
        cash, rev_bal, ar, inv, ap = pre + draw, rev_bal + draw, ar_m, inv_m, ap_m
        equity += ni
        rows[-1].update({"cash": cash, "ar": ar, "inventory": inv, "ppe": ppe, "ap": ap, "revolver": rev_bal,
                         "equity": equity,
                         "balance_check": cash + ar + inv + ppe - ap - rev_bal - equity})
    out = pd.DataFrame(rows).set_index("month")
    out.attrs["opening"] = opening
    return out


def q1_actuals(i: dict, b: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Synthetic Q1 actuals: the approved budget, moved by the stated deviations
    and a little seeded noise. Written once to data/ and read from there."""
    q = A["q1_actuals"]
    rng = np.random.default_rng(q["seed"])
    rows = []
    for _, r in b[b.month <= 3].iterrows():
        cases = r.cases * (1 + q["volume"].get(r.channel, 0.0)) * (1 + rng.normal(0, q["noise_sd"]))
        price = r.price_case * (1 + rng.normal(0, q["noise_sd"] / 4))
        cost = r.cost_case * (1 + q["unit_cost"].get(r.category, 0.0)) * (1 + rng.normal(0, q["noise_sd"] / 2))
        rows.append({"month": int(r.month), "category": r.category, "channel": r.channel,
                     "cases": round(cases, 1), "revenue": round(cases * price, 2), "cogs": round(cases * cost, 2),
                     "freight": round(cases * r.freight_case, 2)})
    ox = opex(i, "v2")
    orows = [{"month": m, "dept": d,
              "amount": round(ox.loc[m, d] * (1 + q["opex"].get(d, 0.0)) * (1 + rng.normal(0, q["noise_sd"])), 2)}
             for m in (1, 2, 3) for d in A["departments"]]
    return pd.DataFrame(rows), pd.DataFrame(orows)


def reforecast(i: dict, b: pd.DataFrame) -> dict:
    act = pd.read_csv(DATA / "q1_sales_actuals.csv")
    act_ox = pd.read_csv(DATA / "q1_opex_actuals.csv")
    bq = b[b.month <= 3]
    vol = act.groupby("channel").cases.sum() / bq.groupby("channel").cases.sum()
    cost = ((act.groupby("category").cogs.sum() / act.groupby("category").cases.sum())
            / (bq.groupby("category").cogs.sum() / bq.groupby("category").cases.sum()))
    ox2 = opex(i, "v2")
    ox_ratio = act_ox.groupby("dept").amount.sum() / ox2.loc[1:3].sum()

    rest = b[b.month > 3].copy()
    rest["cases_rf"] = rest.cases * rest.channel.map(vol)
    rest["revenue_rf"] = rest.cases_rf * rest.price_case
    rest["cogs_rf"] = rest.cases_rf * rest.cost_case * rest.category.map(cost)
    rest["freight_rf"] = rest.cases_rf * rest.freight_case
    opex_rf_rest = sum(ox2.loc[4:12, d].sum() * ox_ratio[d] for d in A["departments"])

    bud = {"revenue": b.revenue_v2.sum(), "gross_margin": b.revenue_v2.sum() - b.cogs.sum() - b.freight.sum(),
           "opex": ox2.to_numpy().sum()}
    bud["ebitda"] = bud["gross_margin"] - bud["opex"]
    q1_gm = act.revenue.sum() - act.cogs.sum() - act.freight.sum()
    rf = {"revenue": act.revenue.sum() + rest.revenue_rf.sum(),
          "gross_margin": q1_gm + (rest.revenue_rf - rest.cogs_rf - rest.freight_rf).sum(),
          "opex": act_ox.amount.sum() + opex_rf_rest}
    rf["ebitda"] = rf["gross_margin"] - rf["opex"]
    margin_case = rest.price_case - rest.cost_case - rest.freight_case
    q1_bud_ebitda = (bq.revenue_v2.sum() - bq.cogs.sum() - bq.freight.sum()) - ox2.loc[1:3].to_numpy().sum()
    bridge = {
        "q1_actual_vs_budget": (q1_gm - act_ox.amount.sum()) - q1_bud_ebitda,
        "volume": ((rest.cases_rf - rest.cases) * margin_case).sum(),
        "unit_cost": -(rest.cases_rf * rest.cost_case * (rest.category.map(cost) - 1)).sum(),
        "opex": -sum(ox2.loc[4:12, d].sum() * (ox_ratio[d] - 1) for d in A["departments"]),
    }
    return {"budget": bud, "reforecast": rf, "bridge": bridge, "volume_ratio": vol.to_dict(),
            "cost_ratio": cost.to_dict(), "opex_ratio": ox_ratio.to_dict()}


def run() -> dict:
    i = inputs()
    b = build(i)
    return {"inputs": i, "build": b, "consolidation": consolidation(i, b), "statements": statements(i, b),
            "reforecast": reforecast(i, b)}


def main(argv: list[str]) -> int:
    if argv[:1] == ["--q1"]:
        i = inputs()
        sales, ox = q1_actuals(i, build(i))
        sales.to_csv(DATA / "q1_sales_actuals.csv", index=False, lineterminator="\n")
        ox.to_csv(DATA / "q1_opex_actuals.csv", index=False, lineterminator="\n")
        print(f"wrote {len(sales)} sales rows and {len(ox)} opex rows")
        return 0
    r = run()
    c, s, rf = r["consolidation"], r["statements"], r["reforecast"]
    for k in ("base", "target", "v1", "v2"):
        print(f"{k:7s} revenue {c[k]['revenue']:>14,.0f}  EBITDA {c[k]['ebitda']:>12,.0f}"
              f"  margin {c[k]['ebitda'] / c[k]['revenue']:.2%}")
    print("bridge V1->V2", {k: round(v) for k, v in c["bridge_v1_v2"].items()})
    print("net income", round(s.net_income.sum()), "closing cash", round(s.cash.iloc[-1]),
          "revolver", round(s.revolver.iloc[-1]), "max |BS check|", abs(s.balance_check).max())
    print("RF", {k: round(v) for k, v in rf["reforecast"].items()}, "bridge",
          {k: round(v) for k, v in rf["bridge"].items()})
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
