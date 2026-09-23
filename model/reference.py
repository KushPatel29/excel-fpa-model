"""An independent recomputation of every figure the workbook publishes.

The workbook does its arithmetic in Excel: SUMIFS over tables, Power Query,
Power Pivot measures, what-if data tables. This module does the same analysis
in pandas from the same CSVs, without looking at the workbook. The tests hold
the two to each other to the cent, so a broken formula, a range that stops one
row short or a mis-wired scenario cannot survive a CI run.

Definitions (they are the ones printed in the workbook):

* Outlook = actuals for closed months + a driver forecast for open months.
* Forecast units = budget units x the category's YTD volume run-rate
  (YTD actual units / YTD budget units) x (1 + volume driver).
* Forecast price, unit cost and freight per case = the trailing three closed
  months, each x (1 + its driver). Opex = budget x the department's YTD
  run-rate x (1 + opex driver).
* Price-volume-mix runs on category x channel segments. Each segment's change
  in cases is valued twice: at the budget's average margin per case (volume)
  and at the segment's margin above or below that average (mix). Price and
  cost are the per-case differences on actual cases. The four effects add up
  exactly to each segment's gross-margin variance, and so to the total.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

AS_OF = pd.Timestamp("2026-08-01")
FY = 2026
TRAIL_MONTHS = 3
WC_DAYS = 91
SHELF_RISK = 0.50

CATEGORIES = ["Beef", "Pork", "Poultry", "Seafood", "Charcuterie", "Cheese & Dairy"]
CHANNELS = ["Restaurants", "Grocery Retail", "Hotels & Institutions", "Online Direct"]
DEPARTMENTS = ["Sales & Marketing", "Warehouse & Logistics", "General & Admin", "Technology"]

# driver: (base, upside, downside)
SCENARIOS = {
    "volume": (0.0, 0.03, -0.04),
    "price": (0.0, 0.015, -0.01),
    "cost": (0.0, -0.01, 0.025),
    "freight": (0.0, -0.02, 0.05),
    "opex": (0.0, -0.01, 0.02),
}
SCENARIO_NAMES = ("Base", "Upside", "Downside")
SENS_PRICE = (-0.03, -0.02, -0.01, 0.0, 0.01, 0.02, 0.03)
SENS_VOLUME = (-0.06, -0.04, -0.02, 0.0, 0.02, 0.04, 0.06)


@dataclass(frozen=True)
class Data:
    sales: pd.DataFrame
    budget: pd.DataFrame
    opex: pd.DataFrame
    balances: pd.DataFrame
    product: pd.DataFrame


def load() -> Data:
    sales = pd.read_csv(DATA / "fact_sales.csv", parse_dates=["month"])
    wide = pd.read_csv(DATA / "budget_units_fy2026.csv")
    rates = pd.read_csv(DATA / "budget_rates_fy2026.csv")
    product = pd.read_csv(DATA / "dim_product.csv")
    keys = ["sku", "channel", "region"]
    budget = wide.melt(id_vars=keys, var_name="month", value_name="units")
    budget["month"] = pd.to_datetime(budget["month"] + "-01")
    budget = budget.merge(rates, on=keys).merge(product[["sku", "category"]], on="sku")
    budget["revenue"] = budget["units"] * budget["net_price"]
    budget["cogs"] = budget["units"] * budget["unit_cost"]
    budget["freight"] = budget["units"] * budget["freight_per_case"]
    opex = pd.read_csv(DATA / "fact_opex.csv", parse_dates=["month"])
    balances = pd.read_csv(DATA / "balances.csv", parse_dates=["month"])
    balances["category"] = balances["category"].fillna("")
    return Data(sales, budget, opex, balances, product)


def fy_months() -> pd.DatetimeIndex:
    return pd.date_range(f"{FY}-01-01", f"{FY}-12-01", freq="MS")


def ytd_mask(month: pd.Series) -> pd.Series:
    return (month >= pd.Timestamp(f"{FY}-01-01")) & (month <= AS_OF)


def trailing_mask(month: pd.Series) -> pd.Series:
    start = AS_OF - pd.DateOffset(months=TRAIL_MONTHS)
    return (month > start) & (month <= AS_OF)


def drivers(scenario: int, sens_price: float = 0.0, sens_volume: float = 0.0) -> dict:
    d = {k: v[scenario - 1] for k, v in SCENARIOS.items()}
    d["price"] += sens_price
    d["volume"] += sens_volume
    return d


def forecast_basis(data: Data) -> pd.DataFrame:
    """Per category: YTD volume run-rate and trailing per-case rates."""
    s, b = data.sales, data.budget
    ytd_a = s[ytd_mask(s["month"])].groupby("category")["units"].sum()
    ytd_b = b[ytd_mask(b["month"])].groupby("category")["units"].sum()
    tr = s[trailing_mask(s["month"])].groupby("category")[
        ["units", "net_revenue", "cogs", "freight"]].sum()
    out = pd.DataFrame(index=CATEGORIES)
    out["ytd_actual_units"] = ytd_a
    out["ytd_budget_units"] = ytd_b
    out["run_rate"] = ytd_a / ytd_b
    out["price0"] = tr["net_revenue"] / tr["units"]
    out["cost0"] = tr["cogs"] / tr["units"]
    out["freight0"] = tr["freight"] / tr["units"]
    return out


def opex_run_rate(data: Data) -> pd.Series:
    o = data.opex
    a = o[(o["scenario"] == "Actual") & ytd_mask(o["month"])].groupby("dept")["amount"].sum()
    b = o[(o["scenario"] == "Budget") & ytd_mask(o["month"])].groupby("dept")["amount"].sum()
    return (a / b).reindex(DEPARTMENTS)


def forecast(data: Data, scenario: int = 1, sens_price: float = 0.0,
             sens_volume: float = 0.0) -> dict[str, pd.DataFrame]:
    """Category x month forecast for open months, plus opex by department."""
    d = drivers(scenario, sens_price, sens_volume)
    basis = forecast_basis(data)
    months = fy_months()
    open_months = months[months > AS_OF]
    budget_units = (data.budget.groupby(["category", "month"])["units"].sum()
                    .unstack("month").reindex(index=CATEGORIES, columns=months).fillna(0))
    u0 = budget_units.mul(basis["run_rate"], axis=0)
    u0.loc[:, months <= AS_OF] = 0.0
    units = u0 * (1 + d["volume"])
    revenue = units.mul(basis["price0"] * (1 + d["price"]), axis=0)
    cogs = units.mul(basis["cost0"] * (1 + d["cost"]), axis=0)
    freight = units.mul(basis["freight0"] * (1 + d["freight"]), axis=0)

    o = data.opex
    ob = (o[o["scenario"] == "Budget"].pivot_table(index="dept", columns="month",
                                                    values="amount", aggfunc="sum")
          .reindex(index=DEPARTMENTS, columns=months).fillna(0))
    opex0 = ob.mul(opex_run_rate(data), axis=0)
    opex0.loc[:, months <= AS_OF] = 0.0
    opex = opex0 * (1 + d["opex"])
    return {"u0": u0, "units": units, "revenue": revenue, "cogs": cogs,
            "freight": freight, "opex0": opex0, "opex": opex, "basis": basis,
            "open_months": open_months}


def monthly_actuals(data: Data, year: int) -> pd.DataFrame:
    s = data.sales[data.sales["month"].dt.year == year]
    out = s.groupby("month")[["units", "net_revenue", "cogs", "freight"]].sum()
    o = data.opex[(data.opex["scenario"] == "Actual") & (data.opex["month"].dt.year == year)]
    opex = o.pivot_table(index="month", columns="dept", values="amount", aggfunc="sum")
    out = out.join(opex[DEPARTMENTS])
    return out.rename(columns={"net_revenue": "revenue"})


def pnl_lines(frame: pd.DataFrame) -> pd.DataFrame:
    """Add the derived P&L lines to a month-indexed frame of base lines."""
    f = frame.copy()
    f["gross_margin"] = f["revenue"] - f["cogs"]
    f["contribution"] = f["gross_margin"] - f["freight"]
    f["opex_total"] = f[DEPARTMENTS].sum(axis=1)
    f["ebitda"] = f["contribution"] - f["opex_total"]
    return f


def outlook(data: Data, scenario: int = 1, sens_price: float = 0.0,
            sens_volume: float = 0.0) -> pd.DataFrame:
    months = fy_months()
    act = monthly_actuals(data, FY).reindex(months).fillna(0)
    fc = forecast(data, scenario, sens_price, sens_volume)
    frame = pd.DataFrame(index=months)
    closed = months <= AS_OF
    frame["units"] = np.where(closed, act["units"], fc["units"].sum(axis=0))
    frame["revenue"] = np.where(closed, act["revenue"], fc["revenue"].sum(axis=0))
    frame["cogs"] = np.where(closed, act["cogs"], fc["cogs"].sum(axis=0))
    frame["freight"] = np.where(closed, act["freight"], fc["freight"].sum(axis=0))
    for dept in DEPARTMENTS:
        frame[dept] = np.where(closed, act[dept], fc["opex"].loc[dept])
    return pnl_lines(frame)


def budget_pnl(data: Data) -> pd.DataFrame:
    months = fy_months()
    b = data.budget.groupby("month")[["units", "revenue", "cogs", "freight"]].sum()
    o = data.opex[data.opex["scenario"] == "Budget"]
    opex = o.pivot_table(index="month", columns="dept", values="amount", aggfunc="sum")
    frame = b.join(opex[DEPARTMENTS]).reindex(months).fillna(0)
    return pnl_lines(frame)


def prior_year_pnl(data: Data) -> pd.DataFrame:
    return pnl_lines(monthly_actuals(data, FY - 1))


def pvm(data: Data, period: str) -> pd.DataFrame:
    """Price-volume-mix by category x channel for 'month' (as-of) or 'ytd'."""
    start = AS_OF if period == "month" else pd.Timestamp(f"{FY}-01-01")
    s = data.sales[(data.sales["month"] >= start) & (data.sales["month"] <= AS_OF)]
    b = data.budget[(data.budget["month"] >= start) & (data.budget["month"] <= AS_OF)]
    keys = ["category", "channel"]
    a = s.groupby(keys)[["units", "net_revenue", "cogs"]].sum().rename(
        columns={"units": "ua", "net_revenue": "ra", "cogs": "ca"})
    bb = b.groupby(keys)[["units", "revenue", "cogs"]].sum().rename(
        columns={"units": "ub", "revenue": "rb", "cogs": "cb"})
    idx = pd.MultiIndex.from_product([CATEGORIES, CHANNELS], names=keys)
    t = bb.join(a).reindex(idx)
    ub, ua = t["ub"].sum(), t["ua"].sum()
    budget_gm = (t["rb"] - t["cb"]).sum()
    m_avg = budget_gm / ub
    pb, pa = t["rb"] / t["ub"], t["ra"] / t["ua"]
    cb, ca = t["cb"] / t["ub"], t["ca"] / t["ua"]
    mb = pb - cb
    t["share_shift"] = t["ua"] / ua - t["ub"] / ub
    t["volume"] = (t["ua"] - t["ub"]) * m_avg
    t["mix"] = (t["ua"] - t["ub"]) * (mb - m_avg)
    t["price"] = t["ua"] * (pa - pb)
    t["cost"] = -t["ua"] * (ca - cb)
    t["effects"] = t[["volume", "mix", "price", "cost"]].sum(axis=1)
    t["gm_variance"] = (t["ra"] - t["ca"]) - (t["rb"] - t["cb"])
    return t


def ebitda_bridge_ytd(data: Data) -> dict[str, float]:
    t = pvm(data, "ytd")
    s = data.sales[ytd_mask(data.sales["month"])]
    b = data.budget[ytd_mask(data.budget["month"])]
    o = data.opex
    opex_a = o[(o["scenario"] == "Actual") & ytd_mask(o["month"])]["amount"].sum()
    opex_b = o[(o["scenario"] == "Budget") & ytd_mask(o["month"])]["amount"].sum()
    budget_ebitda = b["revenue"].sum() - b["cogs"].sum() - b["freight"].sum() - opex_b
    actual_ebitda = s["net_revenue"].sum() - s["cogs"].sum() - s["freight"].sum() - opex_a
    return {
        "budget_ebitda": budget_ebitda,
        "volume": t["volume"].sum(), "mix": t["mix"].sum(),
        "price": t["price"].sum(), "cost": t["cost"].sum(),
        "freight": -(s["freight"].sum() - b["freight"].sum()),
        "opex": -(opex_a - opex_b),
        "actual_ebitda": actual_ebitda,
    }


def scenario_table(data: Data) -> pd.DataFrame:
    rows = []
    for n, name in enumerate(SCENARIO_NAMES, start=1):
        o = outlook(data, n)
        rows.append({"scenario": name, "revenue": o["revenue"].sum(),
                     "gross_margin": o["gross_margin"].sum(), "ebitda": o["ebitda"].sum(),
                     "ebitda_margin": o["ebitda"].sum() / o["revenue"].sum()})
    return pd.DataFrame(rows).set_index("scenario")


def sensitivity_grid(data: Data, scenario: int = 1) -> pd.DataFrame:
    grid = pd.DataFrame(index=SENS_PRICE, columns=SENS_VOLUME, dtype=float)
    for p in SENS_PRICE:
        for v in SENS_VOLUME:
            grid.loc[p, v] = outlook(data, scenario, p, v)["ebitda"].sum()
    return grid


def tornado(data: Data, scenario: int = 1) -> dict[str, float]:
    """Change in FY EBITDA for +1 percentage point on each driver (linear, exact)."""
    d = drivers(scenario)
    fc = forecast(data, scenario)
    b = fc["basis"]
    units = fc["units"].sum(axis=1)
    u0 = fc["u0"].sum(axis=1)
    margin_per_case = (b["price0"] * (1 + d["price"]) - b["cost0"] * (1 + d["cost"])
                       - b["freight0"] * (1 + d["freight"]))
    return {
        "volume": float((u0 * 0.01 * margin_per_case).sum()),
        "price": float((units * b["price0"] * 0.01).sum()),
        "cost": float(-(units * b["cost0"] * 0.01).sum()),
        "freight": float(-(units * b["freight0"] * 0.01).sum()),
        "opex": float(-(fc["opex0"].sum().sum() * 0.01)),
    }


def break_even_uplift(data: Data, scenario: int = 1) -> float:
    """Extra price on open-month volume needed to reach budget FY EBITDA."""
    gap = budget_pnl(data)["ebitda"].sum() - outlook(data, scenario)["ebitda"].sum()
    fc = forecast(data, scenario)
    revenue_at_basis = (fc["units"].sum(axis=1) * fc["basis"]["price0"]).sum()
    return max(gap, 0.0) / revenue_at_basis


def break_even_volume(data: Data, scenario: int = 1) -> float:
    """Extra volume on open months, at today's margins, needed to reach budget FY EBITDA."""
    gap = budget_pnl(data)["ebitda"].sum() - outlook(data, scenario)["ebitda"].sum()
    return max(gap, 0.0) / (tornado(data, scenario)["volume"] / 0.01)


def working_capital_trend(data: Data, months: int = 13) -> pd.DataFrame:
    """Month-end balances and days metrics for the months to AS_OF."""
    rows = []
    for m in pd.date_range(end=AS_OF, periods=months, freq="MS"):
        s = data.sales[(data.sales["month"] > m - pd.DateOffset(months=TRAIL_MONTHS))
                       & (data.sales["month"] <= m)]
        b = data.balances[data.balances["month"] == m].groupby("account")["balance"].sum()
        rev, cogs = s["net_revenue"].sum(), s["cogs"].sum()
        ar, inv, ap = b["Accounts receivable"], b["Inventory"], b["Accounts payable"]
        dso, dio, dpo = ar / rev * WC_DAYS, inv / cogs * WC_DAYS, ap / cogs * WC_DAYS
        rows.append({"month": m, "ar": ar, "inventory": inv, "ap": ap, "nwc": ar + inv - ap,
                     "revenue_3m": rev, "cogs_3m": cogs, "dso": dso, "dio": dio, "dpo": dpo,
                     "ccc": dso + dio - dpo})
    return pd.DataFrame(rows).set_index("month")


def channel_economics(data: Data) -> pd.DataFrame:
    """Year-to-date cases, margin and contribution by channel, against budget."""
    s = data.sales[ytd_mask(data.sales["month"])]
    b = data.budget[ytd_mask(data.budget["month"])]
    a = s.groupby("channel")[["units", "net_revenue", "cogs", "freight"]].sum()
    bb = b.groupby("channel")[["units", "revenue", "cogs", "freight"]].sum()
    out = pd.DataFrame(index=CHANNELS)
    out["cases"] = a["units"]
    out["revenue"] = a["net_revenue"]
    out["gross_margin"] = a["net_revenue"] - a["cogs"]
    out["gm_pct"] = out["gross_margin"] / out["revenue"]
    out["freight"] = a["freight"]
    out["contribution"] = out["gross_margin"] - out["freight"]
    out["contribution_pct"] = out["contribution"] / out["revenue"]
    out["contribution_per_case"] = out["contribution"] / out["cases"]
    out["budget_cases"] = bb["units"]
    out["budget_contribution"] = bb["revenue"] - bb["cogs"] - bb["freight"]
    out["cases_vs_budget"] = (out["cases"] - out["budget_cases"]) / out["budget_cases"]
    out["contribution_vs_budget"] = out["contribution"] - out["budget_contribution"]
    return out


def working_capital(data: Data) -> dict:
    s, bal, prod = data.sales, data.balances, data.product
    tr = s[trailing_mask(s["month"])]
    revenue_3m = tr["net_revenue"].sum()
    cogs_3m = tr["cogs"].sum()
    at = bal[bal["month"] == AS_OF]
    ar = at.loc[at["account"] == "Accounts receivable", "balance"].sum()
    ap = at.loc[at["account"] == "Accounts payable", "balance"].sum()
    inv = at[at["account"] == "Inventory"].set_index("category")["balance"].reindex(CATEGORIES)
    cogs_cat = tr.groupby("category")["cogs"].sum().reindex(CATEGORIES)
    dio_cat = inv / cogs_cat * WC_DAYS
    shelf = prod.groupby("category")["shelf_life_days"].first().reindex(CATEGORIES)
    dso = ar / revenue_3m * WC_DAYS
    dio = inv.sum() / cogs_3m * WC_DAYS
    dpo = ap / cogs_3m * WC_DAYS
    excess = (dio_cat - SHELF_RISK * shelf).clip(lower=0) * cogs_cat / WC_DAYS
    return {"dso": dso, "dio": dio, "dpo": dpo, "ccc": dso + dio - dpo,
            "inventory": inv, "cogs_3m": cogs_cat,
            "dio_by_category": dio_cat, "shelf_life": shelf,
            "shelf_used": dio_cat / shelf, "at_risk": (dio_cat / shelf) > SHELF_RISK,
            "excess": excess}


def budget_accuracy(data: Data) -> dict[str, float]:
    """How well the budget called YTD volume, at category x month grain."""
    s = data.sales[ytd_mask(data.sales["month"])]
    b = data.budget[ytd_mask(data.budget["month"])]
    a = s.groupby(["category", "month"])["units"].sum()
    bb = b.groupby(["category", "month"])["units"].sum()
    return {"wape": float((a - bb).abs().sum() / a.sum()),
            "bias": float((bb - a).sum() / a.sum())}


if __name__ == "__main__":
    data = load()
    ytd = ebitda_bridge_ytd(data)
    print({k: round(v) for k, v in ytd.items()})
    print(scenario_table(data).round(0))
    print({k: round(v) for k, v in tornado(data).items()})
    print(f"break-even uplift {break_even_uplift(data):.2%}, or volume {break_even_volume(data):.1%}")
    wc = working_capital(data)
    print({k: round(float(v), 1) for k, v in wc.items() if not hasattr(v, "index")})
    print(wc["dio_by_category"].round(1).to_dict(), wc["at_risk"].to_dict())
    print(budget_accuracy(data))
