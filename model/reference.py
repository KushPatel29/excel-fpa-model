"""An independent recomputation of every figure the workbook publishes.

The workbook does its arithmetic in Excel: Power Query shapes the public
extracts, SUMIFS and LINEST do the analysis, what-if data tables run the
scenarios. This module does the same analysis in pandas and numpy from the
same files in data/, without reading the workbook, and the tests hold the two
to each other: money to the cent, ratios to 1e-9.

Definitions (the workbook prints the same ones):

* Price is revenue per MWh ($/MWh); the workbook also shows it as cents/kWh.
* Weather model, per customer class (residential, commercial): monthly MWh per
  customer = a + b1 x HDD + b2 x CDD + b3 x t, fitted by least squares on every
  month from January 2017 to December of the year before the plan year, where t
  is years since January 2017. Normal weather is the ten-year average of each
  calendar month's degree days before the plan year.
* Plan for year Y, built only from data through December Y-1: customers grow
  at the Y-1 rate; residential and commercial use per customer comes from the
  weather model at normal weather; industrial and transportation MWh are Y-1's
  same month grown at Y-1's rate; price is Y-1's same-month price grown at
  Y-1's average price change.
* Plan-versus-actual bridge, per class and month (exact by construction):
  customers (C_a - C_p) U_p P_p; weather C_a W P_p, where W is the model's
  weather effect on use per customer; usage C_a (U_a - U_p - W) P_p; price
  MWh_a (P_a - P_p). Industrial and transportation carry the whole volume
  difference in usage.
* Price-volume-mix on customer classes: volume (Q1 - Q0) x average P0; mix
  (Q1 - Q0) x (class P0 - average P0); price Q1 x (P1 - P0). The three add up
  to each class's revenue change.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

CLASSES = ["residential", "commercial", "industrial", "transportation"]
WEATHER_CLASSES = ["residential", "commercial"]
FIT_START = pd.Timestamp("2017-01-01")
NORMAL_YEARS = 10
COMPANY = 250                                   # FERC respondent id, Portland General Electric
PEERS = [250, 303, 162, 182, 216]

# Scenario levers (the workbook's Assumptions and Scenarios sheets)
WEATHER_SCENARIOS = {"Normal": (1.00, 1.00), "Mild": (0.88, 1.00), "Cold": (1.12, 1.00)}
SENS_RATE = (-0.02, -0.01, 0.0, 0.01, 0.02, 0.03, 0.04)
SENS_MW = (0, 50, 100, 150, 200, 250, 300)
TORNADO = {"rate": 0.01, "mw": 50.0, "hdd": -0.10, "cdd": 0.10, "customers": 0.01}

FERC_CLASSES = {"residential": "residential_sales", "commercial": "small_or_commercial",
                "industrial": "large_or_industrial", "lighting": "public_street_and_highway_lighting",
                "public authorities": "other_sales_to_public_authorities"}
POWER_COST_TYPES = ["fuel_steam_power_generation", "nuclear_fuel_expense", "fuel",
                    "fuel_other_renewable_generation", "purchased_power",
                    "power_purchased_for_storage_operations", "storage_fuel_energy_storage_expense"]
OM_FUNCTIONS = {"transmission": ["transmission_expenses"],
                "distribution": ["distribution_expenses"],
                "customer": ["customer_account_expenses", "customer_service_and_information_expenses",
                             "sales_expenses"],
                "administrative & general": ["administrative_and_general_expenses"]}
DA_TYPES = ["depreciation_expense", "depreciation_expense_for_asset_retirement_costs",
            "amortization_and_depletion_of_utility_plant", "amortization_of_other_utility_plant"]


@dataclass(frozen=True)
class Data:
    monthly: pd.DataFrame        # (date, class): revenue ($), mwh, customers
    status: pd.Series            # date -> EIA data status
    weather: pd.DataFrame        # date: hdd, cdd (observed months only)
    revenue: pd.DataFrame        # FERC sched 300, long
    expense: pd.DataFrame        # FERC sched 320, long
    income: pd.DataFrame         # FERC sched 114, long
    utilities: pd.DataFrame


def load() -> Data:
    wide = pd.read_csv(DATA / "eia_pge_monthly.csv")
    wide["date"] = pd.to_datetime(dict(year=wide["year"], month=wide["month"], day=1))
    parts = []
    for cls in CLASSES:
        parts.append(pd.DataFrame({"date": wide["date"], "cls": cls,
                                   "revenue": wide[f"{cls}_revenue_k"] * 1000.0,
                                   "mwh": wide[f"{cls}_mwh"],
                                   "customers": wide[f"{cls}_customers"]}))
    monthly = pd.concat(parts, ignore_index=True).set_index(["date", "cls"]).sort_index()
    status = wide.set_index("date")["data_status"]

    records = []
    for line in (DATA / "noaa_willamette_valley_degree_days.txt").read_text().splitlines():
        element = {"25": "hdd", "26": "cdd"}[line[4:6]]
        year = int(line[6:10])
        for k in range(12):
            value = float(line[10 + 7 * k:17 + 7 * k])
            if value > -9999:
                records.append((pd.Timestamp(year, k + 1, 1), element, value))
    weather = (pd.DataFrame(records, columns=["date", "element", "value"])
               .pivot(index="date", columns="element", values="value").dropna())
    return Data(monthly, status, weather,
                pd.read_csv(DATA / "ferc_revenue.csv"), pd.read_csv(DATA / "ferc_expense.csv"),
                pd.read_csv(DATA / "ferc_income.csv"), pd.read_csv(DATA / "utilities.csv"))


def as_of(data: Data) -> pd.Timestamp:
    return data.monthly.index.get_level_values("date").max()


def fiscal_year(data: Data) -> int:
    return as_of(data).year


def trend(date: pd.Timestamp) -> float:
    return (date.year - FIT_START.year) + (date.month - 1) / 12


# ---------------------------------------------------------------- weather
def normals(data: Data, year: int) -> pd.DataFrame:
    """Average degree days by calendar month over the ten years before `year`."""
    w = data.weather[(data.weather.index.year >= year - NORMAL_YEARS) & (data.weather.index.year < year)]
    return w.groupby(w.index.month).mean()


def regression(data: Data, cls: str, year: int) -> dict[str, float]:
    """Use per customer on HDD, CDD and a time trend, fitted through December year-1."""
    m = data.monthly.xs(cls, level="cls")
    m = m[(m.index >= FIT_START) & (m.index.year < year)]
    w = data.weather.reindex(m.index)
    y = (m["mwh"] / m["customers"]).to_numpy()
    t = np.array([trend(d) for d in m.index])
    X = np.column_stack([np.ones(len(m)), w["hdd"], w["cdd"], t])
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    fitted = X @ coef
    r2 = 1 - ((y - fitted) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    return {"intercept": coef[0], "hdd": coef[1], "cdd": coef[2], "trend": coef[3], "r2": r2, "n": len(m)}


def weather_effect(coef: dict, hdd: float, cdd: float, normal: pd.Series) -> float:
    return coef["hdd"] * (hdd - normal["hdd"]) + coef["cdd"] * (cdd - normal["cdd"])


def weather_impact(data: Data, year: int, months: int) -> pd.DataFrame:
    """MWh and revenue by which weather moved actual sales away from normal, per month."""
    out = []
    n = normals(data, year)
    for cls in WEATHER_CLASSES:
        coef = regression(data, cls, year)
        for k in range(1, months + 1):
            d = pd.Timestamp(year, k, 1)
            w = data.weather.loc[d]
            effect = weather_effect(coef, w["hdd"], w["cdd"], n.loc[k])
            c = data.monthly.loc[(d, cls)]
            out.append({"date": d, "cls": cls, "mwh": c["customers"] * effect,
                        "revenue": c["customers"] * effect * c["revenue"] / c["mwh"]})
    return pd.DataFrame(out)


# ------------------------------------------------------------------- plan
def plan(data: Data, year: int) -> pd.DataFrame:
    """Twelve months of plan per class, built only from data through December year-1."""
    n = normals(data, year)
    rows = []
    for cls in CLASSES:
        m = data.monthly.xs(cls, level="cls")
        c_base = m.loc[pd.Timestamp(year - 1, 12, 1), "customers"]
        c_growth = c_base / m.loc[pd.Timestamp(year - 2, 12, 1), "customers"]
        last, before = m[m.index.year == year - 1], m[m.index.year == year - 2]
        vol_growth = last["mwh"].sum() / before["mwh"].sum()
        price_growth = ((last["revenue"].sum() / last["mwh"].sum())
                        / (before["revenue"].sum() / before["mwh"].sum()))
        coef = regression(data, cls, year) if cls in WEATHER_CLASSES else None
        for k in range(1, 13):
            d = pd.Timestamp(year, k, 1)
            prior = m.loc[pd.Timestamp(year - 1, k, 1)]
            customers = c_base * c_growth ** (k / 12)
            if coef:
                use = (coef["intercept"] + coef["hdd"] * n.loc[k, "hdd"] + coef["cdd"] * n.loc[k, "cdd"]
                       + coef["trend"] * trend(d))
                mwh = customers * use
            else:
                mwh = prior["mwh"] * vol_growth
                use = mwh / customers
            p = prior["revenue"] / prior["mwh"] * price_growth
            rows.append({"date": d, "cls": cls, "customers": customers, "use": use, "mwh": mwh,
                         "price": p, "revenue": mwh * p})
    return pd.DataFrame(rows).set_index(["date", "cls"])


def closed_months(data: Data, year: int) -> int:
    end = as_of(data)
    return 12 if year < end.year else end.month


def plan_bridge_monthly(data: Data, year: int) -> pd.DataFrame:
    """Actual minus plan for each closed month and class of `year`, split four ways."""
    p = plan(data, year)
    n = normals(data, year)
    out = []
    for cls in CLASSES:
        coef = regression(data, cls, year) if cls in WEATHER_CLASSES else None
        for k in range(1, closed_months(data, year) + 1):
            d = pd.Timestamp(year, k, 1)
            a = data.monthly.loc[(d, cls)]
            pl = p.loc[(d, cls)]
            p_a = a["revenue"] / a["mwh"]
            eff = dict.fromkeys(["customers", "weather", "usage", "price"], 0.0)
            if coef:
                w = data.weather.loc[d]
                weff = weather_effect(coef, w["hdd"], w["cdd"], n.loc[k])
                u_a = a["mwh"] / a["customers"]
                eff["customers"] = (a["customers"] - pl["customers"]) * pl["use"] * pl["price"]
                eff["weather"] = a["customers"] * weff * pl["price"]
                eff["usage"] = a["customers"] * (u_a - pl["use"] - weff) * pl["price"]
            else:
                eff["usage"] = (a["mwh"] - pl["mwh"]) * pl["price"]
            eff["price"] = a["mwh"] * (p_a - pl["price"])
            out.append({"date": d, "cls": cls, "plan": pl["revenue"], **eff, "actual": a["revenue"]})
    return pd.DataFrame(out).set_index(["date", "cls"])


def plan_bridge(data: Data, year: int) -> pd.DataFrame:
    """Actual minus plan for the closed months of `year`, split four ways per class."""
    return plan_bridge_monthly(data, year).groupby(level="cls", sort=False).sum()


def backtest(data: Data, year: int) -> dict[str, float]:
    """How far the plan method missed, on the closed months of `year`."""
    p = plan(data, year)
    k = closed_months(data, year)
    idx = [(pd.Timestamp(year, m, 1), c) for m in range(1, k + 1) for c in CLASSES]
    a = data.monthly.loc[idx]
    pl = p.loc[idx]
    monthly_a = a["revenue"].groupby(level="date").sum()
    monthly_p = pl["revenue"].groupby(level="date").sum()
    mwh_a = a["mwh"].groupby(level="date").sum()
    mwh_p = pl["mwh"].groupby(level="date").sum()
    return {"months": k, "plan": pl["revenue"].sum(), "actual": a["revenue"].sum(),
            "variance": a["revenue"].sum() / pl["revenue"].sum() - 1,
            "wape_revenue": (monthly_a - monthly_p).abs().sum() / monthly_a.sum(),
            "wape_mwh": (mwh_a - mwh_p).abs().sum() / mwh_a.sum()}


# ---------------------------------------------------------------- forecast
def forecast(data: Data, weather: str = "Normal", mw: float = 0.0, rate: float = 0.0,
             customers: float = 0.0, hdd_scale: float = 1.0, cdd_scale: float = 1.0) -> pd.DataFrame:
    """The open months of the fiscal year, per class. Weather NOAA has already observed
    is used as observed; the rest is normal, scaled by the weather scenario."""
    year, closed = fiscal_year(data), as_of(data).month
    n = normals(data, year)
    hdd_s, cdd_s = WEATHER_SCENARIOS[weather]
    hdd_s, cdd_s = hdd_s * hdd_scale, cdd_s * cdd_scale
    rows = []
    for cls in CLASSES:
        m = data.monthly.xs(cls, level="cls")
        c_now = m.loc[as_of(data), "customers"]
        c_growth = c_now / m.loc[as_of(data) - pd.DateOffset(years=1), "customers"]
        ytd = m[(m.index.year == year) & (m.index.month <= closed)]
        ytd_prior = m[(m.index.year == year - 1) & (m.index.month <= closed)]
        vol_growth = ytd["mwh"].sum() / ytd_prior["mwh"].sum()
        price_growth = ((ytd["revenue"].sum() / ytd["mwh"].sum())
                        / (ytd_prior["revenue"].sum() / ytd_prior["mwh"].sum()))
        coef = regression(data, cls, year) if cls in WEATHER_CLASSES else None
        for k in range(closed + 1, 13):
            d = pd.Timestamp(year, k, 1)
            prior = m.loc[pd.Timestamp(year - 1, k, 1)]
            cust = c_now * c_growth ** ((k - closed) / 12) * (1 + customers)
            if d in data.weather.index:
                hdd, cdd = data.weather.loc[d, "hdd"], data.weather.loc[d, "cdd"]
            else:
                hdd, cdd = n.loc[k, "hdd"] * hdd_s, n.loc[k, "cdd"] * cdd_s
            if coef:
                use = coef["intercept"] + coef["hdd"] * hdd + coef["cdd"] * cdd + coef["trend"] * trend(d)
                mwh = cust * use
            else:
                mwh = prior["mwh"] * vol_growth * (1 + customers)
            p = prior["revenue"] / prior["mwh"] * price_growth * (1 + rate)
            if cls == "industrial":
                mwh += mw * 24 * d.days_in_month
            rows.append({"date": d, "cls": cls, "customers": cust, "hdd": hdd, "cdd": cdd, "mwh": mwh,
                         "price": p, "revenue": mwh * p})
    return pd.DataFrame(rows).set_index(["date", "cls"])


def outlook(data: Data, **levers) -> dict[str, float]:
    year = fiscal_year(data)
    actual = data.monthly[data.monthly.index.get_level_values("date").year == year]
    f = forecast(data, **levers)
    rev = actual["revenue"].sum() + f["revenue"].sum()
    mwh = actual["mwh"].sum() + f["mwh"].sum()
    return {"revenue": rev, "mwh": mwh, "price": rev / mwh}


def sensitivity_grid(data: Data) -> pd.DataFrame:
    grid = pd.DataFrame(index=SENS_RATE, columns=SENS_MW, dtype=float)
    for r in SENS_RATE:
        for mw in SENS_MW:
            grid.loc[r, mw] = outlook(data, rate=r, mw=mw)["revenue"]
    return grid


def tornado(data: Data) -> dict[str, float]:
    base = outlook(data)["revenue"]
    return {
        "rate": outlook(data, rate=TORNADO["rate"])["revenue"] - base,
        "mw": outlook(data, mw=TORNADO["mw"])["revenue"] - base,
        "hdd": outlook(data, hdd_scale=1 + TORNADO["hdd"])["revenue"] - base,
        "cdd": outlook(data, cdd_scale=1 + TORNADO["cdd"])["revenue"] - base,
        "customers": outlook(data, customers=TORNADO["customers"])["revenue"] - base,
    }


# ------------------------------------------------------------------- FERC
def ferc_wide(data: Data, table: str, utility: int) -> pd.DataFrame:
    frame = {"revenue": data.revenue, "expense": data.expense, "income": data.income}[table]
    column = {"revenue": "revenue_type", "expense": "expense_type", "income": "income_type"}[table]
    f = frame[frame["utility_id_ferc1"] == utility]
    return f.pivot_table(index="report_year", columns=column, values="dollar_value", aggfunc="sum").fillna(0)


def ferc_pnl(data: Data, utility: int = COMPANY) -> pd.DataFrame:
    r = ferc_wide(data, "revenue", utility)
    e = ferc_wide(data, "expense", utility)
    i = ferc_wide(data, "income", utility)
    rv = data.revenue[data.revenue["utility_id_ferc1"] == utility]
    mwh = rv.pivot_table(index="report_year", columns="revenue_type", values="sales_mwh", aggfunc="sum").fillna(0)
    cust = rv.pivot_table(index="report_year", columns="revenue_type", values="avg_customers_per_month",
                          aggfunc="sum").fillna(0)
    pl = pd.DataFrame(index=r.index)
    for label, col in FERC_CLASSES.items():
        pl[label] = r.get(col, 0.0)
    pl["retail"] = r["sales_to_ultimate_consumers"]
    pl["wholesale"] = r["sales_for_resale"]
    pl["refund provision"] = -r["provision_for_rate_refunds"]
    pl["other revenue"] = r["other_operating_revenues"]
    pl["revenue"] = r["electric_operating_revenues"]
    pl["power cost"] = e[[c for c in POWER_COST_TYPES if c in e]].sum(axis=1)
    pl["gross margin"] = pl["revenue"] - pl["power cost"]
    pl["production"] = e["power_production_expenses"] - pl["power cost"]
    for label, cols in OM_FUNCTIONS.items():
        pl[label] = e[[c for c in cols if c in e]].sum(axis=1)
    pl["other O&M"] = e["operations_and_maintenance_expenses_electric"] - pl["power cost"]
    pl["regional market & other"] = pl["other O&M"] - pl[["production", *OM_FUNCTIONS]].sum(axis=1)
    pl["D&A"] = i[[c for c in DA_TYPES if c in i]].sum(axis=1)
    pl["taxes other than income"] = i["taxes_other_than_income_taxes_utility_operating_income"]
    pl["operating income"] = pl["gross margin"] - pl["other O&M"] - pl["D&A"] - pl["taxes other than income"]
    pl["retail mwh"] = mwh["sales_to_ultimate_consumers"]
    pl["total mwh"] = mwh["sales_of_electricity"]
    pl["customers"] = cust["sales_to_ultimate_consumers"]
    for label, col in FERC_CLASSES.items():
        pl[f"{label} mwh"] = mwh.get(col, 0.0)
    return pl


def pvm_table(rows: dict) -> pd.DataFrame:
    t = pd.DataFrame(rows, index=["q0", "r0", "q1", "r1"]).T
    t["p0"], t["p1"] = t["r0"] / t["q0"], t["r1"] / t["q1"]
    avg0 = t["r0"].sum() / t["q0"].sum()
    t["volume"] = (t["q1"] - t["q0"]) * avg0
    t["mix"] = (t["q1"] - t["q0"]) * (t["p0"] - avg0)
    t["price"] = t["q1"] * (t["p1"] - t["p0"])
    t["change"] = t["r1"] - t["r0"]
    return t


def pvm_annual(data: Data, y0: int, y1: int, utility: int = COMPANY) -> pd.DataFrame:
    pl = ferc_pnl(data, utility)
    classes = [c for c in FERC_CLASSES if pl.loc[[y0, y1], f"{c} mwh"].sum() > 0]
    return pvm_table({c: (pl.loc[y0, f"{c} mwh"], pl.loc[y0, c], pl.loc[y1, f"{c} mwh"], pl.loc[y1, c])
                      for c in classes})


def pvm_ytd(data: Data) -> pd.DataFrame:
    year, closed = fiscal_year(data), as_of(data).month
    rows = {}
    for cls in CLASSES:
        m = data.monthly.xs(cls, level="cls")
        a = m[(m.index.year == year) & (m.index.month <= closed)]
        b = m[(m.index.year == year - 1) & (m.index.month <= closed)]
        rows[cls] = (b["mwh"].sum(), b["revenue"].sum(), a["mwh"].sum(), a["revenue"].sum())
    return pvm_table(rows)


def peers(data: Data, y0: int = 2019, y1: int = 2025) -> pd.DataFrame:
    out = {}
    for uid in PEERS:
        pl = ferc_pnl(data, uid)
        years = y1 - y0
        out[uid] = {
            "retail revenue": pl.loc[y1, "retail"],
            "retail mwh": pl.loc[y1, "retail mwh"],
            "customers": pl.loc[y1, "customers"],
            "retail price": pl.loc[y1, "retail"] / pl.loc[y1, "retail mwh"],
            "residential price": pl.loc[y1, "residential"] / pl.loc[y1, "residential mwh"],
            "industrial share": pl.loc[y1, "industrial mwh"] / pl.loc[y1, "retail mwh"],
            "retail mwh cagr": (pl.loc[y1, "retail mwh"] / pl.loc[y0, "retail mwh"]) ** (1 / years) - 1,
            "industrial mwh cagr": (pl.loc[y1, "industrial mwh"] / pl.loc[y0, "industrial mwh"]) ** (1 / years) - 1,
            "retail price cagr": ((pl.loc[y1, "retail"] / pl.loc[y1, "retail mwh"])
                                  / (pl.loc[y0, "retail"] / pl.loc[y0, "retail mwh"])) ** (1 / years) - 1,
            "power cost per mwh": pl.loc[y1, "power cost"] / pl.loc[y1, "total mwh"],
            "other om per customer": pl.loc[y1, "other O&M"] / pl.loc[y1, "customers"],
            "operating margin": pl.loc[y1, "operating income"] / pl.loc[y1, "revenue"],
        }
    return pd.DataFrame(out).T


def eia_ferc_tieout(data: Data) -> pd.DataFrame:
    """PGE's retail year, summed from EIA's monthly survey, against its FERC filing."""
    pl = ferc_pnl(data)
    m = data.monthly.groupby(data.monthly.index.get_level_values("date").year)[["revenue", "mwh"]].sum()
    years = [y for y in pl.index if y in m.index and closed_months(data, y) == 12]
    return pd.DataFrame({"eia_mwh": m.loc[years, "mwh"], "ferc_mwh": pl.loc[years, "retail mwh"],
                         "eia_revenue": m.loc[years, "revenue"], "ferc_revenue": pl.loc[years, "retail"]})


if __name__ == "__main__":
    data = load()
    fy = fiscal_year(data)
    print("as of", as_of(data).date(), "fiscal year", fy)
    for cls in WEATHER_CLASSES:
        print(cls, {k: round(v, 6) for k, v in regression(data, cls, fy).items()})
    print(plan_bridge(data, fy).round(0))
    for y in (fy - 2, fy - 1, fy):
        print("backtest", y, {k: round(v, 4) for k, v in backtest(data, y).items()})
    print("outlook", {k: round(v, 2) for k, v in outlook(data).items()})
    print("tornado", {k: round(v) for k, v in tornado(data).items()})
    print(pvm_annual(data, 2019, 2025)[["volume", "mix", "price", "change"]].round(0))
    print(pvm_ytd(data)[["volume", "mix", "price", "change"]].round(0))
    print((ferc_pnl(data).loc[[2019, 2024, 2025], ["revenue", "power cost", "gross margin", "other O&M", "D&A",
                                                    "operating income"]] / 1e6).round(1))
    print(peers(data).round(4))
    print(eia_ferc_tieout(data).round(0))
