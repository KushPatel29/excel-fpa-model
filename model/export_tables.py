"""Write the tables the Power BI report reads, from the reference model.

    python model/export_tables.py

The workbook and the report answer the same questions, so they must not do
the analysis twice. Excel does its arithmetic in formulas; the tests hold
every published cell to model/reference.py. This module writes that same
reference model's output as tidy tables in tables/, and the Power BI model
(powerbi/model_spec.py) only aggregates them. Two surfaces, one definition:
tests/test_powerbi_tables.py ties the tables back to the workbook's own cells.

Money is in dollars, energy in MWh, rates as fractions. Months are the first
day of the month.
"""
from __future__ import annotations

import calendar
from pathlib import Path

import pandas as pd

import reference as R

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tables"

CLASS_LABELS = {"residential": "Residential", "commercial": "Commercial", "industrial": "Industrial",
                "transportation": "Transportation"}
FERC_LABELS = {"residential": "Residential", "commercial": "Commercial", "industrial": "Industrial",
               "lighting": "Street lighting", "public authorities": "Public authorities"}
PEER_NAMES = {250: "PGE", 303: "PacifiCorp", 162: "Puget Sound Energy", 182: "Avista", 216: "Idaho Power"}
BASE_YEAR = 2019

# The FERC P&L as the workbook's PnL sheet lays it out: (line, section, reference column, sign)
PNL_LINES = [
    ("Retail revenue", "Revenue", "retail", 1),
    ("Wholesale", "Revenue", "wholesale", 1),
    ("Refund provision", "Revenue", "refund provision", 1),
    ("Other operating revenue", "Revenue", "other revenue", 1),
    ("Total operating revenue", "Revenue", "revenue", 1),
    ("Power cost", "Costs", "power cost", -1),
    ("Gross margin", "Margin", "gross margin", 1),
    ("Other O&M", "Costs", "other O&M", -1),
    ("Depreciation and amortization", "Costs", "D&A", -1),
    ("Taxes other than income", "Costs", "taxes other than income", -1),
    ("Operating income", "Margin", "operating income", 1),
]


def month_frame(first: pd.Timestamp, last: pd.Timestamp) -> pd.DataFrame:
    months = pd.date_range(first, last, freq="MS")
    return pd.DataFrame({
        "month": months,
        "month_index": range(len(months)),
        "year": months.year,
        "month_number": months.month,
        "month_name": [calendar.month_abbr[m] for m in months.month],
        "month_label": [d.strftime("%b %Y") for d in months],
        "quarter": [f"Q{(m - 1) // 3 + 1}" for m in months.month],
    })


def build(data: R.Data) -> dict[str, pd.DataFrame]:
    fy = R.fiscal_year(data)
    closed = R.as_of(data).month
    t: dict[str, pd.DataFrame] = {}

    t["dim_month"] = month_frame(R.FIT_START, pd.Timestamp(fy, 12, 1))
    t["dim_class"] = pd.DataFrame({
        "class": R.CLASSES,
        "class_label": [CLASS_LABELS[c] for c in R.CLASSES],
        "class_order": range(len(R.CLASSES)),
        "weather_model": ["yes" if c in R.WEATHER_CLASSES else "no" for c in R.CLASSES],
    })

    m = data.monthly.reset_index().rename(columns={"date": "month", "cls": "class"})
    m["data_status"] = m["month"].map(data.status)
    t["fact_sales"] = m[["month", "class", "revenue", "mwh", "customers", "data_status"]]

    w = data.weather.reindex(t["dim_month"]["month"])
    normal = [R.normals(data, d.year).loc[d.month] for d in w.index]
    t["fact_weather"] = pd.DataFrame({
        "month": w.index, "hdd": w["hdd"].to_numpy(), "cdd": w["cdd"].to_numpy(),
        "normal_hdd": [n["hdd"] for n in normal], "normal_cdd": [n["cdd"] for n in normal],
    })

    impact = R.weather_impact(data, fy, closed).rename(columns={"date": "month", "cls": "class"})
    t["weather_impact"] = impact.rename(columns={"mwh": "weather_mwh", "revenue": "weather_revenue"})

    coef = [{"class": c, **R.regression(data, c, fy)} for c in R.WEATHER_CLASSES]
    t["weather_model"] = pd.DataFrame(coef).rename(columns={
        "hdd": "per_hdd", "cdd": "per_cdd", "trend": "trend_per_year", "r2": "r_squared", "n": "months_fitted"})

    plan = R.plan(data, fy).reset_index().rename(columns={"date": "month", "cls": "class"})
    t["plan_monthly"] = plan.rename(columns={
        "customers": "plan_customers", "use": "plan_use", "mwh": "plan_mwh", "price": "plan_price",
        "revenue": "plan_revenue"})

    steps = [("Plan", "plan"), ("Customers", "customers"), ("Weather", "weather"), ("Usage", "usage"),
             ("Price", "price")]
    bridge = R.plan_bridge_monthly(data, fy).reset_index()
    t["plan_bridge"] = pd.concat([
        pd.DataFrame({"month": bridge["date"], "class": bridge["cls"], "step": label, "step_order": k,
                      "amount": bridge[col]})
        for k, (label, col) in enumerate(steps)], ignore_index=True)

    f = R.forecast(data).reset_index().rename(columns={"date": "month", "cls": "class"})
    f["days"] = f["month"].dt.days_in_month
    t["forecast_monthly"] = f.rename(columns={
        "customers": "forecast_customers", "mwh": "forecast_mwh", "price": "forecast_price",
        "revenue": "forecast_revenue"})[["month", "class", "forecast_customers", "hdd", "cdd", "forecast_mwh",
                                         "forecast_price", "forecast_revenue", "days"]]

    rows = []
    for year in (fy - 3, fy - 2, fy - 1, fy):
        b = R.backtest(data, year)
        rows.append({"plan_year": year, "months": b["months"], "plan_revenue": b["plan"],
                     "actual_revenue": b["actual"], "variance_pct": b["variance"],
                     "wape_revenue_pct": b["wape_revenue"], "wape_mwh_pct": b["wape_mwh"]})
    t["backtest"] = pd.DataFrame(rows)

    labels = {"rate": ("Rate +1 point", "+1 point"), "mw": ("Data-center load +50 MW", "+50 MW"),
              "hdd": ("HDD −10% (milder)", "−10%"), "cdd": ("CDD +10% (hotter)", "+10%"),
              "customers": ("Customers +1%", "+1%")}
    tor = R.tornado(data)
    t["tornado"] = pd.DataFrame([
        {"lever": labels[k][0], "step": labels[k][1], "lever_order": i, "impact": v}
        for i, (k, v) in enumerate(tor.items())])

    grid = R.sensitivity_grid(data)
    t["sensitivity"] = pd.DataFrame([
        {"rate_change": f"{r:+.0%}", "rate_order": i, "extra_mw": f"{mw} MW", "mw_order": j,
         "outlook_revenue": grid.loc[r, mw]}
        for i, r in enumerate(R.SENS_RATE) for j, mw in enumerate(R.SENS_MW)])

    pvm = []
    periods = [(f"{BASE_YEAR} to {fy - 1}", R.pvm_annual(data, BASE_YEAR, fy - 1), FERC_LABELS),
               (f"{fy - 2} to {fy - 1}", R.pvm_annual(data, fy - 2, fy - 1), FERC_LABELS),
               (f"Jan–{calendar.month_abbr[closed]} {fy} vs {fy - 1}", R.pvm_ytd(data), CLASS_LABELS)]
    for order, (period, table, names) in enumerate(periods):
        for cls, row in table.iterrows():
            for k, (step, value) in enumerate((("Opening revenue", row["r0"]), ("Volume", row["volume"]),
                                               ("Mix", row["mix"]), ("Price", row["price"]))):
                pvm.append({"period": period, "period_order": order, "revenue_class": names[cls],
                            "step": step, "step_order": k, "amount": value})
    t["pvm"] = pd.DataFrame(pvm)

    pl = R.ferc_pnl(data)
    t["ferc_pnl"] = pd.DataFrame([
        {"year": y, "line": line, "line_order": k, "section": section, "value": pl.loc[y, col]}
        for k, (line, section, col, _sign) in enumerate(PNL_LINES) for y in pl.index])
    t["ferc_years"] = pd.DataFrame({
        "year": pl.index, "year_label": [str(y) for y in pl.index], "retail_revenue": pl["retail"].to_numpy(), "total_revenue": pl["revenue"].to_numpy(),
        "power_cost": pl["power cost"].to_numpy(), "operating_income": pl["operating income"].to_numpy(),
        "retail_mwh": pl["retail mwh"].to_numpy(), "total_mwh": pl["total mwh"].to_numpy(),
        "customers": pl["customers"].to_numpy()})

    p = R.peers(data, BASE_YEAR, fy - 1)
    t["peers"] = pd.DataFrame({
        "utility": [PEER_NAMES[u] for u in p.index], "is_pge": [u == R.COMPANY for u in p.index],
        "retail_revenue": p["retail revenue"].to_numpy(), "retail_mwh": p["retail mwh"].to_numpy(),
        "customers": p["customers"].to_numpy(),
        "retail_cents": (p["retail price"] / 10).to_numpy(),
        "residential_cents": (p["residential price"] / 10).to_numpy(),
        "industrial_share_pct": p["industrial share"].to_numpy(),
        "retail_mwh_cagr_pct": p["retail mwh cagr"].to_numpy(),
        "industrial_mwh_cagr_pct": p["industrial mwh cagr"].to_numpy(),
        "retail_price_cagr_pct": p["retail price cagr"].to_numpy(),
        "power_cost_per_mwh": p["power cost per mwh"].to_numpy(),
        "other_om_per_customer": p["other om per customer"].to_numpy(),
        "operating_margin_pct": p["operating margin"].to_numpy()})

    tie = R.eia_ferc_tieout(data)
    t["tieout"] = pd.DataFrame({
        "year": tie.index, "year_label": [str(y) for y in tie.index], "eia_mwh": tie["eia_mwh"].to_numpy(), "ferc_mwh": tie["ferc_mwh"].to_numpy(),
        "eia_revenue": tie["eia_revenue"].to_numpy(), "ferc_revenue": tie["ferc_revenue"].to_numpy()})
    return t


def write(tables: dict[str, pd.DataFrame], out: Path = OUT) -> None:
    out.mkdir(parents=True, exist_ok=True)
    for name, frame in tables.items():
        frame = frame.copy()
        for col in frame.columns:
            if pd.api.types.is_datetime64_any_dtype(frame[col]):
                frame[col] = frame[col].dt.strftime("%Y-%m-%d")
        # 15 significant digits: well under a cent on a billion dollars even after
        # a measure sums a few hundred rows, without the last-ulp noise a
        # platform's libm puts in a 17-digit repr.
        frame.to_csv(out / f"{name}.csv", index=False, float_format="%.15g", lineterminator="\n")


if __name__ == "__main__":
    tables = build(R.load())
    write(tables)
    for name, frame in tables.items():
        print(f"tables/{name}.csv  {len(frame):>5} rows")
