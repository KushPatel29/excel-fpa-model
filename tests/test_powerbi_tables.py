"""The Power BI report and the Excel workbook publish one set of numbers.

The report reads tables/ -- written by model/export_tables.py from the same
reference model the workbook is held to -- and only sums them. So the tests
here do two things: prove the committed tables are what the reference model
writes today, and read the *workbook's own cells* against them. A figure on a
Power BI page and the same figure in Excel cannot drift apart without one of
these failing.

They also check the one piece of arithmetic the report does itself: the
Scenario outlook measure, whose DAX is replayed here in pandas and held to the
sensitivity grid the reference model computed by re-running the forecast.
"""
from __future__ import annotations

import pandas as pd
import pytest

import export_tables as X
import layout as L
import reference as R
from workbook_io import ROOT, block, cell, column

CENT = 0.01
TABLES = ROOT / "tables"


@pytest.fixture(scope="module")
def fresh(data):
    return X.build(data)


def read(name: str) -> pd.DataFrame:
    return pd.read_csv(TABLES / f"{name}.csv")


@pytest.mark.parametrize("name", sorted(X.build(R.load())))
def test_committed_table_is_what_the_reference_model_writes(fresh, name, tmp_path):
    """Regenerate every table and compare. Floats are written to 15 significant
    digits, so the comparison is on values, not bytes."""
    X.write({name: fresh[name]}, tmp_path)
    want = pd.read_csv(tmp_path / f"{name}.csv")
    got = read(name)
    assert list(got.columns) == list(want.columns)
    assert len(got) == len(want)
    pd.testing.assert_frame_equal(got, want, check_exact=False, rtol=1e-11, atol=CENT)


def test_every_table_the_model_reads_is_committed():
    from powerbi.model_spec import TABLES as MODEL_TABLES
    missing = [t for t, meta in MODEL_TABLES.items() if not (ROOT / meta["source"] / f"{t}.csv").exists()]
    assert not missing


# ------------------------------------------------ the workbook's own cells
def test_bridge_matches_the_workbook(values):
    b = read("plan_bridge")
    by_step = b.groupby("step")["amount"].sum()
    total = block(values, "plan_bridge_total")[0]
    # plan, customers, weather, usage, price, actual, variance
    assert [by_step[s] for s in ("Plan", "Customers", "Weather", "Usage", "Price")] == pytest.approx(
        total[:5], abs=CENT)
    assert by_step.sum() == pytest.approx(total[5], abs=CENT), "the waterfall's total bar is the actual"


def test_bridge_by_class_matches_the_workbook(values):
    b = read("plan_bridge").pivot_table(index="class", columns="step", values="amount", aggfunc="sum")
    for row, cls in zip(block(values, "plan_bridge"), L.CLASSES):
        r = b.loc[cls]
        assert [r["Plan"], r["Customers"], r["Weather"], r["Usage"], r["Price"]] == pytest.approx(row[:5], abs=CENT)


def test_actual_revenue_to_date_matches_the_workbook(values):
    sales = read("fact_sales")
    fy_rows = sales[sales["month"].str.startswith(str(cell(values, "FiscalYear")))]
    assert fy_rows["revenue"].sum() == pytest.approx(block(values, "plan_bridge_total")[0][5], abs=CENT)


def test_outlook_matches_the_workbook(values):
    sales, forecast = read("fact_sales"), read("forecast_monthly")
    fy = str(cell(values, "FiscalYear"))
    outlook = sales.loc[sales["month"].str.startswith(fy), "revenue"].sum() + forecast["forecast_revenue"].sum()
    assert outlook == pytest.approx(cell(values, "fy_revenue"), abs=CENT)
    assert outlook == pytest.approx(cell(values, "kpi_fy_revenue"), abs=CENT)


def test_plan_matches_the_workbook(values):
    plan = read("plan_monthly")
    total = sum(row[8] for row in block(values, "plan_long"))
    assert plan["plan_revenue"].sum() == pytest.approx(total, abs=CENT)


def test_weather_impact_matches_the_workbook(values):
    wi = read("weather_impact")
    total = block(values, "weather_impact_total")[0]
    assert wi["weather_mwh"].sum() == pytest.approx(total[10], abs=CENT)
    assert wi["weather_revenue"].sum() == pytest.approx(total[11], abs=CENT)
    assert wi["weather_mwh"].sum() / 1000 == pytest.approx(cell(values, "kpi_weather_gwh"), abs=1e-6)


def test_weather_model_matches_linest(values):
    wm = read("weather_model").set_index("class")
    for cls in R.WEATHER_CLASSES:
        got = block(values, f"coef_{cls}")[0]         # intercept, per HDD, per CDD, trend, R²
        r = wm.loc[cls]
        assert [r["intercept"], r["per_hdd"], r["per_cdd"], r["trend_per_year"], r["r_squared"]] == pytest.approx(
            got[:5], abs=1e-9)


def test_tornado_matches_the_workbook(values):
    assert read("tornado")["impact"].tolist() == pytest.approx(column(values, "tornado"), abs=CENT)


def test_sensitivity_matches_the_workbook(values):
    s = read("sensitivity").pivot_table(index="rate_order", columns="mw_order", values="outlook_revenue")
    for got, row in zip(s.to_numpy().tolist(), block(values, "sens_grid")):
        assert got == pytest.approx(row, abs=CENT)


def test_backtest_matches_the_workbook(values):
    bt = read("backtest")
    rows = block(values, "backtest")
    assert bt["plan_year"].tolist() == [r[0] for r in rows]
    for (_, b), row in zip(bt.iterrows(), rows):
        assert [b["plan_revenue"], b["actual_revenue"]] == pytest.approx(row[2:4], abs=CENT)


def test_pvm_matches_the_workbook(values):
    pvm = read("pvm")
    long_view = pvm[pvm["period_order"] == 0].groupby("step")["amount"].sum()
    bridge = column(values, "bridge_long")       # opening, volume, mix, price, closing
    assert [long_view[s] for s in ("Opening revenue", "Volume", "Mix", "Price")] == pytest.approx(
        bridge[:4], abs=CENT)
    assert long_view.sum() == pytest.approx(bridge[4], abs=CENT), "the waterfall's total bar is 2025 revenue"


def test_ferc_pnl_matches_the_workbook(values):
    pnl = read("ferc_pnl").pivot_table(index="line", columns="year", values="value")
    for line, name in (("Retail revenue", "pnl_retail"), ("Total operating revenue", "pnl_revenue"),
                       ("Power cost", "pnl_power"), ("Operating income", "pnl_op_income")):
        assert pnl.loc[line].tolist() == pytest.approx(block(values, name)[0], abs=CENT), line


def test_tieout_matches_the_workbook(values):
    t = read("tieout")
    rows = block(values, "tieout")
    assert t["year"].tolist() == [r[0] for r in rows]
    for (_, r), row in zip(t.iterrows(), rows):
        assert [r["eia_mwh"], r["ferc_mwh"]] == pytest.approx(row[1:3], abs=CENT)


def test_peers_match_the_workbook(values):
    p = read("peers")
    assert p["industrial_mwh_cagr_pct"].tolist() == pytest.approx(block(values, "peers_table")[7], abs=1e-9)


# ------------------------------------------- the report's own arithmetic
def scenario_outlook(rate: float, mw: float) -> float:
    """The Scenario outlook measure's DAX, replayed over the same tables."""
    sales, f = read("fact_sales"), read("forecast_monthly")
    fy = f["month"].str[:4].iloc[0]
    actual = sales.loc[sales["month"].str.startswith(fy), "revenue"].sum()
    industrial = f[f["class"] == "industrial"]
    per_mw = (industrial["days"] * 24 * industrial["forecast_price"]).sum()
    return actual + (f["forecast_revenue"].sum() + mw * per_mw) * (1 + rate)


def test_the_scenario_measure_reproduces_the_sensitivity_grid(data):
    """The what-if sliders drive a DAX measure; the grid was computed by
    re-running the whole forecast. They must agree in every cell."""
    grid = R.sensitivity_grid(data)
    for rate in R.SENS_RATE:
        for mw in R.SENS_MW:
            assert scenario_outlook(rate, mw) == pytest.approx(grid.loc[rate, mw], abs=CENT), (rate, mw)


def test_the_scenario_measure_is_the_dax_in_the_model():
    """Replaying a formula proves nothing if the model's formula is different."""
    from powerbi.model_spec import MEASURES
    dax = next(d for name, d, *_ in MEASURES if name == "Scenario outlook")
    for piece in ('forecast_monthly[class] = "industrial"', "forecast_monthly[days] * 24 * forecast_monthly[forecast_price]",
                  "[Actual this year] + ([Forecast revenue] + vMW * vIndustrial) * (1 + vRate)"):
        assert piece in dax
