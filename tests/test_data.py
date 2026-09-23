"""The dataset: reproducible to the byte, and carrying the story it claims to carry."""
from __future__ import annotations

import filecmp
import json
import sys

import pandas as pd
import pytest

import reference as R
from workbook_io import ROOT

sys.path.insert(0, str(ROOT / "generator"))
import generate_data as G  # noqa: E402

CSVS = ["dim_product.csv", "dim_channel.csv", "dim_region.csv", "fact_sales.csv",
        "budget_units_fy2026.csv", "budget_rates_fy2026.csv", "fact_opex.csv", "balances.csv"]


def test_regenerating_reproduces_the_committed_bytes(tmp_path, monkeypatch):
    monkeypatch.setattr(G, "DATA", tmp_path)
    G.main()
    for name in [*CSVS, "manifest.json"]:
        assert filecmp.cmp(tmp_path / name, ROOT / "data" / name, shallow=False), name


def test_manifest_counts_the_rows():
    manifest = json.loads((ROOT / "data" / "manifest.json").read_text(encoding="utf-8"))
    for table, rows in manifest["rows"].items():
        csv = ROOT / "data" / f"{table}.csv"
        assert sum(1 for _ in csv.open(encoding="utf-8")) - 1 == rows, table


def test_every_segment_every_month(data):
    s = data.sales
    assert len(s) == 32 * 24 * 4 * 3
    assert s.groupby(["month", "sku", "channel", "region"]).size().max() == 1
    assert (s["units"] > 0).all()


def test_budget_is_complete(data):
    assert len(data.budget) == 288 * 12
    assert data.budget[["net_price", "unit_cost", "freight_per_case"]].notna().all().all()


# ---- the planted story: the analysis has to find these, so the data must hold them
def ytd(frame, col="month"):
    return frame[R.ytd_mask(frame[col])]


def test_grocery_volume_runs_well_ahead_of_budget(data):
    a = ytd(data.sales).groupby("channel")["units"].sum()
    b = ytd(data.budget).groupby("channel")["units"].sum()
    assert 0.12 < a["Grocery Retail"] / b["Grocery Retail"] - 1 < 0.22
    assert abs(a["Restaurants"] / b["Restaurants"] - 1) < 0.05


def test_beef_costs_rise_faster_than_budget(data):
    s = data.sales[(data.sales["category"] == "Beef") & (data.sales["month"] >= "2026-05-01")]
    b = data.budget[(data.budget["category"] == "Beef") & (data.budget["month"] >= "2026-05-01")
                    & (data.budget["month"] <= R.AS_OF)]
    actual_cost = s["cogs"].sum() / s["units"].sum()
    budget_cost = b["cogs"].sum() / b["units"].sum()
    assert actual_cost / budget_cost - 1 > 0.05


def test_revenue_beats_budget_but_ebitda_misses(data):
    bridge = R.ebitda_bridge_ytd(data)
    revenue_gap = ytd(data.sales)["net_revenue"].sum() - ytd(data.budget)["revenue"].sum()
    assert revenue_gap > 0
    assert bridge["actual_ebitda"] < bridge["budget_ebitda"]


def test_warehouse_opex_over_budget(data):
    o = ytd(data.opex)
    by = o.groupby(["dept", "scenario"])["amount"].sum().unstack()
    ratio = by["Actual"] / by["Budget"]
    assert 1.05 < ratio["Warehouse & Logistics"] < 1.11
    assert ratio.idxmax() == "Warehouse & Logistics"


def test_cheese_inventory_creeps_toward_shelf_life(data):
    trend = R.working_capital_trend(data)
    assert trend.index[-1] == R.AS_OF
    wc = R.working_capital(data)
    assert wc["at_risk"]["Cheese & Dairy"] and wc["at_risk"].sum() == 1
    assert wc["shelf_used"]["Cheese & Dairy"] > 0.55
    cheese = R.working_capital_trend(data)   # the headline DIO rises with it
    assert cheese["dio"].iloc[-1] > cheese["dio"].iloc[0]


# ---- identities the reference model must honour on its own
@pytest.mark.parametrize("period", ["month", "ytd"])
def test_pvm_adds_up_in_every_segment(data, period):
    t = R.pvm(data, period)
    assert (t["effects"] - t["gm_variance"]).abs().max() < 1e-6


def test_tornado_is_exact_against_the_sensitivity_grid(data):
    grid = R.sensitivity_grid(data)
    t = R.tornado(data)
    assert grid.loc[0.01, 0.0] - grid.loc[0.0, 0.0] == pytest.approx(t["price"], abs=1e-6)
    assert grid.loc[0.0, 0.02] - grid.loc[0.0, 0.0] == pytest.approx(2 * t["volume"], abs=1e-6)


def test_break_even_price_closes_the_gap(data):
    """Push price up by the break-even uplift and the outlook lands on budget."""
    uplift = R.break_even_uplift(data)
    budget = R.budget_pnl(data)["ebitda"].sum()
    assert R.outlook(data, 1, sens_price=uplift)["ebitda"].sum() == pytest.approx(budget, abs=1e-6)


def test_outlook_is_actuals_then_forecast(data):
    o = R.outlook(data)
    act = R.monthly_actuals(data, R.FY)
    closed = o.index <= R.AS_OF
    assert (o.loc[closed, "revenue"] - act["revenue"]).abs().max() < 1e-6
    assert (o.loc[~closed, "revenue"] > 0).all()
    assert isinstance(o.index, pd.DatetimeIndex)
