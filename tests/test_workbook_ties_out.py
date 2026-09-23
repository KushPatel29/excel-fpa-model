"""Every figure the workbook publishes, held to an independent pandas model.

The workbook computes with SUMIFS, Power Query, DAX and what-if data tables;
model/reference.py recomputes the same analysis from the same CSVs without
reading the workbook. Money must agree to the cent, ratios to 1e-9.
"""
from __future__ import annotations

import csv
from datetime import datetime

import pandas as pd
import pytest

import layout as L
import reference as R
from workbook_io import ROOT, block, cell, column, money, signmoney  # noqa: F401

CENT = 0.01
RATIO = 1e-9

LINES = [k for k, *_ in L.PNL_LINES]
KIND = {k: kind for k, _, _, kind in L.PNL_LINES}
IS_COST = {k: c for k, _, c, _ in L.PNL_LINES}
DEPT = dict(zip(["opex_sm", "opex_wl", "opex_ga", "opex_tech"], L.DEPARTMENTS))
MONTHS = R.fy_months()
CLOSED = MONTHS <= R.AS_OF


def tol(key: str) -> float:
    return RATIO if KIND[key] == "%" else CENT


def line(frame: pd.DataFrame, key: str) -> pd.Series:
    if key in DEPT:
        return frame[DEPT[key]]
    if key == "gm_pct":
        return frame["gross_margin"] / frame["revenue"]
    if key == "ebitda_pct":
        return frame["ebitda"] / frame["revenue"]
    return frame[{"cases": "units", "gm": "gross_margin"}.get(key, key)]


def favvar(actual, plan, key):
    if KIND[key] == "$" and IS_COST[key]:
        return plan - actual
    return actual - plan


@pytest.fixture(scope="module")
def ref(data):
    return {"outlook": R.outlook(data), "budget": R.budget_pnl(data), "py": R.prior_year_pnl(data)}


def fy_value(frame: pd.DataFrame, key: str) -> float:
    if key == "gm_pct":
        return frame["gross_margin"].sum() / frame["revenue"].sum()
    if key == "ebitda_pct":
        return frame["ebitda"].sum() / frame["revenue"].sum()
    return float(line(frame, key).sum())


def ytd_value(frame: pd.DataFrame, key: str) -> float:
    return fy_value(frame.iloc[: int(CLOSED.sum())], key)


# ---------------------------------------------------------------------- P&L
@pytest.mark.parametrize("which", ["outlook", "budget", "py"])
@pytest.mark.parametrize("key", LINES)
def test_pnl_month_by_month(values, ref, which, key):
    got = block(values, f"pnl_{which}")[LINES.index(key)]
    want = line(ref[which], key).to_numpy()
    assert got == pytest.approx(list(want), abs=tol(key)), f"{which} {key}"


@pytest.mark.parametrize("key", LINES)
def test_pnl_variance_block(values, ref, key):
    got = block(values, "pnl_variance")[LINES.index(key)]
    want = favvar(line(ref["outlook"], key), line(ref["budget"], key), key)
    assert got == pytest.approx(list(want), abs=tol(key))


@pytest.mark.parametrize("which", ["outlook", "budget", "py"])
def test_pnl_full_year(values, ref, which):
    got = [row[0] for row in block(values, f"pnl_fy_{which}")]
    for key, value in zip(LINES, got):
        assert value == pytest.approx(fy_value(ref[which], key), abs=tol(key)), key


@pytest.mark.parametrize("key", LINES)
def test_pnl_summary_columns(values, ref, key):
    """FY vs budget and PY, and year to date, for every outlook line."""
    row = block(values, "pnl_summary")[LINES.index(key)]
    p, q, r, s, t, u, _, w, x, y, z, aa, ab = row
    o, b, py = ref["outlook"], ref["budget"], ref["py"]
    fy_o, fy_b, fy_py = fy_value(o, key), fy_value(b, key), fy_value(py, key)
    ytd_o, ytd_b, ytd_py = ytd_value(o, key), ytd_value(b, key), ytd_value(py, key)
    t_ = tol(key)
    assert q == pytest.approx(fy_b, abs=t_)
    assert t == pytest.approx(fy_py, abs=t_)
    assert w == pytest.approx(ytd_o, abs=t_)
    assert x == pytest.approx(ytd_b, abs=t_)
    assert aa == pytest.approx(ytd_py, abs=t_)
    if KIND[key] == "%":
        assert r == pytest.approx(fy_o - fy_b, abs=RATIO)
        assert y == pytest.approx(ytd_o - ytd_b, abs=RATIO)
        assert u == pytest.approx(fy_o - fy_py, abs=RATIO)
        assert ab == pytest.approx(ytd_o - ytd_py, abs=RATIO)
        assert s is None and z is None
    else:
        assert r == pytest.approx(favvar(fy_o, fy_b, key), abs=CENT)
        assert y == pytest.approx(favvar(ytd_o, ytd_b, key), abs=CENT)
        assert s == pytest.approx(favvar(fy_o, fy_b, key) / abs(fy_b), abs=RATIO)
        assert z == pytest.approx(favvar(ytd_o, ytd_b, key) / abs(ytd_b), abs=RATIO)
        assert u == pytest.approx((fy_o - fy_py) / abs(fy_py), abs=RATIO)
        assert ab == pytest.approx((ytd_o - ytd_py) / abs(ytd_py), abs=RATIO)


def test_outlook_closed_months_are_the_ledger(values, data):
    revenue = block(values, "pnl_outlook")[LINES.index("revenue")]
    s = data.sales[data.sales["month"].dt.year == R.FY].groupby("month")["net_revenue"].sum()
    for month, got in zip(MONTHS[CLOSED], revenue):
        assert got == pytest.approx(s[month], abs=CENT)


# ------------------------------------------------------------------ Forecast
def test_forecast_basis(values, data):
    basis = R.forecast_basis(data)
    s = data.sales[R.trailing_mask(data.sales["month"])].groupby("category")[
        ["units", "net_revenue", "cogs", "freight"]].sum()
    d = R.drivers(1)
    for row in block(values, "fc_basis"):
        cat = row[0]
        b = basis.loc[cat]
        want = [b["ytd_actual_units"], b["ytd_budget_units"], b["run_rate"],
                s.loc[cat, "units"], s.loc[cat, "net_revenue"], s.loc[cat, "cogs"], s.loc[cat, "freight"],
                b["price0"], b["cost0"], b["freight0"],
                b["price0"] * (1 + d["price"]) - b["cost0"] * (1 + d["cost"]) - b["freight0"] * (1 + d["freight"])]
        assert row[1:] == pytest.approx(want, rel=1e-12, abs=1e-9), cat


@pytest.mark.parametrize("key", ["u0", "units", "revenue", "cogs", "freight", "opex0", "opex"])
def test_forecast_blocks(values, data, key):
    want = R.forecast(data)[key]
    got = block(values, f"fc_{key}")
    assert len(got) == len(want)
    for row, (_, expected) in zip(got, want.iterrows()):
        assert row == pytest.approx(list(expected), abs=CENT)


def test_budget_accuracy(values, data):
    acc = R.budget_accuracy(data)
    assert cell(values, "budget_wape") == pytest.approx(acc["wape"], abs=RATIO)
    assert cell(values, "budget_bias") == pytest.approx(acc["bias"], abs=RATIO)


def test_statistical_cross_check_is_numeric_and_near_the_driver_forecast(values, data):
    """FORECAST.ETS has no pandas twin, so hold it to a sanity band, not a value."""
    ets = [v for v in block(values, "fc_ets")[0] if v not in (None, "")]
    driver = R.forecast(data)["revenue"].sum(axis=0)
    driver = driver[driver > 0]
    assert len(ets) == len(driver) == 12 - int(CLOSED.sum())
    for e, d in zip(ets, driver):
        assert isinstance(e, float) and 0.8 * d < e < 1.2 * d
    assert abs(cell(values, "fc_ets_gap")) < 0.2


# ----------------------------------------------------------------------- PVM
EFFECTS = ["volume", "mix", "price", "cost"]
DRIVER_NAME = dict(zip(EFFECTS, ["Volume", "Mix", "Price", "Unit cost"]))


@pytest.mark.parametrize("period", ["month", "ytd"])
def test_pvm_every_segment(values, data, period):
    t = R.pvm(data, period)
    rows = block(values, f"pvm_{period}")
    assert len(rows) == len(t) == 24
    for row, ((cat, chan), r) in zip(rows, t.iterrows()):
        assert (row[0], row[1]) == (cat, chan)
        want = [r["ub"], r["rb"], r["cb"], r["ua"], r["ra"], r["ca"],
                r["rb"] / r["ub"], r["ra"] / r["ua"], r["cb"] / r["ub"], r["ca"] / r["ua"],
                r["rb"] / r["ub"] - r["cb"] / r["ub"], r["share_shift"],
                r["volume"], r["mix"], r["price"], r["cost"], r["effects"], r["gm_variance"], 0.0]
        assert row[2:21] == pytest.approx(want, abs=CENT), (cat, chan)
        biggest = max(EFFECTS, key=lambda e: abs(r[e]))
        assert row[21] == DRIVER_NAME[biggest]


@pytest.mark.parametrize("period", ["month", "ytd"])
def test_pvm_effects_reconcile_in_every_segment(data, period):
    t = R.pvm(data, period)
    assert (t["effects"] - t["gm_variance"]).abs().max() < 1e-6


def test_ebitda_bridge(values, data):
    b = R.ebitda_bridge_ytd(data)
    want = [b[k] for k in ("budget_ebitda", "volume", "mix", "price", "cost", "freight", "opex",
                           "actual_ebitda")]
    assert column(values, "bridge_ebitda") == pytest.approx(want, abs=CENT)
    assert want[0] + sum(want[1:-1]) == pytest.approx(want[-1], abs=1e-6)


def test_gross_margin_bridge(values, data):
    t = R.pvm(data, "ytd")
    want = [(t["rb"] - t["cb"]).sum(), *(t[e].sum() for e in EFFECTS), (t["ra"] - t["ca"]).sum()]
    assert column(values, "bridge_gm") == pytest.approx(want, abs=CENT)


@pytest.mark.parametrize("level", ["category", "channel"])
def test_pvm_summaries(values, data, level):
    t = R.pvm(data, "ytd").groupby(level=level)
    effects = t[EFFECTS].sum()
    budget_gm = (t["rb"].sum() - t["cb"].sum())
    actual_gm = (t["ra"].sum() - t["ca"].sum())
    for row in block(values, f"pvm_by_{level}"):
        name = row[0]
        want = [*effects.loc[name], effects.loc[name].sum(), budget_gm[name], actual_gm[name],
                (actual_gm[name] - budget_gm[name]) / abs(budget_gm[name])]
        assert row[1:] == pytest.approx(want, abs=CENT), name


# ------------------------------------------------------------------ Scenarios
def test_scenario_table(values, data):
    table = R.scenario_table(data)
    budget = R.budget_pnl(data)["ebitda"].sum()
    for row, name in zip(block(values, "scn_table"), L.SCENARIO_NAMES):
        r = table.loc[name]
        want = [r["revenue"], r["gross_margin"], r["ebitda"], r["ebitda_margin"], r["ebitda"] - budget]
        assert row == pytest.approx(want, abs=CENT), name


def test_scenarios_are_ordered(values):
    ebitda = [row[2] for row in block(values, "scn_table")]
    base, upside, downside = ebitda
    assert downside < base < upside


def test_sensitivity_grid(values, data):
    want = R.sensitivity_grid(data)
    got = block(values, "sens_grid")
    assert column(values, "sens_price_axis") == pytest.approx(list(L.SENS_PRICE))
    assert block(values, "sens_volume_axis")[0] == pytest.approx(list(L.SENS_VOLUME))
    for row, (_, expected) in zip(got, want.iterrows()):
        assert row == pytest.approx(list(expected), abs=CENT)


def test_tornado(values, data):
    t = R.tornado(data)
    assert column(values, "tornado") == pytest.approx([t[k] for k in ("volume", "price", "cost", "freight",
                                                                      "opex")], abs=CENT)


def test_break_even(values, data):
    gap = R.budget_pnl(data)["ebitda"].sum() - R.outlook(data)["ebitda"].sum()
    assert cell(values, "breakeven_gap") == pytest.approx(gap, abs=CENT)
    assert cell(values, "breakeven_uplift") == pytest.approx(R.break_even_uplift(data), abs=RATIO)
    assert cell(values, "breakeven_volume") == pytest.approx(R.break_even_volume(data), abs=RATIO)


# ------------------------------------------------------------ Working capital
WC_LINES = ["ar", "inventory", "ap", "nwc", "revenue_3m", "cogs_3m", "dso", "dio", "dpo", "ccc"]


def test_working_capital_trend(values, data):
    trend = R.working_capital_trend(data)
    months = block(values, "wc_months")[0]
    assert [m.date() for m in months] == [m.date() for m in trend.index]
    for key, row in zip(WC_LINES, block(values, "wc_lines")):
        assert row == pytest.approx(list(trend[key]), abs=CENT), key


def test_working_capital_by_category(values, data):
    wc = R.working_capital(data)
    for row, cat in zip(block(values, "wc_category"), L.CATEGORIES):
        want = [wc["inventory"][cat], wc["cogs_3m"][cat], wc["dio_by_category"][cat], wc["shelf_life"][cat],
                wc["shelf_used"][cat]]
        assert row[:5] == pytest.approx(want, abs=CENT), cat
        assert row[5] == ("At risk" if wc["at_risk"][cat] else "OK")
        assert row[6] == pytest.approx(wc["excess"][cat], abs=CENT)


def test_cheese_is_the_stock_at_risk(values):
    at_risk = [row for row in block(values, "wc_category") if row[5] == "At risk"]
    assert len(at_risk) == 1 and at_risk[0][6] > 0


# ------------------------------------------------------------------- Channels
def test_channel_economics_by_formula(values, data):
    econ = R.channel_economics(data)
    for row, chan in zip(block(values, "ch_formulas"), L.CHANNELS):
        e = econ.loc[chan]
        want = [e["cases"], e["revenue"], e["gross_margin"], e["gm_pct"], e["freight"], e["contribution"],
                e["contribution_pct"], e["contribution_per_case"], e["budget_cases"], e["budget_contribution"],
                e["cases_vs_budget"], e["contribution_vs_budget"], e["revenue"]]
        assert row[:13] == pytest.approx(want, abs=CENT), chan
        assert row[13] == "Yes"


def test_cube_functions_read_the_data_model(values, data):
    """CUBEVALUE results come from the DAX engine, not the formula engine."""
    s = data.sales
    by = s.groupby([s["channel"], s["month"].dt.year])["net_revenue"].sum()
    rows = block(values, "ch_cube_years")
    for row, chan in zip(rows, [*L.CHANNELS, None]):
        for year, got in zip((R.FY - 2, R.FY - 1, R.FY), row):
            want = by.xs(year, level=1).sum() if chan is None else by[(chan, year)]
            assert got == pytest.approx(want, abs=CENT), (chan, year)
    ytd = R.ytd_mask(s["month"])
    py = (s["month"] >= pd.Timestamp(f"{R.FY - 1}-01-01")) & (s["month"] <= R.AS_OF - pd.DateOffset(years=1))
    for row, chan in zip(block(values, "ch_cube_yoy"), [*L.CHANNELS, None]):
        pick = s["channel"] == chan if chan else s["channel"].notna()
        now, before = s[ytd & pick]["net_revenue"].sum(), s[py & pick]["net_revenue"].sum()
        assert row[0] == pytest.approx(now / before - 1, abs=RATIO)
        assert row[1] == pytest.approx(now / before - 1, abs=RATIO)
        assert row[2] == "Yes"


# --------------------------------------------------------------------- Checks
def test_every_check_passes(values):
    status = column(values, "checks_status")
    assert len(status) == 31
    assert status == ["PASS"] * len(status)
    assert cell(values, "checks_summary") == f"All {len(status)} checks pass"


# ---------------------------------------------------------- Dashboard, Cover
def test_dashboard_tiles(values, data, ref):
    o = ref["outlook"]
    assert cell(values, "kpi_revenue_ytd") == pytest.approx(ytd_value(o, "revenue"), abs=CENT)
    assert cell(values, "kpi_gm_pct_ytd") == pytest.approx(ytd_value(o, "gm_pct"), abs=RATIO)
    assert cell(values, "kpi_ebitda_ytd") == pytest.approx(ytd_value(o, "ebitda"), abs=CENT)
    assert cell(values, "kpi_ebitda_fy") == pytest.approx(fy_value(o, "ebitda"), abs=CENT)
    assert cell(values, "kpi_breakeven") == pytest.approx(R.break_even_uplift(data), abs=RATIO)
    assert cell(values, "kpi_ccc") == pytest.approx(R.working_capital(data)["ccc"], abs=1e-9)


def test_dashboard_commentary_states_the_reference_numbers(values, data, ref):
    """The written summary is formulas; its numbers must be the model's numbers."""
    o, b = ref["outlook"], ref["budget"]
    rev_var = ytd_value(o, "revenue") - ytd_value(b, "revenue")
    ebitda_var = ytd_value(o, "ebitda") - ytd_value(b, "ebitda")
    bridge = R.ebitda_bridge_ytd(data)
    first, second, third, fourth, fifth = column(values, "dash_commentary")
    assert f"Revenue is {money(abs(rev_var))} ahead of budget through August" in first
    assert f"EBITDA is {money(abs(ebitda_var))} behind" in first
    assert "The growth has not reached the bottom line." in first
    for effect in EFFECTS:
        label = "unit cost" if effect == "cost" else effect
        assert f"{label} {signmoney(bridge[effect])}" in second
    assert "Grocery Retail, the channel that earns least per case" in second
    assert "Beef costs rose faster than its prices." in second
    assert f"(mostly Warehouse & Logistics) leave EBITDA at {money(bridge['actual_ebitda'])}" in third
    assert f"{R.break_even_uplift(data):.1%} price rise on the 4 open months" in fourth
    assert f"{R.break_even_volume(data):.0%} more volume" in fourth
    assert "Cheese & Dairy holds 52 days of stock" in fifth


def test_dashboard_movers_are_the_largest_segment_variances(values, data):
    t = R.pvm(data, "ytd")
    top = t.reindex(t["gm_variance"].abs().sort_values(ascending=False).index).head(5)
    for row, ((cat, chan), r) in zip(block(values, "dash_movers"), top.iterrows()):
        assert row[0] == f"{cat} · {chan}"
        assert row[5] == DRIVER_NAME[max(EFFECTS, key=lambda e: abs(r[e]))]
        assert row[8] == pytest.approx(r["gm_variance"], abs=CENT)


def test_cover_answers_are_live(values, data):
    answers = column(values, "cover_answers")
    assert f"A {R.break_even_uplift(data):.1%} price rise on the open months" in answers[2]
    assert f"EBITDA {money(R.outlook(data)['ebitda'].sum())} on the Base scenario" in answers[1]
    assert "Grocery Retail grew cases most" in answers[4]


# ------------------------------------------------------ the data in the file
def _csv(name):
    with open(ROOT / "data" / name, newline="", encoding="utf-8") as fh:
        return list(csv.reader(fh))


def _table(values, sheet, name):
    ws = values[sheet]
    ref = ws.tables[name].ref
    return [[c.value for c in row] for row in ws[ref]]


@pytest.mark.parametrize("sheet,table,csv_name", [
    ("Data_Sales", "tbl_Sales", "fact_sales.csv"),
    ("Data_BudgetUnits", "tbl_BudgetUnits", "budget_units_fy2026.csv"),
    ("Data_BudgetRates", "tbl_BudgetRates", "budget_rates_fy2026.csv"),
    ("Data_Opex", "tbl_Opex", "fact_opex.csv"),
    ("Data_Balances", "tbl_Balances", "balances.csv"),
])
def test_workbook_carries_the_committed_data(values, sheet, table, csv_name):
    """A stale workbook built from older CSVs fails here, before any figure is compared."""
    got, want = _table(values, sheet, table), _csv(csv_name)
    assert got[0] == want[0] and len(got) == len(want)
    for g, w in zip(got[1:], want[1:]):
        for gv, wv in zip(g, w):
            if isinstance(gv, datetime):
                assert gv.strftime("%Y-%m-%d") == wv
            elif isinstance(gv, (int, float)):
                assert gv == pytest.approx(float(wv), abs=1e-9)
            else:
                assert (gv or "") == wv


def test_power_query_budget_output(values, data):
    rows = _table(values, "PQ_Budget", "tbl_Budget")
    head, body = rows[0], rows[1:]
    assert len(body) == 288 * 12
    frame = pd.DataFrame(body, columns=head)
    got = frame.groupby("category")[["units", "revenue", "cogs", "freight"]].sum()
    want = data.budget.groupby("category")[["units", "revenue", "cogs", "freight"]].sum()
    for cat in L.CATEGORIES:
        assert list(got.loc[cat]) == pytest.approx(list(want.loc[cat]), abs=CENT)
