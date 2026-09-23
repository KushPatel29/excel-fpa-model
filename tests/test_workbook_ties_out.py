"""Every figure the workbook publishes, held to an independent pandas model.

The workbook computes with Power Query, SUMIFS, LINEST, DAX and what-if data
tables; model/reference.py recomputes the same analysis from the same public
extracts without reading the workbook. Money must agree to the cent, ratios
and coefficients to 1e-9.
"""
from __future__ import annotations

import csv
from datetime import datetime

import pandas as pd
import pytest

import layout as L
import reference as R
from workbook_io import ROOT, block, cell, column, money, signmoney

CENT = 0.01
TIGHT = 1e-9


@pytest.fixture(scope="module")
def fy(data):
    return R.fiscal_year(data)


@pytest.fixture(scope="module")
def pnl_ref(data):
    return R.ferc_pnl(data)


# --------------------------------------------------------------------- PnL
PNL_MAP = {
    "rev_residential_sales": "residential", "rev_small_or_commercial": "commercial",
    "rev_large_or_industrial": "industrial", "rev_public_street_and_highway_lighting": "lighting",
    "rev_other_sales_to_public_authorities": "public authorities", "retail": "retail", "wholesale": "wholesale",
    "refunds": "refund provision", "other_rev": "other revenue", "revenue": "revenue", "power": "power cost",
    "gm": "gross margin", "production": "production", "transmission": "transmission",
    "distribution": "distribution", "customer": "customer", "ag": "administrative & general",
    "regional": "regional market & other", "other_om": "other O&M", "da": "D&A",
    "taxes": "taxes other than income", "op_income": "operating income", "retail_mwh": "retail mwh",
    "residential_mwh": "residential mwh", "industrial_mwh": "industrial mwh", "total_mwh": "total mwh",
    "customers": "customers",
}


def test_pnl_years(values):
    assert block(values, "pnl_years")[0] == L.FERC_YEARS


@pytest.mark.parametrize("key", sorted(PNL_MAP))
def test_pnl_line_every_year(values, pnl_ref, key):
    got = block(values, f"pnl_{key}")[0]
    want = pnl_ref.loc[L.FERC_YEARS, PNL_MAP[key]]
    assert got == pytest.approx(list(want), abs=CENT), key


def test_pnl_ratios(values, pnl_ref):
    r = pnl_ref.loc[L.FERC_YEARS]
    assert block(values, "pnl_op_margin")[0] == pytest.approx(list(r["operating income"] / r["revenue"]), abs=TIGHT)
    assert block(values, "pnl_retail_cpk")[0] == pytest.approx(list(r["retail"] / r["retail mwh"] / 10), abs=TIGHT)
    assert block(values, "pnl_om_per_customer")[0] == pytest.approx(list(r["other O&M"] / r["customers"]), abs=CENT)


# --------------------------------------------------------------------- PVM
PVM_COLS = ["q0", "r0", "q1", "r1", "p0", "p1", "volume", "mix", "price"]


def pvm_rows(values, name):
    return {row[1]: row[2:] for row in block(values, f"pvm_{name}")}


@pytest.mark.parametrize("name,y0,y1", [("long", 2019, 2025), ("short", 2024, 2025)])
def test_pvm_annual(values, data, name, y0, y1):
    ref = R.pvm_annual(data, y0, y1)
    rows = pvm_rows(values, name)
    keys = dict(L.FERC_CLASSES)
    label_of = {v: k for k, v in R.FERC_CLASSES.items()}
    for key, got in rows.items():
        label = label_of[key]
        if label not in ref.index:
            assert all(abs(v) < CENT for v in got), key
            continue
        r = ref.loc[label]
        assert got[:9] == pytest.approx([r[c] for c in PVM_COLS], abs=CENT), label
        assert got[10] == pytest.approx(r["change"], abs=CENT)
        assert abs(got[11]) < 1e-3, "every class reconciles"
    assert keys


def test_pvm_year_to_date(values, data):
    ref = R.pvm_ytd(data)
    for cls, got in pvm_rows(values, "ytd").items():
        r = ref.loc[cls]
        assert got[:9] == pytest.approx([r[c] for c in PVM_COLS], abs=CENT), cls


def test_long_bridge(values, data):
    t = R.pvm_annual(data, 2019, 2025)
    want = [t["r0"].sum(), t["volume"].sum(), t["mix"].sum(), t["price"].sum(), t["r1"].sum()]
    assert column(values, "bridge_long") == pytest.approx(want, abs=CENT)


# ----------------------------------------------------------------- Monthly
def test_year_to_date_by_class(values, data, fy):
    closed = R.as_of(data).month
    m = data.monthly.reset_index()
    this = m[(m.date.dt.year == fy) & (m.date.dt.month <= closed)].groupby("cls")[["revenue", "mwh"]].sum()
    prior = m[(m.date.dt.year == fy - 1) & (m.date.dt.month <= closed)].groupby("cls")[["revenue", "mwh"]].sum()
    for row, cls in zip(block(values, "monthly_ytd"), L.CLASSES):
        assert row[0] == pytest.approx(this.loc[cls, "revenue"], abs=CENT)
        assert row[1] == pytest.approx(prior.loc[cls, "revenue"], abs=CENT)
        assert row[3] == pytest.approx(this.loc[cls, "mwh"], abs=CENT)
        assert row[4] == pytest.approx(prior.loc[cls, "mwh"], abs=CENT)


def test_eia_ferc_tieout(values, data):
    ref = R.eia_ferc_tieout(data)
    rows = block(values, "tieout")
    assert [r[0] for r in rows] == list(ref.index)
    for row, (_, r) in zip(rows, ref.iterrows()):
        assert row[1:3] == pytest.approx([r["eia_mwh"], r["ferc_mwh"]], abs=CENT)
        assert row[4:6] == pytest.approx([r["eia_revenue"], r["ferc_revenue"]], abs=CENT)
        assert abs(r["eia_mwh"] / r["ferc_mwh"] - 1) < 0.001


# ----------------------------------------------------------------- Weather
@pytest.mark.parametrize("cls", L.WEATHER_CLASSES)
def test_linest_matches_least_squares(values, data, fy, cls):
    coef = R.regression(data, cls, fy)
    got = block(values, f"coef_{cls}")[0]
    want = [coef["intercept"], coef["hdd"], coef["cdd"], coef["trend"], coef["r2"]]
    assert got == pytest.approx(want, rel=1e-9, abs=1e-12)


def test_normals(values, data, fy):
    n = R.normals(data, fy)
    assert block(values, "normal_hdd")[0] == pytest.approx(list(n["hdd"]), abs=TIGHT)
    assert block(values, "normal_cdd")[0] == pytest.approx(list(n["cdd"]), abs=TIGHT)


def test_weather_impact(values, data, fy):
    wi = R.weather_impact(data, fy, R.as_of(data).month)
    total = block(values, "weather_impact_total")[0]
    assert total[10] == pytest.approx(wi["mwh"].sum(), abs=CENT)
    assert total[11] == pytest.approx(wi["revenue"].sum(), abs=CENT)


# -------------------------------------------------------------------- Plan
def test_plan_month_by_class(values, data, fy):
    ref = R.plan(data, fy)
    for row in block(values, "plan_long"):
        month, cls = row[0], row[1]
        r = ref.loc[(pd.Timestamp(month), cls)]
        assert [row[3], row[5], row[7], row[8]] == pytest.approx(
            [r["customers"], r["mwh"], r["price"], r["revenue"]], rel=1e-10, abs=CENT), (month, cls)


def test_plan_bridge(values, data, fy):
    ref = R.plan_bridge(data, fy)
    for row, cls in zip(block(values, "plan_bridge"), L.CLASSES):
        r = ref.loc[cls]
        want = [r["plan"], r["customers"], r["weather"], r["usage"], r["price"], r["actual"]]
        assert row[:6] == pytest.approx(want, abs=CENT), cls
    total = block(values, "plan_bridge_total")[0]
    assert total[6] == pytest.approx(ref["actual"].sum() - ref["plan"].sum(), abs=CENT)


def test_backtest_re_runs_the_plan_as_of_each_year(values, data):
    rows = block(values, "backtest")
    assert [r[0] for r in rows] == L.BACKTEST_YEARS
    for row in rows:
        b = R.backtest(data, row[0])
        assert row[1:] == pytest.approx([b["months"], b["plan"], b["actual"], b["variance"], b["wape_revenue"],
                                         b["wape_mwh"]], rel=1e-10, abs=CENT), row[0]


# ---------------------------------------------------------------- Forecast
def test_forecast_open_months(values, data):
    ref = R.forecast(data)
    open_rows = [r for r in block(values, "fc_long") if r[2] is True]
    assert len(open_rows) == len(ref)
    for row in open_rows:
        r = ref.loc[(pd.Timestamp(row[0]), row[1])]
        assert [row[3], row[4], row[5], row[7], row[8], row[9]] == pytest.approx(
            [r["customers"], r["hdd"], r["cdd"], r["mwh"], r["price"], r["revenue"]], rel=1e-10, abs=CENT)


def test_outlook(values, data):
    o = R.outlook(data)
    assert cell(values, "fy_revenue") == pytest.approx(o["revenue"], abs=CENT)
    assert cell(values, "fy_mwh") == pytest.approx(o["mwh"], abs=CENT)


def test_statistical_cross_check_is_sane(values):
    months = block(values, "fc_months")
    ets = [r[5] for r in months if r[5] not in (None, "")]
    driver = [r[2] for r in months if r[2] not in (None, "")]
    assert len(ets) == len(driver) > 0
    for e, d in zip(ets, driver):
        assert 0.8 * d < e < 1.2 * d


# --------------------------------------------------------------- Scenarios
def test_weather_scenarios(values, data):
    for row, (name, *_rest) in zip(block(values, "scn_weather"), L.WEATHER_SCENARIOS):
        o = R.outlook(data, weather=name)
        assert row[0] == pytest.approx(o["revenue"], abs=CENT), name
        assert row[1] == pytest.approx(o["mwh"], abs=CENT), name


def test_sensitivity_grid(values, data):
    want = R.sensitivity_grid(data)
    assert column(values, "sens_rate_axis") == pytest.approx(L.SENS_RATE)
    assert block(values, "sens_mw_axis")[0] == pytest.approx(L.SENS_MW)
    for row, (_, expected) in zip(block(values, "sens_grid"), want.iterrows()):
        assert row == pytest.approx(list(expected), abs=CENT)


def test_tornado(values, data):
    t = R.tornado(data)
    assert column(values, "tornado") == pytest.approx([t[k] for k in ("rate", "mw", "hdd", "cdd", "customers")],
                                                      abs=CENT)


def test_mild_autumn_offset(values, data):
    loss = R.outlook(data)["revenue"] - R.outlook(data, weather="Mild")["revenue"]
    mild_open = R.forecast(data, weather="Mild")["revenue"].sum()
    assert cell(values, "offset_loss") == pytest.approx(loss, abs=CENT)
    assert cell(values, "offset_rate") == pytest.approx(loss / mild_open, abs=TIGHT)


# ------------------------------------------------------------------- Peers
PEER_ROWS = ["retail revenue", "retail mwh", "customers", "retail price", "residential price", "industrial share",
             "retail mwh cagr", "industrial mwh cagr", "retail price cagr", "power cost per mwh",
             "other om per customer", "operating margin"]


def test_peers(values, data):
    ref = R.peers(data)
    assert block(values, "peers_ids")[0] == R.PEERS
    for row, metric in zip(block(values, "peers_table"), PEER_ROWS):
        want = list(ref.loc[R.PEERS, metric])
        if metric in ("retail price", "residential price"):
            want = [w / 10 for w in want]
        assert row == pytest.approx(want, rel=1e-10, abs=CENT), metric


# ------------------------------------------------------- data model, checks
def test_cube_values_come_from_the_data_model(values, data, fy):
    m = data.monthly.reset_index()
    by = m.groupby([m.cls, m.date.dt.year])["revenue"].sum()
    rows = block(values, "explore_cube")
    for row, cls in zip(rows, [*L.CLASSES, None]):
        for year, got in zip(range(fy - 3, fy + 1), row):
            want = by.xs(year, level=1).sum() if cls is None else by[(cls, year)]
            assert got == pytest.approx(want, abs=CENT), (cls, year)
    for dax, formula, agree in block(values, "explore_yoy"):
        assert dax == pytest.approx(formula, abs=1e-9) and agree == "Yes"


def test_every_check_passes(values):
    status = column(values, "checks_status")
    assert len(status) == 31
    assert status == ["PASS"] * len(status)
    assert cell(values, "checks_summary") == f"All {len(status)} checks pass"


# ------------------------------------------------------- dashboard, cover
def test_dashboard_tiles(values, data):
    b = R.plan_bridge(data, R.fiscal_year(data))
    assert cell(values, "kpi_fy_revenue") == pytest.approx(R.outlook(data)["revenue"], abs=CENT)
    assert cell(values, "kpi_ytd_vs_plan") == pytest.approx(b["actual"].sum() - b["plan"].sum(), abs=CENT)
    wi = R.weather_impact(data, R.fiscal_year(data), R.as_of(data).month)
    assert cell(values, "kpi_weather_gwh") == pytest.approx(wi["mwh"].sum() / 1000, abs=1e-6)


def test_dashboard_commentary_states_the_reference_numbers(values, data):
    fy = R.fiscal_year(data)
    b = R.plan_bridge(data, fy).sum()
    wi = R.weather_impact(data, fy, R.as_of(data).month)
    t = R.pvm_annual(data, 2019, 2025).sum()
    first, second, _third, fourth, fifth = column(values, "dash_commentary")
    assert f"came in {money(abs(b['actual'] - b['plan']))}" in first
    assert f"weather {signmoney(b['weather'])}, price {signmoney(b['price'])}" in first
    assert f"{money(abs(wi['revenue'].sum()))} of revenue lost" in second
    assert f"price {signmoney(t['price'])}, volume {signmoney(t['volume'])}, mix {signmoney(t['mix'])}" in fourth
    assert money(R.outlook(data)["revenue"]) in fifth
    assert money(R.tornado(data)["rate"]) in fifth


def test_cover_answers_are_live(values, data):
    answers = column(values, "cover_answers")
    assert money(R.outlook(data)["revenue"]) in answers[3]
    assert "rank 1 of 5" in answers[4]


# --------------------------------------------------------- the data in the file
def _csv(name):
    with open(ROOT / "data" / name, newline="", encoding="utf-8") as fh:
        return list(csv.reader(fh))


@pytest.mark.parametrize("sheet,table,csv_name", [
    ("Data_EIA", "tbl_EIA", "eia_pge_monthly.csv"),
    ("Data_FERC", "tbl_FERC_Revenue", "ferc_revenue.csv"),
    ("Data_FERC", "tbl_FERC_Expense", "ferc_expense.csv"),
    ("Data_FERC", "tbl_FERC_Income", "ferc_income.csv"),
    ("Data_Utilities", "tbl_Utilities", "utilities.csv"),
])
def test_workbook_carries_the_committed_extracts(values, sheet, table, csv_name):
    """A workbook built from older extracts fails here, before any figure is compared."""
    ws = values[sheet]
    got = [[c.value for c in row] for row in ws[ws.tables[table].ref]]
    want = _csv(csv_name)
    assert got[0] == want[0] and len(got) == len(want)
    for g, w in zip(got[1:], want[1:]):
        for gv, wv in zip(g, w):
            if isinstance(gv, datetime):
                raise AssertionError("no dates expected in the raw extracts")
            if isinstance(gv, (int, float)):
                assert gv == pytest.approx(float(wv), abs=1e-9)
            else:
                assert (gv or "") == wv


def test_workbook_carries_the_noaa_lines(values):
    ws = values["Data_NOAA"]
    got = [row[0].value for row in ws[ws.tables["tbl_NOAA"].ref]][1:]
    assert got == (ROOT / "data" / "noaa_willamette_valley_degree_days.txt").read_text().splitlines()
