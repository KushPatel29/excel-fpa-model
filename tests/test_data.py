"""The public extracts: complete, internally consistent, and traceable to their source."""
from __future__ import annotations

import json

import pandas as pd
import pytest

import reference as R
from workbook_io import ROOT

DATA = ROOT / "data"


@pytest.fixture(scope="module")
def manifest():
    return json.loads((DATA / "sources.json").read_text(encoding="utf-8"))


def test_manifest_counts_the_rows(manifest):
    for name, rows in manifest["rows"].items():
        if name == "noaa_lines":
            got = len((DATA / "noaa_willamette_valley_degree_days.txt").read_text().splitlines())
        else:
            got = sum(1 for _ in (DATA / f"{name}.csv").open(encoding="utf-8")) - 1
        assert got == rows, name


def test_every_source_is_recorded_with_a_hash(manifest):
    assert manifest["pudl_release"].startswith("v20")
    for s in manifest["sources"]:
        assert s["url"].startswith("https://") and len(s["sha256"]) == 64 and s["bytes"] > 0


def test_eia_months_are_contiguous():
    eia = pd.read_csv(DATA / "eia_pge_monthly.csv")
    months = pd.to_datetime(dict(year=eia.year, month=eia.month, day=1))
    assert (months.diff().dropna().dt.days.between(28, 31)).all()
    assert not months.duplicated().any()


def test_eia_classes_add_up_to_eias_total():
    eia = pd.read_csv(DATA / "eia_pge_monthly.csv")
    for m in ("revenue_k", "mwh", "customers"):
        parts = sum(eia[f"{c}_{m}"] for c in R.CLASSES)
        assert (parts - eia[f"total_{m}"]).abs().max() < 1e-6, m


def test_ferc_classes_add_up_to_ferc_retail():
    rev = pd.read_csv(DATA / "ferc_revenue.csv")
    w = rev.pivot_table(index=["utility_id_ferc1", "report_year"], columns="revenue_type", values="dollar_value").fillna(0)
    parts = w[list(R.FERC_CLASSES.values())].sum(axis=1)
    gap = parts - w["sales_to_ultimate_consumers"]
    # PGE's five classes add up to FERC's subtotal exactly. Avista books about $2M a
    # year (0.2%) in a class the extract does not carry, so peers are held to 0.5%;
    # the peer metrics read FERC's subtotal, not this sum.
    assert gap.loc[R.COMPANY].abs().max() < 1
    assert (gap.abs() / w["sales_to_ultimate_consumers"]).max() < 0.005


def test_noaa_lines_are_willamette_valley_degree_days():
    lines = (DATA / "noaa_willamette_valley_degree_days.txt").read_text().splitlines()
    assert all(line[:4] == "3502" and line[4:6] in ("25", "26") and len(line) >= 94 for line in lines)


def test_eia_and_ferc_agree_on_volume(data):
    t = R.eia_ferc_tieout(data)
    assert ((t["eia_mwh"] / t["ferc_mwh"] - 1).abs() < 0.001).all()


# ---- identities the reference model must honour on its own
def test_pvm_reconciles_in_every_class(data):
    for t in (R.pvm_annual(data, 2019, 2025), R.pvm_annual(data, 2024, 2025), R.pvm_ytd(data)):
        assert ((t["volume"] + t["mix"] + t["price"]) - t["change"]).abs().max() < 1e-4


def test_plan_bridge_closes(data):
    b = R.plan_bridge(data, R.fiscal_year(data))
    effects = b[["customers", "weather", "usage", "price"]].sum(axis=1)
    assert ((b["plan"] + effects) - b["actual"]).abs().max() < 1e-4


def test_rate_tornado_is_one_percent_of_open_month_revenue(data):
    assert R.tornado(data)["rate"] == pytest.approx(0.01 * R.forecast(data)["revenue"].sum(), rel=1e-12)


def test_weather_model_explains_residential_use(data):
    assert R.regression(data, "residential", R.fiscal_year(data))["r2"] > 0.7
