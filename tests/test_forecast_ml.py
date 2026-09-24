"""The machine-learning backtest: reproducible, leak-free, and reported as it came out."""
from __future__ import annotations

import pandas as pd
import pytest

import forecast_ml as M
import reference as R
from workbook_io import ROOT

TABLES = ROOT / "tables"


@pytest.fixture(scope="module")
def rerun():
    data = R.load()
    scores, monthly, importance = M.backtest(data)
    return scores, M.intervals(monthly), importance


def weighted(scores: pd.DataFrame, cls: str, weather: str, model: str) -> float:
    g = scores[(scores["class"] == cls) & (scores.weather == weather) & (scores.model == model)]
    return (g.wape * g.months).sum() / g.months.sum()


@pytest.mark.parametrize("name", ["ml_backtest", "ml_intervals", "ml_importance"])
def test_committed_results_are_what_the_code_produces(rerun, name):
    got = pd.read_csv(TABLES / f"{name}.csv")
    want = {"ml_backtest": rerun[0], "ml_intervals": rerun[1], "ml_importance": rerun[2]}[name]
    pd.testing.assert_frame_equal(got.reset_index(drop=True), want.reset_index(drop=True),
                                  check_exact=False, rtol=1e-6, atol=1e-9, check_dtype=False)


def test_no_test_month_is_ever_trained_on():
    data = R.load()
    for cls in M.CLASSES:
        x = M.frame(data, cls)
        for year in M.TEST_YEARS:
            train = x[x.index.year < year]
            assert train.index.max() < pd.Timestamp(year, 1, 1)
        # "last year" is the same month twelve months earlier, never the month itself
        assert (x["last_year"].dropna().index.to_series().diff().dropna() > pd.Timedelta(0)).all()
        pd.testing.assert_series_equal(x["last_year"].iloc[12:], x["use"].shift(12).iloc[12:], check_names=False)


def test_the_weather_regression_is_the_workbooks(rerun):
    """Fitted through December of the year before, it must equal reference.regression."""
    data = R.load()
    x = M.frame(data, "residential")
    train = x[x.index.year < 2026]
    import numpy as np
    coef, *_ = np.linalg.lstsq(np.column_stack([np.ones(len(train)), train.hdd, train.cdd, train.trend]),
                               train.use, rcond=None)
    ref = R.regression(data, "residential", 2026)
    assert list(coef) == pytest.approx([ref["intercept"], ref["hdd"], ref["cdd"], ref["trend"]], rel=1e-9)


def test_the_readme_reports_the_results(rerun):
    scores = rerun[0]
    text = (ROOT / "README.md").read_text(encoding="utf-8")
    for cls, weather, model in (("residential", "actual", "gradient boosting"),
                                ("residential", "actual", "weather regression"),
                                ("residential", "normal", "gradient boosting"),
                                ("residential", "normal", "weather regression"),
                                ("industrial", "normal", "linear, same inputs"),
                                ("industrial", "normal", "gradient boosting")):
        assert f"{weighted(scores, cls, weather, model):.1%}" in text, (cls, weather, model)
    cover = rerun[1]
    overall = cover.coverage.mul(cover.months).sum() / cover.months.sum()
    assert f"{overall:.0%}" in text


def test_intervals_are_at_least_as_wide_as_promised(rerun):
    cover = rerun[1]
    overall = cover.coverage.mul(cover.months).sum() / cover.months.sum()
    assert overall >= M.COVERAGE, "conformal intervals under-covered: the promise in the README is false"
