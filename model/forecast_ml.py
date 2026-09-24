"""Does machine learning forecast PGE's load better than the workbook's regression?

    python model/forecast_ml.py          # writes tables/ml_backtest.csv and tables/ml_intervals.csv

The workbook forecasts use per customer with a LINEST regression on degree days
and a trend. This module asks whether a gradient-boosted model does better, and
answers it the way a forecast should be judged: an expanding-window backtest
over five years (2022-2026), each year forecast only from the months before it.

Four models, all predicting monthly MWh per customer, times customers:

* seasonal naive: the same month last year;
* weather regression: the workbook's model (degree days and a trend);
* linear with the same inputs as the boosted model (degree days, trend, season,
  and the same month last year), so the comparison is about the learner, not
  about who was given more information;
* gradient boosting (scikit-learn HistGradientBoostingRegressor) on those inputs.

Each is scored twice: with the weather that happened (model skill), and with
normal weather (what a plan made in December actually knows). The boosted
model also gets 80% prediction intervals by split conformal prediction: the
interval's width is the 80th percentile of the absolute percentage errors on
the previous year's out-of-sample forecasts, and its coverage is then measured
on the year it was not calibrated on.

Industrial load is not weather-driven, so the weather regression is not fitted
to it; the other three models are.
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.inspection import permutation_importance

import reference as R

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tables"
TEST_YEARS = [2022, 2023, 2024, 2025, 2026]
CLASSES = ["residential", "commercial", "industrial"]
FEATURES = ["hdd", "cdd", "trend", "season_sin", "season_cos", "last_year"]
COVERAGE = 0.80
SEED = 0


def frame(data: R.Data, cls: str) -> pd.DataFrame:
    x = data.monthly.xs(cls, level="cls").sort_index().join(data.weather)
    x = x[x.index >= R.FIT_START].copy()
    x["use"] = x["mwh"] / x["customers"]
    x["trend"] = [R.trend(d) for d in x.index]
    x["season_sin"] = np.sin(2 * np.pi * x.index.month / 12)
    x["season_cos"] = np.cos(2 * np.pi * x.index.month / 12)
    x["last_year"] = x["use"].shift(12)
    return x


def with_normal_weather(data: R.Data, test: pd.DataFrame, year: int) -> pd.DataFrame:
    """The test year as a December plan sees it: normal degree days, not actual ones."""
    n = R.normals(data, year)
    t = test.copy()
    t["hdd"] = [n.loc[d.month, "hdd"] for d in t.index]
    t["cdd"] = [n.loc[d.month, "cdd"] for d in t.index]
    return t


def fit_predict(train: pd.DataFrame, test: pd.DataFrame, cls: str) -> dict[str, np.ndarray]:
    out = {"seasonal naive": test["last_year"].to_numpy()}
    if cls in R.WEATHER_CLASSES:
        X = np.column_stack([np.ones(len(train)), train.hdd, train.cdd, train.trend])
        coef, *_ = np.linalg.lstsq(X, train.use, rcond=None)
        out["weather regression"] = np.column_stack([np.ones(len(test)), test.hdd, test.cdd, test.trend]) @ coef
    tr = train.dropna(subset=["last_year"])
    X = np.column_stack([np.ones(len(tr)), tr[FEATURES]])
    coef, *_ = np.linalg.lstsq(X, tr.use, rcond=None)
    out["linear, same inputs"] = np.column_stack([np.ones(len(test)), test[FEATURES]]) @ coef
    gbm = HistGradientBoostingRegressor(max_iter=300, learning_rate=0.05, max_leaf_nodes=8, min_samples_leaf=5,
                                        random_state=SEED).fit(tr[FEATURES], tr.use)
    out["gradient boosting"] = gbm.predict(test[FEATURES])
    out["_model"] = gbm
    return out


def backtest(data: R.Data) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    scores, monthly, importance = [], [], []
    for cls in CLASSES:
        x = frame(data, cls)
        for year in TEST_YEARS:
            train, test = x[x.index.year < year], x[x.index.year == year]
            assert train.index.max() < test.index.min(), "no month of the test year may be trained on"
            for weather, t in (("actual", test), ("normal", with_normal_weather(data, test, year))):
                preds = fit_predict(train, t, cls)
                gbm = preds.pop("_model")
                for model, use in preds.items():
                    mwh = use * t["customers"].to_numpy()
                    err = mwh - t["mwh"].to_numpy()
                    scores.append({"class": cls, "year": year, "weather": weather, "model": model,
                                   "months": len(t), "wape": np.abs(err).sum() / t["mwh"].sum(),
                                   "bias": err.sum() / t["mwh"].sum()})
                    if model == "gradient boosting":
                        monthly += [{"class": cls, "year": year, "weather": weather, "month": d,
                                     "actual_mwh": a, "forecast_mwh": f}
                                    for d, a, f in zip(t.index, t["mwh"], mwh)]
                if weather == "actual" and year == TEST_YEARS[-2]:
                    tr = train.dropna(subset=["last_year"])
                    pi = permutation_importance(gbm, tr[FEATURES], tr.use, n_repeats=20, random_state=SEED)
                    importance += [{"class": cls, "feature": f, "importance": v}
                                   for f, v in zip(FEATURES, pi.importances_mean)]
    return pd.DataFrame(scores), pd.DataFrame(monthly), pd.DataFrame(importance)


def intervals(monthly: pd.DataFrame) -> pd.DataFrame:
    """Split conformal: calibrate on last year's out-of-sample errors, score this year."""
    m = monthly[monthly.weather == "normal"].copy()
    m["ape"] = (m.forecast_mwh - m.actual_mwh).abs() / m.forecast_mwh
    rows = []
    for cls in CLASSES:
        c = m[m["class"] == cls]
        for year in TEST_YEARS[1:]:
            calib = c[c.year == year - 1].ape
            # the finite-sample conformal quantile: ceil((n+1)q)/n
            n = len(calib)
            q = np.quantile(calib, min(1.0, np.ceil((n + 1) * COVERAGE) / n), method="higher")
            test = c[c.year == year]
            lo, hi = test.forecast_mwh * (1 - q), test.forecast_mwh * (1 + q)
            covered = ((test.actual_mwh >= lo) & (test.actual_mwh <= hi)).mean()
            rows.append({"class": cls, "year": year, "half_width": q, "months": len(test), "coverage": covered})
    return pd.DataFrame(rows)


def summary(scores: pd.DataFrame) -> pd.DataFrame:
    return (scores.groupby(["class", "weather", "model"])
            .apply(lambda g: pd.Series({"wape": (g.wape * g.months).sum() / g.months.sum(),
                                        "years_best": 0}), include_groups=False).reset_index())


def main() -> None:
    data = R.load()
    scores, monthly, importance = backtest(data)
    cover = intervals(monthly)
    OUT.mkdir(exist_ok=True)
    fmt = {"float_format": "%.12g", "index": False, "lineterminator": "\n"}
    scores.to_csv(OUT / "ml_backtest.csv", **fmt)
    cover.to_csv(OUT / "ml_intervals.csv", **fmt)
    importance.to_csv(OUT / "ml_importance.csv", **fmt)
    table = scores.groupby(["class", "weather", "model"]).apply(
        lambda g: (g.wape * g.months).sum() / g.months.sum(), include_groups=False).unstack("model")
    print(table.round(4).to_string())
    print(cover.groupby("class")[["half_width", "coverage"]].mean().round(3).to_string())
    print("overall coverage", round(cover.coverage.mul(cover.months).sum() / cover.months.sum(), 3))


if __name__ == "__main__":
    main()
