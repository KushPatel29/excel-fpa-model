"""Download the public sources and write the extracts the workbook is built from.

    pip install -r requirements-pipeline.txt
    python pipeline/fetch_sources.py

Three sources, all public, none needing an account or a key:

* EIA-861M, utility-level monthly retail sales (U.S. Energy Information
  Administration). Portland General Electric (EIA utility 15248, Oregon):
  revenue, megawatt-hours and customers by customer class, every month from
  January 2017. The latest months are marked Preliminary by EIA.
* NOAA NCEI nClimDiv divisional degree days. Oregon climate division 2, the
  Willamette Valley, which is PGE's service territory: heating and cooling
  degree days (base 65 F) by month. Kept as NOAA's own fixed-width lines so
  Power Query does the parsing inside the workbook.
* FERC Form 1, the audited annual report every major US electric utility
  files with its federal regulator, as cleaned and published by Catalyst
  Cooperative's PUDL project (release pinned below). Revenue by customer
  class, operating expenses and the income statement for PGE and four
  Pacific Northwest peers.

The extracts are committed, so the workbook, the reference model and the
tests never touch the network. data/sources.json records every URL, the
retrieval time and each download's SHA-256. EIA revises recent months and
PUDL publishes new releases, so a later run can legitimately produce
different numbers; the commit history is the record of what was used.
"""
from __future__ import annotations

import hashlib
import http.client
import io
import json
import re
import time
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

EIA_UTILITY = 15248                  # Portland General Electric Co, Oregon
FIRST_YEAR, LAST_YEAR = 2017, 2026
EIA_BASE = "https://www.eia.gov/electricity/data/eia861m"

NOAA_URL = "https://www.ncei.noaa.gov/pub/data/cirs/climdiv/"
NOAA_DIVISION = "3502"               # state 35 = Oregon, division 02 = Willamette Valley
NOAA_ELEMENTS = {"25": "hdd", "26": "cdd"}
NOAA_FIRST_YEAR = 2007

PUDL_RELEASE = "v2026.9.0"
PUDL_BASE = f"https://s3.us-west-2.amazonaws.com/pudl.catalyst.coop/{PUDL_RELEASE}"
FERC_YEARS = range(2014, 2026)

# FERC respondent id -> (utility, short name, states served, role)
UTILITIES = {
    250: ("Portland General Electric Company", "PGE", "OR", "company"),
    303: ("PacifiCorp", "PacifiCorp", "OR, WA, UT, WY, ID, CA", "peer"),
    162: ("Puget Sound Energy, Inc.", "Puget Sound Energy", "WA", "peer"),
    182: ("Avista Corporation", "Avista", "WA, ID", "peer"),
    216: ("Idaho Power Company", "Idaho Power", "ID, OR", "peer"),
}
# NorthWestern was dropped: not a Pacific Northwest utility, and its 2024 filing
# shrinks by a fifth because its South Dakota business became a separate FERC
# respondent, a reporting change that would read as lost load.
REVENUE_TYPES = [
    "residential_sales", "small_or_commercial", "large_or_industrial",
    "public_street_and_highway_lighting", "other_sales_to_public_authorities",
    "sales_to_ultimate_consumers", "sales_for_resale", "sales_of_electricity",
    "provision_for_rate_refunds", "revenues_net_of_provision_for_refunds",
    "other_operating_revenues", "electric_operating_revenues",
]
EXPENSE_TYPES = [
    "fuel_steam_power_generation", "nuclear_fuel_expense", "fuel", "fuel_other_renewable_generation",
    "purchased_power", "power_purchased_for_storage_operations", "storage_fuel_energy_storage_expense",
    "power_production_expenses", "transmission_expenses", "distribution_expenses",
    "customer_account_expenses", "customer_service_and_information_expenses", "sales_expenses",
    "administrative_and_general_expenses", "operations_and_maintenance_expenses_electric",
]
INCOME_TYPES = [
    "operating_revenues", "operation_expense", "maintenance_expense", "depreciation_expense",
    "depreciation_expense_for_asset_retirement_costs", "amortization_and_depletion_of_utility_plant",
    "amortization_of_other_utility_plant", "taxes_other_than_income_taxes_utility_operating_income",
]

CLASSES = ["residential", "commercial", "industrial", "transportation", "total"]
MEASURES = ["revenue_k", "mwh", "customers"]

sources: list[dict] = []


def fetch(url: str, label: str, attempts: int = 4) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (excel-fpa-model data refresh)"})
    for attempt in range(1, attempts + 1):
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                body = response.read()
            break
        except (http.client.IncompleteRead, ConnectionError, TimeoutError) as error:
            if attempt == attempts:
                raise
            print(f"  retrying {label} after {type(error).__name__}")
            time.sleep(5 * attempt)
    sources.append({"source": label, "url": url, "bytes": len(body),
                    "sha256": hashlib.sha256(body).hexdigest()})
    return body


def eia_year(year: int) -> bytes:
    """The current year lives in xls/, earlier ones in archive/xls/, and the file
    was renamed from retail_sales to sales_ult_cust in 2021."""
    name = f"retail_sales_{year}.xlsx" if year <= 2020 else f"sales_ult_cust_{year}.xlsx"
    for folder in ("xls", "archive/xls"):
        try:
            body = fetch(f"{EIA_BASE}/{folder}/{name}", f"EIA-861M {year}")
        except urllib.error.HTTPError as error:
            if error.code != 404:
                raise
            continue
        if body[:2] == b"PK":            # an .xlsx is a zip; EIA answers a missing file with a 200 HTML page
            return body
        sources.pop()
    raise RuntimeError(f"EIA-861M {year}: {name} not found in xls/ or archive/xls/")


def eia_monthly() -> pd.DataFrame:
    frames = []
    for year in range(FIRST_YEAR, LAST_YEAR + 1):
        raw = pd.read_excel(io.BytesIO(eia_year(year)), header=None, skiprows=3)
        raw = raw.iloc[:, :22]
        raw.columns = (["year", "month", "utility_number", "utility_name", "state", "ownership",
                        "data_status"] + [f"{c}_{m}" for c in CLASSES for m in MEASURES])
        pge = raw[(pd.to_numeric(raw["utility_number"], errors="coerce") == EIA_UTILITY)
                  & (raw["state"] == "OR")]
        frames.append(pge)
    df = pd.concat(frames, ignore_index=True)
    df["year"] = df["year"].astype(int)
    df["month"] = df["month"].astype(int)
    for c in [f"{c}_{m}" for c in CLASSES for m in MEASURES]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        if c.endswith("_customers"):
            df[c] = df[c].round().astype(int)
    df = df.sort_values(["year", "month"]).reset_index(drop=True)
    assert not df.duplicated(["year", "month"]).any(), "one row per month expected"
    return df[["year", "month", "data_status"] + [f"{c}_{m}" for c in CLASSES for m in MEASURES]]


def noaa_lines() -> list[str]:
    listing = fetch(NOAA_URL, "NOAA nClimDiv file listing").decode("utf-8", "replace")
    lines = []
    for code, element in NOAA_ELEMENTS.items():
        versions = re.findall(rf"climdiv-{element}cdv-(v[\d.]+-\d{{8}})", listing)
        latest = f"climdiv-{element}cdv-" + max(versions, key=lambda v: v.rsplit("-", 1)[1])
        text = fetch(NOAA_URL + latest, f"NOAA nClimDiv {element.upper()} by division").decode("ascii")
        for line in text.splitlines():
            if line[:4] == NOAA_DIVISION and line[4:6] == code and int(line[6:10]) >= NOAA_FIRST_YEAR:
                lines.append(line.rstrip())
    return lines


def ferc(table: str, type_column: str, keep: list[str], values: list[str]) -> pd.DataFrame:
    frame = pd.read_parquet(io.BytesIO(fetch(f"{PUDL_BASE}/{table}.parquet", f"PUDL {PUDL_RELEASE} {table}")))
    frame = frame[frame["utility_id_ferc1"].isin(UTILITIES) & frame["report_year"].isin(FERC_YEARS)
                  & frame[type_column].isin(keep)]
    if "utility_type" in frame:
        frame = frame[frame["utility_type"] == "electric"]
    out = frame[["report_year", "utility_id_ferc1", type_column, *values]].copy()
    out = out.sort_values(["utility_id_ferc1", "report_year", type_column]).reset_index(drop=True)
    assert not out.duplicated(["report_year", "utility_id_ferc1", type_column]).any()
    return out


def write_csv(frame: pd.DataFrame, name: str) -> None:
    frame.to_csv(DATA / name, index=False, lineterminator="\n")


def main() -> None:
    DATA.mkdir(exist_ok=True)
    retrieved = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    monthly = eia_monthly()
    write_csv(monthly, "eia_pge_monthly.csv")

    lines = noaa_lines()
    (DATA / "noaa_willamette_valley_degree_days.txt").write_text("\n".join(lines) + "\n", encoding="ascii",
                                                                  newline="\n")

    revenue = ferc("core_ferc1__yearly_operating_revenues_sched300", "revenue_type", REVENUE_TYPES,
                   ["dollar_value", "sales_mwh", "avg_customers_per_month"])
    expense = ferc("core_ferc1__yearly_operating_expenses_sched320", "expense_type", EXPENSE_TYPES,
                   ["dollar_value"])
    income = ferc("core_ferc1__yearly_income_statements_sched114", "income_type", INCOME_TYPES,
                  ["dollar_value"])
    write_csv(revenue, "ferc_revenue.csv")
    write_csv(expense, "ferc_expense.csv")
    write_csv(income, "ferc_income.csv")
    utilities = pd.DataFrame([(k, *v) for k, v in UTILITIES.items()],
                             columns=["utility_id_ferc1", "utility", "short_name", "states", "role"])
    write_csv(utilities, "utilities.csv")

    last = monthly.iloc[-1]
    manifest = {
        "retrieved_utc": retrieved,
        "as_of_month": f"{int(last['year'])}-{int(last['month']):02d}",
        "pudl_release": PUDL_RELEASE,
        "eia_utility_number": EIA_UTILITY,
        "noaa_division": NOAA_DIVISION,
        "rows": {"eia_pge_monthly": len(monthly), "noaa_lines": len(lines),
                 "ferc_revenue": len(revenue), "ferc_expense": len(expense), "ferc_income": len(income)},
        "sources": sources,
    }
    (DATA / "sources.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8", newline="\n")
    print(json.dumps({k: v for k, v in manifest.items() if k != "sources"}, indent=2))


if __name__ == "__main__":
    main()
