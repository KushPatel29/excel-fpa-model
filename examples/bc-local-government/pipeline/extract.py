"""Validate and normalize B.C. municipal statistics into portfolio-ready CSVs."""

from __future__ import annotations

import csv
import json
import math
import re
import unicodedata
from collections import defaultdict
from pathlib import Path
from statistics import median

from openpyxl import load_workbook

from contracts import (
    ASSET_CATEGORIES,
    ASSET_COLUMNS,
    BLANK_MARKERS,
    KEY_COLUMNS,
    OPTIONAL_FIELDS,
    SCHEDULE_COLUMNS,
)

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
SILVER = ROOT / "data" / "silver"
YEARS = range(2019, 2025)

NAME_RENAMES = {
    "Queen Charlotte": "Daajing Giids",
    "Sechelt Indian Government District": "shíshálh Nation Government District",
    "Sechelt Indian Government": "shíshálh Nation Government District",
    "Vancouver1": "Vancouver",
}

TYPE_LABELS = {
    "C": "City",
    "D": "District",
    "T": "Town",
    "V": "Village",
    "IGD": "Island Municipality",
    "IM": "Island Municipality",
    "M": "Mountain Resort Municipality",
}


def clean_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def numeric(value: object) -> float | None:
    if value in BLANK_MARKERS:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = clean_text(value).replace(",", "").replace("$", "")
    if text in BLANK_MARKERS:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def slug(value: str) -> str:
    folded = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", folded.lower()).strip("-")


def canonical_name(value: object) -> str:
    name = clean_text(value)
    name = re.sub(r"(?<=\D)1$", "", name)
    name = re.sub(r"\s+-\s+(?:C|D|T|V|REGM|IM)$", "", name)
    if "Nation Government District" in name and name.lower().startswith("sh"):
        return "shíshálh Nation Government District"
    return NAME_RENAMES.get(name, name)


def canonical_type(name: str, value: object) -> str:
    type_code = clean_text(value)
    if name == "Mission":
        return "C"
    return type_code


def municipality_id(name: str, type_code: str) -> str:
    return f"{slug(name)}-{slug(type_code)}"


def selection_label(name: str, type_code: str, duplicate_names: set[str]) -> str:
    if name not in duplicate_names:
        return name
    return f"{name} ({TYPE_LABELS.get(type_code, type_code)})"


def resolve(headers: list[str], aliases: tuple[str, ...], schedule: str, field: str) -> int:
    for alias in aliases:
        if alias in headers:
            return headers.index(alias)
    raise ValueError(f"schedule {schedule}: missing {field}; expected one of {aliases}")


def read_schedule(schedule: str, year: int) -> list[dict[str, object]]:
    path = RAW / f"schedule{schedule}_{year}.xlsx"
    sheet = load_workbook(path, read_only=True, data_only=True).active
    raw_headers = next(sheet.iter_rows(min_row=2, max_row=2, values_only=True))
    headers = [clean_text(value) for value in raw_headers]
    for key in KEY_COLUMNS:
        if key not in headers:
            raise ValueError(f"{path.name}: missing key column {key}")
    fields = SCHEDULE_COLUMNS[schedule]
    indexes: dict[str, int | None] = {}
    for field, aliases in fields.items():
        try:
            indexes[field] = resolve(headers, aliases, schedule, field)
        except ValueError:
            if (schedule, field) not in OPTIONAL_FIELDS:
                raise
            indexes[field] = None
    rows: list[dict[str, object]] = []
    for values in sheet.iter_rows(min_row=3, values_only=True):
        name = canonical_name(values[headers.index("Municipalities")])
        if not name or name.lower() in {"total", "totals", "grand total", "grand totals"}:
            continue
        type_code = canonical_type(name, values[headers.index("Type")])
        rd = clean_text(values[headers.index("RD")])
        record: dict[str, object] = {
            "municipality_name": name,
            "municipality_type": type_code,
            "regional_district": rd,
            "year": year,
        }
        for field, index in indexes.items():
            record[field] = None if index is None else numeric(values[index])
        rows.append(record)
    return rows


def read_assets(year: int) -> list[dict[str, object]]:
    path = RAW / f"schedule503_{year}.xlsx"
    sheet = load_workbook(path, read_only=True, data_only=True).active
    headers = [clean_text(value) for value in next(sheet.iter_rows(min_row=2, max_row=2, values_only=True))]
    indexes = {
        field: resolve(headers, aliases, "503", field)
        for field, aliases in ASSET_COLUMNS.items()
    }
    category_index = headers.index("Asset Category")
    rows: list[dict[str, object]] = []
    for values in sheet.iter_rows(min_row=3, values_only=True):
        name = canonical_name(values[headers.index("Municipalities")])
        category = clean_text(values[category_index])
        if not name or not category:
            continue
        if category not in ASSET_CATEGORIES:
            raise ValueError(f"{path.name}: unexpected asset category {category!r}")
        type_code = canonical_type(name, values[headers.index("Type")])
        record: dict[str, object] = {
            "municipality_name": name,
            "municipality_type": type_code,
            "regional_district": clean_text(values[headers.index("RD")]),
            "year": year,
            "asset_category": category,
        }
        for field, index in indexes.items():
            if field == "asset_management_plan":
                record[field] = clean_text(values[index]) or None
            else:
                record[field] = numeric(values[index])
        rows.append(record)
    return rows


def safe_divide(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def combine() -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    by_key: dict[tuple[str, str, int], dict[str, object]] = {}
    type_sets: defaultdict[str, set[str]] = defaultdict(set)
    for schedule in SCHEDULE_COLUMNS:
        for year in YEARS:
            rows = read_schedule(schedule, year)
            for row in rows:
                name = str(row["municipality_name"])
                type_code = str(row["municipality_type"])
                key = (name, type_code, year)
                type_sets[name].add(type_code)
                target = by_key.setdefault(key, {})
                for field, value in row.items():
                    if field in target and target[field] not in (None, "") and value not in (None, "") and target[field] != value:
                        raise ValueError(f"conflicting {field} for {key}: {target[field]} vs {value}")
                    if value not in (None, "") or field not in target:
                        target[field] = value

    assets_all: list[dict[str, object]] = []
    asset_totals: dict[tuple[str, str, int], dict[str, object]] = {}
    for year in YEARS:
        for row in read_assets(year):
            name = str(row["municipality_name"])
            type_code = str(row["municipality_type"])
            type_sets[name].add(type_code)
            if row["asset_category"] == "Total TCA":
                asset_totals[(name, type_code, year)] = row
            if year == 2024:
                assets_all.append(row)

    duplicate_names = {name for name, types in type_sets.items() if len(types) > 1}
    municipal_rows: list[dict[str, object]] = []
    for (name, type_code, year), row in sorted(by_key.items(), key=lambda item: (item[0][0], item[0][1], item[0][2])):
        if name == "Jumbo Glacier" and year > 2020:
            continue
        total_asset = asset_totals.get((name, type_code, year), {})
        for field in (
            "historical_cost",
            "accumulated_amortization",
            "net_book_value",
            "asset_additions",
            "replacement_value",
            "replacement_value_amortization",
        ):
            row[f"tca_{field}"] = total_asset.get(field)
        row["municipality_id"] = municipality_id(name, type_code)
        row["selection_label"] = selection_label(name, type_code, duplicate_names)
        row["key"] = f"{row['municipality_id']}|{year}"
        row["selection_year_key"] = f"{row['selection_label']}|{year}"
        population = row.get("population")
        if population is not None:
            if population < 5_000:
                row["population_band"] = "Under 5,000"
            elif population < 20_000:
                row["population_band"] = "5,000 to 19,999"
            elif population < 100_000:
                row["population_band"] = "20,000 to 99,999"
            else:
                row["population_band"] = "100,000 and over"
        else:
            row["population_band"] = "Unavailable"
        reserves = sum((row.get(k) or 0) for k in ("water_reserves", "sewer_reserves", "other_reserves"))
        row["total_reserves"] = reserves
        row["government_transfers"] = sum(
            (row.get(k) or 0)
            for k in ("federal_transfers", "provincial_transfers", "regional_transfers")
        )
        row["investment_and_other_revenue"] = sum(
            (row.get(k) or 0)
            for k in (
                "investment_income",
                "government_business_income",
                "developer_contributions",
                "gain_on_sale_assets",
                "other_revenue",
            )
        )
        row["cash_operating_expense"] = None
        if row.get("total_expenses") is not None:
            row["cash_operating_expense"] = (
                row["total_expenses"]
                - (row.get("amortization") or 0)
                - (row.get("interest_expense") or 0)
            )
        row["annual_surplus"] = None
        if row.get("total_revenue") is not None and row.get("total_expenses") is not None:
            row["annual_surplus"] = row["total_revenue"] - row["total_expenses"]
        row["operating_margin"] = safe_divide(row.get("annual_surplus"), row.get("total_revenue"))
        row["revenue_per_capita"] = safe_divide(row.get("total_revenue"), population)
        row["expense_per_capita"] = safe_divide(row.get("total_expenses"), population)
        row["net_financial_assets_per_capita"] = safe_divide(row.get("net_financial_assets"), population)
        row["reserves_per_capita"] = safe_divide(reserves, population)
        row["debt_per_capita"] = safe_divide(row.get("closing_debt"), population)
        row["cash_to_liabilities"] = safe_divide(row.get("cash_and_investments"), row.get("total_liabilities"))
        row["capital_reinvestment_rate"] = safe_divide(row.get("tca_asset_additions"), row.get("tca_replacement_value"))
        replacement_value = row.get("tca_replacement_value")
        replacement_amortization = row.get("tca_replacement_value_amortization")
        row["asset_condition_ratio"] = None
        if replacement_value not in (None, 0) and replacement_amortization is not None:
            row["asset_condition_ratio"] = 1 - (replacement_amortization / replacement_value)
        row["debt_service_headroom_pct"] = safe_divide(row.get("debt_servicing_capacity"), row.get("liability_servicing_limit"))
        row["tax_collection_rate"] = safe_divide(row.get("total_taxes_collected"), row.get("taxes_imposed"))
        municipal_rows.append(row)

    # Attach transparent 2024 peer and province medians to each row for workbook lookups.
    benchmark_metrics = (
        "revenue_per_capita",
        "expense_per_capita",
        "net_financial_assets_per_capita",
        "reserves_per_capita",
        "debt_per_capita",
        "cash_to_liabilities",
        "capital_reinvestment_rate",
        "asset_condition_ratio",
        "debt_service_headroom_pct",
        "tax_collection_rate",
        "operating_margin",
    )
    latest = [row for row in municipal_rows if row["year"] == 2024]
    for row in latest:
        peers = [
            candidate
            for candidate in latest
            if candidate["municipality_type"] == row["municipality_type"]
            and candidate["population_band"] == row["population_band"]
        ]
        for metric in benchmark_metrics:
            peer_values = [float(candidate[metric]) for candidate in peers if candidate.get(metric) is not None]
            province_values = [float(candidate[metric]) for candidate in latest if candidate.get(metric) is not None]
            row[f"peer_median_{metric}"] = median(peer_values) if peer_values else None
            row[f"bc_median_{metric}"] = median(province_values) if province_values else None

    id_lookup = {(row["municipality_name"], row["municipality_type"]): row for row in latest}
    for row in assets_all:
        identity = (row["municipality_name"], row["municipality_type"])
        latest_row = id_lookup.get(identity)
        row["municipality_id"] = municipality_id(str(row["municipality_name"]), str(row["municipality_type"]))
        row["selection_label"] = latest_row["selection_label"] if latest_row else selection_label(str(row["municipality_name"]), str(row["municipality_type"]), duplicate_names)
        row["key"] = f"{row['municipality_id']}|{row['asset_category']}"
        row["selection_asset_key"] = f"{row['selection_label']}|{row['asset_category']}"
        row["net_book_value_ratio"] = safe_divide(row.get("net_book_value"), row.get("historical_cost"))
        row["asset_condition_ratio"] = None
        if row.get("replacement_value") not in (None, 0) and row.get("replacement_value_amortization") is not None:
            row["asset_condition_ratio"] = 1 - (row["replacement_value_amortization"] / row["replacement_value"])
    return municipal_rows, assets_all


def write_csv(path: Path, rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    columns: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for key in row:
            if key not in seen:
                columns.append(key)
                seen.add(key)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    municipal_rows, asset_rows = combine()
    write_csv(SILVER / "municipal_finance.csv", municipal_rows)
    write_csv(SILVER / "asset_categories_2024.csv", asset_rows)
    (SILVER / "municipal_finance.json").write_text(
        json.dumps(municipal_rows, ensure_ascii=False), encoding="utf-8"
    )
    (SILVER / "asset_categories_2024.json").write_text(
        json.dumps(asset_rows, ensure_ascii=False), encoding="utf-8"
    )
    latest = [row for row in municipal_rows if row["year"] == 2024]
    summary = {
        "years": list(YEARS),
        "municipality_year_rows": len(municipal_rows),
        "municipalities_2024": len(latest),
        "asset_rows_2024": len(asset_rows),
        "source_files": len(list(RAW.glob("*.xlsx"))),
        "default_municipality": "Vancouver",
        "source_page": "https://www2.gov.bc.ca/gov/content/governments/local-governments/facts-framework/statistics/statistics",
    }
    (SILVER / "summary.json").write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
