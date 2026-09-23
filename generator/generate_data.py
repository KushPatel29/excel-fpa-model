"""Generate the source data for the Kestrel Bay Provisions FP&A model.

Kestrel Bay Provisions is fictional: a BC specialty-food distributor selling
24 products (beef, pork, poultry, seafood, charcuterie, cheese and dairy)
through four channels in three regions. Everything here is synthetic and
seeded, so every figure the workbook shows can be rebuilt byte for byte.

The data carries a story on purpose, so the analysis has something true to
find. Against a budget set in late 2025, fiscal 2026 goes like this:

* a new grocery listing lifts Grocery Retail volume ~20% (budget: +3%). That
  is volume the business wanted, sold at the channel's lower price;
* beef costs rise 6% in January and 10% from May (budget: +3%);
* seafood list prices rise 7% (budget: +4%);
* BC Interior freight rises 12% (budget: +2%);
* warehouse overtime puts Warehouse & Logistics opex ~8% over budget;
* specialty cheese sells ~3% under 2025 while purchasing follows the
  budget, so its days of inventory creep toward its shelf life.

Revenue beats budget; margin does not. The price-volume-mix bridge has to
show why, and the rolling forecast has to price the gap.

Only integer draws and basic arithmetic are used for noise: transcendental
functions can differ in the last bit between platforms, and the CI job
regenerates this data on Linux and requires identical bytes.
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

SEED = 29
AS_OF = "2026-08-01"              # last closed month
FIRST_MONTH = "2024-01-01"
FY_BUDGET = 2026

# sku, product, category, case description, 2024 list price per case,
# cost-to-list ratio, 2024 base cases per month (all channels and regions)
PRODUCTS = [
    ("B01", "Striploin AAA", "Beef", "12 kg case", 268.00, 0.740, 1040),
    ("B02", "Ground Chuck", "Beef", "10 kg case", 118.00, 0.735, 2800),
    ("B03", "Beef Short Rib", "Beef", "10 kg case", 196.00, 0.745, 1220),
    ("B04", "Beef Tenderloin", "Beef", "8 kg case", 342.00, 0.750, 580),
    ("P01", "Pork Belly", "Pork", "10 kg case", 132.00, 0.700, 1520),
    ("P02", "Pork Tenderloin", "Pork", "8 kg case", 104.00, 0.690, 1080),
    ("P03", "Back Ribs", "Pork", "10 kg case", 146.00, 0.705, 960),
    ("P04", "Pork Shoulder", "Pork", "12 kg case", 92.00, 0.695, 1380),
    ("C01", "Whole Chicken", "Poultry", "12 pc case", 78.00, 0.725, 2500),
    ("C02", "Chicken Breast", "Poultry", "10 kg case", 124.00, 0.720, 2200),
    ("C03", "Chicken Thigh", "Poultry", "10 kg case", 96.00, 0.715, 1800),
    ("C04", "Duck Breast", "Poultry", "5 kg case", 188.00, 0.700, 420),
    ("S01", "Atlantic Salmon", "Seafood", "10 kg case", 176.00, 0.755, 1120),
    ("S02", "Halibut Fillet", "Seafood", "5 kg case", 214.00, 0.750, 480),
    ("S03", "Spot Prawns", "Seafood", "5 kg case", 305.00, 0.745, 300),
    ("S04", "Ling Cod", "Seafood", "5 kg case", 138.00, 0.750, 560),
    ("H01", "Prosciutto", "Charcuterie", "6 kg case", 188.00, 0.580, 660),
    ("H02", "Soppressata", "Charcuterie", "5 kg case", 156.00, 0.575, 580),
    ("H03", "Bresaola", "Charcuterie", "4 kg case", 204.00, 0.590, 300),
    ("H04", "Country Pate", "Charcuterie", "3 kg case", 96.00, 0.570, 520),
    ("D01", "Aged Cheddar", "Cheese & Dairy", "10 kg case", 142.00, 0.620, 840),
    ("D02", "Double Cream Brie", "Cheese & Dairy", "6 kg case", 118.00, 0.615, 720),
    ("D03", "Parmigiano Reggiano", "Cheese & Dairy", "10 kg case", 236.00, 0.630, 380),
    ("D04", "Cultured Butter", "Cheese & Dairy", "12 x 454 g", 86.00, 0.610, 1040),
]

CATEGORIES = ["Beef", "Pork", "Poultry", "Seafood", "Charcuterie", "Cheese & Dairy"]

# category, shelf life (days), base days of inventory
CATEGORY_STOCK = {
    "Beef": (35, 16), "Pork": (30, 13), "Poultry": (14, 6),
    "Seafood": (10, 4), "Charcuterie": (120, 44), "Cheese & Dairy": (90, 36),
}

# channel, share of volume, price index vs list, payment terms (days)
CHANNELS = [
    ("Restaurants", 0.44, 1.00, 30),
    ("Grocery Retail", 0.30, 0.88, 45),
    ("Hotels & Institutions", 0.18, 0.95, 35),
    ("Online Direct", 0.08, 1.12, 2),
]

# region, share of volume, 2024 freight per case
REGIONS = [
    ("Lower Mainland", 0.58, 1.60),
    ("Vancouver Island", 0.21, 3.40),
    ("BC Interior", 0.21, 3.90),
]
ONLINE_COURIER_PER_CASE = 8.50

# Seasonality, January..December (normalised below so each averages 1.0).
SEASONALITY = {
    "Beef":           [0.85, 0.85, 0.92, 1.00, 1.10, 1.18, 1.20, 1.15, 1.00, 0.95, 0.90, 0.95],
    "Pork":           [0.95, 0.90, 0.95, 1.00, 1.05, 1.05, 1.05, 1.05, 1.00, 0.98, 1.00, 1.10],
    "Poultry":        [1.00, 0.98, 1.00, 1.00, 1.02, 1.03, 1.03, 1.02, 1.00, 0.98, 1.02, 1.10],
    "Seafood":        [0.80, 0.80, 0.90, 1.00, 1.25, 1.30, 1.20, 1.10, 1.00, 0.90, 0.85, 0.95],
    "Charcuterie":    [0.90, 0.85, 0.90, 0.95, 0.95, 0.95, 0.95, 0.95, 1.00, 1.10, 1.25, 1.40],
    "Cheese & Dairy": [0.95, 0.90, 0.95, 0.95, 0.95, 0.95, 0.95, 0.95, 1.00, 1.05, 1.15, 1.30],
}

# ---- what actually happened ------------------------------------------------
VOLUME_2025 = 1.03                                  # +3% on 2024, every segment
VOLUME_2026_BY_CATEGORY = {"Beef": 0.99, "Pork": 1.02, "Poultry": 1.03,
                           "Seafood": 1.04, "Charcuterie": 1.03, "Cheese & Dairy": 0.97}
GROCERY_LISTING_2026 = 1.20                         # new grocery chain from February
LIST_PRICE_2025 = 1.03
LIST_PRICE_2026 = {"Beef": 1.05, "Seafood": 1.07}   # everything else +4%
COST_2025 = 1.03
BEEF_COST_2026 = (1.06, 1.10)                       # Jan-Apr, then from May
COST_2026_OTHER = 1.03
FREIGHT_2025 = 1.02
FREIGHT_2026 = {"BC Interior": 1.12}                # everywhere else +3%

# ---- what the budget assumed (set in November 2025) -------------------------
BUDGET_VOLUME = {"Beef": 1.02, "Pork": 1.02, "Poultry": 1.03,
                 "Seafood": 1.05, "Charcuterie": 1.04, "Cheese & Dairy": 1.03}
BUDGET_PRICE = 1.04
BUDGET_COST = 1.03
BUDGET_FREIGHT = 1.02

# ---- operating expenses (monthly, 2024 base) --------------------------------
DEPARTMENTS = [
    ("Sales & Marketing", 190_000),
    ("Warehouse & Logistics", 260_000),
    ("General & Admin", 170_000),
    ("Technology", 55_000),
]
OPEX_2025 = 1.03
OPEX_BUDGET_2026 = 1.03                             # on the 2025 run-rate
OPEX_ACTUAL_VS_BUDGET_2026 = {"Sales & Marketing": 0.98, "Warehouse & Logistics": 1.08,
                              "General & Admin": 1.01, "Technology": 1.00}
DAYS_PER_MONTH = 30.4


def noise(rng: np.random.Generator, size, half_width_per_mille: int) -> np.ndarray:
    """Multiplicative noise from integer draws: 1 +/- half_width per mille."""
    return 1 + rng.integers(-half_width_per_mille, half_width_per_mille + 1, size=size) / 1000


def season(category: str, month_number: int) -> float:
    profile = SEASONALITY[category]
    return profile[month_number - 1] * 12 / sum(profile)


def build_sales(rng: np.random.Generator, months: pd.DatetimeIndex) -> pd.DataFrame:
    rows = []
    for month in months:
        year, m = month.year, month.month
        for sku, _name, category, _case, list_2024, cost_ratio, base_cases in PRODUCTS:
            # list price and standard unit cost for the period
            list_price = list_2024
            unit_cost = list_2024 * cost_ratio
            if year >= 2025:
                list_price *= LIST_PRICE_2025
                unit_cost *= COST_2025
            if year >= 2026:
                list_price *= LIST_PRICE_2026.get(category, 1.04)
                if category == "Beef":
                    unit_cost *= BEEF_COST_2026[0] if m <= 4 else BEEF_COST_2026[1]
                else:
                    unit_cost *= COST_2026_OTHER

            volume = base_cases * season(category, m)
            if year >= 2025:
                volume *= VOLUME_2025
            if year >= 2026:
                volume *= VOLUME_2026_BY_CATEGORY[category]

            for channel, ch_share, price_index, _terms in CHANNELS:
                ch_volume = volume * ch_share
                if year >= 2026 and channel == "Grocery Retail" and m >= 2:
                    ch_volume *= GROCERY_LISTING_2026
                for region, rg_share, freight_2024 in REGIONS:
                    rows.append((month, sku, category, channel, region,
                                 ch_volume * rg_share, list_price, price_index,
                                 unit_cost, freight_2024))

    df = pd.DataFrame(rows, columns=["month", "sku", "category", "channel", "region",
                                     "volume", "list_price", "price_index",
                                     "unit_cost", "freight_2024"])
    n = len(df)
    df["units"] = np.rint(df["volume"].to_numpy() * noise(rng, n, 45)).astype(int)
    df["units"] = df["units"].clip(lower=1)

    # promotional depth: grocery runs promotions, other channels discount lightly
    promo = np.where(df["channel"] == "Grocery Retail",
                     rng.integers(0, 41, size=n), rng.integers(0, 16, size=n)) / 1000
    df["discount_pct"] = promo
    df["list_price"] = df["list_price"].round(2)
    net_price = (df["list_price"] * df["price_index"] * (1 - df["discount_pct"])
                 * noise(rng, n, 4))
    df["net_revenue"] = (df["units"] * net_price).round(2)
    df["unit_cost"] = (df["unit_cost"] * noise(rng, n, 3)).round(2)
    df["cogs"] = (df["units"] * df["unit_cost"]).round(2)

    freight = df["freight_2024"].copy()
    freight = np.where(df["month"].dt.year >= 2025, freight * FREIGHT_2025, freight)
    y2026 = df["month"].dt.year >= 2026
    uplift = np.where(df["region"] == "BC Interior", FREIGHT_2026["BC Interior"], 1.03)
    freight = np.where(y2026, freight * uplift, freight)
    freight = freight + np.where(df["channel"] == "Online Direct", ONLINE_COURIER_PER_CASE, 0.0)
    df["freight"] = (df["units"] * freight * noise(rng, n, 30)).round(2)

    df["month"] = df["month"].dt.strftime("%Y-%m-%d")
    return df[["month", "sku", "category", "channel", "region", "units", "list_price",
               "discount_pct", "net_revenue", "unit_cost", "cogs", "freight"]]


def build_budget(sales: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """The FY2026 budget, built the way an FP&A team builds one from FY2025."""
    s25 = sales[sales["month"].str.startswith("2025")].copy()
    s25["m"] = s25["month"].str[5:7].astype(int)
    keys = ["sku", "channel", "region"]
    growth = s25["category"].map(BUDGET_VOLUME)
    s25["budget_units"] = np.rint(s25["units"] * growth).astype(int)

    wide = s25.pivot_table(index=keys, columns="m", values="budget_units", aggfunc="sum")
    wide.columns = [f"{FY_BUDGET}-{m:02d}" for m in wide.columns]
    wide = wide.reset_index()

    agg = s25.groupby(keys).agg(units=("units", "sum"), revenue=("net_revenue", "sum"),
                                cogs=("cogs", "sum"), freight=("freight", "sum")).reset_index()
    rates = agg[keys].copy()
    rates["net_price"] = (agg["revenue"] / agg["units"] * BUDGET_PRICE).round(2)
    rates["unit_cost"] = (agg["cogs"] / agg["units"] * BUDGET_COST).round(2)
    rates["freight_per_case"] = (agg["freight"] / agg["units"] * BUDGET_FREIGHT).round(2)
    return wide, rates


def build_opex(rng: np.random.Generator, months: pd.DatetimeIndex) -> pd.DataFrame:
    rows = []
    run_rate_2025 = {}
    for dept, base in DEPARTMENTS:
        seasonal = {m: (1.06 if (dept == "Warehouse & Logistics" and m in (11, 12)) else 1.0)
                    for m in range(1, 13)}
        amounts_2025 = []
        for month in months:
            amount = base * seasonal[month.month]
            if month.year >= 2025:
                amount *= OPEX_2025
            if month.year >= 2026:
                amount *= OPEX_BUDGET_2026 * OPEX_ACTUAL_VS_BUDGET_2026[dept]
            amount = round(amount * float(noise(rng, 1, 15)[0]), 2)
            rows.append((month.strftime("%Y-%m-%d"), dept, "Actual", amount))
            if month.year == 2025:
                amounts_2025.append(amount)
        run_rate_2025[dept] = sum(amounts_2025) / len(amounts_2025)
        # budget: the 2025 monthly run-rate, grown, with the Q4 warehouse peak kept
        for m in range(1, 13):
            amount = run_rate_2025[dept] * OPEX_BUDGET_2026 * seasonal[m] / (
                sum(seasonal.values()) / 12)
            rows.append((f"{FY_BUDGET}-{m:02d}-01", dept, "Budget", round(amount, 2)))
    return pd.DataFrame(rows, columns=["month", "dept", "scenario", "amount"])


def build_balances(rng: np.random.Generator, sales: pd.DataFrame) -> pd.DataFrame:
    rows = []
    terms = {c[0]: c[3] for c in CHANNELS}
    by_channel = sales.groupby(["month", "channel"])["net_revenue"].sum()
    by_category_cogs = sales.groupby(["month", "category"])["cogs"].sum()
    total_cogs = sales.groupby("month")["cogs"].sum()
    for month in sorted(sales["month"].unique()):
        year, m = int(month[:4]), int(month[5:7])
        ar = sum(by_channel[(month, ch)] * terms[ch] / DAYS_PER_MONTH for ch in terms)
        rows.append((month, "Accounts receivable", "", round(ar * float(noise(rng, 1, 30)[0]), 2)))
        for category in CATEGORIES:
            dio = CATEGORY_STOCK[category][1]
            if category == "Cheese & Dairy" and year >= 2026:
                dio += 2 * m                        # slow sellers, bought to budget
            value = by_category_cogs[(month, category)] * dio / DAYS_PER_MONTH
            rows.append((month, "Inventory", category,
                         round(value * float(noise(rng, 1, 25)[0]), 2)))
        ap = total_cogs[month] * 28 / DAYS_PER_MONTH
        rows.append((month, "Accounts payable", "", round(ap * float(noise(rng, 1, 30)[0]), 2)))
    return pd.DataFrame(rows, columns=["month", "account", "category", "balance"])


def write_csv(df: pd.DataFrame, name: str) -> None:
    df.to_csv(DATA / name, index=False, lineterminator="\n", float_format="%.2f")


def main() -> None:
    DATA.mkdir(exist_ok=True)
    rng = np.random.default_rng(SEED)
    months = pd.date_range(FIRST_MONTH, AS_OF, freq="MS")

    sales = build_sales(rng, months)
    budget_units, budget_rates = build_budget(sales)
    opex = build_opex(rng, months)
    balances = build_balances(rng, sales)

    product = pd.DataFrame(
        [(p[0], p[1], p[2], p[3], CATEGORY_STOCK[p[2]][0]) for p in PRODUCTS],
        columns=["sku", "product", "category", "case_description", "shelf_life_days"])
    channel = pd.DataFrame([(c[0], c[2], c[3]) for c in CHANNELS],
                           columns=["channel", "price_index", "payment_terms_days"])
    region = pd.DataFrame([(r[0], r[2]) for r in REGIONS],
                          columns=["region", "freight_per_case_2024"])

    write_csv(product, "dim_product.csv")
    write_csv(channel, "dim_channel.csv")
    write_csv(region, "dim_region.csv")
    write_csv(sales, "fact_sales.csv")
    write_csv(budget_units, "budget_units_fy2026.csv")
    write_csv(budget_rates, "budget_rates_fy2026.csv")
    write_csv(opex, "fact_opex.csv")
    write_csv(balances, "balances.csv")

    manifest = {
        "company": "Kestrel Bay Provisions (fictional)",
        "seed": SEED,
        "as_of": AS_OF,
        "fiscal_year": FY_BUDGET,
        "rows": {
            "fact_sales": len(sales), "budget_units_fy2026": len(budget_units),
            "budget_rates_fy2026": len(budget_rates), "fact_opex": len(opex),
            "balances": len(balances), "dim_product": len(product),
        },
    }
    (DATA / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n",
                                        encoding="utf-8", newline="\n")
    print(json.dumps(manifest["rows"]))


if __name__ == "__main__":
    main()
