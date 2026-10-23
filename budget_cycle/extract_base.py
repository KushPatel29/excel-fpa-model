"""
Extract the FY2027 budget's starting point from the Kestrel Bay workbook.

    python budget_cycle/extract_base.py

Reads the source tables the Kestrel Bay workbook carries (Data_Sales, Data_Opex,
Data_Balances) and writes the three things a budget is built from:

    base_segments.csv   last twelve months of actuals (Sep 2025 - Aug 2026) by
                        category x channel: cases, net revenue, COGS, freight
    seasonality.csv     each category's share of its calendar-2025 cases by month
    opex_base.csv       last twelve months of operating expenses by department
    opening_wc.csv      receivables, inventory and payables at the latest
                        month-end (Aug 2026), which open the budget's balance sheet

Kestrel Bay is synthetic, so everything downstream is too.
"""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import openpyxl
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "examples" / "kestrel-bay" / "Kestrel_Bay_FPA_Model.xlsx"
OUT = Path(__file__).resolve().parent / "data"
LTM_START, LTM_END = "2025-09-01", "2026-08-01"


def table(wb, name: str) -> pd.DataFrame:
    rows = [r[1:] for r in wb[name].iter_rows(min_row=4, values_only=True)]
    header = [h for h in rows[0] if h is not None]
    return pd.DataFrame([r[: len(header)] for r in rows[1:] if r[0] is not None], columns=header)


def main() -> int:
    wb = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
    sales, opex, bal = table(wb, "Data_Sales"), table(wb, "Data_Opex"), table(wb, "Data_Balances")
    for f in (sales, opex, bal):
        f["month"] = pd.to_datetime(f["month"])

    ltm = sales[(sales.month >= LTM_START) & (sales.month <= LTM_END)]
    seg = (ltm.groupby(["category", "channel"], sort=False)[["units", "net_revenue", "cogs", "freight"]]
           .sum().round(2).reset_index().rename(columns={"units": "cases", "net_revenue": "revenue"}))
    seg.to_csv(OUT / "base_segments.csv", index=False, lineterminator="\n")

    cy = sales[sales.month.dt.year == 2025].assign(m=lambda f: f.month.dt.month)
    by = cy.groupby(["category", "m"])["units"].sum().unstack("m")
    share = by.div(by.sum(axis=1), axis=0).round(6)
    share.columns = [f"m{m:02d}" for m in share.columns]
    # December takes the rounding remainder, so each row sums to exactly 1.
    head = share.iloc[:, :-1].map(lambda v: Decimal(str(v))).sum(axis=1)
    share["m12"] = [float(Decimal(1) - h) for h in head]
    share.reset_index().to_csv(OUT / "seasonality.csv", index=False, lineterminator="\n")

    ox = opex[(opex.scenario == "Actual") & (opex.month >= LTM_START) & (opex.month <= LTM_END)]
    ox.groupby("dept", sort=False)["amount"].sum().round(2).reset_index().to_csv(
        OUT / "opex_base.csv", index=False, lineterminator="\n")

    last = bal[bal.month == bal.month.max()].groupby("account", sort=False)["balance"].sum().round(2)
    last.reset_index().to_csv(OUT / "opening_wc.csv", index=False, lineterminator="\n")
    print(f"{len(seg)} segments, LTM revenue {seg.revenue.sum():,.2f}; opex {ox.amount.sum():,.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
