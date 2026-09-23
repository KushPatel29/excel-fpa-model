"""Build workbook/Kestrel_Bay_FPA_Model.xlsx by driving desktop Excel over COM.

    python model/build_workbook.py [--visible] [--no-pdf]

Needs Windows and Microsoft 365 Excel. Python only lays the workbook out:
every figure in it is an Excel formula, a Power Query step, a Power Pivot
measure or a what-if data table over the CSVs in data/, and tests/ holds the
saved workbook to model/reference.py. The workbook's content is deterministic;
its bytes are not (Excel stamps times), so CI checks values, never bytes.
"""
from __future__ import annotations

import argparse
import csv
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client
from win32com.client import gencache

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import build_analysis as analysis  # noqa: E402
import build_calc as calc  # noqa: E402
import build_setup as setup  # noqa: E402
import build_views as views  # noqa: E402
import layout as L  # noqa: E402
from xl import XL_CALC_AUTOMATIC, XL_CALC_MANUAL, XL_WORKBOOK  # noqa: E402

ROOT = HERE.parent
OUT = ROOT / "workbook" / L.WORKBOOK
PDF = ROOT / "docs" / "board_pack.pdf"


def control_totals() -> dict:
    """Totals from the source CSVs, keyed into the Checks sheet like a system report."""
    with open(ROOT / "data" / "fact_sales.csv", newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    return {
        "ctrl_sales_rows": ("Sales rows", len(rows), "data/fact_sales.csv, row count"),
        "ctrl_sales_units": ("Sales cases", sum(int(r["units"]) for r in rows),
                             "data/fact_sales.csv, sum of units"),
        "ctrl_sales_revenue": ("Sales net revenue", round(sum(float(r["net_revenue"]) for r in rows), 2),
                               "data/fact_sales.csv, sum of net_revenue"),
    }


def start_excel(visible: bool):
    pythoncom.CoInitialize()
    # DispatchEx: a private Excel process, so an Excel the user has open is never touched.
    xl = gencache.EnsureDispatch(win32com.client.DispatchEx("Excel.Application")._oleobj_)
    xl.Visible = visible
    xl.DisplayAlerts = False
    xl.ScreenUpdating = visible
    xl.EnableEvents = False
    return xl


BUSY = (-2147418111, -2147417846)   # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER


def patient(fn, *args, seconds: float = 300):
    """Call into Excel, retrying while it is busy. pywin32 cannot register a COM
    message filter, and CUBE functions keep Excel busy after a calculation."""
    deadline = time.time() + seconds
    while True:
        try:
            return fn(*args)
        except pywintypes.com_error as e:
            if e.hresult not in BUSY or time.time() > deadline:
                raise
            time.sleep(0.5)


def calculate(xl) -> None:
    """Full calculation, then wait for the CUBE functions, which evaluate asynchronously.
    No RefreshAll: every query loaded when its connection was created, and a refresh
    with automatic calculation re-runs the data tables once per query."""
    xl.Calculation = XL_CALC_AUTOMATIC
    patient(xl.CalculateFull)
    patient(xl.CalculateUntilAsyncQueriesDone)
    patient(xl.CalculateFull)
    patient(xl.CalculateUntilAsyncQueriesDone)


def step(label: str, t0: float) -> None:
    print(f"[{time.time() - t0:6.1f}s] {label}", flush=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--visible", action="store_true", help="show Excel while it builds")
    ap.add_argument("--no-pdf", action="store_true", help="skip the board-pack PDF")
    args = ap.parse_args()

    t0 = time.time()
    xl = start_excel(args.visible)
    wb = None
    try:
        wb = xl.Workbooks.Add()
        xl.Calculation = XL_CALC_MANUAL
        setup.create_sheets(wb)
        setup.theme(wb)
        step("sheets", t0)
        setup.data_sheets(wb)
        step("source tables", t0)
        setup.assumptions(wb)
        setup.scenario_inputs(wb)
        setup.lambdas(wb)
        step("inputs, names and LAMBDAs", t0)
        setup.power_query(wb)
        step("Power Query", t0)
        setup.data_model(wb)
        analysis.channel_pivot(wb)
        step("data model, PivotTable and slicers", t0)
        pos: dict = {}
        calc.forecast(wb, pos)
        step("Forecast", t0)
        calc.pnl(wb, pos)
        step("PnL", t0)
        calc.pvm(wb, pos)
        step("PVM", t0)
        analysis.scenarios(wb, pos)
        step("Scenarios", t0)
        analysis.working_capital(wb, pos)
        step("WorkingCapital", t0)
        analysis.channels(wb, pos)
        step("Channels", t0)
        analysis.checks(wb, pos, control_totals())
        step("Checks", t0)
        views.dashboard(wb, pos)
        step("Dashboard", t0)
        views.cover(wb, pos)
        step("Cover", t0)
        calculate(xl)
        step("calculated", t0)
        views.finish(xl, wb, pos)
        wb.BuiltinDocumentProperties("Title").Value = f"{L.COMPANY}: FY2026 FP&A model"
        wb.BuiltinDocumentProperties("Author").Value = "Kush Patel"
        wb.BuiltinDocumentProperties("Subject").Value = (
            "Budget vs actual, price-volume-mix, rolling forecast, scenarios, working capital")
        OUT.parent.mkdir(exist_ok=True)
        wb.SaveAs(str(OUT), XL_WORKBOOK)
        step(f"saved {OUT.relative_to(ROOT)}", t0)
        print("checks:", patient(lambda: wb.Worksheets("Checks").Range("B2").Value))
        if not args.no_pdf:
            PDF.parent.mkdir(exist_ok=True)
            views.export_pdf(wb, PDF)
            step(f"exported {PDF.relative_to(ROOT)}", t0)
        wb.Close(False)
        wb = None
    finally:
        if wb is not None and not args.visible:
            wb.Close(False)
        if not args.visible:
            xl.Quit()


if __name__ == "__main__":
    main()
