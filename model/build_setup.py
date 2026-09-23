"""Workbook skeleton: sheets, source tables, inputs, named functions, Power Query
and the Power Pivot data model. Everything the calculation sheets stand on.
"""
from __future__ import annotations

import csv
from pathlib import Path

import layout as L
from xl import (
    AMBER, BODY_FONT, F_DATE, F_MONTH_LONG, GREY, MUTED, NAVY, TEAL, UNFAV,
    FAV, MSO_THEME_ACCENT, XL_BETWEEN, XL_SRC_MODEL, XL_SRC_RANGE, XL_VALID_ALERT_STOP,
    XL_VALIDATE_DATE, XL_VALIDATE_DECIMAL, XL_VALIDATE_LIST, XL_YES, band,
    font, fx, header, input_cell, numfmt, put, rgb, serial, title, widths,
)

DATA = Path(__file__).resolve().parent.parent / "data"

TAB_COLOURS = {"Cover": NAVY, "Dashboard": NAVY, "Checks": FAV, "Assumptions": AMBER}

# sheet, table, csv, column kinds, per-column formats, title
SOURCES = [
    ("Data_Sales", "tbl_Sales", "fact_sales.csv",
     "date text text text text int num num num num num num",
     {"discount_pct": "0.0%"}, "Sales by month, SKU, channel and region"),
    ("Data_BudgetUnits", "tbl_BudgetUnits", "budget_units_fy2026.csv",
     "text text text" + " int" * 12, {}, "FY2026 budget: cases by SKU, channel and region (wide)"),
    ("Data_BudgetRates", "tbl_BudgetRates", "budget_rates_fy2026.csv",
     "text text text num num num", {}, "FY2026 budget rates per case"),
    ("Data_Opex", "tbl_Opex", "fact_opex.csv", "date text text num", {},
     "Operating expenses by department: actual and budget"),
    ("Data_Balances", "tbl_Balances", "balances.csv", "date text text num", {},
     "Month-end balances: receivables, inventory by category, payables"),
]
DIMS = [  # table, csv, kinds, formats, first column
    ("tbl_Product", "dim_product.csv", "text text text text int", {}, "B"),
    ("tbl_Channel", "dim_channel.csv", "text num int", {"price_index": "0.00"}, "H"),
    ("tbl_Region", "dim_region.csv", "text num", {}, "L"),
]
KIND_FORMAT = {"date": F_DATE, "int": "#,##0", "num": "#,##0.00", "text": "@"}

# Named LAMBDA functions: readable formulas instead of repeated IF(d=0,...) boilerplate.
LAMBDAS = {
    "SAFEDIV": ("=LAMBDA(n,d,IF(d=0,0,n/d))",
                "n ÷ d, or 0 when d is 0."),
    "FAVVAR": ("=LAMBDA(actual,plan,is_cost,IF(is_cost,plan-actual,actual-plan))",
               "Variance signed so favourable is positive, for revenue and cost lines alike."),
    "ISCLOSED": ("=LAMBDA(month,month<=AsOfMonth)",
                 "TRUE when a month is closed (actuals), FALSE when it is forecast."),
    "DAYSOF": ("=LAMBDA(balance,flow,days,IF(flow=0,0,balance/flow*days))",
               "Days of flow a balance represents: DSO, DIO, DPO."),
    "MONEY": ('=LAMBDA(x,IF(ABS(x)>=999500,TEXT(x/1000000,"$#,##0.0")&"M",TEXT(x/1000,"$#,##0")&"K"))',
              "$1.2M / $350K style amounts for written commentary."),
    "SIGNMONEY": ('=LAMBDA(x,IF(x<0,"−","+")&MONEY(ABS(x)))',
                  "MONEY with an explicit sign."),
}

CONN = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={};Extended Properties=""'

M_SALES = """let
    Source = Excel.CurrentWorkbook(){[Name="tbl_Sales"]}[Content],
    Typed = Table.TransformColumnTypes(Source, {
        {"month", type date}, {"sku", type text}, {"category", type text},
        {"channel", type text}, {"region", type text}, {"units", Int64.Type},
        {"list_price", type number}, {"discount_pct", type number},
        {"net_revenue", type number}, {"unit_cost", type number},
        {"cogs", type number}, {"freight", type number}})
in
    Typed"""

M_BUDGET = """// The budget arrives wide (a column per month) and priced separately. Unpivot it,
// attach each row's rate and category, and price it: one long table the P&L,
// the price-volume-mix and the data model can all read.
let
    Units = Excel.CurrentWorkbook(){[Name="tbl_BudgetUnits"]}[Content],
    Long = Table.UnpivotOtherColumns(Units, {"sku", "channel", "region"}, "period", "units"),
    Dated = Table.AddColumn(Long, "month", each Date.FromText([period] & "-01"), type date),
    Rates = Excel.CurrentWorkbook(){[Name="tbl_BudgetRates"]}[Content],
    WithRates = Table.NestedJoin(Dated, {"sku", "channel", "region"}, Rates,
        {"sku", "channel", "region"}, "rate", JoinKind.LeftOuter),
    Rated = Table.ExpandTableColumn(WithRates, "rate", {"net_price", "unit_cost", "freight_per_case"}),
    Products = Excel.CurrentWorkbook(){[Name="tbl_Product"]}[Content],
    WithProduct = Table.NestedJoin(Rated, {"sku"}, Products, {"sku"}, "product", JoinKind.LeftOuter),
    Categorised = Table.ExpandTableColumn(WithProduct, "product", {"category"}),
    Revenue = Table.AddColumn(Categorised, "revenue", each [units] * [net_price], type number),
    Cogs = Table.AddColumn(Revenue, "cogs", each [units] * [unit_cost], type number),
    Freight = Table.AddColumn(Cogs, "freight", each [units] * [freight_per_case], type number),
    Selected = Table.SelectColumns(Freight, {"month", "sku", "category", "channel", "region",
        "units", "net_price", "unit_cost", "freight_per_case", "revenue", "cogs", "freight"}),
    Typed = Table.TransformColumnTypes(Selected, {{"sku", type text}, {"category", type text},
        {"channel", type text}, {"region", type text}, {"units", Int64.Type},
        {"net_price", type number}, {"unit_cost", type number}, {"freight_per_case", type number}}),
    Sorted = Table.Sort(Typed, {{"month", Order.Ascending}, {"sku", Order.Ascending},
        {"channel", Order.Ascending}, {"region", Order.Ascending}})
in
    Sorted"""

M_CALENDAR = """// Month calendar for the data model, three fiscal years ending at FiscalYear.
// status reads AsOfMonth from the Assumptions sheet, so Refresh All re-cuts the
// closed/open split after the as-of month changes.
let
    FiscalYear = Int64.From(Excel.CurrentWorkbook(){[Name="FiscalYear"]}[Content]{0}[Column1]),
    AsOf = Date.From(Excel.CurrentWorkbook(){[Name="AsOfMonth"]}[Content]{0}[Column1]),
    First = #date(FiscalYear - 2, 1, 1),
    Months = List.Transform({0..35}, each Date.AddMonths(First, _)),
    Base = Table.TransformColumnTypes(
        Table.FromList(Months, Splitter.SplitByNothing(), {"month"}), {{"month", type date}}),
    Year = Table.AddColumn(Base, "year", each Date.Year([month]), Int64.Type),
    MonthNo = Table.AddColumn(Year, "month_no", each Date.Month([month]), Int64.Type),
    MonthName = Table.AddColumn(MonthNo, "month_name", each Date.ToText([month], "MMM", "en-US"), type text),
    Quarter = Table.AddColumn(MonthName, "quarter", each "Q" & Text.From(Date.QuarterOfYear([month])), type text),
    Index = Table.AddColumn(Quarter, "month_index", each ([year] - Date.Year(First)) * 12 + [month_no], Int64.Type),
    Status = Table.AddColumn(Index, "status", each if [month] <= AsOf then "Closed" else "Open", type text)
in
    Status"""


def m_dimension(table: str, types: str) -> str:
    return (f'let\n    Source = Excel.CurrentWorkbook(){{[Name="{table}"]}}[Content],\n'
            f'    Typed = Table.TransformColumnTypes(Source, {{{types}}})\nin\n    Typed')


QUERIES = [  # name, M, loads a worksheet table too?
    ("Sales", M_SALES, False),
    ("Product", m_dimension("tbl_Product", '{"sku", type text}, {"product", type text}, '
                            '{"category", type text}, {"case_description", type text}, '
                            '{"shelf_life_days", Int64.Type}'), False),
    ("Channel", m_dimension("tbl_Channel", '{"channel", type text}, {"price_index", type number}, '
                            '{"payment_terms_days", Int64.Type}'), False),
    ("Region", m_dimension("tbl_Region", '{"region", type text}, '
                           '{"freight_per_case_2024", type number}'), False),
    ("Calendar", M_CALENDAR, False),
    ("Budget", M_BUDGET, True),
]

CLOSED = "'Calendar'[status] = \"Closed\""
MEASURES = [  # name, home table, DAX, format, description
    ("Revenue", "Sales", "SUM ( Sales[net_revenue] )", "money", "Net revenue after discounts."),
    ("Gross Margin", "Sales", "[Revenue] - SUM ( Sales[cogs] )", "money", "Revenue less cost of goods."),
    ("GM %", "Sales", "DIVIDE ( [Gross Margin], [Revenue] )", "pct", "Gross margin as a share of revenue."),
    ("Contribution", "Sales", "[Gross Margin] - SUM ( Sales[freight] )", "money",
     "Gross margin less outbound freight: what a channel earns before overheads."),
    ("Contribution %", "Sales", "DIVIDE ( [Contribution], [Revenue] )", "pct", ""),
    ("Cases", "Sales", "SUM ( Sales[units] )", "int", "Cases shipped."),
    ("Contribution per Case", "Sales", "DIVIDE ( [Contribution], [Cases] )", "dec", ""),
    ("Budget Revenue", "Budget", f"CALCULATE ( SUM ( Budget[revenue] ), {CLOSED} )", "money",
     "Budget revenue for closed months only, so it compares like for like with actuals."),
    ("Revenue vs Budget %", "Budget", "DIVIDE ( [Revenue] - [Budget Revenue], [Budget Revenue] )", "pct", ""),
    ("Budget Contribution", "Budget",
     f"CALCULATE ( SUM ( Budget[revenue] ) - SUM ( Budget[cogs] ) - SUM ( Budget[freight] ), {CLOSED} )",
     "money", "Budget contribution for closed months."),
    ("Contribution vs Budget", "Budget", "[Contribution] - [Budget Contribution]", "money", ""),
    ("Revenue PY", "Sales",
     "VAR shifted = SELECTCOLUMNS ( CALCULATETABLE ( VALUES ( 'Calendar'[month_index] ), "
     f"{CLOSED} ), \"month_index\", 'Calendar'[month_index] - 12 ) "
     "RETURN CALCULATE ( [Revenue], ALL ( 'Calendar' ), TREATAS ( shifted, 'Calendar'[month_index] ) )",
     "money", "Revenue for the same closed months one year earlier."),
    ("Revenue YoY %", "Sales", "DIVIDE ( [Revenue] - [Revenue PY], [Revenue PY] )", "pct",
     "Growth against the same months last year."),
]


def read_rows(name: str):
    with open(DATA / name, newline="", encoding="utf-8") as fh:
        rows = list(csv.reader(fh))
    return rows[0], rows[1:]


def convert(value: str, kind: str):
    if kind == "date":
        return serial(value)
    if kind == "int":
        return int(value)
    if kind == "num":
        return float(value)
    return value if value != "" else None


def create_sheets(wb) -> None:
    first = wb.Worksheets(1)
    first.Name = L.SHEETS[0]
    for name in L.SHEETS[1:]:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)).Name = name
    for name in L.SHEETS:
        colour = TAB_COLOURS.get(name, GREY if name.startswith(("Data_", "PQ_")) else TEAL)
        wb.Worksheets(name).Tab.Color = rgb(colour)


def theme(wb) -> None:
    """Body font and the waterfall's colours. Excel's waterfall paints increases,
    decreases and totals with theme accents 1-3, so those carry meaning here."""
    normal = wb.Styles("Normal").Font
    normal.Name = BODY_FONT
    normal.Size = 10
    scheme = wb.Theme.ThemeColorScheme
    for accent, colour in {1: FAV, 2: UNFAV, 3: NAVY, 4: TEAL, 5: AMBER, 6: GREY}.items():
        scheme.Colors(MSO_THEME_ACCENT[accent]).RGB = rgb(colour)


def write_table(ws, first_col: str, first_row: int, name: str, csv_name: str,
                kinds: str, formats: dict):
    head, body = read_rows(csv_name)
    kinds = kinds.split()
    values = tuple(tuple(convert(v, k) for v, k in zip(r, kinds)) for r in body)
    top = ws.Range(f"{first_col}{first_row}")
    hdr = top.GetResize(1, len(head))
    hdr.NumberFormat = "@"                 # '2026-01' must stay a heading, not a date
    hdr.Value = (tuple(head),)
    for j, (h, k) in enumerate(zip(head, kinds)):
        top.GetOffset(1, j).GetResize(len(values), 1).NumberFormat = formats.get(h, KIND_FORMAT[k])
    top.GetOffset(1, 0).GetResize(len(values), len(head)).Value = values
    lo = ws.ListObjects.Add(SourceType=XL_SRC_RANGE, Source=top.GetResize(len(values) + 1, len(head)),
                            XlListObjectHasHeaders=XL_YES)
    lo.Name = name
    lo.TableStyle = "TableStyleLight1"
    return len(values)


def data_sheets(wb) -> dict:
    """Load every CSV into an Excel table. Returns row counts for the control totals."""
    counts = {}
    for sheet, table, csv_name, kinds, formats, caption in SOURCES:
        ws = wb.Worksheets(sheet)
        title(ws, caption, f"Loaded unchanged from data/{csv_name} as the table {table}. "
                           "Formulas and Power Query read the table by name.")
        counts[table] = write_table(ws, "B", 4, table, csv_name, kinds, formats)
        ws.Columns("A").ColumnWidth = 2
        ws.Range("B4").CurrentRegion.Columns.AutoFit()
    ws = wb.Worksheets("Data_Dims")
    title(ws, "Dimensions: products, channels, regions",
          "Loaded from data/dim_*.csv. Shelf life per product drives the inventory risk flags.")
    for table, csv_name, kinds, formats, first in DIMS:
        counts[table] = write_table(ws, first, 4, table, csv_name, kinds, formats)
    ws.Columns("A").ColumnWidth = 2
    ws.Range("B4:N30").Columns.AutoFit()
    return counts


def names(wb, specs: dict) -> None:
    for name, ref in specs.items():
        wb.Names.Add(Name=name, RefersTo="=" + ref)


def assumptions(wb) -> None:
    ws = wb.Worksheets("Assumptions")
    title(ws, "Assumptions and scenario drivers",
          "Blue-on-yellow cells are inputs. Everything else in the workbook is calculated "
          "from them and from the source tables.")
    band(ws, 4, "Model controls", "B", "G")
    header(ws, 5, ["Input", "Value", "Unit", "What it does"], "B", wrap=False)
    ws.Range("E5").HorizontalAlignment = -4131
    controls = [  # label, value, unit, note, is_input, format, name
        ("As-of month (last closed)", serial("2026-08-01"), "month",
         "Months up to here are actuals, later months are forecast. After changing it, "
         "use Data > Refresh All so the data model's calendar follows.", True, F_MONTH_LONG, "AsOfMonth"),
        ("Fiscal year", 2026, "year", "The year the budget and the forecast cover.", True, "0", "FiscalYear"),
        ("Fiscal year starts", "=DATE(FiscalYear,1,1)", "date", "", False, F_DATE, "FYStart"),
        ("Months closed", "=(YEAR(AsOfMonth)-FiscalYear)*12+MONTH(AsOfMonth)", "months", "",
         False, "0", "MonthsClosed"),
        ("Trailing window for rates", 3, "months",
         "Forecast price, cost and freight per case are these closed months' averages.",
         True, "0", "TrailMonths"),
        ("Trailing window starts", "=EDATE(AsOfMonth,1-TrailMonths)", "month", "", False,
         F_MONTH_LONG, "TrailStart"),
        ("Days in the trailing window", 91, "days",
         "Days of flow behind DSO, DIO and DPO: 13 weeks.", True, "0", "WCDays"),
        ("Shelf-life risk line", 0.5, "share",
         "Inventory holding more than this share of its shelf life is flagged at risk.",
         True, "0%", "ShelfRisk"),
        ("Check tolerance", 0.01, "$", "How far two figures may differ before a check fails.",
         True, "0.00", "Tol"),
        ("Last month in the data", "=MAX(tbl_Sales[month])", "month", "", False, F_MONTH_LONG,
         "LastDataMonth"),
    ]
    for i, (label, value, unit, note, is_input, fmt, name) in enumerate(controls):
        r = 6 + i
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", value)
        put(ws, f"D{r}", unit)
        put(ws, f"E{r}", note)
        numfmt(ws.Range(f"C{r}"), fmt)
        if is_input:
            input_cell(ws.Range(f"C{r}"))
        wb.Names.Add(Name=name, RefersTo=f"=Assumptions!$C${r}")
    font(ws.Range("D6:E15"), color=MUTED)

    band(ws, 17, "Scenario drivers: change on top of the trailing-rate forecast", "B", "G")
    header(ws, 18, ["Driver", *L.SCENARIO_NAMES, "Live", "Applies to"], "B", wrap=False)
    ws.Range("G18").HorizontalAlignment = -4131
    drivers = [  # label, base, upside, downside, live formula, note, name
        ("Volume", 0.0, 0.03, -0.04, "=INDEX(C19:E19,ScenarioNo)+sens_volume",
         "Forecast cases, on top of budget cases × YTD run-rate", "sel_volume"),
        ("Price", 0.0, 0.015, -0.01, "=INDEX(C20:E20,ScenarioNo)+sens_price",
         "Net price per case against the trailing months", "sel_price"),
        ("Unit cost", 0.0, -0.01, 0.025, "=INDEX(C21:E21,ScenarioNo)",
         "Cost per case against the trailing months (positive = dearer)", "sel_cost"),
        ("Freight per case", 0.0, -0.02, 0.05, "=INDEX(C22:E22,ScenarioNo)",
         "Freight per case against the trailing months", "sel_freight"),
        ("Operating expenses", 0.0, -0.01, 0.02, "=INDEX(C23:E23,ScenarioNo)",
         "Budget × each department's YTD run-rate", "sel_opex"),
    ]
    for i, (label, base, up, down, live, note, name) in enumerate(drivers):
        r = 19 + i
        put(ws, f"B{r}", label)
        ws.Range(f"C{r}:E{r}").Value = ((base, up, down),)
        input_cell(ws.Range(f"C{r}:E{r}"))
        fx(ws, f"F{r}", live)
        put(ws, f"G{r}", note)
        wb.Names.Add(Name=name, RefersTo=f"=Assumptions!$F${r}")
    numfmt(ws.Range("C19:F23"), '+0.0%;-0.0%;0.0%')
    font(ws.Range("F19:F23"), bold=True)
    font(ws.Range("G19:G23"), color=MUTED)
    wb.Names.Add(Name="ScenarioList", RefersTo="=Assumptions!$C$18:$E$18")
    dv = ws.Range("C19:E23").Validation
    dv.Delete()
    dv.Add(XL_VALIDATE_DECIMAL, XL_VALID_ALERT_STOP, XL_BETWEEN, "-0.5", "0.5")
    dv = ws.Range("C6").Validation
    dv.Delete()
    dv.Add(XL_VALIDATE_DATE, XL_VALID_ALERT_STOP, XL_BETWEEN, "=FYStart", "=LastDataMonth")
    dv.ErrorMessage = "Pick a month in the fiscal year that has actuals."
    widths(ws, {"A": 2, "B": 30, "C:E": 12, "F": 10, "G": 60})
    ws.Range("E6:E15").WrapText = False


def scenario_inputs(wb) -> None:
    """The three cells the what-if data tables vary. A data table's input cells
    must sit on its own sheet, so they live on Scenarios, not Assumptions."""
    ws = wb.Worksheets("Scenarios")
    put(ws, "B4", "Scenario")
    put(ws, "C4", "Base")
    put(ws, "B5", "Price overlay")
    put(ws, "C5", 0)
    put(ws, "B6", "Volume overlay")
    put(ws, "C6", 0)
    fx(ws, "D4", "=MATCH(C4,ScenarioList,0)")
    numfmt(ws.Range("C5:C6"), '+0.0%;-0.0%;0.0%')
    numfmt(ws.Range("D4"), '"scenario "0')
    input_cell(ws.Range("C4:C6"))
    font(ws.Range("B4:B6"), bold=True)
    font(ws.Range("D4"), color=MUTED)
    names(wb, {"ScenarioName": "Scenarios!$C$4", "ScenarioNo": "Scenarios!$D$4",
               "sens_price": "Scenarios!$C$5", "sens_volume": "Scenarios!$C$6"})
    dv = ws.Range("C4").Validation
    dv.Delete()
    dv.Add(XL_VALIDATE_LIST, XL_VALID_ALERT_STOP, XL_BETWEEN, "=ScenarioList")
    dv.InputTitle = "Scenario"
    dv.InputMessage = "Every sheet follows the scenario picked here."
    for addr in ("C5", "C6"):
        dv = ws.Range(addr).Validation
        dv.Delete()
        dv.Add(XL_VALIDATE_DECIMAL, XL_VALID_ALERT_STOP, XL_BETWEEN, "-0.2", "0.2")
        dv.InputMessage = "Added to the scenario's own driver. Try 1% or -2%."


def lambdas(wb) -> None:
    for name, (formula, comment) in LAMBDAS.items():
        n = wb.Names.Add(Name=name, RefersTo=formula)
        n.Comment = comment


def power_query(wb) -> None:
    """Six queries into the data model; Budget also lands on PQ_Budget as tbl_Budget."""
    ws = wb.Worksheets("PQ_Budget")
    title(ws, "Budget, unpivoted and priced by Power Query",
          "Output of the Budget query (Data > Queries & Connections). The P&L, the "
          "price-volume-mix and the data model all read this table, tbl_Budget.")
    for name, m, to_sheet in QUERIES:
        wb.Queries.Add(name, m, f"Kestrel Bay FP&A model: {name}")
        conn = wb.Connections.Add2(f"Query - {name}", f"Connection to the '{name}' query.",
                                   CONN.format(name), f'"{name}"', 6, True, False)
        if to_sheet:
            lo = ws.ListObjects.Add(SourceType=XL_SRC_MODEL, Source=conn, Destination=ws.Range("B4"))
            lo.TableStyle = "TableStyleLight1"
            to = lo.TableObject
            to.RowNumbers = False
            to.PreserveFormatting = True
            to.RefreshStyle = 1
            to.AdjustColumnWidth = True
            lo.DisplayName = "tbl_Budget"
            to.Refresh()
            lo.ListColumns("month").DataBodyRange.NumberFormat = F_DATE
    ws.Columns("A").ColumnWidth = 2


def data_model(wb) -> None:
    model = wb.Model
    tables = model.ModelTables

    def column_(table: str, name: str):
        return tables(table).ModelTableColumns(name)

    for fact in ("Sales", "Budget"):
        for key, dim in (("month", "Calendar"), ("sku", "Product"), ("channel", "Channel"),
                         ("region", "Region")):
            model.ModelRelationships.Add(column_(fact, key), column_(dim, key))
    formats = {
        "money": model.GetModelFormatCurrency("$", 0),
        "pct": model.GetModelFormatPercentageNumber(False, 1),
        "int": model.GetModelFormatWholeNumber(True),
        "dec": model.GetModelFormatDecimalNumber(True, 2),
    }
    for name, home, dax, fmt, description in MEASURES:
        model.ModelMeasures.Add(name, tables(home), dax, formats[fmt], description)


