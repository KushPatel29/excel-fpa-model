"""Workbook skeleton: sheets, source tables, inputs, named functions, Power Query
and the Power Pivot data model. Everything the analysis sheets stand on.
"""
from __future__ import annotations

import csv
from pathlib import Path

import layout as L
from xl import (
    AMBER, BODY_FONT, F_DATE, F_MONTH_LONG, FAV, GREY, MSO_THEME_ACCENT, MUTED, NAVY, TEAL, UNFAV,
    XL_BETWEEN, XL_SRC_MODEL, XL_SRC_RANGE, XL_VALID_ALERT_STOP, XL_VALIDATE_DECIMAL, XL_VALIDATE_LIST,
    XL_YES, band, font, fx, header, input_cell, numfmt, place, put, rgb, title, widths,
)

DATA = Path(__file__).resolve().parent.parent / "data"
TAB_COLOURS = {"Cover": NAVY, "Dashboard": NAVY, "Checks": FAV, "Assumptions": AMBER}
KIND_FORMAT = {"date": F_DATE, "int": "0", "num": "#,##0.000", "text": "@"}

LAMBDAS = {
    "SAFEDIV": ("=LAMBDA(n,d,IF(d=0,0,n/d))", "n ÷ d, or 0 when d is 0."),
    "ISCLOSED": ("=LAMBDA(month,month<=AsOfMonth)", "TRUE for a month with actuals, FALSE for a forecast month."),
    "CPK": ("=LAMBDA(revenue,mwh,SAFEDIV(revenue,mwh)/10)", "Revenue ($) and MWh to cents per kWh."),
    "TRENDYEARS": ("=LAMBDA(month,YEAR(month)-YEAR(FitStart)+(MONTH(month)-1)/12)",
                   "Years since the start of the fit window: the weather model's trend term."),
    "MONEY": ('=LAMBDA(x,IF(ABS(x)>=999500000,TEXT(x/1000000000,"$#,##0.00")&"B",IF(ABS(x)>=999500,'
              'TEXT(x/1000000,"$#,##0.0")&"M",TEXT(x/1000,"$#,##0")&"K")))',
              "$3.01B / $40.4M / $350K style amounts for written commentary."),
    "SIGNMONEY": ('=LAMBDA(x,IF(x<0,"−","+")&MONEY(ABS(x)))', "MONEY with an explicit sign."),
}

CONN = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={};Extended Properties=""'

M_MONTHLY = """// EIA publishes one row per month with the classes side by side. Unpivot it,
// split each column name into class and measure, and pivot the measures back:
// one row per month and customer class, in dollars.
let
    Source = Excel.CurrentWorkbook(){[Name="tbl_EIA"]}[Content],
    Long = Table.UnpivotOtherColumns(Source, {"year", "month", "data_status"}, "field", "value"),
    WithClass = Table.AddColumn(Long, "class", each Text.BeforeDelimiter([field], "_"), type text),
    WithMeasure = Table.AddColumn(WithClass, "measure", each Text.AfterDelimiter([field], "_"), type text),
    Classes = Table.SelectRows(WithMeasure, each [class] <> "total"),
    Slim = Table.RemoveColumns(Classes, {"field"}),
    Wide = Table.Pivot(Slim, List.Sort(List.Distinct(Slim[measure])), "measure", "value", List.Sum),
    Dated = Table.AddColumn(Wide, "month_start", each #date([year], [month], 1), type date),
    Dollars = Table.AddColumn(Dated, "revenue", each [revenue_k] * 1000, type number),
    Selected = Table.SelectColumns(Dollars, {"month_start", "class", "revenue", "mwh", "customers", "data_status"}),
    Renamed = Table.RenameColumns(Selected, {{"month_start", "month"}}),
    Typed = Table.TransformColumnTypes(Renamed, {{"mwh", type number}, {"customers", Int64.Type},
        {"data_status", type text}}),
    Sorted = Table.Sort(Typed, {{"month", Order.Ascending}, {"class", Order.Ascending}})
in
    Sorted"""

M_WEATHER = """// NOAA's nClimDiv file is fixed-width text: division, element, year, then twelve
// seven-character monthly values, -9999 where a month has not happened yet.
// Split by position, unpivot the months, drop the placeholders, and pivot the
// two elements (25 heating, 26 cooling degree days) into columns.
let
    Source = Excel.CurrentWorkbook(){[Name="tbl_NOAA"]}[Content],
    Split = Table.SplitColumn(Source, "line", Splitter.SplitTextByPositions(
        {0, 4, 6, 10, 17, 24, 31, 38, 45, 52, 59, 66, 73, 80, 87}),
        {"division", "element", "year", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"}),
    Long = Table.UnpivotOtherColumns(Split, {"division", "element", "year"}, "month_no", "raw"),
    Parsed = Table.AddColumn(Long, "value",
        each Number.FromText(Text.TrimEnd(Text.Trim([raw]), "."), "en-US"), type number),
    Observed = Table.SelectRows(Parsed, each [value] > -9999),
    Dated = Table.AddColumn(Observed, "month",
        each #date(Number.FromText([year]), Number.FromText([month_no]), 1), type date),
    Named = Table.AddColumn(Dated, "degree_days", each if [element] = "25" then "hdd" else "cdd", type text),
    Slim = Table.SelectColumns(Named, {"month", "degree_days", "value"}),
    Wide = Table.Pivot(Slim, {"hdd", "cdd"}, "degree_days", "value", List.Sum),
    Typed = Table.TransformColumnTypes(Wide, {{"hdd", type number}, {"cdd", type number}}),
    Sorted = Table.Sort(Typed, {{"month", Order.Ascending}})
in
    Sorted"""

M_CALENDAR = """// Month calendar for the data model: the fit window's start to the end of the
// fiscal year. status reads AsOfMonth, so Refresh All re-cuts closed and open.
let
    AsOf = Date.From(Excel.CurrentWorkbook(){[Name="AsOfMonth"]}[Content]{0}[Column1]),
    First = Date.From(Excel.CurrentWorkbook(){[Name="FitStart"]}[Content]{0}[Column1]),
    Last = #date(Date.Year(AsOf), 12, 1),
    Count = (Date.Year(Last) - Date.Year(First)) * 12 + Date.Month(Last) - Date.Month(First) + 1,
    Months = List.Transform({0..Count - 1}, each Date.AddMonths(First, _)),
    Base = Table.TransformColumnTypes(
        Table.FromList(Months, Splitter.SplitByNothing(), {"month"}), {{"month", type date}}),
    Year = Table.AddColumn(Base, "year", each Date.Year([month]), Int64.Type),
    MonthNo = Table.AddColumn(Year, "month_no", each Date.Month([month]), Int64.Type),
    MonthName = Table.AddColumn(MonthNo, "month_name", each Date.ToText([month], "MMM", "en-US"), type text),
    Quarter = Table.AddColumn(MonthName, "quarter", each "Q" & Text.From(Date.QuarterOfYear([month])), type text),
    Index = Table.AddColumn(Quarter, "month_index",
        each ([year] - Date.Year(First)) * 12 + [month_no], Int64.Type),
    Status = Table.AddColumn(Index, "status", each if [month] <= AsOf then "Closed" else "Open", type text)
in
    Status"""

M_CLASS = """let
    Source = Excel.CurrentWorkbook(){[Name="tbl_Class"]}[Content],
    Typed = Table.TransformColumnTypes(Source, {{"class", type text}, {"label", type text},
        {"sort", Int64.Type}, {"weather_model", type text}})
in
    Typed"""

QUERIES = [  # name, M, worksheet to load to (or None), table name
    ("Monthly", M_MONTHLY, "PQ_Monthly", "tbl_Monthly"),
    ("Weather", M_WEATHER, "PQ_Weather", "tbl_Weather"),
    ("Calendar", M_CALENDAR, None, None),
    ("Class", M_CLASS, None, None),
]

CLOSED = "'Calendar'[status] = \"Closed\""
MEASURES = [  # name, home table, DAX, format, description (measure names must not repeat a column name)
    ("Retail Revenue", "Monthly", "SUM ( Monthly[revenue] )", "money", "Retail revenue, dollars."),
    ("Retail MWh", "Monthly", "SUM ( Monthly[mwh] )", "int", "Retail sales, megawatt-hours."),
    ("Price c/kWh", "Monthly", "DIVIDE ( [Retail Revenue], [Retail MWh] ) / 10", "dec",
     "Average retail price, cents per kWh."),
    ("Avg Customers", "Monthly",
     "AVERAGEX ( VALUES ( 'Calendar'[month] ), CALCULATE ( SUM ( Monthly[customers] ) ) )", "int",
     "Average monthly customer count."),
    ("MWh per Customer", "Monthly", "DIVIDE ( [Retail MWh], [Avg Customers] )", "dec", "Use per average customer."),
    ("Retail Revenue PY", "Monthly",
     "VAR shifted = SELECTCOLUMNS ( CALCULATETABLE ( VALUES ( 'Calendar'[month_index] ), "
     f"{CLOSED} ), \"month_index\", 'Calendar'[month_index] - 12 ) "
     "RETURN CALCULATE ( [Retail Revenue], ALL ( 'Calendar' ), TREATAS ( shifted, 'Calendar'[month_index] ) )",
     "money", "Revenue for the same closed months one year earlier."),
    ("Revenue YoY %", "Monthly", "DIVIDE ( [Retail Revenue] - [Retail Revenue PY], [Retail Revenue PY] )", "pct", ""),
    ("Retail MWh PY", "Monthly",
     "VAR shifted = SELECTCOLUMNS ( CALCULATETABLE ( VALUES ( 'Calendar'[month_index] ), "
     f"{CLOSED} ), \"month_index\", 'Calendar'[month_index] - 12 ) "
     "RETURN CALCULATE ( [Retail MWh], ALL ( 'Calendar' ), TREATAS ( shifted, 'Calendar'[month_index] ) )",
     "int", "MWh for the same closed months one year earlier."),
    ("MWh YoY %", "Monthly", "DIVIDE ( [Retail MWh] - [Retail MWh PY], [Retail MWh PY] )", "pct", ""),
    ("Heating Degree Days", "Weather", "SUM ( Weather[hdd] )", "int", "Heating degree days, Willamette Valley."),
    ("Cooling Degree Days", "Weather", "SUM ( Weather[cdd] )", "int", "Cooling degree days, Willamette Valley."),
]

SOURCES = [  # sheet, table, csv, kinds, first column, caption
    ("Data_EIA", "tbl_EIA", "eia_pge_monthly.csv", "int int text" + " num num int" * 5, "B",
     "EIA-861M: PGE monthly retail sales by class, as published"),
    ("Data_FERC", "tbl_FERC_Revenue", "ferc_revenue.csv", "int int text num num num", "B",
     "FERC Form 1 via PUDL: revenue (sched. 300), expenses (sched. 320), income statement (sched. 114)"),
    ("Data_FERC", "tbl_FERC_Expense", "ferc_expense.csv", "int int text num", "J", None),
    ("Data_FERC", "tbl_FERC_Income", "ferc_income.csv", "int int text num", "P", None),
    ("Data_Utilities", "tbl_Utilities", "utilities.csv", "int text text text text", "B",
     "The FERC respondents in the model, and the customer classes"),
]
CLASS_ROWS = [("residential", "Residential", 1, "yes"), ("commercial", "Commercial", 2, "yes"),
              ("industrial", "Industrial", 3, "no"), ("transportation", "Transportation", 4, "no")]


def read_rows(name: str):
    with open(DATA / name, newline="", encoding="utf-8") as fh:
        rows = list(csv.reader(fh))
    return rows[0], rows[1:]


def convert(value: str, kind: str):
    if value == "":
        return None
    if kind == "int":
        return int(float(value))
    if kind == "num":
        return float(value)
    return value


def create_sheets(wb) -> None:
    wb.Worksheets(1).Name = L.SHEETS[0]
    for name in L.SHEETS[1:]:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)).Name = name
    for name in L.SHEETS:
        colour = TAB_COLOURS.get(name, GREY if name.startswith(("Data_", "PQ_")) else TEAL)
        wb.Worksheets(name).Tab.Color = rgb(colour)


def theme(wb) -> None:
    """Body font, and the waterfall's colours: Excel paints increases, decreases and
    totals with theme accents 1-3, so those carry meaning here."""
    normal = wb.Styles("Normal").Font
    normal.Name = BODY_FONT
    normal.Size = 10
    scheme = wb.Theme.ThemeColorScheme
    for accent, colour in {1: FAV, 2: UNFAV, 3: NAVY, 4: TEAL, 5: AMBER, 6: GREY}.items():
        scheme.Colors(MSO_THEME_ACCENT[accent]).RGB = rgb(colour)


def write_table(ws, first_col: str, first_row: int, name: str, head, body, kinds) -> int:
    values = tuple(tuple(convert(v, k) for v, k in zip(r, kinds)) for r in body)
    top = ws.Range(f"{first_col}{first_row}")
    hdr = top.GetResize(1, len(head))
    hdr.NumberFormat = "@"
    hdr.Value = (tuple(head),)
    for j, k in enumerate(kinds):
        top.GetOffset(1, j).GetResize(len(values), 1).NumberFormat = KIND_FORMAT[k]
    top.GetOffset(1, 0).GetResize(len(values), len(head)).Value = values
    lo = ws.ListObjects.Add(SourceType=XL_SRC_RANGE, Source=top.GetResize(len(values) + 1, len(head)),
                            XlListObjectHasHeaders=XL_YES)
    lo.Name = name
    lo.TableStyle = "TableStyleLight1"
    return len(values)


def data_sheets(wb) -> dict:
    counts = {}
    for sheet, table, csv_name, kinds, first, caption in SOURCES:
        ws = wb.Worksheets(sheet)
        if caption:
            title(ws, caption, f"Loaded unchanged from data/{csv_name}. Formulas and Power Query read "
                               "these tables by name; data/sources.json records where each came from.")
        head, body = read_rows(csv_name)
        counts[table] = write_table(ws, first, 4, table, head, body, kinds.split())
    ws = wb.Worksheets("Data_Utilities")
    write_table(ws, "I", 4, "tbl_Class", ["class", "label", "sort", "weather_model"],
                [[str(v) for v in r] for r in CLASS_ROWS], ["text", "text", "int", "text"])
    ws = wb.Worksheets("Data_NOAA")
    title(ws, "NOAA nClimDiv degree days, Oregon division 2 (Willamette Valley), as published",
          "Fixed-width lines: division, element (25 heating, 26 cooling), year, twelve monthly values. "
          "Power Query parses them into tbl_Weather.")
    lines = (DATA / "noaa_willamette_valley_degree_days.txt").read_text().splitlines()
    ws.Range("B4").Value = "line"
    body = ws.Range("B5").GetResize(len(lines), 1)
    body.NumberFormat = "@"
    body.Value = tuple((line,) for line in lines)
    lo = ws.ListObjects.Add(SourceType=XL_SRC_RANGE, Source=ws.Range("B4").GetResize(len(lines) + 1, 1),
                            XlListObjectHasHeaders=XL_YES)
    lo.Name = "tbl_NOAA"
    lo.TableStyle = "TableStyleLight1"
    font(body, name="Consolas", size=9)
    counts["tbl_NOAA"] = len(lines)
    for sheet in ("Data_EIA", "Data_FERC", "Data_Utilities", "Data_NOAA"):
        ws = wb.Worksheets(sheet)
        ws.Columns("A").ColumnWidth = 2
        ws.Range("B4:X4").EntireColumn.AutoFit()
    wb.Worksheets("Data_NOAA").Columns("B").ColumnWidth = 96
    return counts


def assumptions(wb) -> None:
    ws = wb.Worksheets("Assumptions")
    title(ws, "Assumptions and model controls",
          "Blue-on-yellow cells are inputs. The as-of month is read from the data: the latest month EIA has "
          "published for PGE.")
    band(ws, 4, "Model controls", "B", "G")
    header(ws, 5, ["Input", "Value", "Unit", "What it does"], "B", wrap=False)
    ws.Range("D5:E5").HorizontalAlignment = -4131
    controls = [  # label, value, unit, note, is_input, format, name
        ("Latest month with actuals", '=DATE(MAX(tbl_EIA[year]),MAXIFS(tbl_EIA[month],tbl_EIA[year],'
         'MAX(tbl_EIA[year])),1)', "month", "EIA's latest published month for PGE. Everything after it is forecast.",
         False, F_MONTH_LONG, "AsOfMonth"),
        ("Fiscal year", "=YEAR(AsOfMonth)", "year", "PGE reports on the calendar year.", False, "0", "FiscalYear"),
        ("Fiscal year starts", "=DATE(FiscalYear,1,1)", "date", "", False, F_DATE, "FYStart"),
        ("Months closed", "=MONTH(AsOfMonth)", "months", "", False, "0", "MonthsClosed"),
        ("Weather model fit starts", "=DATE(2017,1,1)", "month",
         "First month of the regression window; the window ends the December before the year planned.",
         True, F_MONTH_LONG, "FitStart"),
        ("Years in a weather normal", 10, "years", "Normal weather is this many years' average for each month.",
         True, "0", "NormalYears"),
        ("Company (FERC respondent id)", L.COMPANY_ID, "id", "Portland General Electric Company in FERC Form 1.",
         True, "0", "CompanyId"),
        ("Long-bridge base year", 2019, "year", "The earlier year of the long price-volume-mix bridge.",
         True, "0", "BaseYear"),
        ("Latest FERC year", "=MAX(tbl_FERC_Revenue[report_year])", "year", "The latest annual filing loaded.",
         False, "0", "FercYear"),
        ("Check tolerance", 0.01, "$", "How far two figures may differ before a check fails.", True, "0.00", "Tol"),
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

    band(ws, 17, "Weather scenarios for the months NOAA has not yet observed", "B", "G")
    header(ws, 18, ["Scenario", "HDD scale", "CDD scale", "Meaning"], "B", wrap=False)
    ws.Range("E18").HorizontalAlignment = -4131
    notes = {"Normal": "Ten-year normal degree days.",
             "Mild": "A winter like 2026's: about 12% fewer heating degree days than normal.",
             "Cold": "A cold snap: 12% more heating degree days than normal."}
    for i, (name, hdd, cdd) in enumerate(L.WEATHER_SCENARIOS):
        r = 19 + i
        ws.Range(f"B{r}:D{r}").Value = ((name, hdd, cdd),)
        input_cell(ws.Range(f"C{r}:D{r}"))
        put(ws, f"E{r}", notes[name])
    numfmt(ws.Range("C19:D21"), "0.00")
    font(ws.Range("E19:E21"), color=MUTED)
    wb.Names.Add(Name="ScenarioList", RefersTo="=Assumptions!$B$19:$B$21")
    wb.Names.Add(Name="ScenarioHdd", RefersTo="=Assumptions!$C$19:$C$21")
    wb.Names.Add(Name="ScenarioCdd", RefersTo="=Assumptions!$D$19:$D$21")
    dv = ws.Range("C19:D21").Validation
    dv.Delete()
    dv.Add(XL_VALIDATE_DECIMAL, XL_VALID_ALERT_STOP, XL_BETWEEN, "0.5", "1.5")
    widths(ws, {"A": 2, "B": 30, "C": 14, "D": 11, "E": 70, "F:G": 4})


def scenario_inputs(wb) -> None:
    """The cells the what-if data tables vary. A data table's input cells must be on
    its own sheet, so the levers live on Scenarios, not Assumptions."""
    ws = wb.Worksheets("Scenarios")
    levers = [  # label, value, format, name, validation, note
        ("Weather, months not yet observed", "Normal", "@", "WeatherScenario", None,
         "Normal, Mild or Cold: the multipliers are on Assumptions."),
        ("Rate change on open months", 0, '+0.0%;-0.0%;0.0%', "RateChange", ("-0.2", "0.2"),
         "Added to every class's price from the first open month."),
        ("Extra data-center load", 0, '0" MW"', "ExtraMW", ("0", "2000"),
         "Average megawatts of new industrial load, running around the clock."),
        ("Customer growth overlay", 0, '+0.0%;-0.0%;0.0%', "CustOverlay", ("-0.2", "0.2"),
         "Scales open-month customers (and industrial volume) up or down."),
        ("Heating degree-day overlay", 1, "0.00x", "HddOverlay", ("0.5", "1.5"),
         "Multiplies the scenario's heating degree days."),
        ("Cooling degree-day overlay", 1, "0.00x", "CddOverlay", ("0.5", "1.5"),
         "Multiplies the scenario's cooling degree days."),
    ]
    for i, (label, value, fmt, name, valid, note) in enumerate(levers):
        r = 4 + i
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", value)
        numfmt(ws.Range(f"C{r}"), fmt)
        put(ws, f"E{r}", note)
        input_cell(ws.Range(f"C{r}"))
        wb.Names.Add(Name=name, RefersTo=f"=Scenarios!$C${r}")
        dv = ws.Range(f"C{r}").Validation
        dv.Delete()
        if valid:
            dv.Add(XL_VALIDATE_DECIMAL, XL_VALID_ALERT_STOP, XL_BETWEEN, *valid)
        else:
            dv.Add(XL_VALIDATE_LIST, XL_VALID_ALERT_STOP, XL_BETWEEN, "=ScenarioList")
    font(ws.Range("B4:B9"), bold=True)
    font(ws.Range("E4:E9"), color=MUTED)
    fx(ws, "D4", "=MATCH(WeatherScenario,ScenarioList,0)")
    numfmt(ws.Range("D4"), '"scenario "0')
    font(ws.Range("D4"), color=MUTED)
    wb.Names.Add(Name="ScenarioNo", RefersTo="=Scenarios!$D$4")
    wb.Names.Add(Name="HddScale", RefersTo="=INDEX(ScenarioHdd,ScenarioNo)*HddOverlay")
    wb.Names.Add(Name="CddScale", RefersTo="=INDEX(ScenarioCdd,ScenarioNo)*CddOverlay")


def lambdas(wb) -> None:
    for name, (formula, comment) in LAMBDAS.items():
        n = wb.Names.Add(Name=name, RefersTo=formula)
        n.Comment = comment


def power_query(wb) -> None:
    for sheet, caption in (("PQ_Monthly", "EIA monthly sales, one row per month and class (Power Query)"),
                           ("PQ_Weather", "Degree days by month, parsed from NOAA's fixed-width file (Power Query)")):
        ws = wb.Worksheets(sheet)
        title(ws, caption, "Output of a Power Query query (Data > Queries & Connections). Formulas and the "
                           "data model read this table.")
        ws.Columns("A").ColumnWidth = 2
    for name, m, sheet, table in QUERIES:
        wb.Queries.Add(name, m, f"PGE utility FP&A model: {name}")
        conn = wb.Connections.Add2(f"Query - {name}", f"Connection to the '{name}' query.",
                                   CONN.format(name), f'"{name}"', 6, True, False)
        if sheet:
            ws = wb.Worksheets(sheet)
            lo = ws.ListObjects.Add(SourceType=XL_SRC_MODEL, Source=conn, Destination=ws.Range("B4"))
            lo.TableStyle = "TableStyleLight1"
            to = lo.TableObject
            to.RowNumbers = False
            to.PreserveFormatting = True
            to.RefreshStyle = 1
            to.AdjustColumnWidth = True
            lo.DisplayName = table
            to.Refresh()
            lo.ListColumns("month").DataBodyRange.NumberFormat = F_DATE


def data_model(wb) -> None:
    model = wb.Model
    tables = model.ModelTables

    def column_(table: str, name: str):
        return tables(table).ModelTableColumns(name)

    model.ModelRelationships.Add(column_("Monthly", "month"), column_("Calendar", "month"))
    model.ModelRelationships.Add(column_("Monthly", "class"), column_("Class", "class"))
    model.ModelRelationships.Add(column_("Weather", "month"), column_("Calendar", "month"))
    formats = {
        "money": model.GetModelFormatCurrency("$", 0),
        "pct": model.GetModelFormatPercentageNumber(False, 1),
        "int": model.GetModelFormatWholeNumber(True),
        "dec": model.GetModelFormatDecimalNumber(True, 2),
    }
    for name, home, dax, fmt, description in MEASURES:
        model.ModelMeasures.Add(name, tables(home), dax, formats[fmt], description)


PIVOT_MEASURES = ["Retail Revenue", "Revenue YoY %", "Retail MWh", "MWh YoY %", "Price c/kWh", "Avg Customers"]


def explore_pivot(wb) -> None:
    """The data-model PivotTable and its slicers, created as soon as the model exists:
    Excel refuses a model PivotCache once the sheets hold thousands of uncalculated formulas."""
    ws = wb.Worksheets("Explore")
    pc = wb.PivotCaches().Create(2, wb.Connections("ThisWorkbookDataModel"), 8)
    pt = pc.CreatePivotTable(TableDestination=ws.Range("B16"), TableName="ptExplore")
    pt.ManualUpdate = True
    pt.CubeFields("[Class].[label]").Orientation = 1
    for m in PIVOT_MEASURES:
        pt.CubeFields(f"[Measures].[{m}]").Orientation = 4
    pt.ManualUpdate = False
    pt.RowAxisLayout(1)
    pt.TableStyle2 = "PivotStyleLight16"
    pt.GrandTotalName = "All classes"
    pt.PivotFields("[Class].[label].[label]").Caption = "Class"
    for measure, caption in zip(PIVOT_MEASURES, ("Revenue ", "Revenue YoY", "MWh ", "MWh YoY", "¢/kWh",
                                                 "Customers ")):
        pt.PivotFields(f"[Measures].[{measure}]").Caption = caption
    fiscal_year = str(int(wb.Names("FiscalYear").RefersToRange.Value))
    slicers = [("[Calendar].[year]", "[Calendar].[year].[year]", "Year", "B6", "F13", fiscal_year),
               ("[Calendar].[quarter]", "[Calendar].[quarter].[quarter]", "Quarter", "G6", "I13", None)]
    for field, level, caption, tl, br, select in slicers:
        cache = wb.SlicerCaches.Add2(pt, field, f"Slicer_{caption}")
        left, top, width, height = place(ws, tl, br)
        sl = cache.Slicers.Add(ws, level, caption, caption, top, left, width - 6, height)
        sl.Style = "SlicerStyleLight1"
        sl.NumberOfColumns = 5 if caption == "Year" else 2
        if select:
            items = cache.SlicerCacheLevels(1).SlicerItems
            cache.VisibleSlicerItemsList = [items(i).Name for i in range(1, items.Count + 1)
                                            if items(i).Caption == select]
