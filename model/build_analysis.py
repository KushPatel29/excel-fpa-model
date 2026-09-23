"""Analysis sheets: scenarios and sensitivities, working capital, channel
economics (formulas, a data-model PivotTable and CUBE functions) and the checks.
"""
from __future__ import annotations

import layout as L
from xl import (
    F_CASES, F_DAYS, F_MONEY, F_MONTH, F_PCT, F_PRICE, F_VAR, F_VARPCT, FAIL_FILL, FAV,
    INK, MUTED, NAVY, PASS_FILL, TEAL, UNFAV, XL_BAR_CLUSTERED, XL_CELL_VALUE, XL_EQUAL,
    XL_EXTERNAL, XL_DATA_FIELD, XL_LEFT, XL_LINE, XL_ROW_FIELD, XL_VALID_ALERT_STOP,
    XL_VALIDATE_DECIMAL, XL_BETWEEN, band, column, font, fx, header, input_cell,
    numfmt, place, put, rgb, row, sign_colours, title, total_row, widths,
)
from build_calc import S_YTD, B_YTD

CUBE = '"ThisWorkbookDataModel"'


# ------------------------------------------------------------------- Scenarios
def scenarios(wb, pos: dict) -> None:
    ws = wb.Worksheets("Scenarios")
    fc, p = pos["fc"], pos["pnl"]["outlook"]
    b1, b2, _ = fc["basis"]
    u0, units, opex0 = fc["blocks"]["u0"], fc["blocks"]["units"], fc["blocks"]["opex0"]
    title(ws, "Scenarios, sensitivities and the gap to budget",
          "Pick a scenario in C4 and every sheet follows. The tables below are Excel what-if data "
          "tables: they re-run the whole forecast for each scenario and for every price × volume pair.")
    put(ws, "E4", "Base, Upside or Downside: the driver sets are on Assumptions.")
    put(ws, "E5", "Added to the scenario's price driver.")
    put(ws, "E6", "Added to the scenario's volume driver.")
    font(ws.Range("E4:E6"), color=MUTED)

    band(ws, 8, "Live outputs for the selected scenario, full year", "B", "I")
    outputs = [("FY revenue", f"=PnL!P{p['revenue']}", F_MONEY),
               ("FY gross margin", f"=PnL!P{p['gm']}", F_MONEY),
               ("FY EBITDA", f"=PnL!P{p['ebitda']}", F_MONEY),
               ("FY EBITDA margin", f"=PnL!P{p['ebitda_pct']}", F_PCT),
               ("FY EBITDA vs budget", f"=PnL!R{p['ebitda']}", F_VAR)]
    for i, (label, f, fmt) in enumerate(outputs):
        put(ws, f"B{9 + i}", label)
        fx(ws, f"C{9 + i}", f)
        numfmt(ws.Range(f"C{9 + i}"), fmt)
    font(ws.Range("C9:C13"), bold=True, color="#1E7B34")   # green: links from another sheet
    sign_colours(ws.Range("C13"))

    band(ws, 15, "Scenario comparison, full year: a one-variable data table on C4", "B", "I")
    header(ws, 16, ["Scenario", "Revenue", "Gross margin", "EBITDA", "EBITDA margin", "EBITDA vs budget"])
    put(ws, "B17", "Live selection")
    for c, src in zip("CDEFG", ("C9", "C10", "C11", "C12", "C13")):
        fx(ws, f"{c}17", f"={src}")
    column(ws, "B18", list(L.SCENARIO_NAMES))
    ws.Range("B17:G20").Table(ColumnInput=ws.Range("C4"))
    put(ws, "B21", "Budget")
    for c, key in zip("CDEF", ("revenue", "gm", "ebitda", "ebitda_pct")):
        fx(ws, f"{c}21", f"=PnL!Q{p[key]}")
    for c, fmt in zip("CDEFG", (F_MONEY, F_MONEY, F_MONEY, F_PCT, F_VAR)):
        numfmt(ws.Range(f"{c}17:{c}21"), fmt)
    font(ws.Range("B17:G17"), italic=True, color=MUTED)
    total_row(ws.Range("B21:G21"))
    sign_colours(ws.Range("G18:G20"))

    band(ws, 23, "FY EBITDA sensitivity: price change down the side, volume change across the top "
                 "(two-variable data table on C5 and C6)", "B", "I")
    put(ws, "B24", "Price ↓   Volume →")
    font(ws.Range("B24"), color=MUTED, italic=True)
    fx(ws, "B25", "=C11")
    numfmt(ws.Range("B25"), ";;;")
    row(ws, "C25", L.SENS_VOLUME)
    column(ws, "B26", L.SENS_PRICE)
    ws.Range("B25:I32").Table(RowInput=ws.Range("C6"), ColumnInput=ws.Range("C5"))
    numfmt(ws.Range("C25:I25"), '+0%;-0%;0%')
    numfmt(ws.Range("B26:B32"), '+0%;-0%;0%')
    numfmt(ws.Range("C26:I32"), F_MONEY)
    font(ws.Range("C25:I25"), bold=True)
    font(ws.Range("B26:B32"), bold=True)
    scale = ws.Range("C26:I32").FormatConditions.AddColorScale(3)
    scale.ColorScaleCriteria(1).FormatColor.Color = rgb("#F4C7C0")
    scale.ColorScaleCriteria(2).FormatColor.Color = rgb("#FFFFFF")
    scale.ColorScaleCriteria(3).FormatColor.Color = rgb("#BFE3CB")
    ws.Range("F29").Font.Bold = True

    band(ws, 34, "Which driver matters most: full-year EBITDA impact of one step on each driver", "B", "I")
    put(ws, "B35", "Step")
    put(ws, "C35", 0.01)
    numfmt(ws.Range("C35"), "0.0%")
    input_cell(ws.Range("C35"))
    dv = ws.Range("C35").Validation
    dv.Delete()
    dv.Add(XL_VALIDATE_DECIMAL, XL_VALID_ALERT_STOP, XL_BETWEEN, "0.001", "0.2")
    put(ws, "D35", "Exact, not an estimate: the forecast is linear in each driver.")
    font(ws.Range("D35"), color=MUTED)
    header(ws, 36, ["Driver", "FY EBITDA impact", "Rank"])
    put(ws, "S36", "Chart data: drivers sorted by size (SORTBY + HSTACK)")
    font(ws.Range("S36"), color=MUTED, italic=True)
    pu0 = f"Forecast!$P${u0[0]}:$P${u0[1]}"
    pun = f"Forecast!$P${units[0]}:$P${units[1]}"
    basis = {c: f"Forecast!${c}${b1}:${c}${b2}" for c in "JKLM"}
    drivers = [("Volume", f"=SUMPRODUCT({pu0},{basis['M']})*TornadoStep"),
               ("Price", f"=SUMPRODUCT({pun},{basis['J']})*TornadoStep"),
               ("Unit cost", f"=-SUMPRODUCT({pun},{basis['K']})*TornadoStep"),
               ("Freight per case", f"=-SUMPRODUCT({pun},{basis['L']})*TornadoStep"),
               ("Operating expenses", f"=-Forecast!$P${opex0[2]}*TornadoStep")]
    for i, (label, f) in enumerate(drivers):
        put(ws, f"B{37 + i}", label)
        fx(ws, f"C{37 + i}", f)
    fx(ws, "D37:D41", "=SUMPRODUCT(--(ABS($C$37:$C$41)>ABS(C37)))+1")
    fx(ws, "S37", "=SORTBY(HSTACK(B37:B41,C37:C41),ABS(C37:C41),1)")
    numfmt(ws.Range("C37:C41"), F_VAR)
    numfmt(ws.Range("T37:T41"), F_VAR)
    sign_colours(ws.Range("C37:C41"))
    font(ws.Range("S37:T41"), color=MUTED)

    band(ws, 43, "What would it take to hit budget?", "B", "I")
    gap = [("Full-year EBITDA gap to budget", f"=PnL!Q{p['ebitda']}-PnL!P{p['ebitda']}", F_MONEY),
           ("Open-month revenue at trailing prices", f"=SUMPRODUCT({pun},{basis['J']})", F_MONEY),
           ("Price rise needed on the open months", "=SAFEDIV(MAX(C44,0),C45)", "0.00%"),
           ("…or extra volume at today's margins", f"=SAFEDIV(MAX(C44,0),SUMPRODUCT({pu0},{basis['M']}))",
            "0.0%")]
    for i, (label, f, fmt) in enumerate(gap):
        put(ws, f"B{44 + i}", label)
        fx(ws, f"C{44 + i}", f)
        numfmt(ws.Range(f"C{44 + i}"), fmt)
    font(ws.Range("C46:C47"), bold=True)
    fx(ws, "B48", '=IF(C44<=0,"The "&ScenarioName&" outlook already meets the budget\'s full-year EBITDA.",'
                  '"Budget needs "&MONEY(C44)&" more EBITDA than the "&ScenarioName&" outlook. That is a "'
                  '&TEXT(C46,"0.0%")&" price rise on the "&(12-MonthsClosed)&" open months, or "'
                  '&TEXT(C47,"0%")&" more volume at today\'s margins: price is the lever, volume is not.")')
    ws.Range("B48:I49").Merge()
    ws.Range("B48").WrapText = True
    ws.Range("B48").VerticalAlignment = -4160
    ws.Rows(48).RowHeight = 30
    font(ws.Range("B48"), color=INK, italic=True)

    chart = new_chart(ws, XL_BAR_CLUSTERED, "K34", "Q49", [("FY EBITDA impact", "T37:T41")], "S37:S41",
                      "FY EBITDA impact of one step on each driver")
    chart.HasLegend = False
    chart.Axes(1).TickLabelPosition = -4134        # labels clear of the negative bars
    s = chart.FullSeriesCollection(1)
    s.Format.Fill.ForeColor.RGB = rgb(TEAL)
    s.InvertIfNegative = True
    s.InvertColor = rgb(UNFAV)
    chart.Axes(2).TickLabels.NumberFormat = '$#,##0,"K"'   # values sit in the table beside it
    chart.ChartGroups(1).GapWidth = 60
    style_chart(chart)

    widths(ws, {"A": 2, "B": 34, "C:I": 13, "J": 2, "S": 20, "T": 12})
    names = {"scn_live": "Scenarios!$C$17:$G$17", "scn_table": "Scenarios!$C$18:$G$20",
             "scn_budget": "Scenarios!$C$21:$F$21", "sens_grid": "Scenarios!$C$26:$I$32",
             "sens_price_axis": "Scenarios!$B$26:$B$32", "sens_volume_axis": "Scenarios!$C$25:$I$25",
             "TornadoStep": "Scenarios!$C$35", "tornado": "Scenarios!$C$37:$C$41",
             "breakeven_gap": "Scenarios!$C$44", "breakeven_revenue": "Scenarios!$C$45",
             "breakeven_uplift": "Scenarios!$C$46", "breakeven_volume": "Scenarios!$C$47"}
    for name, ref in names.items():
        wb.Names.Add(Name=name, RefersTo="=" + ref)


def new_chart(ws, chart_type: int, top_left: str, bottom_right: str, series, categories: str,
              title_text: str):
    """A chart with explicitly defined series. Letting Excel guess the source from a
    range fails when the range is a spill Excel has not calculated yet."""
    chart = ws.Shapes.AddChart2(-1, chart_type, *place(ws, top_left, bottom_right)).Chart
    existing = chart.SeriesCollection()
    while existing.Count:
        existing(1).Delete()
    for name, values in series:
        s = chart.SeriesCollection().NewSeries()
        s.Name = name
        s.Values = ws.Range(values)
        s.XValues = ws.Range(categories)
    chart.HasTitle = True
    chart.ChartTitle.Text = title_text
    return chart


def style_chart(chart, size: int = 9) -> None:
    """Quiet chart chrome: no border, small type, light gridlines."""
    chart.ChartArea.Format.Line.Visible = False
    chart.ChartArea.Font.Size = size
    chart.ChartArea.Font.Color = rgb(INK)
    if chart.HasTitle:
        tf = chart.ChartTitle.Format.TextFrame2.TextRange.Font
        tf.Size = 11
        tf.Bold = True
        tf.Fill.ForeColor.RGB = rgb(NAVY)
    try:
        grid = chart.Axes(2).MajorGridlines.Format.Line
        grid.ForeColor.RGB = rgb("#E4E8EE")
    except Exception:  # noqa: BLE001 - some chart types have no value-axis gridlines
        pass


# ------------------------------------------------------------- WorkingCapital
def working_capital(wb, pos: dict) -> None:
    ws = wb.Worksheets("WorkingCapital")
    title(ws, "Working capital: cash tied up, and stock close to its shelf life",
          '="Days are balances over the trailing "&TrailMonths&" months of flow ("&WCDays&'
          '" days). Month-end balances from tbl_Balances; flows from tbl_Sales."')
    band(ws, 4, "Trend: thirteen month-ends to the as-of month", "B", "O")
    put(ws, "B5", "Month-end")
    fx(ws, "C5", "=EDATE(AsOfMonth,-12)")
    fx(ws, "D5:O5", "=EDATE(C5,1)")
    lines = ["Accounts receivable", "Inventory", "Accounts payable", "Net working capital",
             "Trailing revenue", "Trailing COGS", "DSO (days sales outstanding)",
             "DIO (days inventory outstanding)", "DPO (days payables outstanding)",
             "Cash conversion cycle (DSO + DIO − DPO)"]
    column(ws, "B6", lines)
    fx(ws, "C6:O8", "=SUMIFS(tbl_Balances[balance],tbl_Balances[month],C$5,tbl_Balances[account],$B6)")
    fx(ws, "C9:O9", "=C6+C7-C8")
    for r, field in ((10, "net_revenue"), (11, "cogs")):
        fx(ws, f"C{r}:O{r}", f'=SUMIFS(tbl_Sales[{field}],tbl_Sales[month],">="&EDATE(C$5,1-TrailMonths),'
                             f'tbl_Sales[month],"<="&C$5)')
    fx(ws, "C12:O12", "=DAYSOF(C6,C10,WCDays)")
    fx(ws, "C13:O13", "=DAYSOF(C7,C11,WCDays)")
    fx(ws, "C14:O14", "=DAYSOF(C8,C11,WCDays)")
    fx(ws, "C15:O15", "=C12+C13-C14")
    numfmt(ws.Range("C6:O11"), F_MONEY)
    numfmt(ws.Range("C12:O15"), F_DAYS)
    total_row(ws.Range("B9:O9"))
    total_row(ws.Range("B15:O15"), double=True)

    cats = L.CATEGORIES
    band(ws, 17, "Days of inventory by category", "B", "O")
    band(ws, 26, "Share of shelf life used: days of inventory ÷ shelf life", "B", "O")
    for hdr in (18, 27):
        put(ws, f"B{hdr}", "Category")
        fx(ws, f"C{hdr}:O{hdr}", "=C$5")
        font(ws.Range(f"B{hdr}:O{hdr}"), bold=True)
    column(ws, "B19", cats)
    fx(ws, "B28:B33", "=$B19")
    fx(ws, "C19:O24",
       '=DAYSOF(SUMIFS(tbl_Balances[balance],tbl_Balances[month],C$18,tbl_Balances[account],"Inventory",'
       'tbl_Balances[category],$B19),SUMIFS(tbl_Sales[cogs],tbl_Sales[category],$B19,tbl_Sales[month],'
       '">="&EDATE(C$18,1-TrailMonths),tbl_Sales[month],"<="&C$18),WCDays)')
    fx(ws, "C28:O33", "=SAFEDIV(C19,XLOOKUP($B28,tbl_Product[category],tbl_Product[shelf_life_days]))")
    put(ws, "B34", "Risk line")
    fx(ws, "C34:O34", "=ShelfRisk")
    numfmt(ws.Range("C5:O5,C18:O18,C27:O27"), F_MONTH)
    numfmt(ws.Range("C19:O24"), F_DAYS)
    numfmt(ws.Range("C28:O34"), "0%")
    font(ws.Range("B34:O34"), color=MUTED, italic=True)
    c = ws.Range("C28:O33").FormatConditions.Add(XL_CELL_VALUE, 5, "=ShelfRisk")
    c.Font.Color = rgb(UNFAV)
    c.Font.Bold = True

    band(ws, 36, "As-of month: which stock is at risk?", "B", "O")
    header(ws, 37, ["Category", "Inventory", "Trailing COGS", "Days of inventory", "Shelf life (days)",
                    "Share of shelf life used", "Status", "Stock above the risk line"])
    fx(ws, "B38:B43", "=$B19")
    put(ws, "B44", "All categories")
    fx(ws, "C38:C43", '=SUMIFS(tbl_Balances[balance],tbl_Balances[month],AsOfMonth,'
                      'tbl_Balances[account],"Inventory",tbl_Balances[category],$B38)')
    fx(ws, "D38:D43", '=SUMIFS(tbl_Sales[cogs],tbl_Sales[category],$B38,tbl_Sales[month],">="&TrailStart,'
                      'tbl_Sales[month],"<="&AsOfMonth)')
    fx(ws, "C44:D44", "=SUM(C38:C43)")
    fx(ws, "E38:E44", "=DAYSOF(C38,D38,WCDays)")
    fx(ws, "F38:F43", "=XLOOKUP($B38,tbl_Product[category],tbl_Product[shelf_life_days])")
    fx(ws, "G38:G43", "=SAFEDIV(E38,F38)")
    fx(ws, "H38:H43", '=IF(G38>ShelfRisk,"At risk","OK")')
    fx(ws, "I38:I43", "=MAX(0,E38-ShelfRisk*F38)*SAFEDIV(D38,WCDays)")
    fx(ws, "I44", "=SUM(I38:I43)")
    numfmt(ws.Range("C38:D44"), F_MONEY)
    numfmt(ws.Range("E38:E44"), F_DAYS)
    numfmt(ws.Range("F38:F43"), "0")
    numfmt(ws.Range("G38:G43"), "0%")
    numfmt(ws.Range("I38:I44"), F_MONEY)
    total_row(ws.Range("B44:I44"), double=True)
    c = ws.Range("H38:H43").FormatConditions.Add(XL_CELL_VALUE, XL_EQUAL, '="At risk"')
    c.Interior.Color = rgb(FAIL_FILL)
    c.Font.Color = rgb(UNFAV)
    c.Font.Bold = True
    bars = ws.Range("G38:G43").FormatConditions.AddDatabar()
    bars.BarColor.Color = rgb("#9DB4CF")
    bars.MinPoint.Modify(0, 0)          # xlConditionValueNumber=0
    bars.MaxPoint.Modify(0, 1)
    fx(ws, "B46", '=LET(used,G38:G43,cat,B38:B43,worst,XMATCH(MAX(used),used),'
                  'INDEX(cat,worst)&" holds "&TEXT(INDEX(E38:E43,worst),"0")&" days of stock, "'
                  '&TEXT(MAX(used),"0%")&" of its "&INDEX(F38:F43,worst)&"-day shelf life"&'
                  'IF(MAX(used)>ShelfRisk,": "&MONEY(INDEX(I38:I43,worst))&" of it sits above the "'
                  '&TEXT(ShelfRisk,"0%")&" risk line, stock that needs a markdown or a slower '
                  'buy before it becomes a write-off.",", inside the risk line."))')
    ws.Range("B46:I47").Merge()
    ws.Range("B46").WrapText = True
    ws.Range("B46").VerticalAlignment = -4160
    font(ws.Range("B46"), italic=True)
    ws.Rows(46).RowHeight = 30

    chart = new_chart(ws, XL_LINE, "Q4", "Z24",
                      [(name, f"C{r}:O{r}") for r, name in enumerate([*cats, "Risk line"], start=28)],
                      "C27:O27", "Share of shelf life used, by category")
    chart.HasLegend = True
    chart.Legend.Position = -4107
    palette = ["#1F3A5F", "#7A8BA6", "#0F766E", "#5FA8A0", "#B7791F", UNFAV]
    for i, colour in enumerate(palette, start=1):
        line = chart.FullSeriesCollection(i).Format.Line
        line.ForeColor.RGB = rgb(colour)
        line.Weight = 2.25 if colour == UNFAV else 1.5
    risk = chart.FullSeriesCollection(7).Format.Line
    risk.ForeColor.RGB = rgb(MUTED)
    risk.DashStyle = 4
    risk.Weight = 1.25
    chart.Axes(2).TickLabels.NumberFormat = "0%"
    chart.Axes(1).TickLabels.NumberFormat = "mmm yy"
    style_chart(chart)

    widths(ws, {"A": 2, "B": 38, "C:O": 11, "P": 2})
    names = {"wc_months": "WorkingCapital!$C$5:$O$5", "wc_lines": "WorkingCapital!$C$6:$O$15",
             "wc_asof": "WorkingCapital!$O$12:$O$15", "wc_dio_category": "WorkingCapital!$C$19:$O$24",
             "wc_category": "WorkingCapital!$C$38:$I$43", "wc_at_risk_value": "WorkingCapital!$I$44"}
    for name, ref in names.items():
        wb.Names.Add(Name=name, RefersTo="=" + ref)


# ------------------------------------------------------------------- Channels
PIVOT_MEASURES = ["Revenue", "Revenue vs Budget %", "GM %", "Contribution %",
                  "Contribution per Case", "Revenue YoY %"]


def channels(wb, pos: dict) -> None:
    ws = wb.Worksheets("Channels")
    title(ws, "Channel economics: where the growth came from, and what it earned",
          "The same question answered twice: once with worksheet formulas, once from the Power Pivot "
          "data model (DAX measures, a PivotTable with slicers, CUBE functions). The two must agree.")
    band(ws, 4, '="Year to date by channel, "&TEXT(FYStart,"mmm")&"–"&TEXT(AsOfMonth,"mmm yyyy")'
                '&": worksheet formulas"', "B", "P")
    header(ws, 5, ["Channel", "Cases", "Revenue", "Gross margin", "GM %", "Freight", "Contribution",
                   "Contribution %", "Contribution per case", "Budget cases", "Budget contribution",
                   "Cases vs budget", "Contribution vs budget", "Revenue, data model", "Engines agree"])
    column(ws, "B6", L.CHANNELS)
    put(ws, "B10", "All channels")

    def s(field, crit="$B6"):
        return f"SUMIFS(tbl_Sales[{field}],tbl_Sales[channel],{crit},{S_YTD})"

    def b(field, crit="$B6"):
        return f"SUMIFS(tbl_Budget[{field}],tbl_Budget[channel],{crit},{B_YTD})"

    fx(ws, "C6:C9", "=" + s("units"))
    fx(ws, "D6:D9", "=" + s("net_revenue"))
    fx(ws, "E6:E9", "=D6-" + s("cogs"))
    fx(ws, "G6:G9", "=" + s("freight"))
    fx(ws, "K6:K9", "=" + b("units"))
    fx(ws, "L6:L9", f"={b('revenue')}-{b('cogs')}-{b('freight')}")
    fx(ws, "C10:E10", "=SUM(C6:C9)")
    fx(ws, "G10", "=SUM(G6:G9)")
    fx(ws, "K10:L10", "=SUM(K6:K9)")
    fx(ws, "F6:F10", "=SAFEDIV(E6,D6)")
    fx(ws, "H6:H10", "=E6-G6")
    fx(ws, "I6:I10", "=SAFEDIV(H6,D6)")
    fx(ws, "J6:J10", "=SAFEDIV(H6,C6)")
    fx(ws, "M6:M10", "=SAFEDIV(C6-K6,K6)")
    fx(ws, "N6:N10", "=H6-L6")
    year = '"[Calendar].[year].&["&FiscalYear&"]"'
    fx(ws, "O6:O9", f'=CUBEVALUE({CUBE},"[Measures].[Revenue]",{year},"[Calendar].[status].&[Closed]",'
                    '"[Channel].[channel].&["&$B6&"]")')
    fx(ws, "O10", f'=CUBEVALUE({CUBE},"[Measures].[Revenue]",{year},"[Calendar].[status].&[Closed]")')
    fx(ws, "P6:P10", '=IF(ABS(O6-D6)<=Tol,"Yes","No")')
    for rng, fmt in (("C6:C10", F_CASES), ("K6:K10", F_CASES), ("D6:E10", F_MONEY), ("G6:H10", F_MONEY),
                     ("L6:L10", F_MONEY), ("O6:O10", F_MONEY), ("F6:F10", F_PCT), ("I6:I10", F_PCT),
                     ("J6:J10", F_PRICE), ("M6:M10", F_VARPCT), ("N6:N10", F_VAR)):
        numfmt(ws.Range(rng), fmt)
    sign_colours(ws.Range("M6:N10"))
    total_row(ws.Range("B10:P10"), double=True)
    ws.Range("P6:P10").HorizontalAlignment = -4108

    band(ws, 12, "The same model as a PivotTable: Power Pivot measures written in DAX, filtered by slicers",
         "B", "P")
    put(ws, "B13", "Slicers filter the PivotTable only. The measures are defined once in the data model "
                   "(Data > Manage Data Model) and reused by the PivotTable and the CUBE formulas below.")
    font(ws.Range("B13"), color=MUTED)
    band(ws, 32, "Revenue by channel and year: CUBEVALUE formulas against the same data model", "B", "P")
    header(ws, 33, ["Channel", "", "", "", "YoY, DAX measure", "YoY, formulas", "Agree"])
    for c, back in zip("CDE", (2, 1, 0)):
        fx(ws, f"{c}33", f"=FiscalYear-{back}" if back else "=FiscalYear")
    numfmt(ws.Range("C33:E33"), '"FY"0')
    numfmt(ws.Range("E33"), '"FY"0" YTD"')
    fx(ws, "B34:B37", "=$B6")
    put(ws, "B38", "All channels")
    fx(ws, "C34:E37", f'=CUBEVALUE({CUBE},"[Measures].[Revenue]","[Channel].[channel].&["&$B34&"]",'
                      '"[Calendar].[year].&["&C$33&"]")')
    fx(ws, "C38:E38", f'=CUBEVALUE({CUBE},"[Measures].[Revenue]","[Calendar].[year].&["&C$33&"]")')
    fx(ws, "F34:F37", f'=CUBEVALUE({CUBE},"[Measures].[Revenue YoY %]","[Channel].[channel].&["&$B34&"]",'
                      f'{year})')
    fx(ws, "F38", f'=CUBEVALUE({CUBE},"[Measures].[Revenue YoY %]",{year})')
    py = ('SUMIFS(tbl_Sales[net_revenue],tbl_Sales[channel],{c},tbl_Sales[month],">="&EDATE(FYStart,-12),'
          'tbl_Sales[month],"<="&EDATE(AsOfMonth,-12))')
    fx(ws, "G34:G37", f"=SAFEDIV(D6-{py.format(c='$B34')},{py.format(c='$B34')})")
    fx(ws, "G38", f"=SAFEDIV(D10-SUM({py.format(c='B34:B37')}),SUM({py.format(c='B34:B37')}))")
    fx(ws, "H34:H38", '=IF(ABS(F34-G34)<0.00001,"Yes","No")')
    numfmt(ws.Range("C34:E38"), F_MONEY)
    numfmt(ws.Range("F34:G38"), F_VARPCT)
    ws.Range("H34:H38").HorizontalAlignment = -4108
    total_row(ws.Range("B38:H38"), double=True)
    ws.Rows(33).RowHeight = 30
    widths(ws, {"A": 2, "B": 22, "C:P": 12})
    ws.Rows(5).RowHeight = 42
    names = {"ch_formulas": "Channels!$C$6:$P$9", "ch_total": "Channels!$C$10:$P$10",
             "ch_cube_years": "Channels!$C$34:$E$38", "ch_cube_yoy": "Channels!$F$34:$H$38"}
    for name, ref in names.items():
        wb.Names.Add(Name=name, RefersTo="=" + ref)


def channel_pivot(wb) -> None:
    """The data-model PivotTable and its slicers. Created as soon as the model
    exists: Excel refuses a model PivotCache once the sheets hold thousands of
    uncalculated formulas."""
    ws = wb.Worksheets("Channels")
    pc = wb.PivotCaches().Create(XL_EXTERNAL, wb.Connections("ThisWorkbookDataModel"), 8)
    pt = pc.CreatePivotTable(TableDestination=ws.Range("B24"), TableName="ptChannels")
    pt.ManualUpdate = True
    pt.CubeFields("[Channel].[channel]").Orientation = XL_ROW_FIELD
    for m in PIVOT_MEASURES:
        pt.CubeFields(f"[Measures].[{m}]").Orientation = XL_DATA_FIELD
    pt.ManualUpdate = False
    pt.RowAxisLayout(1)
    pt.TableStyle2 = "PivotStyleLight16"
    pt.GrandTotalName = "All channels"
    pt.PivotFields("[Channel].[channel].[channel]").Caption = "Channel"
    for measure, caption in zip(PIVOT_MEASURES, ("Revenue ", "vs budget", "GM % ", "Contrib. %",
                                                 "Per case", "YoY")):
        pt.PivotFields(f"[Measures].[{measure}]").Caption = caption
    fiscal_year = str(int(wb.Names("FiscalYear").RefersToRange.Value))
    slicers = [("[Calendar].[year]", "[Calendar].[year].[year]", "Year", "B15", "D22", fiscal_year),
               ("[Product].[category]", "[Product].[category].[category]", "Category", "E15", "H22", None),
               ("[Region].[region]", "[Region].[region].[region]", "Region", "I15", "K22", None)]
    for field, level, caption, tl, br, select in slicers:
        cache = wb.SlicerCaches.Add2(pt, field, f"Slicer_{caption}")
        left, top, width, height = place(ws, tl, br)
        sl = cache.Slicers.Add(ws, level, caption, caption, top, left, width - 6, height)
        sl.Style = "SlicerStyleLight1"
        sl.NumberOfColumns = 3 if caption == "Year" else 2 if caption == "Category" else 1
        if select:
            items = cache.SlicerCacheLevels(1).SlicerItems
            chosen = [items(i).Name for i in range(1, items.Count + 1) if items(i).Caption == select]
            cache.VisibleSlicerItemsList = chosen


# ----------------------------------------------------------------------- Checks
def checks(wb, pos: dict, controls: dict) -> None:
    ws = wb.Worksheets("Checks")
    p = pos["pnl"]
    fc = pos["fc"]
    pv = pos["pvm"]
    r1, e_last = pv["ebitda_bridge"]
    b1, b2, _ = fc["basis"]
    title(ws, "Checks: the model proves its own numbers")
    ov = "IF(AND(sens_price=0,sens_volume=0),{d}-{c},0)"
    grid = "INDEX(sens_grid,XMATCH({pr},sens_price_axis),XMATCH({vo},sens_volume_axis))"
    items = [
        ("Sales rows loaded = source", "=ctrl_sales_rows", "=ROWS(tbl_Sales)",
         "The load dropped or duplicated nothing."),
        ("Sales cases = source control total", "=ctrl_sales_units", "=SUM(tbl_Sales[units])", ""),
        ("Sales revenue = source control total", "=ctrl_sales_revenue", "=SUM(tbl_Sales[net_revenue])", ""),
        ("Budget unpivot keeps every row", "=ROWS(tbl_BudgetUnits)*12", "=ROWS(tbl_Budget)",
         "Power Query's unpivot turned 12 month columns into 12 rows each."),
        ("Budget cases survive the unpivot", "=SUM(tbl_BudgetUnits[[2026-01]:[2026-12]])",
         "=SUM(tbl_Budget[units])", ""),
        ("Every budget row found its rate and category", "=0",
         "=COUNTBLANK(tbl_Budget[net_price])+COUNTBLANK(tbl_Budget[category])",
         "A failed merge would silently price rows at zero."),
        ("Budget revenue = cases × rate", "=SUMPRODUCT(tbl_Budget[units],tbl_Budget[net_price])",
         "=SUM(tbl_Budget[revenue])", ""),
        ("Balances: receivables, payables and each category every month-end",
         "=(2+ROWS(UNIQUE(tbl_Product[category])))*ROWS(UNIQUE(tbl_Balances[month]))",
         "=ROWS(tbl_Balances)", ""),
        ("As-of month has actuals", "=0", "=MAX(0,AsOfMonth-LastDataMonth)",
         "Forecasting from a month that has not closed would mix actuals and gaps."),
        ("P&L revenue, closed months = ledger", f"=SUMIFS(tbl_Sales[net_revenue],{S_YTD})",
         f"=PnL!W{p['outlook']['revenue']}", ""),
        ("P&L opex, closed months = ledger",
         '=SUMIFS(tbl_Opex[amount],tbl_Opex[scenario],"Actual",tbl_Opex[month],">="&FYStart,'
         'tbl_Opex[month],"<="&AsOfMonth)', f"=PnL!W{p['outlook']['opex_total']}", ""),
        ("P&L budget = budget table (FY revenue)", "=SUM(tbl_Budget[revenue])",
         f"=PnL!Q{p['outlook']['revenue']}", ""),
        ("EBITDA = revenue − COGS − freight − opex (FY)",
         "=PnL!P{revenue}-PnL!P{cogs}-PnL!P{freight}-PnL!P{opex_total}".format(**p["outlook"]),
         f"=PnL!P{p['outlook']['ebitda']}", ""),
        ("Variance block = outlook − budget (FY EBITDA)",
         f"=PnL!P{p['outlook']['ebitda']}-PnL!P{p['budget']['ebitda']}",
         f"=PnL!P{p['variance']['ebitda']}", ""),
        ("Outlook = actuals + forecast (FY revenue)",
         f"=PnL!W{p['outlook']['revenue']}+Forecast!P{fc['blocks']['revenue'][2]}",
         f"=PnL!P{p['outlook']['revenue']}", ""),
        ("Forecast volume = budget × run-rate on open months",
         f'=SUMPRODUCT(SUMIFS(tbl_Budget[units],tbl_Budget[category],Forecast!B{b1}:B{b2},'
         f'tbl_Budget[month],">"&AsOfMonth),Forecast!E{b1}:E{b2})',
         f"=Forecast!P{fc['blocks']['u0'][2]}", "Recomputed a second way, with array SUMIFS."),
        ("Formula engine = data model (YTD revenue)", "=Channels!D10", "=Channels!O10",
         "SUMIFS and the DAX measure are independent engines."),
        ("Formula engine = data model (each channel)", "=0", "=SUMPRODUCT(ABS(Channels!O6:O9-Channels!D6:D9))", ""),
        ("DAX prior-year measure = formula prior year",
         '=SUMIFS(tbl_Sales[net_revenue],tbl_Sales[month],">="&EDATE(FYStart,-12),tbl_Sales[month],'
         '"<="&EDATE(AsOfMonth,-12))',
         f'=CUBEVALUE({CUBE},"[Measures].[Revenue PY]","[Calendar].[year].&["&FiscalYear&"]")',
         "Time intelligence in DAX checked against plain SUMIFS."),
        ("PVM effects explain the month's GM variance", f"=PVM!U{pv['month'][2]}", f"=PVM!T{pv['month'][2]}",
         "Volume + mix + price + cost = actual − budget, with nothing left over."),
        ("PVM effects explain the YTD GM variance", f"=PVM!U{pv['ytd'][2]}", f"=PVM!T{pv['ytd'][2]}", ""),
        ("PVM reconciles in every one of the 48 segment rows", "=0",
         f"=SUMPRODUCT(ABS(PVM!V{pv['month'][0]}:V{pv['month'][1]}))+SUMPRODUCT(ABS(PVM!V{pv['ytd'][0]}:V{pv['ytd'][1]}))",
         "Not just in total: each segment's effects equal its own variance."),
        ("EBITDA bridge closes", f"=PVM!F{e_last}", f"=PVM!F{r1}+SUM(PVM!F{r1 + 1}:F{e_last - 1})", ""),
        ("Scenario table matches the live model", "=Scenarios!E17",
         "=XLOOKUP(ScenarioName,Scenarios!B18:B20,Scenarios!E18:E20)",
         "The data table re-ran the same model the sheets show."),
        ("Sensitivity grid centre = live EBITDA", "=Scenarios!C11",
         "=IFERROR(" + grid.format(pr="sens_price", vo="sens_volume") + ",Scenarios!C11)", ""),
        ("Tornado is exact: one price step", "=Scenarios!C38*0.01/TornadoStep",
         "=" + grid.format(pr="0.01", vo="0") + "-" + grid.format(pr="0", vo="0"),
         "The analytic tornado and the brute-force data table agree."),
        ("Tornado is exact: two volume steps", "=Scenarios!C37*0.02/TornadoStep",
         "=" + grid.format(pr="0", vo="0.02") + "-" + grid.format(pr="0", vo="0"), ""),
        ("Inventory by category = balance-sheet inventory",
         '=SUMIFS(tbl_Balances[balance],tbl_Balances[month],AsOfMonth,tbl_Balances[account],"Inventory")',
         "=WorkingCapital!C44", ""),
        ("Category days of inventory roll up to headline DIO", "=WorkingCapital!O13", "=WorkingCapital!E44", ""),
        ("Dashboard revenue tile = P&L", f"=PnL!W{p['outlook']['revenue']}", "=Dashboard!B5", ""),
        ("Dashboard EBITDA outlook tile = P&L", f"=PnL!P{p['outlook']['ebitda']}", "=Dashboard!K5", ""),
    ]
    band(ws, 4, "Controls", "B", "H")
    header(ws, 5, ["#", "Check", "Expected", "Actual", "Difference", "Status", "Why it matters"])
    ws.Range("C5").HorizontalAlignment = XL_LEFT
    ws.Range("H5").HorizontalAlignment = XL_LEFT
    first = 6
    for i, (label, expected, actual, why) in enumerate(items):
        r = first + i
        put(ws, f"B{r}", i + 1)
        put(ws, f"C{r}", label)
        fx(ws, f"D{r}", expected)
        fx(ws, f"E{r}", actual)
        if label.startswith("Tornado"):
            fx(ws, f"F{r}", "=" + ov.format(d=f"E{r}", c=f"D{r}"))
        else:
            fx(ws, f"F{r}", f"=E{r}-D{r}")
        fx(ws, f"G{r}", f'=IF(ISERROR(F{r}),"FAIL",IF(ABS(F{r})<=Tol,"PASS","FAIL"))')
        put(ws, f"H{r}", why)
    last = first + len(items) - 1
    numfmt(ws.Range(f"D{first}:E{last}"), '#,##0.00;(#,##0.00);"–"')
    numfmt(ws.Range(f"F{first}:F{last}"), '0.00;(0.00);"–"')
    ws.Range(f"G{first}:G{last}").HorizontalAlignment = -4108
    font(ws.Range(f"G{first}:G{last}"), bold=True)
    for text, colour, fore in (("PASS", PASS_FILL, FAV), ("FAIL", FAIL_FILL, UNFAV)):
        c = ws.Range(f"G{first}:G{last}").FormatConditions.Add(XL_CELL_VALUE, XL_EQUAL, f'="{text}"')
        c.Interior.Color = rgb(colour)
        c.Font.Color = rgb(fore)
    font(ws.Range(f"H{first}:H{last}"), color=MUTED)
    font(ws.Range(f"B{first}:B{last}"), color=MUTED)
    status = f"G{first}:G{last}"
    fx(ws, "B2", f'=IF(COUNTIF({status},"FAIL")=0,"All "&COUNTA({status})&" checks pass",'
                 f'COUNTIF({status},"FAIL")&" of "&COUNTA({status})&" checks fail: see column G")')
    font(ws.Range("B2"), bold=True, size=12)
    c = ws.Range("B2").FormatConditions.Add(2, Formula1=f'=COUNTIF({status},"FAIL")=0')
    c.Font.Color = rgb(FAV)
    c = ws.Range("B2").FormatConditions.Add(2, Formula1=f'=COUNTIF({status},"FAIL")>0')
    c.Font.Color = rgb(UNFAV)

    k0 = last + 3
    band(ws, k0, "Source control totals, taken from the CSVs when the workbook was built", "B", "H")
    header(ws, k0 + 1, ["", "Control total", "Value", "Source"])
    for i, (name, (label, value, source)) in enumerate(controls.items()):
        r = k0 + 2 + i
        put(ws, f"C{r}", label)
        put(ws, f"D{r}", value)
        put(ws, f"E{r}", source)
        input_cell(ws.Range(f"D{r}"))
        numfmt(ws.Range(f"D{r}"), "#,##0.00" if isinstance(value, float) else "#,##0")
        wb.Names.Add(Name=name, RefersTo=f"=Checks!$D${r}")
    font(ws.Range(f"E{k0 + 2}:E{k0 + 1 + len(controls)}"), color=MUTED)
    widths(ws, {"A": 2, "B": 5, "C": 52, "D:F": 16, "G": 9, "H": 62})
    wb.Names.Add(Name="checks_status", RefersTo=f"=Checks!$G${first}:$G${last}")
    wb.Names.Add(Name="checks_summary", RefersTo="=Checks!$B$2")
    pos["checks"] = (first, last)


__all__ = ["scenarios", "working_capital", "channels", "channel_pivot", "checks", "style_chart",
           "new_chart"]
