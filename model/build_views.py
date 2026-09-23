"""The two front sheets (Dashboard, Cover), then window, print and export settings."""
from __future__ import annotations

import layout as L
from xl import (
    F_MONEY, F_VARPCT, FAV, GREY, INK, MUTED, NAVY, RULE, TEAL, TILE, TITLE_FONT, UNFAV,
    XL_COLUMN_CLUSTERED, XL_COLUMN_STACKED, XL_EXPRESSION, XL_LANDSCAPE, XL_LINE, XL_SPARKLINE_LINE, XL_TYPE_PDF,
    XL_VCENTER, XL_WATERFALL, band, box, fill, font, fx, header, input_cell, new_chart, numfmt, place, put, rgb, row,
    sign_colours, style_chart, widths,
)

F_BILLIONS = '$#,##0.00,,,"B"'
F_MILLIONS_SIGNED = '+$#,##0.0,,"M";-$#,##0.0,,"M";"–"'


def tile(ws, first: str, label: str, value: str, value_fmt: str, delta: str | None, delta_fmt: str | None,
         spark: str | None, note: str, reverse: bool = False) -> None:
    last = chr(ord(first) + 2)
    for r in range(4, 9):
        ws.Range(f"{first}{r}:{last}{r}").Merge()
    rng = ws.Range(f"{first}4:{last}8")
    fill(rng, TILE)
    box(rng, RULE)
    rng.IndentLevel = 1
    fx(ws, f"{first}4", label) if label.startswith("=") else put(ws, f"{first}4", label)
    font(ws.Range(f"{first}4"), size=9, bold=True, color=MUTED)
    fx(ws, f"{first}5", value)
    numfmt(ws.Range(f"{first}5"), value_fmt)
    font(ws.Range(f"{first}5"), size=20, bold=True, color=NAVY, name=TITLE_FONT)
    ws.Range(f"{first}5").HorizontalAlignment = -4131
    if delta:
        fx(ws, f"{first}6", delta)
        if delta_fmt:
            numfmt(ws.Range(f"{first}6"), delta_fmt)
            sign_colours(ws.Range(f"{first}6"), reverse=reverse)
        font(ws.Range(f"{first}6"), size=9, bold=True)
        ws.Range(f"{first}6").HorizontalAlignment = -4131
    if spark:
        ws.Range(f"{first}7").SparklineGroups.Add(XL_SPARKLINE_LINE, spark)
        sg = ws.Range(f"{first}7").SparklineGroups(1)
        sg.SeriesColor.Color = rgb(TEAL)
        sg.LineWeight = 1.5
    fx(ws, f"{first}8", note)
    font(ws.Range(f"{first}8"), size=9, color=MUTED)


def dashboard(wb, pos: dict) -> None:
    ws = wb.Worksheets("Dashboard")
    fc = pos["forecast"]
    ob, oe, ot = fc["outlook"]
    m1, m2, _ = fc["months"]
    mo_f, mo_last, mo_tot = pos["monthly"]["ytd"]
    pnl = pos["pnl"]
    pv = pos["pvm"]
    widths(ws, {"A": 2, "B:S": 9.4, "T": 2})
    for r, h in {1: 34, 2: 18, 3: 8, 4: 18, 5: 34, 6: 17, 7: 26, 8: 17, 9: 10}.items():
        ws.Rows(r).RowHeight = h
    fx(ws, "B1", '="Portland General Electric: "&FiscalYear&" retail revenue, plan and outlook"')
    font(ws.Range("B1"), size=20, bold=True, color=NAVY, name=TITLE_FONT)
    fx(ws, "B2", '="Actuals from EIA through "&TEXT(AsOfMonth,"mmmm yyyy")&" · FERC Form 1 through "&FercYear&'
                 '" · weather: NOAA, Willamette Valley · scenario: "&WeatherScenario&" · "&checks_summary')
    font(ws.Range("B2"), size=10, color=MUTED)

    ind_row = mo_f + L.CLASSES.index("industrial")
    plan_ind = pos["plan"]["drivers"][0] + L.CLASSES.index("industrial")
    tile(ws, "B", '=UPPER("Retail revenue, "&FiscalYear&" outlook")', "=fy_revenue", F_BILLIONS,
         f"=Forecast!H{ot}", '"▲ "0.0%" on last year";"▼ "0.0%" on last year"', "Dashboard!C72:C83",
         f'="plan "&MONEY(Forecast!I{ot})&" · "&SIGNMONEY(Forecast!J{ot})')
    tile(ws, "E", "YEAR TO DATE AGAINST PLAN", "=INDEX(plan_bridge_total,7)", F_MILLIONS_SIGNED,
         "=INDEX(plan_bridge_total,8)", '"▲ "0.0%" over plan";"▼ "0.0%" under plan";"on plan"', None,
         '="plan from data to Dec "&(PlanYear-1)')
    fx(ws, "E7", '="weather "&SIGNMONEY(INDEX(plan_bridge_total,3))&" · price "&SIGNMONEY(INDEX(plan_bridge_total,5))')
    font(ws.Range("E7"), size=9, color=INK)
    tile(ws, "H", "WEATHER AGAINST NORMAL, YTD", "=INDEX(weather_impact_total,11)/1000",
         '+#,##0" GWh";-#,##0" GWh";"normal"', "=INDEX(weather_impact_total,12)",
         '"▲ "$#,##0.0,,"M of revenue";"▼ "$#,##0.0,,"M of revenue"', None,
         '="heating degree days "&TEXT(SAFEDIV(INDEX(weather_impact_total,1),SUM(INDEX(normal_hdd,1,1):INDEX('
         'normal_hdd,1,MonthsClosed))),"+0%;-0%")&" vs normal"')
    tile(ws, "K", "INDUSTRIAL LOAD, YTD", f"=Monthly!I{ind_row}", '+0.0%;-0.0%', None, None, "pnl_industrial_mwh",
         f'="plan assumed "&TEXT(Plan!I{plan_ind}-1,"+0%;-0%")')
    put(ws, "K6", "megawatt-hours, on last year")
    font(ws.Range("K6"), size=9, color=MUTED)
    tile(ws, "N", "AVERAGE RETAIL PRICE, YTD", f"=Monthly!J{mo_tot}", '0.00" ¢/kWh"', f"=Monthly!L{mo_tot}",
         '"▲ "0.0%" on last year";"▼ "0.0%" on last year"', None,
         f'="classes "&TEXT(MIN(Monthly!L{mo_f}:L{mo_last - 1}),"+0.0%")&" to "&TEXT(MAX(Monthly!L{mo_f}:L'
         f'{mo_last - 1}),"+0.0%")')
    tile(ws, "Q", '=UPPER("Operating margin, FERC "&FercYear)',
         "=INDEX(pnl_op_margin,1,MATCH(FercYear,pnl_years,0))", "0.0%",
         f"=PnL!P{pnl['op_margin']}", '"▲ "0.0%" pts since "0;"▼ "0.0%" pts since base"', "pnl_op_margin",
         '="before income tax, since "&BaseYear', reverse=False)
    numfmt(ws.Range("Q6"), '"▲ "0.0%" pts since base";"▼ "0.0%" pts since base"')

    band(ws, 10, "What the numbers say", "B", "S")
    lt = pv["long"][2]
    lines = [
        '="Retail revenue for "&TEXT(FYStart,"mmm")&"–"&TEXT(AsOfMonth,"mmm")&" came in "&MONEY(ABS(INDEX('
        'plan_bridge_total,7)))&" ("&TEXT(ABS(INDEX(plan_bridge_total,8)),"0.0%")&")"&IF(INDEX(plan_bridge_total,7)<0,'
        '" under"," over")&" plan: weather "&SIGNMONEY(INDEX(plan_bridge_total,3))&", price "&SIGNMONEY(INDEX('
        'plan_bridge_total,5))&", usage "&SIGNMONEY(INDEX(plan_bridge_total,4))&", customers "&SIGNMONEY(INDEX('
        'plan_bridge_total,2))&"."',
        '=LET(dev,SAFEDIV(INDEX(weather_impact_total,1),SUM(INDEX(normal_hdd,1,1):INDEX(normal_hdd,1,MonthsClosed))),'
        '"Heating degree days ran "&TEXT(ABS(dev),"0%")&IF(dev<0," below"," above")&" normal: "&TEXT(ABS(INDEX('
        'weather_impact_total,11))/1000,"#,##0")&" GWh of load and "&MONEY(ABS(INDEX(weather_impact_total,12)))&" of '
        'revenue"&IF(INDEX(weather_impact_total,12)<0," lost"," gained")&" to the thermostat, at actual prices.")',
        f'=LET(p,Monthly!L{mo_f}:L{mo_last - 1},"Every class paid more per kWh ("&TEXT(MIN(p),"+0.0%")&" to "&TEXT(MAX(p),'
        f'"+0.0%")&"), yet the average rose "&TEXT(Monthly!L{mo_tot},"+0.0%;-0.0%")&": the growth is industrial load '
        f'("&TEXT(Monthly!I{ind_row},"+0.0%")&" MWh), billed at "&TEXT(Monthly!J{ind_row},"0.0")&"¢ against "&TEXT('
        f'Monthly!J{mo_f},"0.0")&"¢ residential.")',
        f'="Since "&BaseYear&", retail revenue is up "&MONEY(PVM!N{lt})&": price "&SIGNMONEY(PVM!L{lt})&", volume "'
        f'&SIGNMONEY(PVM!J{lt})&", mix "&SIGNMONEY(PVM!K{lt})&". Power cost rose faster than revenue, and operating '
        'margin fell "&TEXT(ABS(PnL!P' + str(pnl["op_margin"]) + '),"0.0%")&" points."',
        f'="Full year "&FiscalYear&": "&MONEY(fy_revenue)&", "&TEXT(Forecast!H{ot},"+0.0%;-0.0%")&" on "&(FiscalYear-1)&'
        f'IF(ISNUMBER(Forecast!J{ot}),", "&MONEY(ABS(Forecast!J{ot}))&IF(Forecast!J{ot}<0," under"," over")&" plan","")'
        '&". A point of rate on the open months is worth "&MONEY(INDEX(tornado,1))&"; 50 MW of new data-center load, "'
        '&MONEY(INDEX(tornado,2))&"."',
    ]
    for i, f in enumerate(lines):
        r = 11 + i
        ws.Range(f"B{r}:S{r}").Merge()
        fx(ws, f"B{r}", f)
        ws.Range(f"B{r}").WrapText = True
        ws.Range(f"B{r}").VerticalAlignment = XL_VCENTER
        ws.Rows(r).RowHeight = 30
        font(ws.Range(f"B{r}"), size=10, color=INK)
    ws.Rows(16).RowHeight = 8

    # ---- chart data (below the printed page)
    put(ws, "B68", "Chart data: the charts above read these cells, which read the model.")
    font(ws.Range("B68"), italic=True, color=MUTED)
    row(ws, "B71", ["Month", "Actual", "Forecast", "Plan", "Prior year"])
    fx(ws, "B72:B83", f'=TEXT(Forecast!B{m1},"mmm")')
    for c, src in zip("CDEF", "CDEF"):
        fx(ws, f"{c}72:{c}83", f"=Forecast!{src}{m1}")
    c0, c1 = pos["plan"]["chart"]
    fx(ws, "B86:B91", f"=Plan!N{c0}")
    fx(ws, "C86:C91", f"=Plan!O{c0}")
    b0, b1 = pv["bridge"]
    fx(ws, "B94:B98", f"=PVM!B{b0}")
    fx(ws, "C94:C98", f"=PVM!C{b0}")
    numfmt(ws.Range("C72:F83"), F_MONEY)
    numfmt(ws.Range("C86:C91"), '$#,##0.0,,"M";-$#,##0.0,,"M"')
    numfmt(ws.Range("C94:C98"), '$#,##0,,"M";-$#,##0,,"M"')

    for r in range(17, 54):
        ws.Rows(r).RowHeight = 15
    ch = new_chart(ws, XL_COLUMN_STACKED, "B17", "J34",
                   [(n, f"{c}72:{c}83") for n, c in zip(("Actual", "Forecast", "Plan", "Prior year"), "CDEF")],
                   "B72:B83", '="Retail revenue by month"')
    ch.ChartTitle.Text = "Retail revenue by month: actual, forecast, plan and last year"
    for i, colour in enumerate([NAVY, "#9DB4CF", TEAL, GREY], start=1):
        s = ch.FullSeriesCollection(i)
        if i <= 2:
            s.Format.Fill.ForeColor.RGB = rgb(colour)
        else:
            s.ChartType = XL_LINE
            s.Format.Line.ForeColor.RGB = rgb(colour)
            s.Format.Line.Weight = 2.25 if i == 3 else 1.5
            if i == 4:
                s.Format.Line.DashStyle = 4
    ch.ChartGroups(1).GapWidth = 45
    ch.Axes(2).TickLabels.NumberFormat = '$#,##0,,"M"'
    ch.HasLegend = True
    ch.Legend.Position = -4107
    style_chart(ch)

    for data_range, tl, br, totals, caption in (("B86:C91", "K17", "S34", (1, 6), "Plan to actual, year to date"),
                                               ("B94:C98", "B36", "J53", (1, 5), "Retail revenue since the base year")):
        ws.Activate()
        ws.Range(data_range).Select()
        wf = ws.Shapes.AddChart2(-1, XL_WATERFALL, *place(ws, tl, br)).Chart
        series = wf.FullSeriesCollection(1)
        for p in totals:
            series.Points(p).IsTotal = True
        series.HasDataLabels = True
        wf.HasTitle = True
        wf.ChartTitle.Text = caption
        wf.HasLegend = False
        style_chart(wf)
    ws.Range("B1").Select()

    pc = new_chart(ws, XL_COLUMN_CLUSTERED, "K36", "S53", [("Industrial MWh growth a year", "Peers!C14:G14")],
                   "Peers!C4:G4", "Industrial load growth a year: PGE and its peers")
    pc.HasLegend = False
    pc.FullSeriesCollection(1).Format.Fill.ForeColor.RGB = rgb("#9DB4CF")
    pc.FullSeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = rgb(NAVY)
    pc.FullSeriesCollection(1).HasDataLabels = True
    pc.FullSeriesCollection(1).DataLabels().NumberFormat = "0.0%"
    pc.Axes(2).TickLabels.NumberFormat = "0%"
    pc.Axes(1).TickLabelPosition = -4134
    style_chart(pc)

    # ---- bottom tables
    ws.Rows(54).RowHeight = 8
    band(ws, 55, '="Plan against actual by class, "&TEXT(FYStart,"mmm")&"–"&TEXT(AsOfMonth,"mmm")', "B", "J")
    band(ws, 55, '="Full-year outlook by class, "&FiscalYear', "K", "S")
    pb1, pb2, pbt = pos["plan"]["bridge"]
    for c, label in zip("BEGIJ", ["Class", "Plan", "Actual", "Variance", "%"]):
        put(ws, f"{c}56", label)
    for c, label in zip("KMOQS", ["Class", "Actual YTD", "Forecast", "Full year", "Growth"]):
        put(ws, f"{c}56", label)
    for i in range(5):
        r = 57 + i
        src_p = pb1 + i if i < 4 else pbt
        src_o = ob + i if i < 4 else ot
        fx(ws, f"B{r}", f"=Plan!B{src_p}")
        fx(ws, f"E{r}", f"=Plan!D{src_p}")
        fx(ws, f"G{r}", f"=Plan!I{src_p}")
        fx(ws, f"I{r}", f"=Plan!J{src_p}")
        fx(ws, f"J{r}", f"=Plan!K{src_p}")
        fx(ws, f"K{r}", f"=Forecast!B{src_o}")
        fx(ws, f"M{r}", f"=Forecast!D{src_o}")
        fx(ws, f"O{r}", f"=Forecast!E{src_o}")
        fx(ws, f"Q{r}", f"=Forecast!F{src_o}")
        fx(ws, f"S{r}", f"=Forecast!H{src_o}")
    numfmt(ws.Range("E57:G61"), '$#,##0.0,,"M"')
    numfmt(ws.Range("I57:I61"), F_MILLIONS_SIGNED)
    numfmt(ws.Range("J57:J61"), F_VARPCT)
    numfmt(ws.Range("M57:Q61"), '$#,##0.0,,"M"')
    numfmt(ws.Range("S57:S61"), F_VARPCT)
    sign_colours(ws.Range("I57:J61"))
    sign_colours(ws.Range("S57:S61"))
    for rng in ("B56:J56", "K56:S56"):
        font(ws.Range(rng), bold=True, color=INK, size=9)
        ws.Range(rng).Borders(9).Color = rgb(NAVY)
    for c in "EGIJMOQS":
        ws.Range(f"{c}56").HorizontalAlignment = -4152
    font(ws.Range("B61:S61"), bold=True)
    put(ws, "B62", "Independent analysis of public regulatory filings (EIA-861M, FERC Form 1 via Catalyst Cooperative's "
                   "PUDL, NOAA nClimDiv). Not affiliated with or endorsed by Portland General Electric. Built in Excel "
                   "by Kush Patel: github.com/KushPatel29/excel-fpa-model")
    font(ws.Range("B62"), size=8, color=MUTED, italic=True)
    specs = {"kpi_fy_revenue": "$B$5", "kpi_ytd_vs_plan": "$E$5", "kpi_weather_gwh": "$H$5",
             "kpi_industrial_growth": "$K$5", "kpi_price": "$N$5", "kpi_op_margin": "$Q$5",
             "dash_commentary": "$B$11:$B$15", "dash_plan_table": "$B$57:$J$61", "dash_outlook_table": "$K$57:$S$61"}
    for name, ref in specs.items():
        wb.Names.Add(Name=name, RefersTo=f"=Dashboard!{ref}")


SHEET_GUIDE = [
    ("Dashboard", "One page: plan against actual, weather, price mix, the outlook, and a summary written by formulas.",
     "Formula-driven commentary, waterfalls, combo chart, sparklines, LET"),
    ("PnL", "Twelve years of PGE's electric P&L from FERC Form 1, with unit economics.",
     "SUMIFS over a long table with array criteria, named LAMBDAs, CAGR"),
    ("PVM", "Where retail revenue growth came from: volume, mix and price by customer class.",
     "Exact three-effect decomposition, waterfall"),
    ("Monthly", "EIA's monthly survey by class, a revenue heat map, and the EIA-to-FERC tie-out.",
     "SUMIFS, colour scales, cross-source reconciliation"),
    ("Weather", "How much load follows the thermostat; this year's weather impact.",
     "LINEST on FILTER + HSTACK, normals with FILTER/AVERAGE"),
    ("Plan", "A plan built only from data before the year, against actuals, split four ways; backtested.",
     "Month-by-class model, what-if data table on the plan year"),
    ("Forecast", "The rest of the year from drivers and observed weather; a FORECAST.ETS cross-check.",
     "XLOOKUP with fallbacks, FORECAST.ETS"),
    ("Scenarios", "Weather, rates, data-center load: tables that re-run the model, and a tornado.",
     "One- and two-variable data tables, SORTBY + HSTACK"),
    ("Peers", "PGE against four Pacific Northwest utilities on price, growth and cost.",
     "SUMIFS by respondent, RANK.EQ"),
    ("Explore", "The same data as a Power Pivot model: PivotTable, slicers, CUBEVALUE.",
     "Power Pivot, DAX (TREATAS time intelligence), CUBEVALUE"),
    ("Checks", "31 controls that prove the numbers, each with its own tolerance.", "Reconciliations"),
    ("Assumptions", "Model controls and weather scenarios.", "Named inputs, validation"),
    ("PQ_Monthly", "EIA's wide layout unpivoted to one row per month and class.", "Power Query: unpivot, split, pivot"),
    ("PQ_Weather", "NOAA's fixed-width text parsed into degree days by month.", "Power Query: split by position"),
    ("Data_EIA", "The public extracts, exactly as published.", "Excel tables, structured references"),
]


def cover(wb, pos: dict) -> None:
    ws = wb.Worksheets("Cover")
    lt = pos["pvm"]["long"][2]
    ot = pos["forecast"]["outlook"][2]
    widths(ws, {"A": 3, "B": 36, "C": 74, "D": 56})
    put(ws, "B2", L.COMPANY)
    font(ws.Range("B2"), size=28, bold=True, color=NAVY, name=TITLE_FONT)
    ws.Rows(2).RowHeight = 40
    put(ws, "B3", "Utility FP&A model on public filings: FERC, EIA and NOAA")
    font(ws.Range("B3"), size=15, color=INK, name=TITLE_FONT)
    put(ws, "B4", "Plan against actual · weather normalization · price-volume-mix · rolling forecast · scenarios · peers")
    font(ws.Range("B4"), size=10, color=TEAL, bold=True)
    put(ws, "B6", "Every figure is a live formula, Power Query step or DAX measure over public regulatory data: PGE's "
                  "monthly returns to the U.S. Energy Information Administration, its audited annual FERC Form 1 "
                  "(processed by Catalyst Cooperative's PUDL project), and NOAA's degree days for the Willamette Valley.")
    ws.Range("B6:D6").Merge()
    ws.Range("B6").WrapText = True
    ws.Rows(6).RowHeight = 30
    put(ws, "B7", "Independent analysis. Not affiliated with or endorsed by Portland General Electric.")
    font(ws.Range("B6:B7"), size=10, color=MUTED)
    put(ws, "B8", "Built by Kush Patel · github.com/KushPatel29/excel-fpa-model")
    ws.Hyperlinks.Add(Anchor=ws.Range("B8"), Address="https://github.com/KushPatel29/excel-fpa-model",
                      TextToDisplay="Built by Kush Patel · github.com/KushPatel29/excel-fpa-model")
    font(ws.Range("B8"), size=10)

    put(ws, "B10", "Model status")
    font(ws.Range("B10"), bold=True)
    fx(ws, "C10", "=checks_summary")
    ws.Hyperlinks.Add(Anchor=ws.Range("D10"), Address="", SubAddress="'Checks'!A1", TextToDisplay="Open the checks →")
    font(ws.Range("C10"), bold=True)
    for formula, colour in (('=COUNTIF(checks_status,"FAIL")=0', FAV), ('=COUNTIF(checks_status,"FAIL")>0', UNFAV)):
        c = ws.Range("C10").FormatConditions.Add(XL_EXPRESSION, Formula1=formula)
        c.Font.Color = rgb(colour)

    band(ws, 12, "Five questions it answers, live from the model", "B", "D")
    qa = [
        ("Is the year on plan, and if not, why?",
         '="Year to date "&MONEY(ABS(INDEX(plan_bridge_total,7)))&IF(INDEX(plan_bridge_total,7)<0," under"," over")&'
         '" plan ("&TEXT(INDEX(plan_bridge_total,8),"+0.0%;-0.0%")&"): weather "&SIGNMONEY(INDEX(plan_bridge_total,3))&'
         '", price "&SIGNMONEY(INDEX(plan_bridge_total,5))&", usage "&SIGNMONEY(INDEX(plan_bridge_total,4))&"."',
         "Plan"),
        ("How much of that is the weather?",
         '="Weather against normal: "&TEXT(INDEX(weather_impact_total,11)/1000,"+#,##0;-#,##0")&" GWh and "&SIGNMONEY('
         'INDEX(weather_impact_total,12))&" of revenue this year."', "Weather"),
        ("Where did revenue growth come from?",
         f'="Since "&BaseYear&": price "&SIGNMONEY(PVM!L{lt})&", volume "&SIGNMONEY(PVM!J{lt})&", mix "&SIGNMONEY('
         f'PVM!K{lt})&"."', "PVM"),
        ("Where will the year land?",
         f'="Retail revenue "&MONEY(fy_revenue)&", "&TEXT(Forecast!H{ot},"+0.0%;-0.0%")&" on "&(FiscalYear-1)&"; a '
         'point of rate on the open months is worth "&MONEY(INDEX(tornado,1))&"."', "Forecast"),
        ("How does PGE compare with its neighbours?",
         '="Industrial load growing "&TEXT(INDEX(peers_table,8,1),"0.0%")&" a year, rank "&INDEX(peers_rank,8)&" of 5; '
         'retail price "&TEXT(INDEX(peers_table,4,1),"0.0")&"¢/kWh, rank "&INDEX(peers_rank,4)&" of 5."', "Peers"),
    ]
    for i, (q, answer, sheet) in enumerate(qa):
        r = 13 + i
        put(ws, f"B{r}", q)
        fx(ws, f"C{r}", answer)
        ws.Hyperlinks.Add(Anchor=ws.Range(f"D{r}"), Address="", SubAddress=f"'{sheet}'!A1",
                          TextToDisplay=f"{sheet} →")
        font(ws.Range(f"B{r}"), bold=True)
        ws.Range(f"C{r}").WrapText = True
        ws.Range(f"B{r}:D{r}").VerticalAlignment = -4160
        ws.Rows(r).RowHeight = 28
    wb.Names.Add(Name="cover_answers", RefersTo="=Cover!$C$13:$C$17")

    band(ws, 19, "Sheets", "B", "D")
    header(ws, 20, ["Sheet", "What it shows", "Excel on show"], "B", wrap=False)
    ws.Range("C20:D20").HorizontalAlignment = -4131
    for i, (sheet, what, how) in enumerate(SHEET_GUIDE):
        r = 21 + i
        ws.Hyperlinks.Add(Anchor=ws.Range(f"B{r}"), Address="", SubAddress=f"'{sheet}'!A1", TextToDisplay=sheet)
        put(ws, f"C{r}", what)
        put(ws, f"D{r}", how)
    last = 21 + len(SHEET_GUIDE) - 1
    font(ws.Range(f"D21:D{last}"), color=MUTED)
    ws.Range(f"C21:D{last}").WrapText = True

    h = last + 2
    band(ws, h, "How to use it", "B", "D")
    steps = [
        ("1  Move a lever", "Scenarios!C4:C9: weather, a rate change, new data-center load. Every sheet follows."),
        ("2  Re-run the plan as of another year", "Plan!C4. The backtest on the same sheet does it for four years."),
        ("3  Explore the data model", "The slicers on Explore filter a PivotTable built on DAX measures."),
        ("4  Refresh the data", "python pipeline/fetch_sources.py downloads the latest public releases; rebuild."),
    ]
    for i, (step, how) in enumerate(steps):
        put(ws, f"B{h + 1 + i}", step)
        put(ws, f"C{h + 1 + i}", how)
    font(ws.Range(f"B{h + 1}:B{h + len(steps)}"), bold=True)

    k = h + len(steps) + 2
    band(ws, k, "Sources and conventions", "B", "D")
    notes = [
        ("EIA-861M", "U.S. Energy Information Administration, monthly electric utility sales and revenue. Public domain."),
        ("FERC Form 1", "Federal Energy Regulatory Commission annual report, via Catalyst Cooperative's PUDL "
                        "(CC BY 4.0)."),
        ("NOAA nClimDiv", "NOAA National Centers for Environmental Information, divisional degree days. Public domain."),
    ]
    for i, (name, text) in enumerate(notes):
        put(ws, f"B{k + 1 + i}", name)
        put(ws, f"C{k + 1 + i}", text)
    put(ws, f"B{k + 4}", 0.01)
    numfmt(ws.Range(f"B{k + 4}"), "+0.0%")
    input_cell(ws.Range(f"B{k + 4}"))
    put(ws, f"C{k + 4}", "Input: blue on yellow. Change these; everything else is calculated.")
    put(ws, f"B{k + 5}", "+1,234   (1,234)")
    put(ws, f"C{k + 5}", "Variances are signed so favourable is positive (green) and unfavourable is in brackets (red).")
    ws.Range(f"B{k + 4}:B{k + 5}").HorizontalAlignment = -4108
    font(ws.Range(f"C{k + 1}:C{k + 5}"), color=MUTED)
    font(ws.Range(f"B{k + 1}:B{k + 3}"), bold=True)
    pos["cover_last"] = k + 5


def finish(xl, wb, pos: dict) -> None:
    """Gridlines, zoom, frozen panes, print setup, and every sheet opened at A1."""
    freeze = {"PnL": "D5", "Data_EIA": "A5", "Data_FERC": "A5", "PQ_Monthly": "A5", "PQ_Weather": "A5",
              "Data_NOAA": "A5"}
    xl.ScreenUpdating = True
    for ws in wb.Worksheets:
        ws.Activate()
        win = xl.ActiveWindow
        win.DisplayGridlines = ws.Name.startswith(("Data_", "PQ_"))
        win.Zoom = 90 if ws.Name not in ("Dashboard", "Cover") else 100
        if ws.Name in freeze:
            ws.Range(freeze[ws.Name]).Select()
            win.FreezePanes = True
        ws.Range("A1").Select()
        win.ScrollRow = 1
        win.ScrollColumn = 1
    xl.ScreenUpdating = False
    pvm_end = pos["pvm"]["bridge"][1] + 13
    plan_end = pos["plan"]["backtest"][1] + 1
    wt = pos["weather"]["end"]
    mt = pos["forecast"]["months"][2]
    chk = pos["checks"][1] + 8
    tie = pos["monthly"]["tieout"][1] + 1
    areas = {  # sheet: (print area, fit on one page tall)
        "Cover": (f"$A$1:$D${pos['cover_last']}", True),
        "Dashboard": ("$A$1:$T$62", True),
        "PnL": (f"$B$1:$R${max(pos['pnl'].values())}", True),
        "PVM": (f"$B$1:$O${pvm_end}", True),
        "Plan": (f"$B$1:$U${plan_end}", 2),
        "Weather": (f"$B$1:$R${wt}", True),
        "Forecast": (f"$B$1:$N${mt}", 2),
        "Scenarios": ("$B$1:$T$48", True),
        "Peers": ("$B$1:$J$38", True),
        "Monthly": (f"$B$1:$Q${tie}", True),
        "Checks": (f"$B$1:$I${chk}", True),
    }
    xl.PrintCommunication = False
    for name, (area, one_page) in areas.items():
        ps = wb.Worksheets(name).PageSetup
        ps.PrintArea = area
        ps.Orientation = XL_LANDSCAPE
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = one_page if isinstance(one_page, int) and one_page is not True else (1 if one_page else False)
        ps.CenterHorizontally = True
        ps.LeftMargin = ps.RightMargin = xl.InchesToPoints(0.35)
        ps.TopMargin = xl.InchesToPoints(0.45)
        ps.BottomMargin = xl.InchesToPoints(0.5)
        ps.FooterMargin = xl.InchesToPoints(0.2)
        ps.LeftFooter = "&8Portland General Electric utility FP&&A model · public data (EIA, FERC via PUDL, NOAA)"
        ps.CenterFooter = "&8&A"
        ps.RightFooter = "&8Page &P of &N"
    xl.PrintCommunication = True
    wb.Worksheets("Cover").Activate()


PDF_ORDER = ["Cover", "Dashboard", "PnL", "PVM", "Plan", "Weather", "Forecast", "Scenarios", "Peers", "Monthly", "Checks"]


def export_pdf(wb, path) -> None:
    wb.Worksheets(PDF_ORDER).Select()
    wb.ActiveSheet.ExportAsFixedFormat(XL_TYPE_PDF, str(path), 0, True, False)
    wb.Worksheets("Cover").Select()
