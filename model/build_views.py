"""The two front sheets (Dashboard, Cover), then window, print and export settings."""
from __future__ import annotations

import layout as L
from build_analysis import new_chart, style_chart
from xl import (
    AMBER, F_MONEY, F_MONEY_M, F_PCT, F_VAR, FAV, GREY, INK, MUTED, NAVY, RULE, TEAL, TILE,
    TITLE_FONT, UNFAV, XL_BAR_STACKED, XL_COLUMN_CLUSTERED, XL_COLUMN_STACKED,
    XL_EXPRESSION, XL_LANDSCAPE, XL_LINE, XL_SPARKLINE_LINE, XL_TYPE_PDF, XL_VCENTER,
    XL_WATERFALL, band, box, fill, font, fx, header, input_cell, numfmt, place, put, rgb, row,
    sign_colours, widths,
)

TILES = "BEHKNQ"   # first column of each dashboard tile; each tile is three columns wide


def tile(ws, first: str, label: str, value: str, value_fmt: str, delta: str | None,
         delta_fmt: str | None, spark: str | None, note: str, reverse: bool = False) -> None:
    last = chr(ord(first) + 2)
    for r in range(4, 9):
        ws.Range(f"{first}{r}:{last}{r}").Merge()
    rng = ws.Range(f"{first}4:{last}8")
    fill(rng, TILE)
    box(rng, RULE)
    rng.IndentLevel = 1
    put(ws, f"{first}4", label)
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
    p = pos["pnl"]["outlook"]
    pv = pos["pvm"]
    y1, y2, yt = pv["ytd"]
    c1, c2, ct = pv["category"]
    h1, h2, _ = pv["channel"]
    widths(ws, {"A": 2, "B:S": 9.4, "T": 2})
    for r, h in {1: 34, 2: 18, 3: 8, 4: 18, 5: 34, 6: 17, 7: 26, 8: 17, 9: 10}.items():
        ws.Rows(r).RowHeight = h
    fx(ws, "B1", '="Kestrel Bay Provisions: FY"&FiscalYear&" performance and outlook"')
    font(ws.Range("B1"), size=20, bold=True, color=NAVY, name=TITLE_FONT)
    fx(ws, "B2", '="Actuals through "&TEXT(AsOfMonth,"mmmm yyyy")&IF(MonthsClosed<12," · forecast "'
                 '&TEXT(EDATE(AsOfMonth,1),"mmmm")&" to December","")&" · scenario: "&ScenarioName'
                 '&" · "&checks_summary')
    font(ws.Range("B2"), size=10, color=MUTED)

    up_down = '"▲ "$#,##0,"K vs budget";"▼ "$#,##0,"K vs budget";"on budget"'
    tile(ws, "B", "REVENUE, YEAR TO DATE", f"=PnL!W{p['revenue']}", F_MONEY_M,
         f"=PnL!Y{p['revenue']}", up_down, f"PnL!D{p['revenue']}:O{p['revenue']}",
         f'="vs prior year "&TEXT(PnL!AB{p["revenue"]},"+0.0%;-0.0%")')
    tile(ws, "E", "GROSS MARGIN %, YTD", f"=PnL!W{p['gm_pct']}", "0.0%",
         f"=PnL!Y{p['gm_pct']}*100", '"▲ "0.0" pts vs budget";"▼ "0.0" pts vs budget";"on budget"',
         f"PnL!D{p['gm_pct']}:O{p['gm_pct']}", f'="budget "&TEXT(PnL!X{p["gm_pct"]},"0.0%")')
    tile(ws, "H", "EBITDA, YEAR TO DATE", f"=PnL!W{p['ebitda']}", F_MONEY_M,
         f"=PnL!Y{p['ebitda']}", up_down, f"PnL!D{p['ebitda']}:O{p['ebitda']}",
         f'="margin "&TEXT(PnL!W{p["ebitda_pct"]},"0.0%")&" vs "&TEXT(PnL!X{p["ebitda_pct"]},"0.0%")'
         '&" budget"')
    tile(ws, "K", "EBITDA, FULL-YEAR OUTLOOK", f"=PnL!P{p['ebitda']}", F_MONEY_M,
         f"=PnL!R{p['ebitda']}", up_down, None, f'="budget "&MONEY(PnL!Q{p["ebitda"]})')
    fx(ws, "K7", '="scenario: "&ScenarioName')
    font(ws.Range("K7"), size=9, color=MUTED)
    tile(ws, "N", "PRICE RISE TO REACH BUDGET", "=breakeven_uplift", "0.0%",
         '="on the "&(12-MonthsClosed)&" open months"', None, None,
         '="or "&TEXT(breakeven_volume,"0%")&" more volume"')
    font(ws.Range("N6"), color=INK)
    tile(ws, "Q", "CASH CONVERSION CYCLE", "=WorkingCapital!O15", '0.0" days"',
         "=WorkingCapital!O15-WorkingCapital!C15",
         '"▲ "0.0" days on a year ago";"▼ "0.0" days on a year ago";"flat on a year ago"',
         "WorkingCapital!C15:O15",
         '=LET(u,WorkingCapital!G38:G43,c,WorkingCapital!B38:B43,INDEX(c,XMATCH(MAX(u),u))&": "'
         '&TEXT(MAX(u),"0%")&" of shelf life")', reverse=True)

    band(ws, 10, "What the numbers say", "B", "S")
    cat_cost = f"PVM!$F${c1}:$F${c2}"
    lines = [
        f'="Revenue is "&MONEY(ABS(PnL!Y{p["revenue"]}))&IF(PnL!Y{p["revenue"]}>=0," ahead of"," behind")'
        f'&" budget through "&TEXT(AsOfMonth,"mmmm")&" ("&TEXT(PnL!Z{p["revenue"]},"+0.0%;-0.0%")&"), '
        f'yet EBITDA is "&MONEY(ABS(PnL!Y{p["ebitda"]}))&IF(PnL!Y{p["ebitda"]}>=0," ahead"," behind")'
        f'&" ("&TEXT(PnL!Z{p["ebitda"]},"+0.0%;-0.0%")&")."&IF(AND(PnL!Y{p["revenue"]}>0,'
        f'PnL!Y{p["ebitda"]}<0)," The growth has not reached the bottom line.","")',
        f'=LET(v,PVM!C{ct},m,PVM!D{ct},pr,PVM!E{ct},c,PVM!F{ct},chm,PVM!D{h1}:D{h2},chn,PVM!B{h1}:B{h2},'
        f'mixch,INDEX(chn,XMATCH(MIN(chm),chm)),low,INDEX(Channels!B6:B9,XMATCH(MIN(Channels!J6:J9),'
        f'Channels!J6:J9)),cc,{cat_cost},cp,PVM!$E${c1}:$E${c2},cn,PVM!$B${c1}:$B${c2},k,XMATCH(MIN(cc),cc),'
        '"Gross margin, year to date: volume "&SIGNMONEY(v)&", price "&SIGNMONEY(pr)&", mix "&SIGNMONEY(m)'
        '&", unit cost "&SIGNMONEY(c)&". Mix is negative because volume shifted toward "&mixch'
        '&IF(mixch=low,", the channel that earns least per case","")&"; "&INDEX(cn,k)&" costs rose"'
        '&IF(INDEX(cc,k)+INDEX(cp,k)<0," faster than its prices.","."))',
        f'="Below gross margin, freight "&SIGNMONEY(PnL!Y{p["freight"]})&" and operating expenses "'
        f'&SIGNMONEY(PnL!Y{p["opex_total"]})&" (mostly "&INDEX(TRIM(PnL!B{p["opex_sm"]}:B{p["opex_tech"]}),'
        f'XMATCH(MIN(PnL!Y{p["opex_sm"]}:Y{p["opex_tech"]}),PnL!Y{p["opex_sm"]}:Y{p["opex_tech"]}))'
        f'&") leave EBITDA at "&MONEY(PnL!W{p["ebitda"]})&" against a budget of "&MONEY(PnL!X{p["ebitda"]})&"."',
        f'=IF(breakeven_gap>0,"Full year on the "&ScenarioName&" scenario: EBITDA "&MONEY(PnL!P{p["ebitda"]})'
        f'&" against a budget of "&MONEY(PnL!Q{p["ebitda"]})&". Closing the "&MONEY(breakeven_gap)'
        '&" gap takes a "&TEXT(breakeven_uplift,"0.0%")&" price rise on the "&(12-MonthsClosed)'
        '&" open months, or "&TEXT(breakeven_volume,"0%")&" more volume: price is the realistic lever.",'
        f'"Full year on the "&ScenarioName&" scenario: EBITDA "&MONEY(PnL!P{p["ebitda"]})&", "'
        '&MONEY(-breakeven_gap)&" above budget.")',
        '="Cash: DSO "&TEXT(WorkingCapital!O12,"0")&" days, DIO "&TEXT(WorkingCapital!O13,"0")'
        '&", DPO "&TEXT(WorkingCapital!O14,"0")&", a "&TEXT(WorkingCapital!O15,"0.0")&"-day cycle. "'
        '&WorkingCapital!B46',
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

    # ---- chart data, below the printed page
    put(ws, "B68", "Chart data: the charts above read these cells, which read the model.")
    font(ws.Range("B68"), italic=True, color=MUTED)
    row(ws, "B71", ["Month", "Actual", "Forecast", "Budget", "Prior year"])
    rv, bv, pyv = (f"PnL!$D${r}:$O${r}" for r in (p["revenue"], pos["pnl"]["budget"]["revenue"],
                                                   pos["pnl"]["py"]["revenue"]))
    fx(ws, "B72", '=TRANSPOSE(TEXT(pnl_months,"mmm"))')
    fx(ws, "C72", f'=TRANSPOSE(IF(pnl_months<=AsOfMonth,{rv},""))')
    fx(ws, "D72", f'=TRANSPOSE(IF(pnl_months>AsOfMonth,{rv},""))')
    fx(ws, "E72", f"=TRANSPOSE({bv})")
    fx(ws, "F72", f"=TRANSPOSE({pyv})")
    er1, er2 = pv["ebitda_bridge"]
    fx(ws, "B86:C93", f"=PVM!E{er1}")
    fx(ws, "C86:C93", f"=PVM!F{er1}")
    row(ws, "B96", ["Category", "Volume", "Mix", "Price", "Unit cost"])
    fx(ws, "B97:F102", f"=PVM!B{c1}")
    row(ws, "B105", ["Scenario", "FY EBITDA"])
    put(ws, "B106", "Budget")
    fx(ws, "C106", f"=PnL!Q{p['ebitda']}")
    fx(ws, "B107:B109", "=Scenarios!B18")
    fx(ws, "C107:C109", "=Scenarios!E18")
    numfmt(ws.Range("C72:F83,C97:F102,C106:C109"), F_MONEY)
    numfmt(ws.Range("C86:C93"), '$#,##0,"K";-$#,##0,"K"')

    # ---- charts
    for r in range(17, 54):
        ws.Rows(r).RowHeight = 15
    ch = new_chart(ws, XL_COLUMN_STACKED, "B17", "J34",
                   [(n, f"{c}72:{c}83") for n, c in zip(("Actual", "Forecast", "Budget", "Prior year"), "CDEF")],
                   "B72:B83", "Revenue by month: actual, forecast, budget and prior year")
    colours = [NAVY, "#9DB4CF", TEAL, GREY]
    for i, colour in enumerate(colours, start=1):
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
    ch.Axes(2).TickLabels.NumberFormat = '$#,##0.0,,"M"'
    ch.HasLegend = True
    ch.Legend.Position = -4107
    style_chart(ch)

    ws.Activate()
    ws.Range("B86:C93").Select()
    wf = ws.Shapes.AddChart2(-1, XL_WATERFALL, *place(ws, "K17", "S34")).Chart
    series = wf.FullSeriesCollection(1)
    series.Points(1).IsTotal = True
    series.Points(8).IsTotal = True
    wf.HasTitle = True
    wf.ChartTitle.Text = "EBITDA bridge, budget to actual, year to date"
    wf.HasLegend = False
    series.HasDataLabels = True     # waterfall labels take their format from the source cells
    try:
        style_chart(wf)
    except Exception:  # noqa: BLE001 - a waterfall exposes only part of the classic chart model
        pass
    ws.Range("B1").Select()

    pc = new_chart(ws, XL_BAR_STACKED, "B36", "J53",
                   [(n, f"{c}97:{c}102") for n, c in zip(("Volume", "Mix", "Price", "Unit cost"), "CDEF")],
                   "B97:B102", "Gross-margin variance by category, year to date")
    for i, colour in enumerate([TEAL, AMBER, NAVY, UNFAV], start=1):
        pc.FullSeriesCollection(i).Format.Fill.ForeColor.RGB = rgb(colour)
    pc.Axes(1).ReversePlotOrder = True
    pc.Axes(1).TickLabelPosition = -4134          # category names clear of the negative bars
    pc.Axes(2).TickLabels.NumberFormat = '$#,##0,"K";-$#,##0,"K"'
    pc.ChartGroups(1).GapWidth = 50
    pc.HasLegend = True
    pc.Legend.Position = -4107
    style_chart(pc)

    sc = new_chart(ws, XL_COLUMN_CLUSTERED, "K36", "S53", [("FY EBITDA", "C106:C109")], "B106:B109",
                   "Full-year EBITDA by scenario, against budget")
    sc.HasLegend = False
    s = sc.FullSeriesCollection(1)
    for i, colour in enumerate([GREY, NAVY, TEAL, UNFAV], start=1):
        s.Points(i).Format.Fill.ForeColor.RGB = rgb(colour)
    s.HasDataLabels = True
    s.DataLabels().NumberFormat = '$#,##0.00,,"M"'
    sc.Axes(2).TickLabels.NumberFormat = '$#,##0.0,,"M"'
    sc.ChartGroups(1).GapWidth = 70
    style_chart(sc)

    # ---- bottom tables
    ws.Rows(54).RowHeight = 8
    band(ws, 55, "Largest gross-margin movers against budget, year to date", "B", "J")
    band(ws, 55, "Full year by scenario", "K", "S")
    row(ws, "B56", ["Segment"])
    put(ws, "G56", "Main driver")
    put(ws, "J56", "GM vs budget")
    ws.Range("J56").HorizontalAlignment = -4152
    seg = f"PVM!$B${y1}:$B${y2}&\" · \"&PVM!$C${y1}:$C${y2}"
    v = f"PVM!$U${y1}:$U${y2}"
    fx(ws, "B57:B61", f"=LET(seg,{seg},v,{v},INDEX(SORTBY(seg,ABS(v),-1),ROWS(B$57:B57)))")
    fx(ws, "G57:G61", f"=LET(d,PVM!$W${y1}:$W${y2},v,{v},INDEX(SORTBY(d,ABS(v),-1),ROWS(G$57:G57)))")
    fx(ws, "J57:J61", f"=LET(v,{v},INDEX(SORTBY(v,ABS(v),-1),ROWS(J$57:J57)))")
    numfmt(ws.Range("J57:J61"), F_VAR)
    sign_colours(ws.Range("J57:J61"))
    for c, label in zip("KMOQS", ["Scenario", "Revenue", "EBITDA", "Margin", "vs budget"]):
        put(ws, f"{c}56", label)
        if c != "K":
            ws.Range(f"{c}56").HorizontalAlignment = -4152
    fx(ws, "K57:K59", "=Scenarios!B18")
    for c, src in zip("MOQS", "CEFG"):
        fx(ws, f"{c}57:{c}59", f"=Scenarios!{src}18")
    put(ws, "K60", "Budget")
    for c, src in zip("MOQ", "CEF"):
        fx(ws, f"{c}60", f"=Scenarios!{src}21")
    for c, fmt in zip("MOQS", (F_MONEY_M, F_MONEY_M, F_PCT, '+$#,##0,"K";($#,##0,"K");"–"')):
        numfmt(ws.Range(f"{c}57:{c}60"), fmt)
    sign_colours(ws.Range("S57:S59"))
    hi = ws.Range("K57:S59").FormatConditions.Add(XL_EXPRESSION, Formula1="=$K57=ScenarioName")
    hi.Font.Bold = True
    hi.Interior.Color = rgb("#EAF3FB")
    for rng in ("B56:J56", "K56:S56"):
        font(ws.Range(rng), bold=True, color=INK, size=9)
        ws.Range(rng).Borders(9).Color = rgb(NAVY)
    font(ws.Range("B57:S60"), size=10)
    font(ws.Range("K60:S60"), color=MUTED)
    names = {"kpi_revenue_ytd": "$B$5", "kpi_gm_pct_ytd": "$E$5", "kpi_ebitda_ytd": "$H$5",
             "kpi_ebitda_fy": "$K$5", "kpi_breakeven": "$N$5", "kpi_ccc": "$Q$5",
             "dash_commentary": "$B$11:$B$15", "dash_movers": "$B$57:$J$61",
             "dash_scenarios": "$K$57:$S$60"}
    for name, ref in names.items():
        wb.Names.Add(Name=name, RefersTo=f"=Dashboard!{ref}")
    put(ws, "B62", "Kestrel Bay Provisions is fictional and every figure is synthetic. Built in Excel "
                   "by Kush Patel: github.com/KushPatel29/excel-fpa-model")
    font(ws.Range("B62"), size=8, color=MUTED, italic=True)


SHEET_GUIDE = [
    ("Dashboard", "One page for the board: KPIs, bridges, scenarios and a written summary that "
                  "rewrites itself.", "Formula-driven narrative, waterfall and combo charts, sparklines, "
                                      "LET, SORTBY, named LAMBDAs"),
    ("PnL", "Monthly P&L: outlook (actuals + forecast), budget, prior year, variances, YTD.",
     "SUMIFS over tables, FAVVAR / SAFEDIV LAMBDAs, conditional formatting, sparklines"),
    ("PVM", "Price-volume-mix on 24 segments; gross-margin and EBITDA bridges.",
     "Exact four-effect decomposition, XMATCH driver tags, waterfall"),
    ("Forecast", "8+4 driver forecast by category, budget accuracy, statistical cross-check.",
     "XLOOKUP, WAPE and bias, FORECAST.ETS"),
    ("Scenarios", "Base / upside / downside, price × volume grid, tornado, gap to budget.",
     "What-if data tables (1- and 2-variable), data validation, SORTBY + HSTACK"),
    ("WorkingCapital", "DSO, DIO, DPO and the cash cycle; stock against shelf life.",
     "DAYSOF LAMBDA, trailing windows, data bars, risk flags"),
    ("Channels", "Channel economics by formula and by data model, reconciled.",
     "Power Pivot + DAX (TREATAS), PivotTable with slicers, CUBEVALUE"),
    ("Checks", "31 controls that prove the model ties out.", "Reconciliations with a tolerance"),
    ("Assumptions", "Every input and scenario driver in one place.", "Named inputs, validation"),
    ("PQ_Budget", "The budget, unpivoted and priced.", "Power Query (M): unpivot, merges, types"),
    ("Data_Sales", "Source tables, loaded unchanged from the CSVs.", "Excel tables, structured references"),
]


def cover(wb, pos: dict) -> None:
    ws = wb.Worksheets("Cover")
    p = pos["pnl"]["outlook"]
    yt = pos["pvm"]["ytd"][2]
    widths(ws, {"A": 3, "B": 36, "C": 70, "D": 58})
    put(ws, "B2", L.COMPANY)
    font(ws.Range("B2"), size=28, bold=True, color=NAVY, name=TITLE_FONT)
    ws.Rows(2).RowHeight = 40
    fx(ws, "B3", '="FY"&FiscalYear&" financial planning & analysis model"')
    font(ws.Range("B3"), size=15, color=INK, name=TITLE_FONT)
    put(ws, "B4", "Budget vs actual · price-volume-mix · rolling 8+4 forecast · scenarios · "
                  "working capital · channel economics")
    font(ws.Range("B4"), size=10, color=TEAL, bold=True)
    put(ws, "B6", "A fictional BC specialty-food distributor: 24 products, 4 channels, 3 regions, "
                  "32 months of synthetic data. Every figure is a live formula, Power Query step or DAX "
                  "measure over the source tables. Nothing is pasted in.")
    ws.Range("B6:D6").Merge()
    ws.Range("B6").WrapText = True
    ws.Rows(6).RowHeight = 30
    put(ws, "B7", "Built by Kush Patel · github.com/KushPatel29/excel-fpa-model")
    ws.Hyperlinks.Add(Anchor=ws.Range("B7"), Address="https://github.com/KushPatel29/excel-fpa-model",
                      TextToDisplay="Built by Kush Patel · github.com/KushPatel29/excel-fpa-model")
    font(ws.Range("B6:B7"), size=10, color=MUTED)

    put(ws, "B9", "Model status")
    font(ws.Range("B9"), bold=True)
    fx(ws, "C9", "=checks_summary")
    ws.Hyperlinks.Add(Anchor=ws.Range("D9"), Address="", SubAddress="'Checks'!A1",
                      TextToDisplay="Open the checks →")
    font(ws.Range("C9"), bold=True)
    for formula, colour in (('=COUNTIF(checks_status,"FAIL")=0', FAV), ('=COUNTIF(checks_status,"FAIL")>0', UNFAV)):
        c = ws.Range("C9").FormatConditions.Add(XL_EXPRESSION, Formula1=formula)
        c.Font.Color = rgb(colour)

    band(ws, 11, "Five questions it answers, live from the model", "B", "D")
    qa = [
        ("Are we on budget, and if not, why?",
         f'="YTD revenue "&TEXT(PnL!Z{p["revenue"]},"+0.0%;-0.0%")&" against budget, EBITDA "'
         f'&TEXT(PnL!Z{p["ebitda"]},"+0.0%;-0.0%")&". Biggest gross-margin drag: "&INDEX({{"volume","mix",'
         f'"price","unit cost"}},XMATCH(MIN(PVM!P{yt}:S{yt}),PVM!P{yt}:S{yt}))&"."', "PVM"),
        ("Where will the year land?",
         f'="EBITDA "&MONEY(PnL!P{p["ebitda"]})&" on the "&ScenarioName&" scenario, against a budget of "'
         f'&MONEY(PnL!Q{p["ebitda"]})&"."', "Forecast"),
        ("What would it take to hit budget?",
         '=IF(breakeven_gap>0,"A "&TEXT(breakeven_uplift,"0.0%")&" price rise on the open months, or "'
         '&TEXT(breakeven_volume,"0%")&" more volume.","Nothing: the outlook is at or above budget.")',
         "Scenarios"),
        ("Where is cash tied up?",
         '="A "&TEXT(WorkingCapital!O15,"0.0")&"-day cash cycle; "&MONEY(wc_at_risk_value)'
         '&" of stock sits above its shelf-life risk line."', "WorkingCapital"),
        ("Which channels earn the growth?",
         '=LET(c,Channels!B6:B9,g,Channels!M6:M9,j,Channels!J6:J9,k,XMATCH(MAX(g),g),INDEX(c,k)'
         '&" grew cases most ("&TEXT(MAX(g),"+0%")&" vs budget) at "&TEXT(INDEX(j,k),"$0.00")'
         '&" contribution a case; "&INDEX(c,XMATCH(MAX(j),j))&" earns most, "&TEXT(MAX(j),"$0.00")&".")',
         "Channels"),
    ]
    for i, (q, answer, sheet) in enumerate(qa):
        r = 12 + i
        put(ws, f"B{r}", q)
        fx(ws, f"C{r}", answer)
        ws.Hyperlinks.Add(Anchor=ws.Range(f"D{r}"), Address="", SubAddress=f"'{sheet}'!A1",
                          TextToDisplay=f"{sheet} →")
        font(ws.Range(f"B{r}"), bold=True)
        ws.Range(f"C{r}").WrapText = True
        ws.Range(f"B{r}:D{r}").VerticalAlignment = -4160
        ws.Rows(r).RowHeight = 28

    wb.Names.Add(Name="cover_answers", RefersTo="=Cover!$C$12:$C$16")
    band(ws, 18, "Sheets", "B", "D")
    header(ws, 19, ["Sheet", "What it shows", "Excel on show"], "B", wrap=False)
    ws.Range("C19:D19").HorizontalAlignment = -4131
    for i, (sheet, what, how) in enumerate(SHEET_GUIDE):
        r = 20 + i
        ws.Hyperlinks.Add(Anchor=ws.Range(f"B{r}"), Address="", SubAddress=f"'{sheet}'!A1",
                          TextToDisplay=sheet)
        put(ws, f"C{r}", what)
        put(ws, f"D{r}", how)
    last = 20 + len(SHEET_GUIDE) - 1
    font(ws.Range(f"D20:D{last}"), color=MUTED)
    ws.Range(f"C20:D{last}").WrapText = True

    h = last + 2
    band(ws, h, "How to use it", "B", "D")
    steps = [
        ("1  Pick a scenario", "Scenarios!C4. Every sheet, chart and sentence follows."),
        ("2  Test your own view", "Price and volume overlays in Scenarios!C5:C6, on top of the scenario."),
        ("3  Re-cut from an earlier close", "Change the as-of month on Assumptions, then Data > Refresh All."),
        ("4  Explore the data model", "The slicers on Channels filter a PivotTable built on DAX measures."),
    ]
    for i, (step, how) in enumerate(steps):
        put(ws, f"B{h + 1 + i}", step)
        put(ws, f"C{h + 1 + i}", how)
    font(ws.Range(f"B{h + 1}:B{h + len(steps)}"), bold=True)

    k = h + len(steps) + 2
    band(ws, k, "Conventions", "B", "D")
    put(ws, f"B{k + 1}", 0.015)
    numfmt(ws.Range(f"B{k + 1}"), "+0.0%")
    input_cell(ws.Range(f"B{k + 1}"))
    put(ws, f"C{k + 1}", "Input: blue on yellow. Change these; everything else is calculated.")
    put(ws, f"B{k + 2}", "1,234")
    put(ws, f"C{k + 2}", "Calculated: black. Links from another sheet are green on Scenarios.")
    put(ws, f"B{k + 3}", "Forecast month")
    fill(ws.Range(f"B{k + 3}"), "#EAF3FB")
    put(ws, f"C{k + 3}", "Blue shading marks forecast months; unshaded months are actuals.")
    put(ws, f"B{k + 4}", "+1,234   (1,234)")
    put(ws, f"C{k + 4}", "Variances are signed so favourable is positive (green) and unfavourable "
                         "is in brackets (red), for revenue and cost lines alike.")
    ws.Range(f"B{k + 1}:B{k + 4}").HorizontalAlignment = -4108
    font(ws.Range(f"C{k + 1}:C{k + 4}"), color=MUTED)
    pos["cover_last"] = k + 4


def finish(xl, wb, pos: dict) -> None:
    """Gridlines, zoom, frozen panes, print setup, and every sheet opened at A1."""
    freeze = {"PnL": "D6", "PVM": "D3", "Forecast": "C3", "WorkingCapital": "C3",
              "Data_Sales": "A5", "Data_BudgetUnits": "A5", "Data_BudgetRates": "A5",
              "Data_Opex": "A5", "Data_Balances": "A5", "PQ_Budget": "A5"}
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

    pvm_end = pos["pvm"]["bridge_check"]
    fc = pos["fc"]
    areas = {
        "Cover": (f"$A$1:$D${pos['cover_last']}", True),
        "Dashboard": ("$A$1:$T$62", True),
        "PnL": ("$B$1:$AB$71", False),
        "PVM": (f"$B$1:$W${pvm_end}", False),
        "Forecast": (f"$B$1:$P${fc['ets'] + 5}", False),
        "Scenarios": ("$B$1:$Q$49", True),
        "WorkingCapital": ("$B$1:$Z$47", False),
        "Channels": ("$B$1:$P$38", False),
        "Checks": (f"$B$1:$H${pos['checks'][1] + 8}", True),
        "Assumptions": ("$B$1:$G$23", True),
    }
    xl.PrintCommunication = False
    for name, (area, one_page) in areas.items():
        ps = wb.Worksheets(name).PageSetup
        ps.PrintArea = area
        ps.Orientation = XL_LANDSCAPE
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1 if one_page else 2 if name == "Forecast" else False
        ps.CenterHorizontally = True
        ps.LeftMargin = ps.RightMargin = xl.InchesToPoints(0.35)
        ps.TopMargin = xl.InchesToPoints(0.45)
        ps.BottomMargin = xl.InchesToPoints(0.5)
        ps.FooterMargin = xl.InchesToPoints(0.2)
        ps.LeftFooter = "&8Kestrel Bay Provisions (fictional) · FP&&A model"
        ps.CenterFooter = "&8&A"
        ps.RightFooter = "&8Page &P of &N"
        if name == "PnL":
            ps.PrintTitleColumns = "$B:$B"
    xl.PrintCommunication = True
    wb.Worksheets("Cover").Activate()


def export_pdf(wb, path) -> None:
    order = ["Cover", "Dashboard", "PnL", "PVM", "Forecast", "Scenarios", "WorkingCapital",
             "Channels", "Checks"]
    wb.Worksheets(order).Select()
    wb.ActiveSheet.ExportAsFixedFormat(XL_TYPE_PDF, str(path), 0, True, False)
    wb.Worksheets("Cover").Select()


__all__ = ["dashboard", "cover", "finish", "export_pdf"]
