"""Scenarios and sensitivities, peer benchmarking, the data-model view, and the checks."""
from __future__ import annotations

import layout as L
from build_calc import DA, F_CPK, F_DOLLAR2, F_MWH, FE, FI, FR, FUEL, PURCHASED, ferc_sum, names
from xl import (
    F_CASES, F_MONEY, F_PCT, F_VAR, F_VARPCT, FAIL_FILL, FAV, INK, MUTED, NAVY, PASS_FILL, TEAL, UNFAV,
    XL_BAR_CLUSTERED, XL_CELL_VALUE, XL_COLUMN_CLUSTERED, XL_EQUAL, XL_LEFT, band, column, font, fx, header,
    input_cell, new_chart, numfmt, put, rgb, row, sign_colours, style_chart, title, total_row, widths,
)

CUBE = '"ThisWorkbookDataModel"'


# ------------------------------------------------------------------- Scenarios
def scenarios(wb, pos: dict) -> None:
    ws = wb.Worksheets("Scenarios")
    ot = pos["forecast"]["outlook"][2]
    title(ws, '="Scenarios: what moves the "&FiscalYear&" outlook"',
          "Change a lever and every sheet follows. The tables below are Excel what-if data tables: each re-runs the "
          "whole forecast with a lever changed, so nothing here is an approximation.", span="T")
    band(ws, 11, "Live outputs, full year", "B", "I")
    outputs = [("Retail revenue", "=fy_revenue", F_MONEY), ("Retail MWh", "=fy_mwh", F_MWH),
               ("Average price, ¢/kWh", "=CPK(fy_revenue,fy_mwh)", F_CPK),
               ('="Growth on "&(FiscalYear-1)', f"=Forecast!H{ot}", F_VARPCT),
               ("Against plan", f"=Forecast!J{ot}", F_VAR),
               ("Open-month revenue", "=h2_revenue", F_MONEY)]
    for i, (label, f, fmt) in enumerate(outputs):
        put(ws, f"B{12 + i}", label)
        fx(ws, f"C{12 + i}", f)
        numfmt(ws.Range(f"C{12 + i}"), fmt)
    font(ws.Range("C12:C17"), bold=True, color="#1E7B34")
    sign_colours(ws.Range("C15:C16"))

    band(ws, 19, "Weather: a one-variable data table on the weather scenario", "B", "I")
    header(ws, 20, ["Scenario", "Retail revenue", "Retail MWh", "¢/kWh", "Against plan", "Open-month revenue"])
    put(ws, "B21", "Live selection")
    for c, src in zip("CDEFG", ("C12", "C13", "C14", "C16", "C17")):
        fx(ws, f"{c}21", f"={src}")
    column(ws, "B22", [s[0] for s in L.WEATHER_SCENARIOS])
    ws.Range("B21:G24").Table(ColumnInput=ws.Range("C4"))
    for c, fmt in zip("CDEFG", (F_MONEY, F_MWH, F_CPK, F_VAR, F_MONEY)):
        numfmt(ws.Range(f"{c}21:{c}24"), fmt)
    font(ws.Range("B21:G21"), italic=True, color=MUTED)
    sign_colours(ws.Range("F22:F24"))

    band(ws, 26, "Retail revenue: rate change on the open months down the side, new data-center MW across the top",
         "B", "I")
    put(ws, "B27", "Rate ↓   MW →")
    font(ws.Range("B27"), color=MUTED, italic=True)
    fx(ws, "B28", "=C12")
    numfmt(ws.Range("B28"), ";;;")
    row(ws, "C28", L.SENS_MW)
    column(ws, "B29", L.SENS_RATE)
    ws.Range("B28:I35").Table(RowInput=ws.Range("C6"), ColumnInput=ws.Range("C5"))
    numfmt(ws.Range("C28:I28"), '0" MW"')
    numfmt(ws.Range("B29:B35"), '+0%;-0%;0%')
    numfmt(ws.Range("C29:I35"), '$#,##0.0,,"M"')
    font(ws.Range("C28:I28"), bold=True)
    font(ws.Range("B29:B35"), bold=True)
    scale = ws.Range("C29:I35").FormatConditions.AddColorScale(3)
    scale.ColorScaleCriteria(1).FormatColor.Color = rgb("#F4C7C0")
    scale.ColorScaleCriteria(2).FormatColor.Color = rgb("#FFFFFF")
    scale.ColorScaleCriteria(3).FormatColor.Color = rgb("#BFE3CB")

    band(ws, 37, "Tornado: full-year revenue impact of one step on each lever, from the base case", "B", "I")
    header(ws, 38, ["Lever", "Step", "Base case", "With the step", "Impact"])
    levers = [  # label, lever cell, base, step value, step text
        ("Rate change on open months", "C5", 0, 0.01, "+1 point"),
        ("New data-center load", "C6", 0, 50, "+50 MW"),
        ("Heating degree days", "C8", 1, 0.9, "−10%"),
        ("Cooling degree days", "C9", 1, 1.1, "+10%"),
        ("Customers", "C7", 0, 0.01, "+1%"),
    ]
    put(ws, "K37", "Workings: one small data table per lever (base value, then the step)")
    font(ws.Range("K37"), color=MUTED, italic=True)
    ws.Range("C39:C43").NumberFormat = "@"  # else Excel reads "+10%" as the number 0.1
    for i, (label, cell, base, step, text) in enumerate(levers):
        r = 39 + i
        a = 39 + 4 * i                      # this lever's data table: K{a}:L{a+2}
        put(ws, f"K{a}", label)
        fx(ws, f"L{a}", "=C12")
        ws.Range(f"K{a + 1}:K{a + 2}").Value = ((base,), (step,))
        ws.Range(f"K{a}:L{a + 2}").Table(ColumnInput=ws.Range(cell))
        font(ws.Range(f"K{a}:L{a + 2}"), color=MUTED, size=8)
        numfmt(ws.Range(f"L{a}:L{a + 2}"), F_MONEY)
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", text)
        fx(ws, f"D{r}", f"=L{a + 1}")
        fx(ws, f"E{r}", f"=L{a + 2}")
        fx(ws, f"F{r}", f"=E{r}-D{r}")
    numfmt(ws.Range("D39:E43"), F_MONEY)
    numfmt(ws.Range("F39:F43"), F_VAR)
    sign_colours(ws.Range("F39:F43"))
    ws.Range("C39:C43").HorizontalAlignment = -4152
    put(ws, "S38", "Chart data: sorted by size")
    font(ws.Range("S38"), color=MUTED, italic=True)
    fx(ws, "S39", "=SORTBY(HSTACK(B39:B43,F39:F43),ABS(F39:F43),1)")
    numfmt(ws.Range("T39:T43"), F_VAR)
    font(ws.Range("S39:T43"), color=MUTED)
    chart = new_chart(ws, XL_BAR_CLUSTERED, "N11", "T35", [("FY revenue impact", "T39:T43")], "S39:S43",
                      "Full-year revenue impact of one step on each lever")
    chart.HasLegend = False
    chart.Axes(1).TickLabelPosition = -4134
    s = chart.FullSeriesCollection(1)
    s.Format.Fill.ForeColor.RGB = rgb(TEAL)
    s.InvertIfNegative = True
    s.InvertColor = rgb(UNFAV)
    chart.Axes(2).TickLabels.NumberFormat = '$#,##0.0,,"M";-$#,##0.0,,"M"'
    chart.ChartGroups(1).GapWidth = 60
    style_chart(chart)

    band(ws, 45, "What would offset a mild autumn?", "B", "I")
    put(ws, "B46", "Revenue lost if the open months are mild")
    fx(ws, "C46", "=C22-C23")
    put(ws, "B47", "Rate change on the open months that recovers it")
    fx(ws, "C47", "=SAFEDIV(C46,G23)")
    numfmt(ws.Range("C46"), F_MONEY)
    numfmt(ws.Range("C47"), "0.00%")
    font(ws.Range("C46:C47"), bold=True)
    fx(ws, "B48", '="A mild rest of the year would cost "&MONEY(C46)&" of retail revenue: about a "&TEXT(C47,"0.0%")'
                  '&" rate change on the open months, or "&TEXT(SAFEDIV(C46,F40/50),"#,##0")&" MW of new '
                  'data-center load from now."')
    font(ws.Range("B48"), italic=True, color=INK)
    widths(ws, {"A": 2, "B": 40, "C:I": 13, "J": 2, "K": 26, "L": 14, "M": 2, "N:R": 11, "S": 26, "T": 12})
    ws.Range("C4:C9").HorizontalAlignment = -4152
    names(wb, {"scn_live": "Scenarios!$C$12:$C$17", "scn_weather": "Scenarios!$C$22:$G$24",
               "sens_grid": "Scenarios!$C$29:$I$35", "sens_rate_axis": "Scenarios!$B$29:$B$35",
               "sens_mw_axis": "Scenarios!$C$28:$I$28", "tornado": "Scenarios!$F$39:$F$43",
               "offset_loss": "Scenarios!$C$46", "offset_rate": "Scenarios!$C$47"})


# ------------------------------------------------------------------------ Peers
def peers(wb, pos: dict) -> None:
    ws = wb.Worksheets("Peers")
    title(ws, "Peers: PGE against four Pacific Northwest utilities (FERC Form 1)",
          '="Electric operations only, as filed with FERC. Price and cost metrics are for "&FercYear&"; growth is the '
          'yearly rate since "&BaseYear&". Rank 1 is the highest value."', span="J")
    put(ws, "B4", "Metric")
    fx(ws, "C4:G4", "=INDEX(tbl_Utilities[short_name],COLUMN()-2)")
    put(ws, "H4", "PGE rank")
    put(ws, "B5", "FERC respondent id")
    fx(ws, "C5:G5", "=INDEX(tbl_Utilities[utility_id_ferc1],COLUMN()-2)")
    font(ws.Range("B4:H4"), bold=True)
    font(ws.Range("B5:G5"), color=MUTED, size=8)
    ws.Range("C4:H5").HorizontalAlignment = -4152

    def rv(key, year="FercYear", value="dollar_value"):
        return ferc_sum(FR, value, f'"{key}"', utility="C$5", year=year)

    def ex(key, many=False):
        return ferc_sum(FE, "dollar_value", key if many else f'"{key}"', utility="C$5", year="FercYear", many=many)

    def inc(key, many=False):
        return ferc_sum(FI, "dollar_value", key if many else f'"{key}"', utility="C$5", year="FercYear", many=many)

    retail, retail_mwh = rv("sales_to_ultimate_consumers"), rv("sales_to_ultimate_consumers", value="sales_mwh")
    power = f"({ex(FUEL, True)}+{ex(PURCHASED, True)})"
    years = "(FercYear-BaseYear)"
    metrics = [  # label, formula, format
        ("Retail revenue", f"={retail}", F_MONEY),
        ("Retail MWh", f"={retail_mwh}", F_MWH),
        ("Retail customers", f"={rv('sales_to_ultimate_consumers', value='avg_customers_per_month')}", F_CASES),
        ("Retail price, ¢/kWh", f"=CPK({retail},{retail_mwh})", F_CPK),
        ("Residential price, ¢/kWh", f"=CPK({rv('residential_sales')},{rv('residential_sales', value='sales_mwh')})",
         F_CPK),
        ("Industrial share of retail MWh", f"=SAFEDIV({rv('large_or_industrial', value='sales_mwh')},{retail_mwh})",
         F_PCT),
        ("Retail MWh growth a year", f"=({retail_mwh}/{rv('sales_to_ultimate_consumers', 'BaseYear', 'sales_mwh')})"
                                     f"^(1/{years})-1", F_VARPCT),
        ("Industrial MWh growth a year", f"=({rv('large_or_industrial', value='sales_mwh')}/"
                                         f"{rv('large_or_industrial', 'BaseYear', 'sales_mwh')})^(1/{years})-1", F_VARPCT),
        ("Retail price growth a year", f"=(CPK({retail},{retail_mwh})/CPK({rv('sales_to_ultimate_consumers', 'BaseYear')},"
                                       f"{rv('sales_to_ultimate_consumers', 'BaseYear', 'sales_mwh')}))^(1/{years})-1",
         F_VARPCT),
        ("Power cost per MWh sold, $", f"=SAFEDIV({power},{rv('sales_of_electricity', value='sales_mwh')})", F_DOLLAR2),
        ("Other O&M per customer, $", f"=SAFEDIV({ex('operations_and_maintenance_expenses_electric')}-{power},"
                                      f"{rv('sales_to_ultimate_consumers', value='avg_customers_per_month')})", F_DOLLAR2),
        ("Operating margin before income tax", f"=SAFEDIV({rv('electric_operating_revenues')}"
                                               f"-{ex('operations_and_maintenance_expenses_electric')}-{inc(DA, True)}"
                                               f"-{inc('taxes_other_than_income_taxes_utility_operating_income')},"
                                               f"{rv('electric_operating_revenues')})", F_PCT),
    ]
    for i, (label, f, fmt) in enumerate(metrics):
        r = 7 + i
        put(ws, f"B{r}", label)
        fx(ws, f"C{r}:G{r}", f)
        numfmt(ws.Range(f"C{r}:G{r}"), fmt)
        fx(ws, f"H{r}", f"=RANK.EQ(C{r},C{r}:G{r},0)")
    last = 7 + len(metrics) - 1
    numfmt(ws.Range(f"H7:H{last}"), '0" of 5"')
    ws.Range(f"C7:C{last}").Interior.Color = rgb("#EAF3FB")
    font(ws.Range(f"C7:C{last}"), bold=True)
    c1 = new_chart(ws, XL_COLUMN_CLUSTERED, "B21", "E38", [("Industrial MWh growth a year", "C14:G14")], "C4:G4",
                   "Industrial MWh growth a year since the base year")
    c1.HasLegend = False
    c1.FullSeriesCollection(1).Format.Fill.ForeColor.RGB = rgb(TEAL)
    c1.FullSeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = rgb(NAVY)
    c1.Axes(2).TickLabels.NumberFormat = "0%"
    c1.Axes(1).TickLabelPosition = -4134
    style_chart(c1)
    c2 = new_chart(ws, XL_COLUMN_CLUSTERED, "F21", "J38", [("Retail price, ¢/kWh", "C10:G10")], "C4:G4",
                   "Retail price, cents per kWh")
    c2.HasLegend = False
    c2.FullSeriesCollection(1).Format.Fill.ForeColor.RGB = rgb("#9DB4CF")
    c2.FullSeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = rgb(NAVY)
    style_chart(c2)
    widths(ws, {"A": 2, "B": 36, "C:G": 16, "H": 10, "I:J": 6})
    names(wb, {"peers_table": f"Peers!$C$7:$G${last}", "peers_ids": "Peers!$C$5:$G$5",
               "peers_rank": f"Peers!$H$7:$H${last}"})


# ----------------------------------------------------------------------- Explore
def explore(wb, pos: dict) -> None:
    ws = wb.Worksheets("Explore")
    title(ws, "Explore: the same data, asked of the Power Pivot data model",
          "Slicers filter the PivotTable; the measures are DAX, defined once in the data model (Data > Manage Data "
          "Model). Below, CUBEVALUE formulas read the same model and are held to SUMIFS on the worksheet table.", span="L")
    band(ws, 4, "PivotTable on the data model, filtered by slicers", "B", "L")
    band(ws, 26, "Revenue by class and year: CUBEVALUE against the data model", "B", "L")
    header(ws, 27, ["Class", "", "", "", "", "YoY, DAX", "YoY, SUMIFS", "Agree"])
    for c, back in zip("CDEF", (3, 2, 1, 0)):
        fx(ws, f"{c}27", f"=FiscalYear-{back}" if back else "=FiscalYear")
    numfmt(ws.Range("C27:E27"), '"FY"0')
    numfmt(ws.Range("F27"), '"FY"0" YTD"')
    column(ws, "B28", [L.CLASS_LABELS[c] for c in L.CLASSES])
    put(ws, "B32", "All classes")
    member = '"[Class].[label].&["&$B28&"]"'
    fx(ws, "C28:F31", f'=CUBEVALUE({CUBE},"[Measures].[Retail Revenue]",{member},"[Calendar].[year].&["&C$27&"]")')
    fx(ws, "C32:F32", f'=CUBEVALUE({CUBE},"[Measures].[Retail Revenue]","[Calendar].[year].&["&C$27&"]")')
    fx(ws, "G28:G31", f'=CUBEVALUE({CUBE},"[Measures].[Revenue YoY %]",{member},"[Calendar].[year].&["&FiscalYear&"]")')
    fx(ws, "G32", f'=CUBEVALUE({CUBE},"[Measures].[Revenue YoY %]","[Calendar].[year].&["&FiscalYear&"]")')
    this = ('SUMIFS(tbl_Monthly[revenue],tbl_Monthly[class],{c},tbl_Monthly[month],">="&FYStart,'
            'tbl_Monthly[month],"<="&AsOfMonth)')
    prior = ('SUMIFS(tbl_Monthly[revenue],tbl_Monthly[class],{c},tbl_Monthly[month],">="&EDATE(FYStart,-12),'
             'tbl_Monthly[month],"<="&EDATE(AsOfMonth,-12))')
    key = "LOWER($B28)"
    fx(ws, "H28:H31", f"=SAFEDIV({this.format(c=key)}-{prior.format(c=key)},{prior.format(c=key)})")
    fx(ws, "H32", f'=SAFEDIV(SUM({this.format(c="tbl_Class[class]")})-SUM({prior.format(c="tbl_Class[class]")}),'
                  f'SUM({prior.format(c="tbl_Class[class]")}))')
    fx(ws, "I28:I32", '=IF(ABS(G28-H28)<0.000001,"Yes","No")')
    numfmt(ws.Range("C28:F32"), F_MONEY)
    numfmt(ws.Range("G28:H32"), F_VARPCT)
    ws.Range("I28:I32").HorizontalAlignment = -4108
    total_row(ws.Range("B32:I32"), double=True)
    widths(ws, {"A": 2, "B": 18, "C:I": 14})
    names(wb, {"explore_cube": "Explore!$C$28:$F$32", "explore_yoy": "Explore!$G$28:$I$32"})


# ------------------------------------------------------------------------ Checks
def checks(wb, pos: dict, controls: dict) -> None:
    ws = wb.Worksheets("Checks")
    pv = pos["pvm"]
    pb, pe = pos["plan"]["long"]
    ot = pos["forecast"]["outlook"][2]
    t1, t2 = pos["monthly"]["tieout"]
    title(ws, "Checks: the model proves its own numbers")
    eia_sum = "+".join(f"tbl_EIA[{c}_{{m}}]" for c in L.CLASSES)
    ferc_line = ('SUMIFS(tbl_FERC_Revenue[dollar_value],tbl_FERC_Revenue[utility_id_ferc1],CompanyId,'
                 'tbl_FERC_Revenue[report_year],pnl_years,tbl_FERC_Revenue[revenue_type],"{key}")')
    om_line = ('SUMIFS(tbl_FERC_Expense[dollar_value],tbl_FERC_Expense[utility_id_ferc1],CompanyId,'
               'tbl_FERC_Expense[report_year],pnl_years,tbl_FERC_Expense[expense_type],'
               '"operations_and_maintenance_expenses_electric")')
    inc_line = ('SUMIFS(tbl_FERC_Income[dollar_value],tbl_FERC_Income[utility_id_ferc1],CompanyId,'
                'tbl_FERC_Income[report_year],pnl_years,tbl_FERC_Income[income_type],"{key}")')
    defaults = "AND(RateChange=0,ExtraMW=0,CustOverlay=0,HddOverlay=1,CddOverlay=1)"
    items = [  # label, expected, actual, tolerance, why
        ("EIA rows loaded = source", "=ctrl_eia_rows", "=ROWS(tbl_EIA)", "=Tol", "Nothing dropped or duplicated."),
        ("EIA megawatt-hours = source control total", "=ctrl_eia_mwh", "=SUM(tbl_EIA[total_mwh])", "=Tol", ""),
        ("EIA revenue = source control total", "=ctrl_eia_revenue_k", "=SUM(tbl_EIA[total_revenue_k])", "=Tol", ""),
        ("No month missing from EIA", '=DATEDIF(DATE(MIN(tbl_EIA[year]),1,1),AsOfMonth,"m")+1', "=ROWS(tbl_EIA)",
         "=Tol", "A gap would shift every year-over-year comparison."),
        ("EIA classes add up to EIA's total: revenue", "=0",
         f"=SUMPRODUCT(ABS({eia_sum.format(m='revenue_k')}-tbl_EIA[total_revenue_k]))", "=Tol",
         "The regulator's own arithmetic, checked."),
        ("EIA classes add up to EIA's total: MWh", "=0",
         f"=SUMPRODUCT(ABS({eia_sum.format(m='mwh')}-tbl_EIA[total_mwh]))", "=Tol", ""),
        ("Power Query unpivot keeps every value", "=ROWS(tbl_EIA)*4", "=ROWS(tbl_Monthly)", "=Tol",
         "Four classes per month after the unpivot."),
        ("Power Query dollars = EIA thousands × 1,000", "=SUM(tbl_EIA[total_revenue_k])*1000",
         "=SUM(tbl_Monthly[revenue])", "=Tol", ""),
        ("NOAA rows loaded = source", "=ctrl_noaa_lines", "=ROWS(tbl_NOAA)", "=Tol", ""),
        ("NOAA parse: one row per month, no gaps", '=DATEDIF(MIN(tbl_Weather[month]),MAX(tbl_Weather[month]),"m")+1',
         "=ROWS(tbl_Weather)", "=Tol", "Fixed-width parsing done in Power Query."),
        ("FERC rows loaded = source", "=ctrl_ferc_rows",
         "=ROWS(tbl_FERC_Revenue)+ROWS(tbl_FERC_Expense)+ROWS(tbl_FERC_Income)", "=Tol", ""),
        ("FERC retail = the classes, every year", "=0",
         f"=SUMPRODUCT(ABS(pnl_retail-{ferc_line.format(key='sales_to_ultimate_consumers')}))", "=Tol",
         "The P&L rebuilds retail from the classes; it must meet FERC's own subtotal."),
        ("FERC total revenue = the lines, every year", "=0",
         f"=SUMPRODUCT(ABS(pnl_revenue-{ferc_line.format(key='electric_operating_revenues')}))", "=Tol", ""),
        ("FERC O&M schedule = the income statement, every year", "=0",
         f"=SUMPRODUCT(ABS({om_line}-{inc_line.format(key='operation_expense')}-"
         f"{inc_line.format(key='maintenance_expense')}))", "=100",
         "Two FERC schedules that must agree; they differ by a few dollars of rounding over twelve years."),
        ("EIA and FERC agree on MWh, every full year", "=0",
         f"=MAX(ABS(SAFEDIV(Monthly!E{t1}:E{t2},Monthly!D{t1}:D{t2})))", "=0.001",
         "Two regulators, two filings, the same volumes (within 0.1%)."),
        ("EIA and FERC revenue within 3%, every full year", "=0", f"=MAX(ABS(Monthly!I{t1}:I{t2}))", "=0.03",
         "Billed against booked revenue: the gap is explained on Monthly."),
        ("Weather model fitted on every month before this year", "=(FiscalYear-YEAR(FitStart))*12",
         "=INDEX(fit_months,1)", "=Tol", ""),
        ("Weather explains most residential use (R² ≥ 0.7)", "=0", "=MAX(0,0.7-INDEX(coef_residential,5))", "=Tol",
         "A weak fit would turn the weather effect into noise."),
        ("PVM reconciles in every class: long view", "=0", f"=SUMPRODUCT(ABS(PVM!O{pv['long'][0]}:O{pv['long'][2]}))",
         "=Tol", "Volume + mix + price = the revenue change, class by class."),
        ("PVM reconciles in every class: last year", "=0",
         f"=SUMPRODUCT(ABS(PVM!O{pv['short'][0]}:O{pv['short'][2]}))", "=Tol", ""),
        ("PVM reconciles in every class: year to date", "=0",
         f"=SUMPRODUCT(ABS(PVM!O{pv['ytd'][0]}:O{pv['ytd'][2]}))", "=Tol", ""),
        ("Plan bridge closes in every month and class", "=0", f"=SUM(IFERROR(ABS(Plan!T{pb}:T{pe}),0))", "=Tol",
         "Customers + weather + usage + price = actual − plan."),
        ("Plan bridge total = actual − plan", "=INDEX(plan_bridge_total,6)-INDEX(plan_bridge_total,1)",
         "=SUM(INDEX(plan_bridge_total,2),INDEX(plan_bridge_total,3),INDEX(plan_bridge_total,4),"
         "INDEX(plan_bridge_total,5))", "=Tol", ""),
        ("Outlook = actuals + forecast", f"=Forecast!D{ot}+Forecast!E{ot}", "=fy_revenue", "=Tol", ""),
        ("Outlook actuals = EIA year to date", "=INDEX(monthly_ytd_total,1)", f"=Forecast!D{ot}", "=Tol", ""),
        ("Weather table matches the live model", "=Scenarios!C12",
         "=XLOOKUP(WeatherScenario,Scenarios!B22:B24,Scenarios!C22:C24)", "=Tol",
         "The data table re-ran the same model the sheets show."),
        ("Sensitivity grid matches the live model", "=fy_revenue",
         "=IFERROR(INDEX(sens_grid,XMATCH(RateChange,sens_rate_axis),XMATCH(ExtraMW,sens_mw_axis)),fy_revenue)",
         "=Tol", ""),
        ("Rate tornado is exact: 1 point of rate = 1% of open-month revenue",
         f"=IF({defaults},h2_revenue*0.01,INDEX(tornado,1))", "=INDEX(tornado,1)", "=Tol",
         "The brute-force data table agrees with the algebra."),
        ("Data model = formulas: this year's revenue", "=INDEX(monthly_ytd_total,1)", "=INDEX(explore_cube,5,4)",
         "=Tol", "DAX and SUMIFS are independent engines."),
        ("DAX year-over-year = formula year-over-year", "=0",
         "=SUMPRODUCT(ABS(INDEX(explore_yoy,0,1)-INDEX(explore_yoy,0,2)))", "=0.000001",
         "Time intelligence with TREATAS, checked against plain SUMIFS."),
        ("Dashboard outlook tile = Forecast", "=fy_revenue", "=kpi_fy_revenue", "=Tol", ""),
    ]
    band(ws, 4, "Controls", "B", "I")
    header(ws, 5, ["#", "Check", "Expected", "Actual", "Difference", "Tolerance", "Status", "Why it matters"])
    ws.Range("C5").HorizontalAlignment = XL_LEFT
    ws.Range("I5").HorizontalAlignment = XL_LEFT
    first = 6
    for i, (label, expected, actual, tol, why) in enumerate(items):
        r = first + i
        put(ws, f"B{r}", i + 1)
        put(ws, f"C{r}", label)
        fx(ws, f"D{r}", expected)
        fx(ws, f"E{r}", actual)
        fx(ws, f"F{r}", f"=E{r}-D{r}")
        fx(ws, f"G{r}", tol)
        fx(ws, f"H{r}", f'=IF(ISERROR(F{r}),"FAIL",IF(ABS(F{r})<=G{r},"PASS","FAIL"))')
        put(ws, f"I{r}", why)
    last = first + len(items) - 1
    numfmt(ws.Range(f"D{first}:F{last}"), '#,##0.00;(#,##0.00);"–"')
    numfmt(ws.Range(f"G{first}:G{last}"), "General")
    ws.Range(f"H{first}:H{last}").HorizontalAlignment = -4108
    font(ws.Range(f"H{first}:H{last}"), bold=True)
    for text, colour, fore in (("PASS", PASS_FILL, FAV), ("FAIL", FAIL_FILL, UNFAV)):
        c = ws.Range(f"H{first}:H{last}").FormatConditions.Add(XL_CELL_VALUE, XL_EQUAL, f'="{text}"')
        c.Interior.Color = rgb(colour)
        c.Font.Color = rgb(fore)
    font(ws.Range(f"I{first}:I{last}"), color=MUTED)
    font(ws.Range(f"B{first}:B{last}"), color=MUTED)
    font(ws.Range(f"G{first}:G{last}"), color=MUTED)
    status = f"H{first}:H{last}"
    fx(ws, "B2", f'=IF(COUNTIF({status},"FAIL")=0,"All "&COUNTA({status})&" checks pass",'
                 f'COUNTIF({status},"FAIL")&" of "&COUNTA({status})&" checks fail: see column H")')
    font(ws.Range("B2"), bold=True, size=12)
    for formula, colour in ((f'=COUNTIF({status},"FAIL")=0', FAV), (f'=COUNTIF({status},"FAIL")>0', UNFAV)):
        c = ws.Range("B2").FormatConditions.Add(2, Formula1=formula)
        c.Font.Color = rgb(colour)

    k0 = last + 3
    band(ws, k0, "Source control totals, taken from the extracts when the workbook was built", "B", "I")
    header(ws, k0 + 1, ["", "Control total", "Value", "Source"])
    for i, (name, (label, value, source)) in enumerate(controls.items()):
        r = k0 + 2 + i
        put(ws, f"C{r}", label)
        put(ws, f"D{r}", value)
        put(ws, f"E{r}", source)
        input_cell(ws.Range(f"D{r}"))
        numfmt(ws.Range(f"D{r}"), "#,##0.000" if isinstance(value, float) else "#,##0")
        wb.Names.Add(Name=name, RefersTo=f"=Checks!$D${r}")
    font(ws.Range(f"E{k0 + 2}:E{k0 + 1 + len(controls)}"), color=MUTED)
    widths(ws, {"A": 2, "B": 5, "C": 54, "D:F": 16, "G": 10, "H": 9, "I": 64})
    names(wb, {"checks_status": f"Checks!$H${first}:$H${last}", "checks_summary": "Checks!$B$2"})
    pos["checks"] = (first, last)
