"""The plan, the variance bridge, the backtest and the rolling forecast.

The plan for a year is built only from data through the December before it, so
it can be judged honestly against what happened; a what-if data table re-runs
it as of four year-ends to show how well the method has done.
"""
from __future__ import annotations

import layout as L
from build_calc import F_CPK, F_DOLLAR2, F_MWH, names
from xl import (
    F_CASES, F_MONEY, F_MONTH, F_VAR, F_VARPCT, FORECAST_FILL, MUTED, NAVY, XL_EDGE_BOTTOM, XL_EXPRESSION,
    XL_THIN, XL_WATERFALL, band, column, edge, font, fx, header, input_cell, numfmt, place, put, rgb, sign_colours,
    style_chart, title, total_row, widths,
)

CLASS_ORDER = L.CLASSES
WEATHER_FLAG = 'XLOOKUP($C{r},tbl_Class[class],tbl_Class[weather_model])="yes"'


def month_sum(field: str, cls: str, month: str) -> str:
    return f"SUMIFS(tbl_Monthly[{field}],tbl_Monthly[class],{cls},tbl_Monthly[month],{month})"


def range_sum(field: str, cls: str, first: str, last: str) -> str:
    return (f'SUMIFS(tbl_Monthly[{field}],tbl_Monthly[class],{cls},tbl_Monthly[month],">="&{first},'
            f'tbl_Monthly[month],"<="&{last})')


def class_rows(ws, r0: int) -> tuple[int, int]:
    column(ws, f"B{r0}", [L.CLASS_LABELS[c] for c in CLASS_ORDER])
    column(ws, f"C{r0}", CLASS_ORDER)
    font(ws.Range(f"C{r0}:C{r0 + 3}"), color=MUTED, size=8)
    return r0, r0 + 3


# ------------------------------------------------------------------- Plan
def plan(wb, pos: dict) -> None:
    ws = wb.Worksheets("Plan")
    hist = pos["weather"]["history"]
    title(ws, '="Plan against actual: the "&PlanYear&" plan, built only from data through December "&(PlanYear-1)',
          "Customers grow at last year's rate; residential and commercial use comes from the weather model at normal "
          "weather; industrial volume and every class's price grow at last year's rate. The bridge splits actual minus "
          "plan into customers, weather, usage and price, exactly.", span="U")
    put(ws, "B4", "Plan year")
    put(ws, "C4", 2026)
    input_cell(ws.Range("C4"))
    numfmt(ws.Range("C4"), "0")
    wb.Names.Add(Name="PlanYear", RefersTo="=Plan!$C$4")
    put(ws, "B5", "Months closed")
    fx(ws, "C5", "=IF(PlanYear<FiscalYear,12,IF(PlanYear=FiscalYear,MonthsClosed,0))")
    wb.Names.Add(Name="PlanClosed", RefersTo="=Plan!$C$5")
    put(ws, "D4", "Change it to re-run the plan as of an earlier year-end; the backtest below does exactly that.")
    font(ws.Range("B4:B5"), bold=True)
    font(ws.Range("D4"), color=MUTED)

    # --- drivers
    band(ws, 7, '="Drivers, from "&(PlanYear-2)&" and "&(PlanYear-1)', "B", "L")
    header(ws, 8, ["Class", "Key", "Customers, Dec before", "Dec two before", "Customer growth", "MWh, year before",
                   "Two before", "Volume growth", "$/MWh, year before", "Two before", "Price growth"])
    d1, d2 = class_rows(ws, 9)
    y1 = "DATE(PlanYear-1,1,1)", "DATE(PlanYear-1,12,1)"
    y2 = "DATE(PlanYear-2,1,1)", "DATE(PlanYear-2,12,1)"
    fx(ws, f"D{d1}:D{d2}", "=" + month_sum("customers", f"$C{d1}", "DATE(PlanYear-1,12,1)"))
    fx(ws, f"E{d1}:E{d2}", "=" + month_sum("customers", f"$C{d1}", "DATE(PlanYear-2,12,1)"))
    fx(ws, f"F{d1}:F{d2}", f"=SAFEDIV(D{d1},E{d1})")
    fx(ws, f"G{d1}:G{d2}", "=" + range_sum("mwh", f"$C{d1}", *y1))
    fx(ws, f"H{d1}:H{d2}", "=" + range_sum("mwh", f"$C{d1}", *y2))
    fx(ws, f"I{d1}:I{d2}", f"=SAFEDIV(G{d1},H{d1})")
    fx(ws, f"J{d1}:J{d2}", f"=SAFEDIV({range_sum('revenue', f'$C{d1}', *y1)},G{d1})")
    fx(ws, f"K{d1}:K{d2}", f"=SAFEDIV({range_sum('revenue', f'$C{d1}', *y2)},H{d1})")
    fx(ws, f"L{d1}:L{d2}", f"=SAFEDIV(J{d1},K{d1})")
    numfmt(ws.Range(f"D{d1}:E{d2}"), F_CASES)
    numfmt(ws.Range(f"G{d1}:H{d2}"), F_MWH)
    numfmt(ws.Range(f"J{d1}:K{d2}"), F_DOLLAR2)
    for c in "FIL":
        numfmt(ws.Range(f"{c}{d1}:{c}{d2}"), '0.0000"x"')

    # --- weather model for the plan
    hb, he = hist
    band(ws, 14, '="Weather model for the plan, fitted to Dec "&(PlanYear-1)', "B", "L")
    header(ws, 15, ["Class", "Key", "Intercept", "Per HDD", "Per CDD", "Trend per year", "R²", "Months"])
    window = (f"(Weather!$B${hb}:$B${he}<=AsOfMonth)*(YEAR(Weather!$B${hb}:$B${he})<PlanYear)"
              f"*(Weather!$B${hb}:$B${he}>=FitStart)")

    def filt(c):
        return f"FILTER(Weather!${c}${hb}:${c}${he},{window})"

    for i, (cls, ycol) in enumerate((("residential", "F"), ("commercial", "G"))):
        r = 16 + i
        put(ws, f"B{r}", L.CLASS_LABELS[cls])
        put(ws, f"C{r}", cls)
        fx(ws, f"D{r}", f"=LET(s,LINEST({filt(ycol)},HSTACK({filt('D')},{filt('E')},{filt('C')}),TRUE,TRUE),"
                        "HSTACK(INDEX(s,1,4),INDEX(s,1,3),INDEX(s,1,2),INDEX(s,1,1),INDEX(s,3,1)))")
        fx(ws, f"I{r}", f"=ROWS({filt(ycol)})")
    numfmt(ws.Range("D16:G17"), "0.000000")
    numfmt(ws.Range("H16:H17"), "0.000")
    font(ws.Range("C16:C17"), color=MUTED, size=8)

    # --- normal and actual weather for the plan year
    band(ws, 19, '="Normal weather for "&PlanYear&" ("&(PlanYear-NormalYears)&"–"&(PlanYear-1)&") and what happened"',
         "B", "O")
    put(ws, "B20", "Month")
    fx(ws, "C20:N20", '=TEXT(DATE(2000,COLUMN()-2,1),"mmm")')
    for r, label in ((21, "Normal HDD"), (22, "Normal CDD"), (23, "Actual HDD"), (24, "Actual CDD")):
        put(ws, f"B{r}", label)
    for r, field in ((21, "hdd"), (22, "cdd")):
        fx(ws, f"C{r}:N{r}", f"=AVERAGE(FILTER(tbl_Weather[{field}],(YEAR(tbl_Weather[month])>=PlanYear-NormalYears)"
                             f"*(YEAR(tbl_Weather[month])<PlanYear)*(MONTH(tbl_Weather[month])=COLUMN()-2)))")
    for r, field in ((23, "hdd"), (24, "cdd")):
        fx(ws, f"C{r}:N{r}", f'=XLOOKUP(DATE(PlanYear,COLUMN()-2,1),tbl_Weather[month],tbl_Weather[{field}],"")')
    numfmt(ws.Range("C21:N24"), "0.0")
    font(ws.Range("B20:N20"), bold=True)
    ws.Range("C20:N20").HorizontalAlignment = -4152

    # --- month by class
    t0 = 26
    band(ws, t0, "Plan and actual, month by class", "B", "U")
    header(ws, t0 + 1, ["Month", "Class", "Has actuals", "Plan customers", "Model MWh per customer", "Plan MWh",
                        "Plan MWh per customer", "Plan $/MWh", "Plan revenue", "Actual customers", "Actual MWh",
                        "Actual revenue", "Actual $/MWh", "Weather effect per customer", "Customers", "Weather",
                        "Usage", "Price", "Check"])
    t1 = t0 + 2
    t2 = t1 + 48 - 1
    months = [(k, cls) for cls in CLASS_ORDER for k in range(1, 13)]
    ws.Range(f"B{t1}:C{t2}").Value = tuple((None, cls) for _, cls in months)
    fx(ws, f"B{t1}:B{t2}", f"=DATE(PlanYear,MOD(ROWS($B${t1}:B{t1})-1,12)+1,1)")
    r = t1
    wf = WEATHER_FLAG.format(r=r)
    k = f"MONTH($B{r})"
    driver = f"XLOOKUP($C{r},$C${d1}:$C${d2},{{col}})"
    coef = f"XLOOKUP($C{r},$C$16:$C$17,{{col}})"
    fx(ws, f"D{t1}:D{t2}", f"=MONTH($B{r})<=PlanClosed")
    fx(ws, f"E{t1}:E{t2}", f"={driver.format(col=f'$D${d1}:$D${d2}')}*{driver.format(col=f'$F${d1}:$F${d2}')}^({k}/12)")
    fx(ws, f"F{t1}:F{t2}", f'=IF({wf},{coef.format(col="$D$16:$D$17")}+{coef.format(col="$E$16:$E$17")}'
                           f'*INDEX($C$21:$N$21,{k})+{coef.format(col="$F$16:$F$17")}*INDEX($C$22:$N$22,{k})'
                           f'+{coef.format(col="$G$16:$G$17")}*TRENDYEARS($B{r}),"")')
    prior_mwh = month_sum("mwh", f"$C{r}", f"EDATE($B{r},-12)")
    prior_rev = month_sum("revenue", f"$C{r}", f"EDATE($B{r},-12)")
    fx(ws, f"G{t1}:G{t2}", f"=IF({wf},E{r}*F{r},{prior_mwh}*{driver.format(col=f'$I${d1}:$I${d2}')})")
    fx(ws, f"H{t1}:H{t2}", f"=SAFEDIV(G{r},E{r})")
    fx(ws, f"I{t1}:I{t2}", f"=SAFEDIV({prior_rev},{prior_mwh})*{driver.format(col=f'$L${d1}:$L${d2}')}")
    fx(ws, f"J{t1}:J{t2}", f"=G{r}*I{r}")
    fx(ws, f"K{t1}:K{t2}", f'=IF($D{r},{month_sum("customers", f"$C{r}", f"$B{r}")},"")')
    fx(ws, f"L{t1}:L{t2}", f'=IF($D{r},{month_sum("mwh", f"$C{r}", f"$B{r}")},"")')
    fx(ws, f"M{t1}:M{t2}", f'=IF($D{r},{month_sum("revenue", f"$C{r}", f"$B{r}")},"")')
    fx(ws, f"N{t1}:N{t2}", f'=IF($D{r},SAFEDIV(M{r},L{r}),"")')
    fx(ws, f"O{t1}:O{t2}", f'=IF(AND($D{r},{wf}),{coef.format(col="$E$16:$E$17")}*(INDEX($C$23:$N$23,{k})'
                           f'-INDEX($C$21:$N$21,{k}))+{coef.format(col="$F$16:$F$17")}*(INDEX($C$24:$N$24,{k})'
                           f'-INDEX($C$22:$N$22,{k})),IF($D{r},0,""))')
    fx(ws, f"P{t1}:P{t2}", f'=IF($D{r},IF({wf},(K{r}-E{r})*H{r}*I{r},0),"")')
    fx(ws, f"Q{t1}:Q{t2}", f'=IF($D{r},IF({wf},K{r}*O{r}*I{r},0),"")')
    fx(ws, f"R{t1}:R{t2}", f'=IF($D{r},IF({wf},K{r}*(SAFEDIV(L{r},K{r})-H{r}-O{r})*I{r},(L{r}-G{r})*I{r}),"")')
    fx(ws, f"S{t1}:S{t2}", f'=IF($D{r},L{r}*(N{r}-I{r}),"")')
    fx(ws, f"T{t1}:T{t2}", f'=IF($D{r},P{r}+Q{r}+R{r}+S{r}-(M{r}-J{r}),"")')
    numfmt(ws.Range(f"B{t1}:B{t2}"), F_MONTH)
    numfmt(ws.Range(f"E{t1}:E{t2}"), F_CASES)
    numfmt(ws.Range(f"K{t1}:K{t2}"), F_CASES)
    numfmt(ws.Range(f"F{t1}:F{t2}"), "0.0000")
    numfmt(ws.Range(f"H{t1}:H{t2}"), "0.0000")
    numfmt(ws.Range(f"O{t1}:O{t2}"), "+0.0000;-0.0000;0")
    numfmt(ws.Range(f"G{t1}:G{t2}"), F_MWH)
    numfmt(ws.Range(f"L{t1}:L{t2}"), F_MWH)
    numfmt(ws.Range(f"I{t1}:I{t2}"), F_DOLLAR2)
    numfmt(ws.Range(f"N{t1}:N{t2}"), F_DOLLAR2)
    numfmt(ws.Range(f"J{t1}:J{t2}"), F_MONEY)
    numfmt(ws.Range(f"M{t1}:M{t2}"), F_MONEY)
    numfmt(ws.Range(f"P{t1}:S{t2}"), F_VAR)
    numfmt(ws.Range(f"T{t1}:T{t2}"), '[>=0.5]+#,##0;[<=-0.5](#,##0);"–"')
    sign_colours(ws.Range(f"P{t1}:S{t2}"))
    font(ws.Range(f"T{t1}:T{t2}"), color=MUTED)
    c = ws.Range(f"B{t1}:T{t2}").FormatConditions.Add(XL_EXPRESSION, Formula1=f"=NOT($D{t1})")
    c.Interior.Color = rgb(FORECAST_FILL)
    for i in range(1, 4):
        edge(ws.Range(f"B{t1 + 12 * i - 1}:T{t1 + 12 * i - 1}"), XL_EDGE_BOTTOM, XL_THIN, NAVY)

    # --- bridge by class
    b0 = t2 + 3
    band(ws, b0, '="Plan to actual, "&TEXT(DATE(PlanYear,1,1),"mmm")&"–"&TEXT(DATE(PlanYear,MAX(PlanClosed,1),1),'
                 '"mmm yyyy")', "B", "L")
    header(ws, b0 + 1, ["Class", "Key", "Plan", "Customers", "Weather", "Usage", "Price", "Actual", "Actual vs plan",
                        "%", "Check"])
    b1, b2 = class_rows(ws, b0 + 2)
    bt = b2 + 1
    put(ws, f"B{bt}", "All classes")
    closed = f"$C${t1}:$C${t2},$C{b1},$D${t1}:$D${t2},TRUE"
    for col, src in zip("DEFGHI", "JPQRSM"):
        fx(ws, f"{col}{b1}:{col}{b2}", f"=SUMIFS(${src}${t1}:${src}${t2},{closed})")
    fx(ws, f"D{bt}:I{bt}", f"=SUM(D{b1}:D{b2})")
    fx(ws, f"J{b1}:J{bt}", f"=I{b1}-D{b1}")
    fx(ws, f"K{b1}:K{bt}", f"=SAFEDIV(J{b1},D{b1})")
    fx(ws, f"L{b1}:L{bt}", f"=SUM(E{b1}:H{b1})-J{b1}")
    numfmt(ws.Range(f"D{b1}:D{bt}"), F_MONEY)
    numfmt(ws.Range(f"I{b1}:I{bt}"), F_MONEY)
    numfmt(ws.Range(f"E{b1}:H{bt}"), F_VAR)
    numfmt(ws.Range(f"J{b1}:J{bt}"), F_VAR)
    numfmt(ws.Range(f"K{b1}:K{bt}"), F_VARPCT)
    numfmt(ws.Range(f"L{b1}:L{bt}"), '[>=0.5]+#,##0;[<=-0.5](#,##0);"–"')
    sign_colours(ws.Range(f"E{b1}:K{bt}"))
    total_row(ws.Range(f"B{bt}:L{bt}"), double=True)
    font(ws.Range(f"L{b1}:L{bt}"), color=MUTED)

    # --- accuracy by month (all classes) and the bridge chart's data
    a0 = bt + 3
    band(ws, a0, "Accuracy: monthly totals, all classes", "B", "L")
    header(ws, a0 + 1, ["Month", "Has actuals", "Plan revenue", "Actual revenue", "Absolute miss", "Plan MWh",
                        "Actual MWh", "Absolute miss"])
    a1 = a0 + 2
    a2 = a1 + 11
    fx(ws, f"B{a1}:B{a2}", f"=DATE(PlanYear,ROWS($B${a1}:B{a1}),1)")
    fx(ws, f"C{a1}:C{a2}", f"=MONTH(B{a1})<=PlanClosed")
    fx(ws, f"D{a1}:D{a2}", f"=SUMIFS($J${t1}:$J${t2},$B${t1}:$B${t2},B{a1})")
    fx(ws, f"E{a1}:E{a2}", f'=IF(C{a1},SUMIFS($M${t1}:$M${t2},$B${t1}:$B${t2},B{a1}),"")')
    fx(ws, f"F{a1}:F{a2}", f'=IF(C{a1},ABS(E{a1}-D{a1}),"")')
    fx(ws, f"G{a1}:G{a2}", f"=SUMIFS($G${t1}:$G${t2},$B${t1}:$B${t2},B{a1})")
    fx(ws, f"H{a1}:H{a2}", f'=IF(C{a1},SUMIFS($L${t1}:$L${t2},$B${t1}:$B${t2},B{a1}),"")')
    fx(ws, f"I{a1}:I{a2}", f'=IF(C{a1},ABS(H{a1}-G{a1}),"")')
    numfmt(ws.Range(f"B{a1}:B{a2}"), F_MONTH)
    numfmt(ws.Range(f"D{a1}:F{a2}"), F_MONEY)
    numfmt(ws.Range(f"G{a1}:I{a2}"), F_MWH)
    s0 = a2 + 2
    stats = [("Months with actuals", "=PlanClosed", "0"),
             ("Plan revenue, months with actuals", f"=SUMIFS(D{a1}:D{a2},C{a1}:C{a2},TRUE)", F_MONEY),
             ("Actual revenue", f"=SUM(E{a1}:E{a2})", F_MONEY),
             ("Actual against plan", f"=SAFEDIV(C{s0 + 2}-C{s0 + 1},C{s0 + 1})", F_VARPCT),
             ("Monthly revenue WAPE", f"=SAFEDIV(SUM(F{a1}:F{a2}),SUM(E{a1}:E{a2}))", "0.0%"),
             ("Monthly MWh WAPE", f"=SAFEDIV(SUM(I{a1}:I{a2}),SUM(H{a1}:H{a2}))", "0.0%")]
    for i, (label, f, fmt) in enumerate(stats):
        put(ws, f"B{s0 + i}", label)
        fx(ws, f"C{s0 + i}", f)
        numfmt(ws.Range(f"C{s0 + i}"), fmt)
    font(ws.Range(f"C{s0}:C{s0 + 5}"), bold=True)

    # --- backtest: the whole plan re-run as of four year-ends
    k0 = s0 + 8
    band(ws, k0, "Backtest: the same method, re-run as of each year-end (a what-if data table on Plan year)", "B", "L")
    header(ws, k0 + 1, ["Plan year", "Months", "Plan revenue", "Actual revenue", "Actual vs plan", "Revenue WAPE",
                        "MWh WAPE"])
    k1 = k0 + 2
    put(ws, f"B{k1}", "Live")
    for c, i in zip("CDEFGH", range(6)):
        fx(ws, f"{c}{k1}", f"=C{s0 + i}")
    column(ws, f"B{k1 + 1}", L.BACKTEST_YEARS)
    k2 = k1 + len(L.BACKTEST_YEARS)
    ws.Range(f"B{k1}:H{k2}").Table(ColumnInput=ws.Range("C4"))
    numfmt(ws.Range(f"C{k1}:C{k2}"), "0")
    numfmt(ws.Range(f"D{k1}:E{k2}"), F_MONEY)
    numfmt(ws.Range(f"F{k1}:F{k2}"), F_VARPCT)
    numfmt(ws.Range(f"G{k1}:H{k2}"), "0.0%")
    sign_colours(ws.Range(f"F{k1 + 1}:F{k2}"))
    font(ws.Range(f"B{k1}:H{k1}"), italic=True, color=MUTED)
    put(ws, f"B{k2 + 1}", "Volume is forecastable: MWh misses by about 4% a month. Price is not, from history: PGE's "
                          "rates move in steps set by rate cases, so a plan that extrapolates last year's increase "
                          "overshoots when a big step does not repeat. Plan price from the rate case calendar.")
    font(ws.Range(f"B{k2 + 1}"), italic=True, color=MUTED)

    # --- bridge chart (plan -> effects -> actual)
    c0 = b0 + 1
    labels = [('="Plan "&PlanYear'), "Customers", "Weather", "Usage", "Price", "Actual"]
    for i, lab in enumerate(labels):
        put(ws, f"N{c0 + i}", lab)
    for i, src in enumerate(["D", "E", "F", "G", "H", "I"]):
        fx(ws, f"O{c0 + i}", f"={src}{bt}")
    numfmt(ws.Range(f"O{c0}:O{c0 + 5}"), '$#,##0.0,,"M";-$#,##0.0,,"M"')
    font(ws.Range(f"N{c0}:O{c0 + 5}"), color=MUTED, size=8)
    ws.Activate()
    ws.Range(f"N{c0}:O{c0 + 5}").Select()
    wfc = ws.Shapes.AddChart2(-1, XL_WATERFALL, *place(ws, f"N{a0}", f"U{a0 + 18}")).Chart
    s = wfc.FullSeriesCollection(1)
    s.Points(1).IsTotal = True
    s.Points(6).IsTotal = True
    s.HasDataLabels = True
    wfc.HasTitle = True
    wfc.ChartTitle.Text = "Plan to actual, year to date"
    wfc.HasLegend = False
    style_chart(wfc)
    ws.Range("A1").Select()

    widths(ws, {"A": 2, "B": 14, "C": 13, "D:T": 12, "U": 2})
    ws.Rows(t0 + 1).RowHeight = 42
    pos["plan"] = {"drivers": (d1, d2), "long": (t1, t2), "bridge": (b1, b2, bt), "accuracy": (a1, a2, s0),
                   "backtest": (k1, k2), "chart": (c0, c0 + 5)}
    names(wb, {"plan_drivers": f"Plan!$D${d1}:$L${d2}", "plan_coef": "Plan!$D$16:$I$17",
               "plan_normals": "Plan!$C$21:$N$22", "plan_long": f"Plan!$B${t1}:$T${t2}",
               "plan_bridge": f"Plan!$D${b1}:$L${b2}", "plan_bridge_total": f"Plan!$D${bt}:$L${bt}",
               "plan_stats": f"Plan!$C${s0}:$C${s0 + 5}", "backtest": f"Plan!$B${k1 + 1}:$H${k2}",
               "plan_chart": f"Plan!$N${c0}:$O${c0 + 5}"})


# --------------------------------------------------------------- Forecast
def forecast(wb, pos: dict) -> None:
    ws = wb.Worksheets("Forecast")
    title(ws, '="Rolling forecast: "&TEXT(AsOfMonth,"mmmm yyyy")&" actuals, "&(12-MonthsClosed)&" months forecast"',
          "Customers and industrial volume grow at the year-to-date rate; residential and commercial use comes from "
          "the weather model, with NOAA's observed degree days where the month has passed and normal (times the "
          "scenario) where it has not. Price grows at the year-to-date rate, plus any rate change on Scenarios.",
          span="Q")
    band(ws, 4, "Drivers, year to date", "B", "M")
    header(ws, 5, ["Class", "Key", "Customers, latest", "A year earlier", "Customer growth", "MWh YTD",
                   "Prior year YTD", "Volume growth", "$/MWh YTD", "Prior year YTD", "Price growth"])
    d1, d2 = class_rows(ws, 6)
    fx(ws, f"D{d1}:D{d2}", "=" + month_sum("customers", f"$C{d1}", "AsOfMonth"))
    fx(ws, f"E{d1}:E{d2}", "=" + month_sum("customers", f"$C{d1}", "EDATE(AsOfMonth,-12)"))
    fx(ws, f"F{d1}:F{d2}", f"=SAFEDIV(D{d1},E{d1})")
    fx(ws, f"G{d1}:G{d2}", "=" + range_sum("mwh", f"$C{d1}", "FYStart", "AsOfMonth"))
    fx(ws, f"H{d1}:H{d2}", "=" + range_sum("mwh", f"$C{d1}", "EDATE(FYStart,-12)", "EDATE(AsOfMonth,-12)"))
    fx(ws, f"I{d1}:I{d2}", f"=SAFEDIV(G{d1},H{d1})")
    fx(ws, f"J{d1}:J{d2}", f"=SAFEDIV({range_sum('revenue', f'$C{d1}', 'FYStart', 'AsOfMonth')},G{d1})")
    fx(ws, f"K{d1}:K{d2}", f"=SAFEDIV({range_sum('revenue', f'$C{d1}', 'EDATE(FYStart,-12)', 'EDATE(AsOfMonth,-12)')},"
                           f"H{d1})")
    fx(ws, f"L{d1}:L{d2}", f"=SAFEDIV(J{d1},K{d1})")
    numfmt(ws.Range(f"D{d1}:E{d2}"), F_CASES)
    numfmt(ws.Range(f"G{d1}:H{d2}"), F_MWH)
    numfmt(ws.Range(f"J{d1}:K{d2}"), F_DOLLAR2)
    for c in "FIL":
        numfmt(ws.Range(f"{c}{d1}:{c}{d2}"), '0.0000"x"')

    # --- weather used for each month
    band(ws, 11, "Degree days used: NOAA where observed, otherwise normal × scenario", "B", "O")
    put(ws, "B12", "Month")
    fx(ws, "C12:N12", "=DATE(FiscalYear,COLUMN()-2,1)")
    numfmt(ws.Range("C12:N12"), "mmm")
    put(ws, "B13", "HDD used")
    put(ws, "B14", "CDD used")
    put(ws, "B15", "Source")
    fx(ws, "C13:N13", '=XLOOKUP(C$12,tbl_Weather[month],tbl_Weather[hdd],INDEX(normal_hdd,COLUMN()-2)*HddScale)')
    fx(ws, "C14:N14", '=XLOOKUP(C$12,tbl_Weather[month],tbl_Weather[cdd],INDEX(normal_cdd,COLUMN()-2)*CddScale)')
    fx(ws, "C15:N15", '=IF(ISNUMBER(XMATCH(C$12,tbl_Weather[month])),"NOAA","normal")')
    numfmt(ws.Range("C13:N14"), "0.0")
    font(ws.Range("C15:N15"), color=MUTED, size=8)
    font(ws.Range("B12:N12"), bold=True)
    ws.Range("C12:N15").HorizontalAlignment = -4152

    # --- month by class
    t0 = 17
    band(ws, t0, "Forecast, month by class (open months only)", "B", "M")
    header(ws, t0 + 1, ["Month", "Class", "Open", "Customers", "HDD", "CDD", "Model MWh per customer", "MWh",
                        "$/MWh", "Revenue"])
    t1 = t0 + 2
    t2 = t1 + 47
    ws.Range(f"C{t1}:C{t2}").Value = tuple((cls,) for cls in CLASS_ORDER for _ in range(12))
    fx(ws, f"B{t1}:B{t2}", f"=DATE(FiscalYear,MOD(ROWS($B${t1}:B{t1})-1,12)+1,1)")
    r = t1
    wf = WEATHER_FLAG.format(r=r)
    drv = f"XLOOKUP($C{r},$C${d1}:$C${d2},{{col}})"
    fx(ws, f"D{t1}:D{t2}", f"=$B{r}>AsOfMonth")
    fx(ws, f"E{t1}:E{t2}", f'=IF($D{r},{drv.format(col=f"$D${d1}:$D${d2}")}*{drv.format(col=f"$F${d1}:$F${d2}")}'
                           f'^((MONTH($B{r})-MonthsClosed)/12)*(1+CustOverlay),"")')
    fx(ws, f"F{t1}:F{t2}", f'=IF($D{r},INDEX($C$13:$N$13,MONTH($B{r})),"")')
    fx(ws, f"G{t1}:G{t2}", f'=IF($D{r},INDEX($C$14:$N$14,MONTH($B{r})),"")')
    coef = "XLOOKUP($C{r},Weather!$I$6:$I$7,Weather!{col}$6:{col}$7)"
    fx(ws, f"H{t1}:H{t2}", f'=IF(AND($D{r},{wf}),{coef.format(r=r, col="$C")}+{coef.format(r=r, col="$D")}*F{r}'
                           f'+{coef.format(r=r, col="$E")}*G{r}+{coef.format(r=r, col="$F")}*TRENDYEARS($B{r}),"")')
    prior_mwh = month_sum("mwh", f"$C{r}", f"EDATE($B{r},-12)")
    prior_rev = month_sum("revenue", f"$C{r}", f"EDATE($B{r},-12)")
    fx(ws, f"I{t1}:I{t2}", f'=IF($D{r},IF({wf},E{r}*H{r},{prior_mwh}*{drv.format(col=f"$I${d1}:$I${d2}")}'
                           f'*(1+CustOverlay))+IF($C{r}="industrial",ExtraMW*24*DAY(EOMONTH($B{r},0)),0),"")')
    fx(ws, f"J{t1}:J{t2}", f'=IF($D{r},SAFEDIV({prior_rev},{prior_mwh})*{drv.format(col=f"$L${d1}:$L${d2}")}'
                           f'*(1+RateChange),"")')
    fx(ws, f"K{t1}:K{t2}", f'=IF($D{r},I{r}*J{r},"")')
    numfmt(ws.Range(f"B{t1}:B{t2}"), F_MONTH)
    numfmt(ws.Range(f"E{t1}:E{t2}"), F_CASES)
    numfmt(ws.Range(f"F{t1}:G{t2}"), "0.0")
    numfmt(ws.Range(f"H{t1}:H{t2}"), "0.0000")
    numfmt(ws.Range(f"I{t1}:I{t2}"), F_MWH)
    numfmt(ws.Range(f"J{t1}:J{t2}"), F_DOLLAR2)
    numfmt(ws.Range(f"K{t1}:K{t2}"), F_MONEY)
    c = ws.Range(f"B{t1}:K{t2}").FormatConditions.Add(XL_EXPRESSION, Formula1=f"=NOT($D{t1})")
    c.Font.Color = rgb("#B8C0CC")
    for i in range(1, 4):
        edge(ws.Range(f"B{t1 + 12 * i - 1}:K{t1 + 12 * i - 1}"), XL_EDGE_BOTTOM, XL_THIN, NAVY)

    # --- outlook by class
    ob = t2 + 3
    band(ws, ob, '="Full-year outlook, "&FiscalYear', "B", "M")
    header(ws, ob + 1, ["Class", "Key", "Actual YTD", "Forecast", "Full year", '=FiscalYear-1', "Growth",
                        '="Plan "&FiscalYear', "Against plan", "MWh, full year", "¢/kWh"])
    fx(ws, f"G{ob + 1}", '="Actual "&(FiscalYear-1)')
    fx(ws, f"I{ob + 1}", '="Plan "&FiscalYear')
    o1, o2 = class_rows(ws, ob + 2)
    ot = o2 + 1
    put(ws, f"B{ot}", "All classes")
    pl = pos["plan"]["long"]
    fx(ws, f"D{o1}:D{o2}", "=" + range_sum("revenue", f"$C{o1}", "FYStart", "AsOfMonth"))
    fx(ws, f"E{o1}:E{o2}", f"=SUMIFS($K${t1}:$K${t2},$C${t1}:$C${t2},$C{o1},$D${t1}:$D${t2},TRUE)")
    fx(ws, f"F{o1}:F{o2}", f"=D{o1}+E{o1}")
    fx(ws, f"G{o1}:G{o2}", "=" + range_sum("revenue", f"$C{o1}", "DATE(FiscalYear-1,1,1)", "DATE(FiscalYear-1,12,1)"))
    fx(ws, f"H{o1}:H{ot}", f"=SAFEDIV(F{o1}-G{o1},G{o1})")
    fx(ws, f"I{o1}:I{o2}", f'=IF(PlanYear=FiscalYear,SUMIFS(Plan!$J${pl[0]}:$J${pl[1]},Plan!$C${pl[0]}:$C${pl[1]},'
                           f'$C{o1}),NA())')
    fx(ws, f"J{o1}:J{ot}", f"=IFERROR(F{o1}-I{o1},0)")
    fx(ws, f"K{o1}:K{o2}", f"={range_sum('mwh', f'$C{o1}', 'FYStart', 'AsOfMonth')}+SUMIFS($I${t1}:$I${t2},"
                           f"$C${t1}:$C${t2},$C{o1},$D${t1}:$D${t2},TRUE)")
    fx(ws, f"L{o1}:L{ot}", f"=CPK(F{o1},K{o1})")
    for col in "DEFGIK":
        fx(ws, f"{col}{ot}", f"=SUM({col}{o1}:{col}{o2})")
    numfmt(ws.Range(f"D{o1}:G{ot}"), F_MONEY)
    numfmt(ws.Range(f"I{o1}:I{ot}"), F_MONEY)
    numfmt(ws.Range(f"H{o1}:H{ot}"), F_VARPCT)
    numfmt(ws.Range(f"J{o1}:J{ot}"), F_VAR)
    numfmt(ws.Range(f"K{o1}:K{ot}"), F_MWH)
    numfmt(ws.Range(f"L{o1}:L{ot}"), F_CPK)
    sign_colours(ws.Range(f"H{o1}:H{ot}"))
    sign_colours(ws.Range(f"J{o1}:J{ot}"))
    total_row(ws.Range(f"B{ot}:L{ot}"), double=True)

    # --- monthly totals: actual, forecast, plan, prior year, and the statistical cross-check
    m0 = ot + 3
    band(ws, m0, "Monthly totals, all classes, and a statistical cross-check (FORECAST.ETS)", "B", "M")
    header(ws, m0 + 1, ["Month", "Actual", "Forecast", "Plan", "Prior year", "FORECAST.ETS", "Driver vs ETS"])
    m1 = m0 + 2
    m2 = m1 + 11
    fx(ws, f"B{m1}:B{m2}", f"=DATE(FiscalYear,ROWS($B${m1}:B{m1}),1)")
    fx(ws, f"C{m1}:C{m2}", f'=IF(ISCLOSED(B{m1}),SUMIFS(tbl_Monthly[revenue],tbl_Monthly[month],B{m1}),"")')
    fx(ws, f"D{m1}:D{m2}", f'=IF(ISCLOSED(B{m1}),"",SUMIFS($K${t1}:$K${t2},$B${t1}:$B${t2},B{m1}))')
    fx(ws, f"E{m1}:E{m2}", f'=IF(PlanYear=FiscalYear,SUMIFS(Plan!$J${pl[0]}:$J${pl[1]},Plan!$B${pl[0]}:$B${pl[1]},'
                           f'B{m1}),"")')
    fx(ws, f"F{m1}:F{m2}", f"=SUMIFS(tbl_Monthly[revenue],tbl_Monthly[month],EDATE(B{m1},-12))")
    h0 = m2 + 4
    h1 = h0 + 2
    h2 = h1 + 12 * 10 - 1
    months_hist = f"$B${h1}:$B${h2}"
    closed_n = f"MATCH(AsOfMonth,{months_hist},0)"
    fx(ws, f"G{m1}:G{m2}", f'=IF(ISCLOSED(B{m1}),"",FORECAST.ETS(MATCH(B{m1},{months_hist},0),'
                           f"$D${h1}:INDEX($D${h1}:$D${h2},{closed_n}),$C${h1}:INDEX($C${h1}:$C${h2},{closed_n}),12))")
    fx(ws, f"H{m1}:H{m2}", f'=IF(ISCLOSED(B{m1}),"",SAFEDIV(D{m1}-G{m1},G{m1}))')
    mt = m2 + 1
    put(ws, f"B{mt}", "Full year")
    for col in "CDEFG":
        fx(ws, f"{col}{mt}", f"=SUM({col}{m1}:{col}{m2})")
    fx(ws, f"H{mt}", f"=SAFEDIV(D{mt}-G{mt},G{mt})")
    numfmt(ws.Range(f"B{m1}:B{m2}"), F_MONTH)
    numfmt(ws.Range(f"C{m1}:G{mt}"), F_MONEY)
    numfmt(ws.Range(f"H{m1}:H{mt}"), '+0.0%;-0.0%;0.0%')
    total_row(ws.Range(f"B{mt}:H{mt}"))
    band(ws, h0, "History feeding FORECAST.ETS: total retail revenue by month", "B", "M")
    put(ws, f"B{h0 + 1}", "Month")
    put(ws, f"C{h0 + 1}", "Index")
    put(ws, f"D{h0 + 1}", "Revenue")
    font(ws.Range(f"B{h0 + 1}:D{h0 + 1}"), bold=True)
    fx(ws, f"B{h1}", "=FitStart")
    fx(ws, f"B{h1 + 1}:B{h2}", f"=EDATE(B{h1},1)")
    fx(ws, f"C{h1}:C{h2}", f"=ROWS($B${h1}:B{h1})")
    fx(ws, f"D{h1}:D{h2}", f'=IF(ISCLOSED(B{h1}),SUMIFS(tbl_Monthly[revenue],tbl_Monthly[month],B{h1}),"")')
    numfmt(ws.Range(f"B{h1}:B{h2}"), F_MONTH)
    numfmt(ws.Range(f"D{h1}:D{h2}"), F_MONEY)

    widths(ws, {"A": 2, "B": 14, "C:N": 12, "O:Q": 12})
    ws.Rows(5).RowHeight = 30
    ws.Rows(t0 + 1).RowHeight = 30
    pos["forecast"] = {"drivers": (d1, d2), "long": (t1, t2), "outlook": (o1, o2, ot), "months": (m1, m2, mt),
                       "history": (h1, h2)}
    names(wb, {"fc_drivers": f"Forecast!$D${d1}:$L${d2}", "fc_degree_days": "Forecast!$C$13:$N$14",
               "fc_long": f"Forecast!$B${t1}:$K${t2}", "fc_outlook": f"Forecast!$D${o1}:$L${o2}",
               "fc_outlook_total": f"Forecast!$D${ot}:$L${ot}", "fc_months": f"Forecast!$B${m1}:$H${m2}",
               "fc_months_total": f"Forecast!$C${mt}:$H${mt}", "fy_revenue": f"Forecast!$F${ot}",
               "fy_mwh": f"Forecast!$K${ot}", "h2_revenue": f"Forecast!$E${ot}"})
