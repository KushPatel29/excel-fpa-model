"""Calculation sheets: the driver forecast, the P&L and the price-volume-mix.

Every figure is a formula over the source tables (tbl_Sales, tbl_Budget,
tbl_Opex) and the named inputs. Formulas are written once per row or block
and filled with relative references, the way an analyst would enter them.
"""
from __future__ import annotations

import layout as L
from xl import (
    F_CASES, F_MONEY, F_MONTH, F_MULT, F_PCT, F_PRICE, F_VAR, F_VARPCT, FORECAST_FILL,
    MUTED, NAVY, TEAL, XL_EDGE_BOTTOM, XL_EXPRESSION, XL_LEFT, XL_SPARKLINE_LINE,
    XL_THIN, band, column, edge, fill, font, fx, header, numfmt, put, rgb, row,
    sign_colours, title, total_row, widths,
)

S_YTD = 'tbl_Sales[month],">="&FYStart,tbl_Sales[month],"<="&AsOfMonth'
B_YTD = 'tbl_Budget[month],">="&FYStart,tbl_Budget[month],"<="&AsOfMonth'
S_TRAIL = 'tbl_Sales[month],">="&TrailStart,tbl_Sales[month],"<="&AsOfMonth'

LINE_KEYS = [k for k, *_ in L.PNL_LINES]
IDX = {k: i for i, k in enumerate(LINE_KEYS)}
IS_COST = {k: c for k, _, c, _ in L.PNL_LINES}
KIND = {k: kind for k, _, _, kind in L.PNL_LINES}
OPEX_KEYS = ["opex_sm", "opex_wl", "opex_ga", "opex_tech"]


def prow(block: str, key: str) -> int:
    """Worksheet row of a P&L line in a block."""
    return L.PNL_BLOCKS[block] + IDX[key]


def subhead(ws, r: int, text: str, c1: str = "B", c2: str = "P") -> None:
    put(ws, f"{c1}{r}", text)
    font(ws.Range(f"{c1}{r}"), bold=True, color=NAVY)
    edge(ws.Range(f"{c1}{r}:{c2}{r}"), XL_EDGE_BOTTOM, XL_THIN, TEAL)


def shade_forecast(ws, rng: str, month_row: int, first_col: str = "D") -> None:
    c = ws.Range(rng).FormatConditions.Add(XL_EXPRESSION, Formula1=f"={first_col}${month_row}>AsOfMonth")
    c.Interior.Color = rgb(FORECAST_FILL)


def month_header(ws, r: int, first: str = "D", last: str = "O") -> None:
    fx(ws, f"{first}{r}", "=FYStart")
    nxt = chr(ord(first) + 1)
    fx(ws, f"{nxt}{r}:{last}{r}", f"=EDATE({first}{r},1)")
    numfmt(ws.Range(f"{first}{r}:{last}{r}"), F_MONTH)
    fx(ws, f"{first}{r + 1}:{last}{r + 1}", f'=IF(ISCLOSED({first}${r}),"Actual","Forecast")')
    rng = ws.Range(f"{first}{r}:{last}{r + 1}")
    rng.HorizontalAlignment = -4152
    font(ws.Range(f"{first}{r}:{last}{r}"), bold=True)
    font(ws.Range(f"{first}{r + 1}:{last}{r + 1}"), color=MUTED, size=9)


# ---------------------------------------------------------------------- Forecast
def forecast(wb, pos: dict) -> None:
    ws = wb.Worksheets("Forecast")
    title(ws, "Rolling forecast by driver",
          '="Actuals through "&TEXT(AsOfMonth,"mmmm yyyy")&"; the other "&(12-MonthsClosed)'
          '&" months are forecast from drivers. Scenario: "&ScenarioName&"."')
    cats = L.CATEGORIES
    b1, b2, bt = 6, 5 + len(cats), 6 + len(cats)
    band(ws, 4, "Forecast basis by category: year-to-date run-rate and trailing unit economics")
    header(ws, 5, ["Category", "YTD cases", "YTD budget cases", "Volume run-rate",
                   "Trailing cases", "Trailing revenue", "Trailing COGS", "Trailing freight",
                   "Price per case", "Cost per case", "Freight per case",
                   "Margin per case, live drivers"])
    column(ws, f"B{b1}", cats)
    put(ws, f"B{bt}", "All categories")
    fx(ws, f"C{b1}:C{b2}", f"=SUMIFS(tbl_Sales[units],tbl_Sales[category],$B{b1},{S_YTD})")
    fx(ws, f"D{b1}:D{b2}", f"=SUMIFS(tbl_Budget[units],tbl_Budget[category],$B{b1},{B_YTD})")
    for c, field in zip("FGHI", ("units", "net_revenue", "cogs", "freight")):
        fx(ws, f"{c}{b1}:{c}{b2}",
           f"=SUMIFS(tbl_Sales[{field}],tbl_Sales[category],$B{b1},{S_TRAIL})")
    fx(ws, f"C{bt}:D{bt}", f"=SUM(C{b1}:C{b2})")
    fx(ws, f"F{bt}:I{bt}", f"=SUM(F{b1}:F{b2})")
    fx(ws, f"E{b1}:E{bt}", f"=SAFEDIV(C{b1},D{b1})")
    fx(ws, f"J{b1}:L{bt}", f"=SAFEDIV(G{b1},$F{b1})")
    fx(ws, f"M{b1}:M{bt}", f"=J{b1}*(1+sel_price)-K{b1}*(1+sel_cost)-L{b1}*(1+sel_freight)")
    for rng, fmt in ((f"C{b1}:D{bt}", F_CASES), (f"F{b1}:F{bt}", F_CASES), (f"E{b1}:E{bt}", F_MULT),
                     (f"G{b1}:I{bt}", F_MONEY), (f"J{b1}:M{bt}", F_PRICE)):
        numfmt(ws.Range(rng), fmt)
    total_row(ws.Range(f"B{bt}:M{bt}"))

    o1, o2, ot = 16, 19, 20
    band(ws, 14, "Operating expense run-rate by department")
    header(ws, 15, ["Department", "YTD actual", "YTD budget", "Run-rate"])
    column(ws, f"B{o1}", L.DEPARTMENTS)
    put(ws, f"B{ot}", "All departments")
    for c, scenario in (("C", "Actual"), ("D", "Budget")):
        fx(ws, f"{c}{o1}:{c}{o2}",
           f'=SUMIFS(tbl_Opex[amount],tbl_Opex[dept],$B{o1},tbl_Opex[scenario],"{scenario}",'
           f'tbl_Opex[month],">="&FYStart,tbl_Opex[month],"<="&AsOfMonth)')
    fx(ws, f"C{ot}:D{ot}", f"=SUM(C{o1}:C{o2})")
    fx(ws, f"E{o1}:E{ot}", f"=SAFEDIV(C{o1},D{o1})")
    numfmt(ws.Range(f"C{o1}:D{ot}"), F_MONEY)
    numfmt(ws.Range(f"E{o1}:E{ot}"), F_MULT)
    total_row(ws.Range(f"B{ot}:E{ot}"))

    mr = 23
    band(ws, 21, "Monthly forecast, category by month. Closed months are zero here: the P&L uses actuals for them")
    put(ws, f"B{mr}", "Month")
    put(ws, f"P{mr}", "Full year")
    month_header(ws, mr)
    font(ws.Range(f"B{mr}:P{mr}"), bold=True)
    spec = [
        ("u0", cats, b1, "Baseline cases: budget cases × the category's YTD volume run-rate"),
        ("units", cats, b1, "Forecast cases: baseline × (1 + volume driver)"),
        ("revenue", cats, b1, "Forecast revenue: cases × trailing price per case × (1 + price driver)"),
        ("cogs", cats, b1, "Forecast cost of goods: cases × trailing cost per case × (1 + cost driver)"),
        ("freight", cats, b1, "Forecast freight: cases × trailing freight per case × (1 + freight driver)"),
        ("opex0", L.DEPARTMENTS, o1, "Baseline operating expenses: budget × the department's YTD run-rate"),
        ("opex", L.DEPARTMENTS, o1, "Forecast operating expenses: baseline × (1 + opex driver)"),
    ]
    blocks = {}
    t = mr + 3
    for key, labels, label_row, caption in spec:
        f, last, tot = t + 1, t + len(labels), t + len(labels) + 1
        blocks[key] = (f, last, tot)
        subhead(ws, t, caption)
        fx(ws, f"B{f}:B{last}", f"=$B{label_row}")
        put(ws, f"B{tot}", "Total")
        t = tot + 3
    lk = f"$B${b1}:$B${b2}"
    formulas = {
        "u0": (f"=IF(ISCLOSED(D${mr}),0,SUMIFS(tbl_Budget[units],tbl_Budget[category],$B{{f}},"
               f"tbl_Budget[month],D${mr})*XLOOKUP($B{{f}},{lk},$E${b1}:$E${b2}))"),
        "units": "=D{u0}*(1+sel_volume)",
        "revenue": f"=D{{units}}*(XLOOKUP($B{{f}},{lk},$J${b1}:$J${b2})*(1+sel_price))",
        "cogs": f"=D{{units}}*(XLOOKUP($B{{f}},{lk},$K${b1}:$K${b2})*(1+sel_cost))",
        "freight": f"=D{{units}}*(XLOOKUP($B{{f}},{lk},$L${b1}:$L${b2})*(1+sel_freight))",
        "opex0": (f'=IF(ISCLOSED(D${mr}),0,SUMIFS(tbl_Opex[amount],tbl_Opex[dept],$B{{f}},'
                  f'tbl_Opex[scenario],"Budget",tbl_Opex[month],D${mr})'
                  f"*XLOOKUP($B{{f}},$B${o1}:$B${o2},$E${o1}:$E${o2}))"),
        "opex": "=D{opex0}*(1+sel_opex)",
    }
    for key, (f, last, tot) in blocks.items():
        refs = {"f": f, "u0": blocks["u0"][0], "units": blocks["units"][0], "opex0": blocks["opex0"][0]}
        fx(ws, f"D{f}:O{last}", formulas[key].format(**refs))
        fx(ws, f"P{f}:P{last}", f"=SUM(D{f}:O{f})")
        fx(ws, f"D{tot}:P{tot}", f"=SUM(D{f}:D{last})")
        numfmt(ws.Range(f"D{f}:P{tot}"), F_CASES if key in ("u0", "units") else F_MONEY)
        total_row(ws.Range(f"B{tot}:P{tot}"))
        font(ws.Range(f"P{f}:P{tot}"), bold=True)
    last_block = blocks["opex"][2]
    shade_forecast(ws, f"D{mr}:O{last_block}", mr)

    # --- how good was the budget?
    a0 = last_block + 3
    band(ws, a0, "How well did the budget call volume? Category × closed month")
    a_hdr, a1 = a0 + 1, a0 + 2
    a2 = a1 + len(cats) - 1
    c_hdr, c1 = a2 + 2, a2 + 3
    c2 = c1 + len(cats) - 1
    for hdr_row, caption in ((a_hdr, "Actual cases"), (c_hdr, "Budget cases")):
        put(ws, f"B{hdr_row}", caption)
        fx(ws, f"D{hdr_row}:O{hdr_row}", f"=D${mr}")
        numfmt(ws.Range(f"D{hdr_row}:O{hdr_row}"), F_MONTH)
        font(ws.Range(f"B{hdr_row}:P{hdr_row}"), bold=True)
        edge(ws.Range(f"B{hdr_row}:P{hdr_row}"), XL_EDGE_BOTTOM, XL_THIN, NAVY)
    put(ws, f"P{a_hdr}", "WAPE")
    fx(ws, f"B{a1}:B{a2}", f"=$B{b1}")
    fx(ws, f"B{c1}:B{c2}", f"=$B{b1}")
    fx(ws, f"D{a1}:O{a2}", f"=IF(ISCLOSED(D${mr}),SUMIFS(tbl_Sales[units],tbl_Sales[category],$B{a1},"
                           f"tbl_Sales[month],D${mr}),0)")
    fx(ws, f"D{c1}:O{c2}", f"=IF(ISCLOSED(D${mr}),SUMIFS(tbl_Budget[units],tbl_Budget[category],$B{c1},"
                           f"tbl_Budget[month],D${mr}),0)")
    fx(ws, f"P{a1}:P{a2}", f"=SAFEDIV(SUMPRODUCT(ABS(D{a1}:O{a1}-D{c1}:O{c1})),SUM(D{a1}:O{a1}))")
    numfmt(ws.Range(f"D{a1}:O{c2}"), F_CASES)
    numfmt(ws.Range(f"P{a1}:P{a2}"), F_PCT)
    w = c2 + 2
    put(ws, f"B{w}", "Budget accuracy, all categories: WAPE")
    put(ws, f"B{w + 1}", "Bias: (budget − actual) ÷ actual")
    fx(ws, f"D{w}", f"=SAFEDIV(SUMPRODUCT(ABS(D{a1}:O{a2}-D{c1}:O{c2})),SUM(D{a1}:O{a2}))")
    fx(ws, f"D{w + 1}", f"=SAFEDIV(SUM(D{c1}:O{c2})-SUM(D{a1}:O{a2}),SUM(D{a1}:O{a2}))")
    numfmt(ws.Range(f"D{w}:D{w + 1}"), '+0.0%;-0.0%;0.0%')
    numfmt(ws.Range(f"D{w}"), "0.0%")
    font(ws.Range(f"B{w}:D{w + 1}"), bold=True)
    put(ws, f"F{w}", "Weighted absolute percentage error: the share of cases the budget "
                     "placed in the wrong category-month.")
    put(ws, f"F{w + 1}", '="Below zero means the budget under-called volume"&IF(D'
                         f'{w + 1}<0,": the year sold more cases than planned.",".")')
    font(ws.Range(f"F{w}:F{w + 1}"), color=MUTED)

    # --- statistical cross-check
    e0 = w + 4
    band(ws, e0, "Cross-check: a statistical forecast of total revenue (FORECAST.ETS, 12-month seasonality)")
    h0 = e0 + 8
    h1, h2 = h0 + 2, h0 + 37
    put(ws, f"B{e0 + 1}", "Month")
    fx(ws, f"D{e0 + 1}:O{e0 + 1}", f"=D${mr}")
    numfmt(ws.Range(f"D{e0 + 1}:O{e0 + 1}"), F_MONTH)
    put(ws, f"P{e0 + 1}", "Open months")
    font(ws.Range(f"B{e0 + 1}:P{e0 + 1}"), bold=True)
    edge(ws.Range(f"B{e0 + 1}:P{e0 + 1}"), XL_EDGE_BOTTOM, XL_THIN, NAVY)
    rev_total = blocks["revenue"][2]
    months = f"$B${h1}:$B${h2}"
    closed_n = f"MATCH(AsOfMonth,{months},0)"
    put(ws, f"B{e0 + 2}", "Driver forecast")
    put(ws, f"B{e0 + 3}", "FORECAST.ETS")
    put(ws, f"B{e0 + 4}", "Driver vs statistical")
    fx(ws, f"D{e0 + 2}:O{e0 + 2}", f'=IF(ISCLOSED(D${mr}),"",D${rev_total})')
    fx(ws, f"D{e0 + 3}:O{e0 + 3}",
       f'=IF(ISCLOSED(D${mr}),"",FORECAST.ETS(MATCH(D${mr},{months},0),'
       f"$D${h1}:INDEX($D${h1}:$D${h2},{closed_n}),$C${h1}:INDEX($C${h1}:$C${h2},{closed_n}),12))")
    fx(ws, f"D{e0 + 4}:O{e0 + 4}", f'=IF(ISCLOSED(D${mr}),"",SAFEDIV(D{e0 + 2}-D{e0 + 3},D{e0 + 3}))')
    fx(ws, f"P{e0 + 2}:P{e0 + 3}", f"=SUM(D{e0 + 2}:O{e0 + 2})")
    fx(ws, f"P{e0 + 4}", f"=SAFEDIV(P{e0 + 2}-P{e0 + 3},P{e0 + 3})")
    numfmt(ws.Range(f"D{e0 + 2}:P{e0 + 3}"), F_MONEY)
    numfmt(ws.Range(f"D{e0 + 4}:P{e0 + 4}"), '+0.0%;-0.0%;0.0%')
    put(ws, f"B{e0 + 5}",
        f'="Over the open months the driver forecast sits "&TEXT(ABS(P{e0 + 4}),"0.0%")&IF(P{e0 + 4}>=0,'
        f'" above"," below")&" a purely statistical projection of the same history. A gap this size says '
        'the drivers, not the trend, are doing the work: check them before trusting either."')
    font(ws.Range(f"B{e0 + 5}"), color=MUTED)
    subhead(ws, h0, "History feeding FORECAST.ETS: total revenue by month")
    row(ws, f"B{h0 + 1}", ["Month", "Index", "Revenue"])
    font(ws.Range(f"B{h0 + 1}:D{h0 + 1}"), bold=True)
    fx(ws, f"B{h1}", "=EDATE(FYStart,-24)")
    fx(ws, f"B{h1 + 1}:B{h2}", f"=EDATE(B{h1},1)")
    fx(ws, f"C{h1}:C{h2}", f"=ROWS($B${h1}:B{h1})")
    fx(ws, f"D{h1}:D{h2}", f'=IF(B{h1}<=AsOfMonth,SUMIFS(tbl_Sales[net_revenue],tbl_Sales[month],B{h1}),"")')
    numfmt(ws.Range(f"B{h1}:B{h2}"), F_MONTH)
    numfmt(ws.Range(f"D{h1}:D{h2}"), F_MONEY)

    widths(ws, {"A": 2, "B": 30, "C": 11, "D:O": 11, "P": 12})
    ws.Range("C5:M5").RowHeight = 42
    pos["fc"] = {"mr": mr, "basis": (b1, b2, bt), "opex_basis": (o1, o2, ot), "blocks": blocks,
                 "accuracy": (a1, a2, c1, c2, w), "ets": e0, "history": (h1, h2)}
    names = {
        "fc_basis": f"Forecast!$B${b1}:$M${b2}",
        "fc_opex_basis": f"Forecast!$B${o1}:$E${o2}",
        "budget_wape": f"Forecast!$D${w}", "budget_bias": f"Forecast!$D${w + 1}",
        "fc_ets": f"Forecast!$D${e0 + 3}:$O${e0 + 3}", "fc_ets_gap": f"Forecast!$P${e0 + 4}",
    }
    for key, (f, last, tot) in blocks.items():
        names[f"fc_{key}"] = f"Forecast!$D${f}:$O${last}"
    for name, ref in names.items():
        wb.Names.Add(Name=name, RefersTo="=" + ref)


# --------------------------------------------------------------------------- P&L
def pnl(wb, pos: dict) -> None:
    ws = wb.Worksheets("PnL")
    fc = pos["fc"]["blocks"]
    title(ws, '="Profit and loss, FY"&FiscalYear&": outlook against budget and prior year"',
          '="Actuals through "&TEXT(AsOfMonth,"mmmm yyyy")&", forecast after (shaded). Scenario: "'
          '&ScenarioName&". Favourable variances are positive; unfavourable ones in brackets."')
    put(ws, "B4", "Month")
    month_header(ws, 4)
    summary = ["FY outlook", "FY budget", "Fav / (unfav)", "Var %", "Prior year", "vs PY",
               "Trend", "YTD actual", "YTD budget", "Fav / (unfav)", "Var %",
               "YTD prior year", "YTD vs PY"]
    row(ws, "P4", summary)
    font(ws.Range("B4:AB4"), bold=True)
    ws.Range("P4:AB4").HorizontalAlignment = -4152
    ws.Range("P4:AB4").WrapText = True
    ws.Rows(4).RowHeight = 28
    put(ws, "P5", "all 12 months")
    put(ws, "W5", '="Jan–"&TEXT(AsOfMonth,"mmm")')
    font(ws.Range("P5:AB5"), color=MUTED, size=9)
    ws.Range("P5:AB5").HorizontalAlignment = -4152

    captions = {"outlook": "Outlook: actuals to the as-of month, driver forecast after",
                "budget": "Budget", "py": "Prior year (actuals)",
                "variance": "Variance: outlook against budget, favourable = positive"}
    for block, first in L.PNL_BLOCKS.items():
        subhead(ws, first - 1, captions[block], "B", "AB" if block == "outlook" else "P")
        for key, label, *_ in L.PNL_LINES:
            r = first + IDX[key]
            put(ws, f"B{r}", ("   " if key.startswith("opex_") and key != "opex_total" else "") + label)

    def r_(block, key):
        return prow(block, key)

    actual = {"cases": "units", "revenue": "net_revenue", "cogs": "cogs", "freight": "freight"}
    fc_total = {"cases": fc["units"][2], "revenue": fc["revenue"][2], "cogs": fc["cogs"][2],
                "freight": fc["freight"][2]}
    o_first, o_last = fc["opex"][0], fc["opex"][1]
    for block in L.PNL_BLOCKS:
        g = {k: r_(block, k) for k in LINE_KEYS}
        for key in LINE_KEYS:
            r = g[key]
            if key in actual and block == "outlook":
                f = (f"=IF(ISCLOSED(D$4),SUMIFS(tbl_Sales[{actual[key]}],tbl_Sales[month],D$4),"
                     f"Forecast!D${fc_total[key]})")
            elif key in actual and block == "budget":
                f = f"=SUMIFS(tbl_Budget[{'units' if key == 'cases' else key}],tbl_Budget[month],D$4)"
            elif key in actual and block == "py":
                f = f"=SUMIFS(tbl_Sales[{actual[key]}],tbl_Sales[month],EDATE(D$4,-12))"
            elif key in OPEX_KEYS and block == "outlook":
                f = (f'=IF(ISCLOSED(D$4),SUMIFS(tbl_Opex[amount],tbl_Opex[month],D$4,tbl_Opex[dept],'
                     f'TRIM($B{r}),tbl_Opex[scenario],"Actual"),XLOOKUP(TRIM($B{r}),'
                     f"Forecast!$B${o_first}:$B${o_last},Forecast!D${o_first}:D${o_last}))")
            elif key in OPEX_KEYS and block == "budget":
                f = (f'=SUMIFS(tbl_Opex[amount],tbl_Opex[month],D$4,tbl_Opex[dept],TRIM($B{r}),'
                     f'tbl_Opex[scenario],"Budget")')
            elif key in OPEX_KEYS and block == "py":
                f = (f'=SUMIFS(tbl_Opex[amount],tbl_Opex[month],EDATE(D$4,-12),tbl_Opex[dept],'
                     f'TRIM($B{r}),tbl_Opex[scenario],"Actual")')
            elif block == "variance":
                o, b = r_("outlook", key), r_("budget", key)
                if KIND[key] == "$":
                    f = f"=FAVVAR(D{o},D{b},{str(IS_COST[key]).upper()})"
                else:
                    f = f"=D{o}-D{b}"
            elif key == "gm":
                f = f"=D{g['revenue']}-D{g['cogs']}"
            elif key == "gm_pct":
                f = f"=SAFEDIV(D{g['gm']},D{g['revenue']})"
            elif key == "contribution":
                f = f"=D{g['gm']}-D{g['freight']}"
            elif key == "opex_total":
                f = f"=SUM(D{g['opex_sm']}:D{g['opex_tech']})"
            elif key == "ebitda":
                f = f"=D{g['contribution']}-D{g['opex_total']}"
            elif key == "ebitda_pct":
                f = f"=SAFEDIV(D{g['ebitda']},D{g['revenue']})"
            fx(ws, f"D{r}:O{r}", f)
            # full-year column
            if KIND[key] == "%" and block != "variance":
                num = g["gm"] if key == "gm_pct" else g["ebitda"]
                fx(ws, f"P{r}", f"=SAFEDIV(P{num},P{g['revenue']})")
            elif KIND[key] == "%":
                fx(ws, f"P{r}", f"=P{r_('outlook', key)}-P{r_('budget', key)}")
            else:
                fx(ws, f"P{r}", f"=SUM(D{r}:O{r})")
            is_var = block == "variance"
            fmt = {"$": F_VAR if is_var else F_MONEY, "cases": F_VAR if is_var else F_CASES,
                   "%": F_VARPCT if is_var else F_PCT}[KIND[key]]
            numfmt(ws.Range(f"D{r}:P{r}"), fmt)
        for key in ("gm", "contribution", "opex_total"):
            total_row(ws.Range(f"B{g[key]}:P{g[key]}"))
        total_row(ws.Range(f"B{g['ebitda']}:P{g['ebitda']}"), double=True)
        for key in ("gm_pct", "ebitda_pct"):
            font(ws.Range(f"B{g[key]}:P{g[key]}"), italic=True, color=MUTED)
        font(ws.Range(f"P{g['cases']}:P{g['ebitda_pct']}"), bold=True)
        if block == "variance":
            sign_colours(ws.Range(f"D{g['cases']}:P{g['ebitda_pct']}"))

    # summary columns on the outlook block
    ytd = '$D$4:$O$4,"<="&AsOfMonth'
    for key in LINE_KEYS:
        r, b, p = r_("outlook", key), r_("budget", key), r_("py", key)
        cost = str(IS_COST[key]).upper()
        pct = KIND[key] == "%"
        rev_o = r_("outlook", "revenue")
        num_key = "gm" if key == "gm_pct" else "ebitda"
        cells = {"Q": f"=P{b}", "T": f"=P{p}"}
        if pct:
            n = r_("outlook", num_key)
            cells.update({
                "R": f"=P{r}-Q{r}", "U": f"=P{r}-T{r}",
                "W": f"=SAFEDIV(W{n},W{rev_o})", "X": f"=SAFEDIV(X{n},X{rev_o})",
                "Y": f"=W{r}-X{r}", "AA": f"=SAFEDIV(AA{n},AA{rev_o})", "AB": f"=W{r}-AA{r}",
            })
        else:
            var = (lambda a, c: f"=FAVVAR({a}{r},{c}{r},{cost})") if KIND[key] == "$" else \
                  (lambda a, c: f"={a}{r}-{c}{r}")
            cells.update({
                "R": var("P", "Q"), "S": f"=SAFEDIV(R{r},ABS(Q{r}))",
                "U": f"=SAFEDIV(P{r}-T{r},ABS(T{r}))",
                "W": f"=SUMIFS(D{r}:O{r},{ytd})", "X": f"=SUMIFS(D{b}:O{b},{ytd})",
                "Y": var("W", "X"), "Z": f"=SAFEDIV(Y{r},ABS(X{r}))",
                "AA": f"=SUMIFS(D{p}:O{p},{ytd})", "AB": f"=SAFEDIV(W{r}-AA{r},ABS(AA{r}))",
            })
        for c, f in cells.items():
            fx(ws, f"{c}{r}", f)
        base = {"$": F_MONEY, "cases": F_CASES, "%": F_PCT}[KIND[key]]
        for c in ("Q", "T", "W", "X", "AA"):
            numfmt(ws.Range(f"{c}{r}"), base)
        for c in ("R", "Y"):
            numfmt(ws.Range(f"{c}{r}"), F_VARPCT if pct else F_VAR)
        for c in ("S", "U", "Z", "AB"):
            numfmt(ws.Range(f"{c}{r}"), F_VARPCT)
    o1, o2 = L.PNL_BLOCKS["outlook"], L.PNL_BLOCKS["outlook"] + len(LINE_KEYS) - 1
    for c in ("R", "S", "Y", "Z"):
        sign_colours(ws.Range(f"{c}{o1}:{c}{o2}"))
    for key in ("gm", "contribution", "opex_total"):
        total_row(ws.Range(f"Q{r_('outlook', key)}:AB{r_('outlook', key)}"))
    total_row(ws.Range(f"Q{r_('outlook', 'ebitda')}:AB{r_('outlook', 'ebitda')}"), double=True)
    for key in ("gm_pct", "ebitda_pct"):
        font(ws.Range(f"Q{r_('outlook', key)}:AB{r_('outlook', key)}"), italic=True, color=MUTED)
    ws.Range(f"V{o1}:V{o2}").SparklineGroups.Add(XL_SPARKLINE_LINE, f"PnL!D{o1}:O{o2}")
    sg = ws.Range(f"V{o1}").SparklineGroups(1)
    sg.SeriesColor.Color = rgb(NAVY)
    sg.Points.Highpoint.Visible = True
    sg.Points.Highpoint.Color.Color = rgb(TEAL)
    fill(ws.Range(f"W{o1 - 1}:AB{o2}"), "#F7F9FC")
    shade_forecast(ws, "D4:O71", 4)
    widths(ws, {"A": 2, "B": 26, "C": 1, "D:O": 10.5, "P:U": 11, "V": 12, "W:AB": 11})
    ws.Columns("B").HorizontalAlignment = XL_LEFT

    names = {"pnl_months": "PnL!$D$4:$O$4"}
    for block, first in L.PNL_BLOCKS.items():
        last = first + len(LINE_KEYS) - 1
        names[f"pnl_{block}"] = f"PnL!$D${first}:$O${last}"
        names[f"pnl_fy_{block}"] = f"PnL!$P${first}:$P${last}"
    names["pnl_summary"] = f"PnL!$P${o1}:$AB${o2}"
    for name, ref in names.items():
        wb.Names.Add(Name=name, RefersTo="=" + ref)
    pos["pnl"] = {k: {key: r_(k, key) for key in LINE_KEYS} for k in L.PNL_BLOCKS}


# ------------------------------------------------------------------------- PVM
PVM_HEADER = ["Category", "Channel", "Budget cases", "Budget revenue", "Budget COGS",
              "Actual cases", "Actual revenue", "Actual COGS", "Budget price / case",
              "Actual price / case", "Budget cost / case", "Actual cost / case",
              "Budget margin / case", "Share shift", "Volume", "Mix", "Price", "Unit cost",
              "Effects, total", "GM variance", "Unexplained", "Main driver"]


def pvm_table(ws, band_row: int, caption: str, start: str) -> tuple[int, int, int]:
    band(ws, band_row, caption, "B", "W")
    header(ws, band_row + 1, PVM_HEADER)
    ws.Range(f"C{band_row + 1}").HorizontalAlignment = XL_LEFT
    f = band_row + 2
    segs = [(c, ch) for c in L.CATEGORIES for ch in L.CHANNELS]
    last, tot = f + len(segs) - 1, f + len(segs)
    ws.Range(f"B{f}:C{last}").Value = tuple(segs)
    put(ws, f"B{tot}", "All segments")
    window = f'">="&{start},{{t}}[month],"<="&AsOfMonth'
    for c, (table, field) in zip("DEFGHI", (("tbl_Budget", "units"), ("tbl_Budget", "revenue"),
                                            ("tbl_Budget", "cogs"), ("tbl_Sales", "units"),
                                            ("tbl_Sales", "net_revenue"), ("tbl_Sales", "cogs"))):
        fx(ws, f"{c}{f}:{c}{last}",
           f"=SUMIFS({table}[{field}],{table}[category],$B{f},{table}[channel],$C{f},"
           f"{table}[month],{window.format(t=table)})")
    fx(ws, f"D{tot}:I{tot}", f"=SUM(D{f}:D{last})")
    fx(ws, f"J{f}:J{tot}", f"=SAFEDIV(E{f},D{f})")
    fx(ws, f"K{f}:K{tot}", f"=SAFEDIV(H{f},G{f})")
    fx(ws, f"L{f}:L{tot}", f"=SAFEDIV(F{f},D{f})")
    fx(ws, f"M{f}:M{tot}", f"=SAFEDIV(I{f},G{f})")
    fx(ws, f"N{f}:N{tot}", f"=J{f}-L{f}")
    fx(ws, f"O{f}:O{last}", f"=SAFEDIV(G{f},G${tot})-SAFEDIV(D{f},D${tot})")
    fx(ws, f"O{tot}", f"=SUM(O{f}:O{last})")
    fx(ws, f"P{f}:P{last}", f"=(G{f}-D{f})*N${tot}")
    fx(ws, f"Q{f}:Q{last}", f"=(G{f}-D{f})*(N{f}-N${tot})")
    fx(ws, f"R{f}:R{last}", f"=G{f}*(K{f}-J{f})")
    fx(ws, f"S{f}:S{last}", f"=-G{f}*(M{f}-L{f})")
    fx(ws, f"P{tot}:S{tot}", f"=SUM(P{f}:P{last})")
    fx(ws, f"T{f}:T{tot}", f"=SUM(P{f}:S{f})")
    fx(ws, f"U{f}:U{tot}", f"=(H{f}-I{f})-(E{f}-F{f})")
    fx(ws, f"V{f}:V{tot}", f"=T{f}-U{f}")
    fx(ws, f"W{f}:W{tot}",
       f'=INDEX({{"Volume","Mix","Price","Unit cost"}},XMATCH(MAX(ABS(P{f}:S{f})),ABS(P{f}:S{f})))')
    for rng, fmt in ((f"D{f}:D{tot}", F_CASES), (f"G{f}:G{tot}", F_CASES),
                     (f"E{f}:F{tot}", F_MONEY), (f"H{f}:I{tot}", F_MONEY),
                     (f"J{f}:N{tot}", F_PRICE), (f"O{f}:O{tot}", '+0.0%;-0.0%;"–"'), (f"P{f}:V{tot}", F_VAR)):
        numfmt(ws.Range(rng), fmt)
    sign_colours(ws.Range(f"P{f}:U{tot}"))
    total_row(ws.Range(f"B{tot}:W{tot}"), double=True)
    font(ws.Range(f"T{f}:U{tot}"), bold=True)
    font(ws.Range(f"V{f}:V{tot}"), color=MUTED)
    numfmt(ws.Range(f"V{f}:V{tot}"), '[>=0.5]+#,##0;[<=-0.5](#,##0);"–"')   # float dust shows as a dash
    return f, last, tot


def summary_table(ws, band_row: int, caption: str, labels, key_col: str, src: tuple) -> tuple:
    f0, l0, _ = src
    band(ws, band_row, caption, "B", "J")
    header(ws, band_row + 1, [caption.split(" by ")[-1].capitalize(), "Volume", "Mix", "Price",
                              "Unit cost", "Effects, total", "Budget GM", "Actual GM", "GM vs budget"])
    f = band_row + 2
    last, tot = f + len(labels) - 1, f + len(labels)
    column(ws, f"B{f}", labels)
    put(ws, f"B{tot}", "Total")
    key = f"${key_col}${f0}:${key_col}${l0}"
    fx(ws, f"C{f}:F{last}", f"=SUMIFS(P${f0}:P${l0},{key},$B{f})")
    fx(ws, f"H{f}:H{last}", f"=SUMIFS($E${f0}:$E${l0},{key},$B{f})-SUMIFS($F${f0}:$F${l0},{key},$B{f})")
    fx(ws, f"I{f}:I{last}", f"=SUMIFS($H${f0}:$H${l0},{key},$B{f})-SUMIFS($I${f0}:$I${l0},{key},$B{f})")
    fx(ws, f"C{tot}:F{tot}", f"=SUM(C{f}:C{last})")
    fx(ws, f"H{tot}:I{tot}", f"=SUM(H{f}:H{last})")
    fx(ws, f"G{f}:G{tot}", f"=SUM(C{f}:F{f})")
    fx(ws, f"J{f}:J{tot}", f"=SAFEDIV(I{f}-H{f},ABS(H{f}))")
    numfmt(ws.Range(f"C{f}:G{tot}"), F_VAR)
    numfmt(ws.Range(f"H{f}:I{tot}"), F_MONEY)
    numfmt(ws.Range(f"J{f}:J{tot}"), F_VARPCT)
    sign_colours(ws.Range(f"C{f}:G{tot}"))
    sign_colours(ws.Range(f"J{f}:J{tot}"))
    total_row(ws.Range(f"B{tot}:J{tot}"), double=True)
    return f, last, tot


def pvm(wb, pos: dict) -> None:
    ws = wb.Worksheets("PVM")
    title(ws, "Price-volume-mix: why margin moved against budget",
          "Gross-margin variance split four ways on 24 category × channel segments. Each segment's change "
          "in cases is valued at the budget's average margin per case (volume) and at its own margin's "
          "distance from that average (mix); price and cost are per-case changes on actual cases. The four "
          "effects add up to every segment's variance exactly: column V is zero.", span="W")
    m = pvm_table(ws, 4, '="Month: "&TEXT(AsOfMonth,"mmmm yyyy")&" against budget"', "AsOfMonth")
    y = pvm_table(ws, m[2] + 3, '="Year to date: "&TEXT(FYStart,"mmmm")&" to "'
                                '&TEXT(AsOfMonth,"mmmm yyyy")&" against budget"', "FYStart")
    cat = summary_table(ws, y[2] + 3, "YTD gross-margin bridge by category", L.CATEGORIES, "B", y)
    chan = summary_table(ws, cat[2] + 3, "YTD gross-margin bridge by channel", L.CHANNELS, "C", y)

    b0 = chan[2] + 3
    band(ws, b0, "Year-to-date bridges", "B", "W")
    yt = y[2]
    p = pos["pnl"]["outlook"]
    gm = [("Budget gross margin", f"=E{yt}-F{yt}"), ("Volume", f"=P{yt}"), ("Mix", f"=Q{yt}"),
          ("Price", f"=R{yt}"), ("Unit cost", f"=S{yt}"), ("Actual gross margin", f"=H{yt}-I{yt}")]
    eb = [("Budget EBITDA", f"=PnL!X{p['ebitda']}"), ("Volume", f"=P{yt}"), ("Mix", f"=Q{yt}"),
          ("Price", f"=R{yt}"), ("Unit cost", f"=S{yt}"), ("Freight", f"=PnL!Y{p['freight']}"),
          ("Operating expenses", f"=PnL!Y{p['opex_total']}"), ("Actual EBITDA", f"=PnL!W{p['ebitda']}")]
    r1 = b0 + 2
    row(ws, f"B{b0 + 1}", ["Gross margin", "$"])
    row(ws, f"E{b0 + 1}", ["EBITDA", "$"])
    font(ws.Range(f"B{b0 + 1}:F{b0 + 1}"), bold=True)
    for i, (label, f) in enumerate(gm):
        put(ws, f"B{r1 + i}", label)
        fx(ws, f"C{r1 + i}", f)
    for i, (label, f) in enumerate(eb):
        put(ws, f"E{r1 + i}", label)
        fx(ws, f"F{r1 + i}", f)
    g_last, e_last = r1 + len(gm) - 1, r1 + len(eb) - 1
    numfmt(ws.Range(f"C{r1}:C{g_last}"), F_MONEY)
    numfmt(ws.Range(f"F{r1}:F{e_last}"), F_MONEY)
    numfmt(ws.Range(f"C{r1 + 1}:C{g_last - 1}"), F_VAR)
    numfmt(ws.Range(f"F{r1 + 1}:F{e_last - 1}"), F_VAR)
    sign_colours(ws.Range(f"C{r1 + 1}:C{g_last - 1}"))
    sign_colours(ws.Range(f"F{r1 + 1}:F{e_last - 1}"))
    for rng in (f"B{g_last}:C{g_last}", f"E{e_last}:F{e_last}"):
        total_row(ws.Range(rng), double=True)
    chk = e_last + 2
    put(ws, f"B{chk}", "Bridge closes (zero)")
    fx(ws, f"C{chk}", f"=C{r1}+SUM(C{r1 + 1}:C{g_last - 1})-C{g_last}")
    put(ws, f"E{chk}", "Bridge closes (zero)")
    fx(ws, f"F{chk}", f"=F{r1}+SUM(F{r1 + 1}:F{e_last - 1})-F{e_last}")
    numfmt(ws.Range(f"C{chk}"), "0.00")
    numfmt(ws.Range(f"F{chk}"), "0.00")
    font(ws.Range(f"B{chk}:F{chk}"), color=MUTED, italic=True)

    widths(ws, {"A": 2, "B": 19, "C": 19, "D:I": 11, "J:N": 10, "O": 8, "P:V": 11, "W": 10})
    ws.Rows(5).RowHeight = 42
    ws.Rows(y[0] - 1).RowHeight = 42
    pos["pvm"] = {"month": m, "ytd": y, "category": cat, "channel": chan,
                  "gm_bridge": (r1, g_last), "ebitda_bridge": (r1, e_last), "bridge_check": chk,
                  "gm_labels": [x[0] for x in gm], "ebitda_labels": [x[0] for x in eb]}
    names = {
        "pvm_month": f"PVM!$B${m[0]}:$W${m[1]}", "pvm_month_total": f"PVM!$B${m[2]}:$W${m[2]}",
        "pvm_ytd": f"PVM!$B${y[0]}:$W${y[1]}", "pvm_ytd_total": f"PVM!$B${y[2]}:$W${y[2]}",
        "pvm_by_category": f"PVM!$B${cat[0]}:$J${cat[1]}",
        "pvm_by_channel": f"PVM!$B${chan[0]}:$J${chan[1]}",
        "bridge_gm": f"PVM!$C${r1}:$C${g_last}", "bridge_ebitda": f"PVM!$F${r1}:$F${e_last}",
    }
    for name, ref in names.items():
        wb.Names.Add(Name=name, RefersTo="=" + ref)


__all__ = ["forecast", "pnl", "pvm", "prow", "IDX"]
