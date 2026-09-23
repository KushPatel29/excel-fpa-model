"""Calculation sheets built on the public filings: the FERC P&L, price-volume-mix,
the EIA monthly view with its FERC tie-out, and the weather model.

Every figure is a formula over the source tables (tbl_FERC_*, tbl_Monthly,
tbl_Weather) and the named inputs, written once per row or block and filled
with relative references the way an analyst would enter them.
"""
from __future__ import annotations

import layout as L
from xl import (
    F_CASES, F_MONEY, F_MONTH, F_PCT, F_VAR, F_VARPCT, INK, MUTED, NAVY, TEAL, XL_EDGE_BOTTOM, XL_LEFT,
    XL_LINE, XL_THIN, XL_WATERFALL, band, column, edge, font, fx, header, new_chart, numfmt, place, put,
    rgb, sign_colours, style_chart, title, total_row, widths,
)

F_CPK = '0.00" ¢"'
F_DOLLAR2 = '$#,##0.00'
F_MWH = '#,##0;(#,##0);"–"'
F_MILLIONS = '#,##0.0,,;(#,##0.0,,);"–"'
F_MILLIONS_VAR = '+#,##0.0,,;(#,##0.0,,);"–"'

FR, FE, FI = "tbl_FERC_Revenue", "tbl_FERC_Expense", "tbl_FERC_Income"
FUEL = '{"fuel_steam_power_generation","nuclear_fuel_expense","fuel","fuel_other_renewable_generation"}'
PURCHASED = '{"purchased_power","power_purchased_for_storage_operations","storage_fuel_energy_storage_expense"}'
CUSTOMER = '{"customer_account_expenses","customer_service_and_information_expenses","sales_expenses"}'
DA = ('{"depreciation_expense","depreciation_expense_for_asset_retirement_costs",'
      '"amortization_and_depletion_of_utility_plant","amortization_of_other_utility_plant"}')


def names(wb, specs: dict) -> None:
    for name, ref in specs.items():
        wb.Names.Add(Name=name, RefersTo="=" + ref)


def ferc_sum(table: str, value: str, key: str, utility: str = "CompanyId", year: str = "D$4",
             many: bool = False) -> str:
    """SUMIFS over a long FERC table for one utility, year and line; `many` sums an array of lines."""
    kind = {FR: "revenue_type", FE: "expense_type", FI: "income_type"}[table]
    inner = (f"SUMIFS({table}[{value}],{table}[utility_id_ferc1],{utility},{table}[report_year],{year},"
             f"{table}[{kind}],{key})")
    return f"SUM({inner})" if many else inner


def fill_tokens(formula: str, rows: dict[str, int]) -> str:
    """Replace {key} tokens with row numbers, leaving array constants alone."""
    out = formula
    for key, r in sorted(rows.items(), key=lambda kv: -len(kv[0])):
        out = out.replace("{" + key + "}", str(r))
    return out


# ------------------------------------------------------------------- PnL
PNL_ROWS = [  # key, label, kind, formula for column D (tokens in braces), style
    ("band", "Revenue"),
    *[(f"rev_{k}", lbl, "$", "=" + ferc_sum(FR, "dollar_value", f'"{k}"'), "class") for lbl, k in L.FERC_CLASSES],
    ("retail", "Retail revenue", "$", "=SUM(D{rev_residential_sales}:D{rev_other_sales_to_public_authorities})",
     "total"),
    ("wholesale", "Wholesale (sales for resale)", "$", "=" + ferc_sum(FR, "dollar_value", '"sales_for_resale"'), ""),
    ("refunds", "Refund provision", "$", "=-" + ferc_sum(FR, "dollar_value", '"provision_for_rate_refunds"'), ""),
    ("other_rev", "Other operating revenue", "$", "=" + ferc_sum(FR, "dollar_value", '"other_operating_revenues"'), ""),
    ("revenue", "Total operating revenue", "$", "=SUM(D{retail}:D{other_rev})", "total"),
    ("band", "Power supply"),
    ("fuel", "Fuel", "$", "=" + ferc_sum(FE, "dollar_value", FUEL, many=True), ""),
    ("purchased", "Purchased power", "$", "=" + ferc_sum(FE, "dollar_value", PURCHASED, many=True), ""),
    ("power", "Power cost", "$", "=D{fuel}+D{purchased}", "total"),
    ("gm", "Gross margin", "$", "=D{revenue}-D{power}", "total"),
    ("gm_pct", "Gross margin %", "%", "=SAFEDIV(D{gm},D{revenue})", "pct"),
    ("band", "Operating expenses, excluding power"),
    ("production", "Production, excluding fuel and purchased power", "$",
     "=" + ferc_sum(FE, "dollar_value", '"power_production_expenses"') + "-D{power}", ""),
    ("transmission", "Transmission", "$", "=" + ferc_sum(FE, "dollar_value", '"transmission_expenses"'), ""),
    ("distribution", "Distribution", "$", "=" + ferc_sum(FE, "dollar_value", '"distribution_expenses"'), ""),
    ("customer", "Customer accounts, service and sales", "$", "=" + ferc_sum(FE, "dollar_value", CUSTOMER, many=True),
     ""),
    ("ag", "Administrative and general", "$",
     "=" + ferc_sum(FE, "dollar_value", '"administrative_and_general_expenses"'), ""),
    ("regional", "Regional market and other", "$", "=D{other_om}-SUM(D{production}:D{ag})", ""),
    ("other_om", "Other O&M", "$",
     "=" + ferc_sum(FE, "dollar_value", '"operations_and_maintenance_expenses_electric"') + "-D{power}", "total"),
    ("da", "Depreciation and amortization", "$", "=" + ferc_sum(FI, "dollar_value", DA, many=True), ""),
    ("taxes", "Taxes other than income", "$",
     "=" + ferc_sum(FI, "dollar_value", '"taxes_other_than_income_taxes_utility_operating_income"'), ""),
    ("op_income", "Operating income before income taxes", "$", "=D{gm}-D{other_om}-D{da}-D{taxes}", "grand"),
    ("op_margin", "Operating margin %", "%", "=SAFEDIV(D{op_income},D{revenue})", "pct"),
    ("band", "Volumes and unit economics"),
    ("retail_mwh", "Retail MWh", "mwh", "=" + ferc_sum(FR, "sales_mwh", '"sales_to_ultimate_consumers"'), ""),
    ("residential_mwh", "Residential MWh", "mwh", "=" + ferc_sum(FR, "sales_mwh", '"residential_sales"'), ""),
    ("industrial_mwh", "Industrial MWh", "mwh", "=" + ferc_sum(FR, "sales_mwh", '"large_or_industrial"'), ""),
    ("total_mwh", "Total MWh sold, including wholesale", "mwh",
     "=" + ferc_sum(FR, "sales_mwh", '"sales_of_electricity"'), ""),
    ("customers", "Retail customers (average)", "mwh",
     "=" + ferc_sum(FR, "avg_customers_per_month", '"sales_to_ultimate_consumers"'), ""),
    ("retail_cpk", "Retail price, ¢/kWh", "cpk", "=CPK(D{retail},D{retail_mwh})", ""),
    ("residential_cpk", "Residential price, ¢/kWh", "cpk", "=CPK(D{rev_residential_sales},D{residential_mwh})", ""),
    ("industrial_cpk", "Industrial price, ¢/kWh", "cpk", "=CPK(D{rev_large_or_industrial},D{industrial_mwh})", ""),
    ("power_per_mwh", "Power cost per MWh sold, $", "unit", "=SAFEDIV(D{power},D{total_mwh})", ""),
    ("om_per_customer", "Other O&M per customer, $", "unit", "=SAFEDIV(D{other_om},D{customers})", ""),
    ("industrial_share", "Industrial share of retail MWh", "%", "=SAFEDIV(D{industrial_mwh},D{retail_mwh})", "pct"),
]


def pnl(wb, pos: dict) -> None:
    ws = wb.Worksheets("PnL")
    title(ws, "Profit and loss: Portland General Electric, electric utility (FERC Form 1)",
          "Audited annual figures as filed with FERC, via Catalyst Cooperative's PUDL, in $ millions. Column C names "
          "the FERC line each revenue row reads. Operating income is before income taxes and interest.", span="R")
    put(ws, "B4", "Year")
    put(ws, "C4", "FERC line")
    fx(ws, "D4", "=FercYear-11")
    fx(ws, "E4:O4", "=D4+1")
    numfmt(ws.Range("D4:O4"), "0")
    fx(ws, "P4", '="CAGR "&BaseYear&"–"&RIGHT(FercYear,2)')
    fx(ws, "Q4", '="Δ vs "&(FercYear-1)')
    put(ws, "R4", "% change")
    font(ws.Range("B4:R4"), bold=True)
    ws.Range("D4:R4").HorizontalAlignment = -4152
    edge(ws.Range("B4:R4"), XL_EDGE_BOTTOM, XL_THIN, NAVY)

    r = 5
    rows: dict[str, int] = {}
    planned = []
    for item in PNL_ROWS:
        if item[0] == "band":
            planned.append((r, item))
            r += 1
            continue
        rows[item[0]] = r
        planned.append((r, item))
        r += 1
    last = r - 1
    for r, item in planned:
        if item[0] == "band":
            band(ws, r, item[1], "B", "R")
            continue
        key, label, kind, formula, style = item
        put(ws, f"B{r}", ("   " + label) if style == "class" else label)
        if key.startswith("rev_"):
            put(ws, f"C{r}", key[4:])
        fx(ws, f"D{r}:O{r}", fill_tokens(formula, rows))
        fmt = {"$": F_MILLIONS, "%": F_PCT, "mwh": F_MWH, "cpk": F_CPK, "unit": F_DOLLAR2}[kind]
        numfmt(ws.Range(f"D{r}:O{r}"), fmt)
        latest = f"INDEX($D{r}:$O{r},MATCH(FercYear,$D$4:$O$4,0))"
        prior = f"INDEX($D{r}:$O{r},MATCH(FercYear-1,$D$4:$O$4,0))"
        base = f"INDEX($D{r}:$O{r},MATCH(BaseYear,$D$4:$O$4,0))"
        if kind == "%":
            fx(ws, f"P{r}", f"={latest}-{base}")
            fx(ws, f"Q{r}", f"={latest}-{prior}")
            numfmt(ws.Range(f"P{r}:Q{r}"), F_VARPCT)
        else:
            fx(ws, f"P{r}", f"=IFERROR(({latest}/{base})^(1/(FercYear-BaseYear))-1,0)")
            fx(ws, f"Q{r}", f"={latest}-{prior}")
            fx(ws, f"R{r}", f"=SAFEDIV(Q{r},ABS({prior}))")
            numfmt(ws.Range(f"P{r}"), F_VARPCT)
            numfmt(ws.Range(f"Q{r}"), {"$": F_MILLIONS_VAR, "mwh": F_VAR}.get(kind, '+#,##0.00;(#,##0.00);"–"'))
            numfmt(ws.Range(f"R{r}"), F_VARPCT)
        if style in ("total", "grand"):
            total_row(ws.Range(f"B{r}:R{r}"), double=style == "grand")
        if style == "pct":
            font(ws.Range(f"B{r}:R{r}"), italic=True, color=MUTED)
    font(ws.Range(f"C5:C{last}"), color=MUTED, size=8)
    widths(ws, {"A": 2, "B": 40, "C": 24, "D:O": 10, "P": 13, "Q:R": 11})
    pos["pnl"] = rows
    specs = {"pnl_years": "PnL!$D$4:$O$4", "pnl_summary": f"PnL!$P$5:$R${last}"}
    for key, r in rows.items():
        specs[f"pnl_{key}"] = f"PnL!$D${r}:$O${r}"
    names(wb, specs)


# ------------------------------------------------------------------- PVM
PVM_HEADER = ["Class", "Source line", "MWh before", "Revenue before", "MWh after", "Revenue after",
              "$/MWh before", "$/MWh after", "Volume", "Mix", "Price", "Effects, total", "Revenue change",
              "Unexplained"]


def pvm_block(ws, r0: int, caption: str, rows: list[tuple], q0: str, v0: str, q1: str, v1: str):
    """rows: (label, key). q0/v0/q1/v1: formula templates with a {key} placeholder."""
    band(ws, r0, caption, "B", "O")
    header(ws, r0 + 1, PVM_HEADER)
    ws.Range(f"C{r0 + 1}").HorizontalAlignment = XL_LEFT
    f, last = r0 + 2, r0 + 1 + len(rows)
    tot = last + 1
    for i, (label, key) in enumerate(rows):
        r = f + i
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", key)
        for c, template in zip("DEFG", (q0, v0, q1, v1)):
            fx(ws, f"{c}{r}", template.replace("{key}", f"$C{r}"))
    put(ws, f"B{tot}", "All classes")
    fx(ws, f"D{tot}:G{tot}", f"=SUM(D{f}:D{last})")
    fx(ws, f"H{f}:H{tot}", f"=SAFEDIV(E{f},D{f})")
    fx(ws, f"I{f}:I{tot}", f"=SAFEDIV(G{f},F{f})")
    fx(ws, f"J{f}:J{last}", f"=(F{f}-D{f})*$H${tot}")
    fx(ws, f"K{f}:K{last}", f"=(F{f}-D{f})*(H{f}-$H${tot})")
    fx(ws, f"L{f}:L{last}", f"=F{f}*(I{f}-H{f})")
    fx(ws, f"J{tot}:L{tot}", f"=SUM(J{f}:J{last})")
    fx(ws, f"M{f}:M{tot}", f"=SUM(J{f}:L{f})")
    fx(ws, f"N{f}:N{tot}", f"=G{f}-E{f}")
    fx(ws, f"O{f}:O{tot}", f"=M{f}-N{f}")
    numfmt(ws.Range(f"D{f}:D{tot}"), F_MWH)
    numfmt(ws.Range(f"F{f}:F{tot}"), F_MWH)
    numfmt(ws.Range(f"E{f}:E{tot}"), F_MONEY)
    numfmt(ws.Range(f"G{f}:G{tot}"), F_MONEY)
    numfmt(ws.Range(f"H{f}:I{tot}"), F_DOLLAR2)
    numfmt(ws.Range(f"J{f}:N{tot}"), F_VAR)
    numfmt(ws.Range(f"O{f}:O{tot}"), '[>=0.5]+#,##0;[<=-0.5](#,##0);"–"')
    sign_colours(ws.Range(f"J{f}:N{tot}"))
    font(ws.Range(f"C{f}:C{last}"), color=MUTED, size=8)
    font(ws.Range(f"O{f}:O{tot}"), color=MUTED)
    total_row(ws.Range(f"B{tot}:O{tot}"), double=True)
    return f, last, tot


def pvm(wb, pos: dict) -> None:
    ws = wb.Worksheets("PVM")
    title(ws, "Price, volume and mix: where retail revenue growth came from",
          "Each class's change in MWh is valued at the earlier period's average price (volume) and at its own price's "
          "distance from that average (mix); price is the change in $/MWh on the later period's MWh. The three effects "
          "add up to every class's revenue change: column O is zero.", span="O")
    ferc = list(L.FERC_CLASSES)

    def f_(value, year):
        return "=" + ferc_sum(FR, value, "{key}", year=year)

    long = pvm_block(ws, 4, '="Long view: "&BaseYear&" to "&FercYear&", FERC Form 1"', ferc,
                     f_("sales_mwh", "BaseYear"), f_("dollar_value", "BaseYear"),
                     f_("sales_mwh", "FercYear"), f_("dollar_value", "FercYear"))
    short = pvm_block(ws, long[2] + 3, '="Last year: "&(FercYear-1)&" to "&FercYear&", FERC Form 1"', ferc,
                      f_("sales_mwh", "FercYear-1"), f_("dollar_value", "FercYear-1"),
                      f_("sales_mwh", "FercYear"), f_("dollar_value", "FercYear"))
    eia = [(L.CLASS_LABELS[c], c) for c in L.CLASSES]

    def y(field, start, end):
        return (f'=SUMIFS(tbl_Monthly[{field}],tbl_Monthly[class],{{key}},tbl_Monthly[month],">="&{start},'
                f'tbl_Monthly[month],"<="&{end})')

    ytd = pvm_block(ws, short[2] + 3, '="Year to date: "&TEXT(FYStart,"mmm")&"–"&TEXT(AsOfMonth,"mmm yyyy")&'
                    '" against the same months of "&(FiscalYear-1)&", EIA-861M"', eia,
                    y("mwh", "EDATE(FYStart,-12)", "EDATE(AsOfMonth,-12)"),
                    y("revenue", "EDATE(FYStart,-12)", "EDATE(AsOfMonth,-12)"),
                    y("mwh", "FYStart", "AsOfMonth"), y("revenue", "FYStart", "AsOfMonth"))
    b0 = ytd[2] + 3
    band(ws, b0, "Bridge for the chart: long view", "B", "O")
    lt = long[2]
    labels = ['=BaseYear&" retail revenue"', "Volume", "Mix", "Price", '=FercYear&" retail revenue"']
    values = [f"=E{lt}", f"=J{lt}", f"=K{lt}", f"=L{lt}", f"=G{lt}"]
    for i, (lab, val) in enumerate(zip(labels, values)):
        put(ws, f"B{b0 + 1 + i}", lab)
        fx(ws, f"C{b0 + 1 + i}", val)
    numfmt(ws.Range(f"C{b0 + 1}:C{b0 + 5}"), '$#,##0,,"M";-$#,##0,,"M"')
    ws.Activate()
    ws.Range(f"B{b0 + 1}:C{b0 + 5}").Select()
    wf = ws.Shapes.AddChart2(-1, XL_WATERFALL, *place(ws, f"E{b0 + 1}", f"K{b0 + 18}")).Chart
    s = wf.FullSeriesCollection(1)
    s.Points(1).IsTotal = True
    s.Points(5).IsTotal = True
    s.HasDataLabels = True
    wf.HasTitle = True
    wf.ChartTitle.Text = "Retail revenue bridge, long view"
    wf.HasLegend = False
    style_chart(wf)
    ws.Range("A1").Select()
    widths(ws, {"A": 2, "B": 22, "C": 26, "D:G": 13, "H:I": 11, "J:N": 13, "O": 11})
    for block in (long, short, ytd):
        ws.Rows(block[0] - 1).RowHeight = 30
    pos["pvm"] = {"long": long, "short": short, "ytd": ytd, "bridge": (b0 + 1, b0 + 5)}
    specs = {"bridge_long": f"PVM!$C${b0 + 1}:$C${b0 + 5}"}
    for name, (f, last, tot) in (("long", long), ("short", short), ("ytd", ytd)):
        specs[f"pvm_{name}"] = f"PVM!$B${f}:$O${last}"
        specs[f"pvm_{name}_total"] = f"PVM!$B${tot}:$O${tot}"
    names(wb, specs)


# --------------------------------------------------------------- Monthly
def monthly(wb, pos: dict) -> None:
    ws = wb.Worksheets("Monthly")
    title(ws, "Monthly retail sales: EIA-861M, and how it ties to the FERC filing",
          '="PGE\'s monthly returns to EIA, "&TEXT(FitStart,"mmm yyyy")&" to "&TEXT(AsOfMonth,"mmm yyyy")&". EIA marks '
          'the latest months Preliminary and revises them."', span="Q")
    cls = [(L.CLASS_LABELS[c], c) for c in L.CLASSES]
    band(ws, 4, '="Year to date by class: "&TEXT(FYStart,"mmm")&"–"&TEXT(AsOfMonth,"mmm yyyy")&" against "'
                '&(FiscalYear-1)', "B", "Q")
    header(ws, 5, ["Class", "Key", "Revenue", "Prior year", "Change", "MWh", "Prior year", "Change", "¢/kWh",
                   "Prior year", "Change", "Customers, latest", "A year earlier", "Change", "MWh per customer",
                   "Prior year"])
    f = 6
    last = f + len(cls) - 1
    tot = last + 1
    column(ws, f"B{f}", [c[0] for c in cls])
    column(ws, f"C{f}", [c[1] for c in cls])
    put(ws, f"B{tot}", "All classes")

    def ytd(field, a, b):
        return (f'=SUMIFS(tbl_Monthly[{field}],tbl_Monthly[class],$C{f},tbl_Monthly[month],">="&{a},'
                f'tbl_Monthly[month],"<="&{b})')

    fx(ws, f"D{f}:D{last}", ytd("revenue", "FYStart", "AsOfMonth"))
    fx(ws, f"E{f}:E{last}", ytd("revenue", "EDATE(FYStart,-12)", "EDATE(AsOfMonth,-12)"))
    fx(ws, f"G{f}:G{last}", ytd("mwh", "FYStart", "AsOfMonth"))
    fx(ws, f"H{f}:H{last}", ytd("mwh", "EDATE(FYStart,-12)", "EDATE(AsOfMonth,-12)"))
    fx(ws, f"M{f}:M{last}", f"=SUMIFS(tbl_Monthly[customers],tbl_Monthly[class],$C{f},tbl_Monthly[month],AsOfMonth)")
    fx(ws, f"N{f}:N{last}", f"=SUMIFS(tbl_Monthly[customers],tbl_Monthly[class],$C{f},tbl_Monthly[month],"
                            "EDATE(AsOfMonth,-12))")
    for c in "DEGHMN":
        fx(ws, f"{c}{tot}", f"=SUM({c}{f}:{c}{last})")
    fx(ws, f"F{f}:F{tot}", f"=SAFEDIV(D{f}-E{f},E{f})")
    fx(ws, f"I{f}:I{tot}", f"=SAFEDIV(G{f}-H{f},H{f})")
    fx(ws, f"J{f}:J{tot}", f"=CPK(D{f},G{f})")
    fx(ws, f"K{f}:K{tot}", f"=CPK(E{f},H{f})")
    fx(ws, f"L{f}:L{tot}", f"=SAFEDIV(J{f}-K{f},K{f})")
    fx(ws, f"O{f}:O{tot}", f"=SAFEDIV(M{f}-N{f},N{f})")
    fx(ws, f"P{f}:P{tot}", f"=SAFEDIV(G{f},M{f})")
    fx(ws, f"Q{f}:Q{tot}", f"=SAFEDIV(H{f},N{f})")
    for rng, fmt in ((f"D{f}:E{tot}", F_MONEY), (f"G{f}:H{tot}", F_MWH), (f"M{f}:N{tot}", F_CASES),
                     (f"J{f}:K{tot}", F_CPK), (f"P{f}:Q{tot}", "0.00")):
        numfmt(ws.Range(rng), fmt)
    for c in "FILO":
        numfmt(ws.Range(f"{c}{f}:{c}{tot}"), F_VARPCT)
        sign_colours(ws.Range(f"{c}{f}:{c}{tot}"))
    font(ws.Range(f"C{f}:C{last}"), color=MUTED, size=8)
    total_row(ws.Range(f"B{tot}:Q{tot}"), double=True)

    # revenue grid: year x month
    g0 = tot + 3
    band(ws, g0, "Retail revenue by month, all classes ($M)", "B", "Q")
    put(ws, f"B{g0 + 1}", "Year")
    fx(ws, f"C{g0 + 1}:N{g0 + 1}", '=TEXT(DATE(2000,COLUMN()-2,1),"mmm")')
    put(ws, f"O{g0 + 1}", "Year")
    font(ws.Range(f"B{g0 + 1}:O{g0 + 1}"), bold=True)
    ws.Range(f"C{g0 + 1}:O{g0 + 1}").HorizontalAlignment = -4152
    g1 = g0 + 2
    g2 = g1 + 9
    fx(ws, f"B{g1}:B{g2}", f"=YEAR(FitStart)+ROWS($B${g1}:B{g1})-1")
    fx(ws, f"C{g1}:N{g2}", f"=SUMIFS(tbl_Monthly[revenue],tbl_Monthly[month],DATE($B{g1},COLUMN()-2,1))/1000000")
    fx(ws, f"O{g1}:O{g2}", f"=SUM(C{g1}:N{g1})")
    numfmt(ws.Range(f"C{g1}:O{g2}"), '0.0;(0.0);""')
    numfmt(ws.Range(f"B{g1}:B{g2}"), "0")
    scale = ws.Range(f"C{g1}:N{g2}").FormatConditions.AddColorScale(2)
    scale.ColorScaleCriteria(1).FormatColor.Color = rgb("#FFFFFF")
    scale.ColorScaleCriteria(2).FormatColor.Color = rgb("#9DB4CF")
    font(ws.Range(f"O{g1}:O{g2}"), bold=True)

    # EIA against FERC
    t0 = g2 + 3
    band(ws, t0, "Two regulators, one utility: EIA's monthly survey summed to a year, against the FERC filing", "B", "Q")
    header(ws, t0 + 1, ["Year", "EIA MWh", "FERC MWh", "Difference", "EIA revenue", "FERC retail revenue", "Gap",
                        "Gap %"])
    t1 = t0 + 2
    t2 = t1 + 8
    fx(ws, f"B{t1}", "=YEAR(FitStart)")
    fx(ws, f"B{t1 + 1}:B{t2}", f"=B{t1}+1")
    window = f'tbl_Monthly[month],">="&DATE($B{t1},1,1),tbl_Monthly[month],"<="&DATE($B{t1},12,1)'
    fx(ws, f"C{t1}:C{t2}", f"=SUMIFS(tbl_Monthly[mwh],{window})")
    fx(ws, f"D{t1}:D{t2}", "=" + ferc_sum(FR, "sales_mwh", '"sales_to_ultimate_consumers"', year=f"$B{t1}"))
    fx(ws, f"E{t1}:E{t2}", f"=C{t1}-D{t1}")
    fx(ws, f"F{t1}:F{t2}", f"=SUMIFS(tbl_Monthly[revenue],{window})")
    fx(ws, f"G{t1}:G{t2}", "=" + ferc_sum(FR, "dollar_value", '"sales_to_ultimate_consumers"', year=f"$B{t1}"))
    fx(ws, f"H{t1}:H{t2}", f"=F{t1}-G{t1}")
    fx(ws, f"I{t1}:I{t2}", f"=SAFEDIV(H{t1},G{t1})")
    numfmt(ws.Range(f"B{t1}:B{t2}"), "0")
    numfmt(ws.Range(f"C{t1}:E{t2}"), F_MWH)
    numfmt(ws.Range(f"F{t1}:H{t2}"), F_MONEY)
    numfmt(ws.Range(f"I{t1}:I{t2}"), F_VARPCT)
    put(ws, f"B{t2 + 1}", "Volumes agree to within a few MWh a year. EIA's revenue runs 1–2.5% under FERC's: the monthly "
                          "survey reports billed sales, the annual filing books accounting revenue (unbilled accruals, "
                          "deferral amortization and other adjustments).")
    font(ws.Range(f"B{t2 + 1}"), color=MUTED, italic=True)
    widths(ws, {"A": 2, "B": 16, "C": 12, "D:Q": 12})
    ws.Rows(5).RowHeight = 30
    pos["monthly"] = {"ytd": (f, last, tot), "grid": (g1, g2), "tieout": (t1, t2)}
    names(wb, {"monthly_ytd": f"Monthly!$D${f}:$Q${last}", "monthly_ytd_total": f"Monthly!$D${tot}:$Q${tot}",
               "revenue_grid": f"Monthly!$C${g1}:$N${g2}", "tieout": f"Monthly!$B${t1}:$I${t2}"})


# --------------------------------------------------------------- Weather
def weather(wb, pos: dict) -> None:
    ws = wb.Worksheets("Weather")
    title(ws, "Weather: how much of the load is the thermostat",
          "Monthly MWh per customer regressed with LINEST on heating and cooling degree days (Willamette Valley, base "
          "65°F) and a time trend, fitted from the fit start to the December before the fiscal year.", span="R")
    # --- the regression's data: one row per month
    h0 = 60
    band(ws, h0, "Regression data: one row per month", "B", "R")
    header(ws, h0 + 1, ["Month", "Trend (years)", "HDD", "CDD", "Residential MWh/customer", "Commercial MWh/customer",
                        "Residential fitted", "Commercial fitted", "In fit window"])
    h1 = h0 + 2
    h2 = h1 + 12 * 10 - 1
    fx(ws, f"B{h1}", "=FitStart")
    fx(ws, f"B{h1 + 1}:B{h2}", f"=EDATE(B{h1},1)")
    fx(ws, f"C{h1}:C{h2}", f"=TRENDYEARS(B{h1})")
    fx(ws, f"D{h1}:D{h2}", f'=XLOOKUP(B{h1},tbl_Weather[month],tbl_Weather[hdd],"")')
    fx(ws, f"E{h1}:E{h2}", f'=XLOOKUP(B{h1},tbl_Weather[month],tbl_Weather[cdd],"")')
    for c, cls in zip("FG", L.WEATHER_CLASSES):
        fx(ws, f"{c}{h1}:{c}{h2}",
           f'=IF(ISCLOSED(B{h1}),SAFEDIV(SUMIFS(tbl_Monthly[mwh],tbl_Monthly[class],"{cls}",tbl_Monthly[month],B{h1}),'
           f'SUMIFS(tbl_Monthly[customers],tbl_Monthly[class],"{cls}",tbl_Monthly[month],B{h1})),"")')
    fx(ws, f"J{h1}:J{h2}", f"=AND(ISCLOSED(B{h1}),YEAR(B{h1})<FiscalYear)")
    numfmt(ws.Range(f"B{h1}:B{h2}"), F_MONTH)
    numfmt(ws.Range(f"C{h1}:C{h2}"), "0.000")
    numfmt(ws.Range(f"D{h1}:E{h2}"), "0")
    numfmt(ws.Range(f"F{h1}:I{h2}"), "0.0000")

    # --- coefficients
    band(ws, 4, '="Weather model, fitted "&TEXT(FitStart,"mmm yyyy")&" to Dec "&(FiscalYear-1)', "B", "I")
    header(ws, 5, ["Class", "Intercept", "Per HDD", "Per CDD", "Trend per year", "R²", "Months", "Key"])

    def rng(c):
        return f"FILTER({c}${h1}:{c}${h2},$J${h1}:$J${h2})"

    for i, (c, cls) in enumerate(zip("FG", L.WEATHER_CLASSES)):
        r = 6 + i
        put(ws, f"B{r}", L.CLASS_LABELS[cls])
        put(ws, f"I{r}", cls)
        fx(ws, f"C{r}", f"=LET(s,LINEST({rng(c)},HSTACK({rng('D')},{rng('E')},{rng('C')}),TRUE,TRUE),"
                        "HSTACK(INDEX(s,1,4),INDEX(s,1,3),INDEX(s,1,2),INDEX(s,1,1),INDEX(s,3,1)))")
        fx(ws, f"H{r}", f"=ROWS({rng(c)})")
    numfmt(ws.Range("C6:F7"), "0.000000")
    numfmt(ws.Range("G6:G7"), "0.000")
    font(ws.Range("I6:I7"), color=MUTED, size=8)
    put(ws, "B8", "MWh per customer per month: per degree day, and per year of trend. LINEST returns the coefficients "
                  "last variable first; the formula puts them back in reading order.")
    font(ws.Range("B8"), color=MUTED, italic=True)
    for c, r in zip("HI", (6, 7)):
        fx(ws, f"{c}{h1}:{c}{h2}", f'=IF(ISNUMBER(D{h1}),$C${r}+$D${r}*D{h1}+$E${r}*E{h1}+$F${r}*C{h1},"")')

    # --- normals
    band(ws, 10, '="Normal weather: the average of "&(FiscalYear-NormalYears)&"–"&(FiscalYear-1)&", by month"', "B", "O")
    put(ws, "B11", "Month")
    fx(ws, "C11:N11", '=TEXT(DATE(2000,COLUMN()-2,1),"mmm")')
    put(ws, "O11", "Year")
    for r, label in ((12, "Normal HDD"), (13, "Normal CDD"), (14, '="Actual HDD, "&FiscalYear'),
                     (15, '="Actual CDD, "&FiscalYear')):
        put(ws, f"B{r}", label)
    for r, field in ((12, "hdd"), (13, "cdd")):
        fx(ws, f"C{r}:N{r}", f"=AVERAGE(FILTER(tbl_Weather[{field}],(YEAR(tbl_Weather[month])>=FiscalYear-NormalYears)"
                             f"*(YEAR(tbl_Weather[month])<FiscalYear)*(MONTH(tbl_Weather[month])=COLUMN()-2)))")
    for r, field in ((14, "hdd"), (15, "cdd")):
        fx(ws, f"C{r}:N{r}", f'=XLOOKUP(DATE(FiscalYear,COLUMN()-2,1),tbl_Weather[month],tbl_Weather[{field}],"")')
    fx(ws, "O12:O15", "=SUM(C12:N12)")
    numfmt(ws.Range("C12:O15"), "0")
    font(ws.Range("B11:O11"), bold=True)
    ws.Range("C11:O11").HorizontalAlignment = -4152
    names(wb, {"normal_hdd": "Weather!$C$12:$N$12", "normal_cdd": "Weather!$C$13:$N$13",
               "actual_hdd": "Weather!$C$14:$N$14", "actual_cdd": "Weather!$C$15:$N$15"})

    # --- this year's weather impact
    w0 = 17
    band(ws, w0, '="Weather impact, "&TEXT(FYStart,"mmm")&"–"&TEXT(AsOfMonth,"mmm yyyy")&": actual against normal"',
         "B", "N")
    header(ws, w0 + 1, ["Month", "HDD vs normal", "CDD vs normal", "Residential MWh per customer", "Residential "
                        "customers", "Residential MWh", "Residential $", "Commercial MWh per customer",
                        "Commercial customers", "Commercial MWh", "Commercial $", "Weather MWh", "Weather $"])
    w1 = w0 + 2
    w2 = w1 + 11
    fx(ws, f"B{w1}:B{w2}", f"=DATE(FiscalYear,ROWS($B${w1}:B{w1}),1)")
    fx(ws, f"C{w1}:C{w2}", f'=IF(ISCLOSED(B{w1}),INDEX(actual_hdd,MONTH(B{w1}))-INDEX(normal_hdd,MONTH(B{w1})),"")')
    fx(ws, f"D{w1}:D{w2}", f'=IF(ISCLOSED(B{w1}),INDEX(actual_cdd,MONTH(B{w1}))-INDEX(normal_cdd,MONTH(B{w1})),"")')
    for c0, cls, coef_row in (("E", "residential", 6), ("I", "commercial", 7)):
        c1, c2, c3 = (chr(ord(c0) + j) for j in (1, 2, 3))
        fx(ws, f"{c0}{w1}:{c0}{w2}", f'=IF(ISCLOSED(B{w1}),$D${coef_row}*C{w1}+$E${coef_row}*D{w1},"")')
        fx(ws, f"{c1}{w1}:{c1}{w2}", f'=IF(ISCLOSED(B{w1}),SUMIFS(tbl_Monthly[customers],tbl_Monthly[class],"{cls}",'
                                     f'tbl_Monthly[month],B{w1}),"")')
        fx(ws, f"{c2}{w1}:{c2}{w2}", f'=IF(ISCLOSED(B{w1}),{c0}{w1}*{c1}{w1},"")')
        fx(ws, f"{c3}{w1}:{c3}{w2}", f'=IF(ISCLOSED(B{w1}),{c2}{w1}*SAFEDIV(SUMIFS(tbl_Monthly[revenue],'
                                     f'tbl_Monthly[class],"{cls}",tbl_Monthly[month],B{w1}),SUMIFS(tbl_Monthly[mwh],'
                                     f'tbl_Monthly[class],"{cls}",tbl_Monthly[month],B{w1})),"")')
    fx(ws, f"M{w1}:M{w2}", f'=IF(ISCLOSED(B{w1}),G{w1}+K{w1},"")')
    fx(ws, f"N{w1}:N{w2}", f'=IF(ISCLOSED(B{w1}),H{w1}+L{w1},"")')
    wt = w2 + 1
    put(ws, f"B{wt}", "Year to date")
    for c in "CDGHKLMN":
        fx(ws, f"{c}{wt}", f"=SUM({c}{w1}:{c}{w2})")
    numfmt(ws.Range(f"B{w1}:B{w2}"), F_MONTH)
    numfmt(ws.Range(f"C{w1}:D{wt}"), '+0;-0;0')
    numfmt(ws.Range(f"E{w1}:E{wt}"), '+0.0000;-0.0000;0')
    numfmt(ws.Range(f"I{w1}:I{wt}"), '+0.0000;-0.0000;0')
    numfmt(ws.Range(f"F{w1}:F{wt}"), F_CASES)
    numfmt(ws.Range(f"J{w1}:J{wt}"), F_CASES)
    for c in "GHKLMN":
        numfmt(ws.Range(f"{c}{w1}:{c}{wt}"), F_VAR)
    for c in "HLN":
        sign_colours(ws.Range(f"{c}{w1}:{c}{wt}"))
    total_row(ws.Range(f"B{wt}:N{wt}"), double=True)
    fx(ws, f"B{wt + 1}", f'=IF(N{wt}<0,"Weather took ","Weather added ")&TEXT(ABS(M{wt})/1000,"#,##0")&" GWh and "'
                         f'&MONEY(ABS(N{wt}))&IF(N{wt}<0," off"," to")&" retail revenue this year, valued at each month\'s '
                         'actual price."')
    font(ws.Range(f"B{wt + 1}"), italic=True, color=INK)

    # The chart reads names that stop at the last reported month: a line chart plots
    # the "" of a month not yet reported as zero.
    reported = f"COUNT(Weather!$F${h1}:$F${h2})"
    names(wb, {f"wx_{k}": f"Weather!${c}${h1}:INDEX(Weather!${c}${h1}:${c}${h2},{reported})"
               for k, c in (("months", "B"), ("actual", "F"), ("fitted", "H"))})
    c0 = wt + 3
    ch = new_chart(ws, XL_LINE, f"B{c0}", f"N{c0 + 19}", [("Actual", f"F{h1}:F{h2}"), ("Weather model", f"H{h1}:H{h2}")],
                   f"B{h1}:B{h2}", "Residential MWh per customer: actual against the weather model")
    book = wb.Name.replace("'", "''")
    for i, key in ((1, "actual"), (2, "fitted")):
        ch.FullSeriesCollection(i).Values = f"='{book}'!wx_{key}"
        ch.FullSeriesCollection(i).XValues = f"='{book}'!wx_months"
    ch.FullSeriesCollection(1).Format.Line.ForeColor.RGB = rgb(NAVY)
    ch.FullSeriesCollection(2).Format.Line.ForeColor.RGB = rgb(TEAL)
    ch.FullSeriesCollection(2).Format.Line.DashStyle = 4
    ch.Axes(1).TickLabels.NumberFormat = "yyyy"
    ch.Axes(1).CategoryType = 2        # text axis, so each year is labelled once
    ch.Axes(1).TickLabelSpacing = 12
    ch.Axes(1).TickMarkSpacing = 12
    ch.Axes(1).TickLabels.Orientation = 0
    ch.HasLegend = True
    ch.Legend.Position = -4107
    style_chart(ch)

    widths(ws, {"A": 2, "B": 18, "C:N": 11, "O:R": 11})
    ws.Rows(18).RowHeight = 42
    pos["weather"] = {"history": (h1, h2), "impact": (w1, w2, wt), "end": c0 + 19}
    names(wb, {"coef_residential": "Weather!$C$6:$G$6", "coef_commercial": "Weather!$C$7:$G$7",
               "fit_months": "Weather!$H$6:$H$7", "weather_impact": f"Weather!$B${w1}:$N${w2}",
               "weather_impact_total": f"Weather!$C${wt}:$N${wt}", "weather_history": f"Weather!$B${h1}:$J${h2}"})
