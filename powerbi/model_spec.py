"""
What the semantic model contains: tables, relationships and measures.

Kept apart from the writer in :mod:`powerbi.build_pbip` so the *shape* of the
model can be read, reviewed and asserted on without wading through TMDL.

Nothing here repeats an analysis. The tables in ``tables/`` are written by
``model/export_tables.py`` from the same reference model the Excel workbook is
held to, and every measure below either sums one of their columns or divides
two such sums. The weather regression, the plan, the bridge and the
price-volume-mix are computed once; a DAX re-derivation of any of them would
be a third implementation of a definition that already has two, tested
against each other.
"""

from __future__ import annotations

# --------------------------------------------------------------------------
# Tables. `source` is the directory the CSV lives in, relative to DataPath.
# --------------------------------------------------------------------------

TABLES: dict[str, dict] = {
    # Dimensions. The month dimension is deliberately NOT marked as a date
    # table: Power BI needs a contiguous *daily* column for that, and every
    # fact here is monthly. Comparisons shift `month_index`, which is
    # contiguous by construction.
    "dim_month": {"source": "tables", "kind": "dimension"},
    "dim_class": {"source": "tables", "kind": "dimension"},

    # EIA-861M monthly sales and NOAA degree days
    "fact_sales": {"source": "tables", "kind": "fact"},
    "fact_weather": {"source": "tables", "kind": "fact"},

    # The reference model's analysis
    "weather_impact": {"source": "tables", "kind": "analysis"},
    "weather_model": {"source": "tables", "kind": "analysis"},
    "plan_monthly": {"source": "tables", "kind": "analysis"},
    "plan_bridge": {"source": "tables", "kind": "analysis"},
    "forecast_monthly": {"source": "tables", "kind": "analysis"},
    "backtest": {"source": "tables", "kind": "analysis"},
    "tornado": {"source": "tables", "kind": "analysis"},
    "sensitivity": {"source": "tables", "kind": "analysis"},
    "pvm": {"source": "tables", "kind": "analysis"},

    # FERC Form 1, via PUDL
    "ferc_pnl": {"source": "tables", "kind": "analysis"},
    "ferc_years": {"source": "tables", "kind": "analysis"},
    "peers": {"source": "tables", "kind": "analysis"},
    "tieout": {"source": "tables", "kind": "analysis"},
}

# Columns whose order carries meaning, and the column that orders them. A text
# column on an axis otherwise sorts alphabetically: Apr before Jan, and a
# bridge that opens on "Customers" instead of the plan.
SORT_BY: dict[str, dict[str, str]] = {
    "dim_month": {"month_name": "month_number", "month_label": "month_index"},
    "dim_class": {"class_label": "class_order"},
    "plan_bridge": {"step": "step_order"},
    "tornado": {"lever": "lever_order"},
    "sensitivity": {"rate_change": "rate_order", "extra_mw": "mw_order"},
    "pvm": {"step": "step_order", "period": "period_order"},
    "ferc_pnl": {"line": "line_order"},
    "ferc_years": {"year_label": "year"},
    "tieout": {"year_label": "year"},
}

# --------------------------------------------------------------------------
# Relationships. All single-direction many-to-one.
# --------------------------------------------------------------------------

RELATIONSHIPS: list[tuple[str, str, str, str]] = [
    ("fact_sales", "month", "dim_month", "month"),
    ("fact_sales", "class", "dim_class", "class"),
    ("fact_weather", "month", "dim_month", "month"),
    ("weather_impact", "month", "dim_month", "month"),
    ("weather_impact", "class", "dim_class", "class"),
    ("weather_model", "class", "dim_class", "class"),
    ("plan_monthly", "month", "dim_month", "month"),
    ("plan_monthly", "class", "dim_class", "class"),
    ("plan_bridge", "month", "dim_month", "month"),
    ("plan_bridge", "class", "dim_class", "class"),
    ("forecast_monthly", "month", "dim_month", "month"),
    ("forecast_monthly", "class", "dim_class", "class"),
]

# Tables intentionally left unrelated, with the reason.
UNRELATED: dict[str, str] = {
    "backtest": "One row per plan year: the plan method re-run as of each year-end. A plan "
                "year is not a calendar filter -- slicing it by month would score a "
                "different plan.",
    "tornado": "One row per scenario lever, each a full-year outlook re-run. Slicing it by "
               "class would report a lever nobody moved.",
    "sensitivity": "A rate-by-megawatt grid of full-year outlooks, not an observation of the "
                   "business.",
    "pvm": "Price-volume-mix over FERC's revenue classes, which include street lighting and "
           "public authorities that EIA's classes do not. Its own `period` is the slicer.",
    "ferc_pnl": "Annual, from FERC Form 1. The monthly calendar has no year before 2017 and "
                "the P&L runs from 2014.",
    "ferc_years": "Annual FERC totals, for the same reason as ferc_pnl.",
    "peers": "One row per utility. Only PGE reports to EIA in this model.",
    "tieout": "One row per full year: EIA's monthly survey summed against the FERC filing.",
}

# --------------------------------------------------------------------------
# What-if parameters. (table, column, min, max, step, format, measure)
#
# Calculated tables, deliberately unrelated to everything: a parameter joined to
# a fact would filter the fact to the parameter's value.
# --------------------------------------------------------------------------

WHATIF_PARAMETERS: tuple[tuple[str, str, float, float, float, str, str], ...] = (
    ("RateChange", "Rate change %", -0.02, 0.04, 0.005, "0.0%", "Rate change value"),
    ("LoadChange", "Industrial load change (MW)", -100, 300, 25, "+#,0;-#,0;0", "Load change value"),
)

# Field parameters are deliberately absent: Desktop would not bind a
# hand-authored one (see the pricing repository's model_spec for the history).
FIELD_PARAMETERS: tuple[tuple[str, str, tuple[tuple[str, str], ...]], ...] = ()


def field_parameter_columns() -> dict[str, set[str]]:
    return {table: {column, f"{column} Fields", f"{column} Order"}
            for table, column, _ in FIELD_PARAMETERS}


def whatif_columns() -> dict[str, set[str]]:
    return {table: {column} for table, column, *_ in WHATIF_PARAMETERS}


# --------------------------------------------------------------------------
# Measures. (name, DAX, format string, folder)
#
# Every VAR is prefixed `v`: Power BI reserves far more VAR names than it
# documents, and a measure that trips one fails at runtime, not at load.
# --------------------------------------------------------------------------

_FY = "VAR vFY = [Fiscal year]\n"
_CLOSED = "VAR vClosed = [Months closed]\n"

MEASURES: list[tuple[str, str, str, str]] = [
    # --- The fiscal year the data reaches ----------------------------------
    ("Fiscal year", "YEAR(CALCULATE(MAX(fact_sales[month]), REMOVEFILTERS()))", "0", "00 Calendar"),
    ("Months closed", "MONTH(CALCULATE(MAX(fact_sales[month]), REMOVEFILTERS()))", "0", "00 Calendar"),

    # --- EIA monthly sales -------------------------------------------------
    ("Retail revenue", "SUM(fact_sales[revenue])", "\\$#,0", "01 Sales"),
    ("Retail MWh", "SUM(fact_sales[mwh])", "#,0", "01 Sales"),
    ("Average customers",
     "AVERAGEX(VALUES(dim_month[month]), CALCULATE(SUM(fact_sales[customers])))", "#,0", "01 Sales"),
    ("Retail price (cents/kWh)", "DIVIDE([Retail revenue], [Retail MWh]) / 10", "0.00", "01 Sales"),
    ("MWh per customer", "DIVIDE([Retail MWh], [Average customers])", "0.000", "01 Sales"),
    ("Retail revenue last year",
     "VAR vShift = 12\n"
     "RETURN CALCULATE(\n"
     "    [Retail revenue],\n"
     "    ALL(dim_month),\n"
     "    TREATAS(\n"
     "        SELECTCOLUMNS(VALUES(dim_month[month_index]),\n"
     "                      \"month_index\", dim_month[month_index] - vShift),\n"
     "        dim_month[month_index]\n"
     "    )\n"
     ")",
     "\\$#,0", "01 Sales"),
    ("Revenue YoY %",
     "DIVIDE([Retail revenue] - [Retail revenue last year], [Retail revenue last year])",
     "+0.0%;-0.0%;0.0%", "01 Sales"),

    # --- This fiscal year, whatever the calendar filter --------------------
    ("Actual this year", _FY + "RETURN CALCULATE([Retail revenue], dim_month[year] = vFY)",
     "\\$#,0", "02 Plan"),
    ("Actual last year", _FY + "RETURN CALCULATE([Retail revenue], dim_month[year] = vFY - 1)",
     "\\$#,0", "02 Plan"),
    ("MWh this year", _FY + "RETURN CALCULATE([Retail MWh], dim_month[year] = vFY)", "#,0", "02 Plan"),
    ("MWh last year to date",
     _FY + _CLOSED + "RETURN CALCULATE([Retail MWh], dim_month[year] = vFY - 1, "
                     "dim_month[month_number] <= vClosed)",
     "#,0", "02 Plan"),
    ("Industrial load growth",
     "CALCULATE(DIVIDE([MWh this year], [MWh last year to date]) - 1, dim_class[class] = \"industrial\")",
     "+0.0%;-0.0%;0.0%", "02 Plan"),
    ("Plan revenue", "SUM(plan_monthly[plan_revenue])", "\\$#,0", "02 Plan"),
    ("Plan revenue to date",
     _CLOSED + "RETURN CALCULATE([Plan revenue], dim_month[month_number] <= vClosed)", "\\$#,0", "02 Plan"),
    ("Against plan", "[Actual this year] - [Plan revenue to date]", "\\$#,0", "02 Plan"),
    ("Against plan %", "DIVIDE([Against plan], [Plan revenue to date])", "+0.0%;-0.0%;0.0%", "02 Plan"),
    ("Forecast revenue", "SUM(forecast_monthly[forecast_revenue])", "\\$#,0", "02 Plan"),
    ("Outlook revenue", "[Actual this year] + [Forecast revenue]", "\\$#,0", "02 Plan"),
    ("Outlook against plan", "[Outlook revenue] - [Plan revenue]", "\\$#,0", "02 Plan"),
    ("Outlook growth %", "DIVIDE([Outlook revenue], [Actual last year]) - 1", "+0.0%;-0.0%;0.0%",
     "02 Plan"),

    # --- The plan-to-actual bridge -----------------------------------------
    ("Bridge amount", "SUM(plan_bridge[amount])", "\\$#,0", "03 Bridge"),
    # The four effects without the plan they start from: a $1.5B opening bar
    # would flatten a $100M bridge to a line.
    ("Bridge effect", "CALCULATE([Bridge amount], plan_bridge[step] <> \"Plan\")", "\\$#,0", "03 Bridge"),
    ("Customer effect", "CALCULATE([Bridge amount], plan_bridge[step] = \"Customers\")", "\\$#,0",
     "03 Bridge"),
    ("Weather effect", "CALCULATE([Bridge amount], plan_bridge[step] = \"Weather\")", "\\$#,0",
     "03 Bridge"),
    ("Usage effect", "CALCULATE([Bridge amount], plan_bridge[step] = \"Usage\")", "\\$#,0", "03 Bridge"),
    ("Price effect", "CALCULATE([Bridge amount], plan_bridge[step] = \"Price\")", "\\$#,0", "03 Bridge"),
    ("Backtest variance", "AVERAGE(backtest[variance_pct])", "+0.0%;-0.0%;0.0%", "03 Bridge"),
    ("Revenue WAPE", "AVERAGE(backtest[wape_revenue_pct])", "0.0%", "03 Bridge"),
    ("MWh WAPE", "AVERAGE(backtest[wape_mwh_pct])", "0.0%", "03 Bridge"),

    # --- Weather -----------------------------------------------------------
    ("Heating degree days", "SUM(fact_weather[hdd])", "#,0", "04 Weather"),
    ("Normal heating degree days", "SUM(fact_weather[normal_hdd])", "#,0", "04 Weather"),
    ("Cooling degree days", "SUM(fact_weather[cdd])", "#,0", "04 Weather"),
    ("HDD this year", _FY + "RETURN CALCULATE([Heating degree days], dim_month[year] = vFY)", "#,0",
     "04 Weather"),
    ("Normal HDD this year",
     _FY + "RETURN CALCULATE([Normal heating degree days], dim_month[year] = vFY)", "#,0", "04 Weather"),
    ("HDD against normal %",
     _FY + _CLOSED + "VAR vActual = CALCULATE([Heating degree days], dim_month[year] = vFY, "
                     "dim_month[month_number] <= vClosed)\n"
                     "VAR vNormal = CALCULATE([Normal heating degree days], dim_month[year] = vFY, "
                     "dim_month[month_number] <= vClosed)\n"
                     "RETURN DIVIDE(vActual - vNormal, vNormal)",
     "+0%;-0%;0%", "04 Weather"),
    ("Weather load (GWh)", "SUM(weather_impact[weather_mwh]) / 1000", "+#,0;-#,0;0", "04 Weather"),
    ("Weather revenue", "SUM(weather_impact[weather_revenue])", "\\$#,0", "04 Weather"),
    ("Residential MWh per customer",
     "CALCULATE([MWh per customer], dim_class[class] = \"residential\")", "0.000", "04 Weather"),
    ("Residential R²", "CALCULATE(MAX(weather_model[r_squared]), dim_class[class] = \"residential\")",
     "0.000", "04 Weather"),

    # --- Scenarios ---------------------------------------------------------
    ("Rate change value", "SELECTEDVALUE(RateChange[Rate change %], 0)", "+0.0%;-0.0%;0.0%",
     "05 Scenarios"),
    ("Load change value", "SELECTEDVALUE(LoadChange[Industrial load change (MW)], 0)", "+#,0;-#,0;0",
     "05 Scenarios"),
    # The same arithmetic as reference.forecast, not a new model of it: open-month
    # revenue is MWh x price, a rate change scales the price, and a steady load of
    # M megawatts adds M x 24 x days MWh to industrial at industrial's price.
    ("Scenario outlook",
     "VAR vRate = [Rate change value]\n"
     "VAR vMW = [Load change value]\n"
     "VAR vIndustrial = SUMX(FILTER(forecast_monthly, forecast_monthly[class] = \"industrial\"),\n"
     "    forecast_monthly[days] * 24 * forecast_monthly[forecast_price])\n"
     "RETURN [Actual this year] + ([Forecast revenue] + vMW * vIndustrial) * (1 + vRate)",
     "\\$#,0", "05 Scenarios"),
    ("Scenario against plan", "[Scenario outlook] - [Plan revenue]", "\\$#,0", "05 Scenarios"),
    ("Tornado impact", "SUM(tornado[impact])", "\\$#,0", "05 Scenarios"),
    ("Sensitivity outlook", "SUM(sensitivity[outlook_revenue])", "\\$#,0", "05 Scenarios"),

    # --- Price, volume and mix ---------------------------------------------
    # One period at a time: summed across periods the opening bar would add three
    # different years' revenue. With no period chosen, the long view.
    ("PVM amount",
     "VAR vOrder = CALCULATE(MIN(pvm[period_order]), REMOVEFILTERS(pvm[revenue_class], pvm[step]))\n"
     "RETURN CALCULATE(SUM(pvm[amount]), KEEPFILTERS(pvm[period_order] = vOrder))",
     "\\$#,0", "06 PVM"),
    ("PVM opening revenue", "CALCULATE([PVM amount], pvm[step] = \"Opening revenue\")", "\\$#,0", "06 PVM"),
    ("PVM volume", "CALCULATE([PVM amount], pvm[step] = \"Volume\")", "\\$#,0", "06 PVM"),
    ("PVM mix", "CALCULATE([PVM amount], pvm[step] = \"Mix\")", "\\$#,0", "06 PVM"),
    ("PVM price", "CALCULATE([PVM amount], pvm[step] = \"Price\")", "\\$#,0", "06 PVM"),

    # --- FERC Form 1 -------------------------------------------------------
    ("P&L ($M)", "DIVIDE(SUM(ferc_pnl[value]), 1000000)", "#,0.0", "07 FERC"),
    ("Operating margin %",
     "DIVIDE(SUM(ferc_years[operating_income]), SUM(ferc_years[total_revenue]))", "0.0%", "07 FERC"),
    ("FERC retail price (cents/kWh)",
     "DIVIDE(SUM(ferc_years[retail_revenue]), SUM(ferc_years[retail_mwh])) / 10", "0.00", "07 FERC"),
    ("Power cost per MWh", "DIVIDE(SUM(ferc_years[power_cost]), SUM(ferc_years[total_mwh]))",
     "\\$#,0.00", "07 FERC"),
    ("Operating margin, latest year",
     "VAR vYear = CALCULATE(MAX(ferc_years[year]), REMOVEFILTERS(ferc_years))\n"
     "RETURN CALCULATE([Operating margin %], ferc_years[year] = vYear)", "0.0%", "07 FERC"),
    ("Retail price, latest year (cents/kWh)",
     "VAR vYear = CALCULATE(MAX(ferc_years[year]), REMOVEFILTERS(ferc_years))\n"
     "RETURN CALCULATE([FERC retail price (cents/kWh)], ferc_years[year] = vYear)", "0.00", "07 FERC"),
    ("Power cost per MWh, latest year",
     "VAR vYear = CALCULATE(MAX(ferc_years[year]), REMOVEFILTERS(ferc_years))\n"
     "RETURN CALCULATE([Power cost per MWh], ferc_years[year] = vYear)", "\\$#,0.00", "07 FERC"),
    ("PGE industrial load growth",
     "CALCULATE(AVERAGE(peers[industrial_mwh_cagr_pct]), REMOVEFILTERS(peers), peers[is_pge] = TRUE())",
     "0.0%", "07 FERC"),
    ("Industrial load growth a year", "AVERAGE(peers[industrial_mwh_cagr_pct])", "0.0%", "07 FERC"),
    ("Peer ¢/kWh", "AVERAGE(peers[retail_cents])", "0.00", "07 FERC"),
    ("Peer margin", "AVERAGE(peers[operating_margin_pct])", "0.0%", "07 FERC"),
    ("Peer power $/MWh", "AVERAGE(peers[power_cost_per_mwh])", "\\$#,0.00", "07 FERC"),

    # --- EIA against FERC --------------------------------------------------
    ("EIA MWh", "SUM(tieout[eia_mwh])", "#,0", "08 Reconciliation"),
    ("EIA revenue", "SUM(tieout[eia_revenue])", "\\$#,0", "08 Reconciliation"),
    ("FERC revenue", "SUM(tieout[ferc_revenue])", "\\$#,0", "08 Reconciliation"),
    ("FERC MWh", "SUM(tieout[ferc_mwh])", "#,0", "08 Reconciliation"),
    # Rounded to the shown precision, so a gap of one MWh in 17 million reads
    # 0.000%, not -0.000%.
    ("MWh gap %", "ROUND(DIVIDE(SUM(tieout[eia_mwh]) - SUM(tieout[ferc_mwh]), SUM(tieout[ferc_mwh])), 5)",
     "+0.000%;-0.000%;0.000%", "08 Reconciliation"),
    ("Revenue gap %",
     "DIVIDE(SUM(tieout[eia_revenue]) - SUM(tieout[ferc_revenue]), SUM(tieout[ferc_revenue]))",
     "+0.000%;-0.000%;0.000%", "08 Reconciliation"),

    # --- Tile captions -----------------------------------------------------
    ("Outlook reference",
     "\"plan \" & FORMAT([Plan revenue] / 1000000000, \"$0.00\") & \"B · \" "
     "& FORMAT([Outlook growth %], \"+0.0%;-0.0%\") & \" on last year\"",
     "", "09 Captions"),
    ("Against plan reference",
     _FY + _CLOSED + "RETURN FORMAT([Against plan %], \"+0.0%;-0.0%\") & \" · Jan–\" "
                     "& FORMAT(DATE(vFY, vClosed, 1), \"mmm yyyy\")",
     "", "09 Captions"),
    ("Weather reference",
     "FORMAT([Weather revenue] / 1000000, \"$#,0.0;-$#,0.0\") & \"M · HDD \" "
     "& FORMAT([HDD against normal %], \"+0%;-0%\") & \" vs normal\"",
     "", "09 Captions"),
    ("Industrial reference", "\"MWh against the same months last year\"", "", "09 Captions"),
    ("Scenario reference",
     "FORMAT([Scenario against plan] / 1000000, \"+$#,0.0;-$#,0.0\") & \"M against plan\"",
     "", "09 Captions"),
]


# Power BI has no thousands-scaling in a format string (Excel's `#,0,,` prints
# "$3,007,664,836.8M" there, and axes read "$0MM"), and a full-dollar figure is
# too wide for a table column. Charts auto-scale their axes, so the base measures
# stay in dollars; tables read these, in $ millions.
TABLE_MILLIONS = {  # base measure -> the short name a table column wears
    "Actual this year": "Actual ($M)",
    "Forecast revenue": "Forecast ($M)",
    "Outlook revenue": "Outlook ($M)",
    "Plan revenue": "Full-year plan ($M)",
    "Outlook against plan": "Outlook vs plan ($M)",
    "Plan revenue to date": "Plan to date ($M)",
    "Customer effect": "Customers ($M)",
    "Weather effect": "Weather ($M)",
    "Usage effect": "Usage ($M)",
    "Price effect": "Price ($M)",
    "PVM amount": "Amount ($M)",
    "Sensitivity outlook": "Full-year revenue ($M)",
    "EIA revenue": "EIA revenue ($M)",
    "FERC revenue": "FERC revenue ($M)",
}
MEASURES += [(label, f"DIVIDE([{name}], 1000000)", "\\$#,0.0", "10 Tables ($M)")
             for name, label in TABLE_MILLIONS.items()]
