"""
What the report contains: seven pages, and every visual on them.

The same questions the Excel workbook answers, asked of the same numbers: the
tables come from the reference model the workbook is held to, so a figure on a
page here and a cell there are one definition, not two that happen to agree.

Visual shorthand::

    card(measure, subtitle=measure)
    bar / column / stacked_column / line / scatter / table / matrix / slicer / waterfall

Fields are written ``table[column]`` for a column and ``[Measure]`` for a
measure, the notation Desktop shows in the field well.
"""

from __future__ import annotations

# Canvas is 1280x720 at FitToPage. report_chrome adds the header, the page
# buttons and the filter panel, and reflows what is written here around them.
CARD_Y = 20
CARD_H = 118
ROW1_Y = 152
ROW2_Y = 442
CARDS = ((20, 300), (330, 296), (634, 300), (942, 318))


def cards(*specs: tuple) -> list[dict]:
    """Four KPI tiles across the top: (measure, subtitle or None, alt text)."""
    out = []
    for (x, width), (field, subtitle, alt) in zip(CARDS, specs):
        spec = {"type": "card", "field": field, "pos": (x, CARD_Y, width, CARD_H), "alt": alt}
        if subtitle:
            spec["subtitle"] = subtitle
        out.append(spec)
    return out


CLASS_SLICER = {"type": "slicer", "field": "dim_class[class_label]", "title": "Customer class",
                "pos": (1068, ROW2_Y, 192, 76),
                "alt": "Slicer. Filters the page by customer class label."}

PAGES: list[dict] = [
    # ----------------------------------------------------------------- 1
    {
        "name": "section_overview",
        "display": "Overview",
        "visuals": [
            *cards(
                ("[Outlook revenue]", "[Outlook reference]",
                 "Card. Outlook revenue: this year's actual months plus the forecast for the rest."),
                ("[Against plan]", "[Against plan reference]",
                 "Card. Against plan: actual revenue to date minus the plan for the same months."),
                ("[Weather load (GWh)]", "[Weather reference]",
                 "Card. Weather load (GWh): the load the weather added or took away against normal."),
                ("[Industrial load growth]", "[Industrial reference]",
                 "Card. Industrial load growth against the same months last year."),
            ),
            {"type": "line", "x": "dim_month[month_name]",
             "y": ["[Actual this year]", "[Forecast revenue]", "[Plan revenue]", "[Actual last year]"],
             "sort": ("dim_month[month_name]", "Ascending"),
             "title": "Retail revenue by month: actual, forecast, plan and last year",
             "pos": (20, ROW1_Y, 760, 272),
             "alt": "Line chart titled Retail revenue by month. Plots Actual this year, Forecast "
                    "revenue, Plan revenue and Actual last year by month name."},
            {"type": "waterfall", "x": "plan_bridge[step]", "y": ["[Bridge effect]"],
             "sort": ("plan_bridge[step]", "Ascending"),
             "title": "Actual against plan, year to date: what moved it",
             "pos": (794, ROW1_Y, 466, 272),
             "alt": "Waterfall chart titled Actual against plan, what moved it. Plots Bridge effect "
                    "by step: the customer, weather, usage and price effects; the total bar is "
                    "actual minus plan."},
            {"type": "bar", "x": "dim_class[class_label]", "y": ["[Against plan]"],
             "sort": ("dim_class[class_label]", "Ascending"),
             "title": "Against plan by customer class, year to date",
             "pos": (20, ROW2_Y, 620, 258),
             "alt": "Bar chart titled Against plan by customer class. Plots Against plan by class "
                    "label for each customer class."},
            {"type": "table",
             "columns": ["dim_class[class_label]", "[Actual ($M)]", "[Forecast ($M)]",
                         "[Outlook ($M)]", "[Outlook vs plan ($M)]"],
             "title": "Full-year outlook by class",
             "pos": (654, ROW2_Y, 410, 258),
             "alt": "Table titled Full-year outlook by class. Lists class label, actual, forecast, "
                    "outlook and outlook against plan, in $ millions."},
            dict(CLASS_SLICER),
        ],
    },
    # ----------------------------------------------------------------- 2
    {
        "name": "section_plan",
        "display": "Plan and bridge",
        "visuals": [
            *cards(
                ("[Plan revenue to date]", None,
                 "Card. Plan revenue to date: the plan built from last year's data, for the closed "
                 "months."),
                ("[Actual this year]", "[Against plan reference]",
                 "Card. Actual this year: retail revenue for the closed months."),
                ("[Weather effect]", None,
                 "Card. Weather effect: what weather against normal did to revenue, at plan prices."),
                ("[Price effect]", None,
                 "Card. Price effect: actual MWh times the difference between actual and plan price."),
            ),
            {"type": "stacked_column", "x": "dim_month[month_name]",
             "y": ["[Customer effect]", "[Weather effect]", "[Usage effect]", "[Price effect]"],
             "sort": ("dim_month[month_name]", "Ascending"),
             "title": "What moved revenue against plan, month by month",
             "pos": (20, ROW1_Y, 620, 272),
             "alt": "Stacked column chart titled What moved revenue against plan. Plots Customer "
                    "effect, Weather effect, Usage effect and Price effect by month name."},
            {"type": "table",
             "columns": ["dim_class[class_label]", "[Plan to date ($M)]", "[Customers ($M)]",
                         "[Weather ($M)]", "[Usage ($M)]", "[Price ($M)]", "[Actual ($M)]"],
             "title": "The bridge by class: plan, four effects, actual",
             "pos": (654, ROW1_Y, 606, 272),
             "alt": "Table titled The bridge by class. Lists class label, plan revenue to date, "
                    "customer, weather, usage and price effect, and actual this year."},
            {"type": "column", "x": "backtest[plan_year]", "y": ["[Backtest variance]"],
             "sort": ("backtest[plan_year]", "Ascending"),
             "title": "The same plan method, re-run as of each year-end: actual against plan",
             "pos": (20, ROW2_Y, 620, 258),
             "alt": "Column chart titled The same plan method re-run as of each year-end. Plots "
                    "Backtest variance by plan year."},
            {"type": "table",
             "columns": ["backtest[plan_year]", "backtest[months]", "[Backtest variance]",
                         "[Revenue WAPE]", "[MWh WAPE]"],
             "title": "Backtest: load is forecastable, price is not",
             "totals": False,
             "pos": (654, ROW2_Y, 410, 258),
             "alt": "Table titled Backtest. Lists plan year, months, backtest variance, revenue "
                    "WAPE and MWh WAPE."},
            dict(CLASS_SLICER),
        ],
    },
    # ----------------------------------------------------------------- 3
    {
        "name": "section_weather",
        "display": "Weather",
        "visuals": [
            *cards(
                ("[HDD against normal %]", None,
                 "Card. HDD against normal: this year's heating degree days against the ten-year "
                 "normal."),
                ("[Weather load (GWh)]", None,
                 "Card. Weather load (GWh): the load weather moved against normal."),
                ("[Weather revenue]", None,
                 "Card. Weather revenue: that load valued at each month's actual price."),
                ("[Residential R²]", None,
                 "Card. Residential R squared: how much of residential use per customer the "
                 "weather model explains."),
            ),
            {"type": "line", "x": "dim_month[month_name]", "y": ["[HDD this year]", "[Normal HDD this year]"],
             "sort": ("dim_month[month_name]", "Ascending"),
             "title": "Heating degree days, Willamette Valley: this year against normal",
             "pos": (20, ROW1_Y, 620, 272),
             "alt": "Line chart titled Heating degree days this year against normal. Plots HDD "
                    "this year and Normal HDD this year by month name."},
            {"type": "column", "x": "dim_month[month_name]", "y": ["[Weather revenue]"],
             "sort": ("dim_month[month_name]", "Ascending"),
             "title": "Revenue the weather moved, by month",
             "pos": (654, ROW1_Y, 606, 272),
             "alt": "Column chart titled Revenue the weather moved. Plots Weather revenue by month "
                    "name."},
            {"type": "scatter", "x": "[Heating degree days]", "y": ["[Residential MWh per customer]"],
             "category": "dim_month[month_label]",
             "title": "Residential use per customer against heating degree days, every month",
             "pos": (20, ROW2_Y, 620, 258),
             "alt": "Scatter chart titled Residential use per customer against heating degree "
                    "days. Plots Residential MWh per customer against Heating degree days for each "
                    "month label."},
            {"type": "table",
             "totals": False,
             "columns": ["dim_class[class_label]", "weather_model[intercept]", "weather_model[per_hdd]",
                         "weather_model[per_cdd]", "weather_model[trend_per_year]",
                         "weather_model[r_squared]", "weather_model[months_fitted]"],
             "title": "The regression: MWh per customer on degree days and trend",
             "pos": (654, ROW2_Y, 410, 258),
             "alt": "Table titled The regression. Lists class label, intercept, per HDD, per CDD, trend "
                    "per year, R squared and months fitted."},
            {"type": "slicer", "field": "dim_month[year]", "title": "Year",
             "pos": (1068, ROW2_Y, 192, 76),
             "alt": "Slicer. Filters the scatter chart's months by year."},
        ],
    },
    # ----------------------------------------------------------------- 4
    {
        "name": "section_pvm",
        "display": "Price, volume and mix",
        "visuals": [
            *cards(
                ("[PVM opening revenue]", None,
                 "Card. PVM opening revenue: retail revenue at the start of the selected period."),
                ("[PVM volume]", None, "Card. PVM volume: the change in MWh at the opening average "
                                       "price."),
                ("[PVM mix]", None, "Card. PVM mix: the change in MWh at each class's distance from "
                                    "the average price."),
                ("[PVM price]", None, "Card. PVM price: the change in price on the closing MWh."),
            ),
            {"type": "waterfall", "x": "pvm[step]", "y": ["[PVM amount]"],
             "sort": ("pvm[step]", "Ascending"),
             "title": "Retail revenue: opening, then volume, mix and price",
             "pos": (20, ROW1_Y, 760, 272),
             "alt": "Waterfall chart titled Retail revenue opening then volume mix and price. Plots "
                    "PVM amount by step; the total bar is the closing revenue."},
            {"type": "matrix", "rows": "pvm[revenue_class]", "columns_by": "pvm[step]",
             "values": ["[Amount ($M)]"],
             "title": "By revenue class",
             "pos": (794, ROW1_Y, 466, 272),
             "alt": "Matrix titled By revenue class. Shows PVM amount for each revenue class by "
                    "step."},
            {"type": "line", "x": "dim_month[year]", "y": ["[Retail price (cents/kWh)]"],
             "series": "dim_class[class_label]",
             "title": "Retail price by class, cents/kWh (EIA)",
             "pos": (20, ROW2_Y, 620, 258),
             "alt": "Line chart titled Retail price by class. Plots Retail price in cents per kWh "
                    "by year for each class label."},
            {"type": "stacked_column", "x": "dim_month[year]", "y": ["[Retail MWh]"],
             "series": "dim_class[class_label]",
             "title": "Retail MWh by class (EIA; this year to date)",
             "pos": (654, ROW2_Y, 410, 258),
             "alt": "Stacked column chart titled Retail MWh by class. Plots Retail MWh by year for "
                    "each class label."},
            {"type": "slicer", "field": "pvm[period]", "title": "Period",
             "pos": (1068, ROW2_Y, 192, 76),
             "alt": "Slicer. Chooses the period the price-volume-mix bridge covers; with none "
                    "chosen the page shows the long view."},
        ],
    },
    # ----------------------------------------------------------------- 5
    {
        "name": "section_scenarios",
        "display": "Forecast and scenarios",
        "visuals": [
            *cards(
                ("[Outlook revenue]", "[Outlook reference]",
                 "Card. Outlook revenue on normal weather for the open months."),
                ("[Scenario outlook]", "[Scenario reference]",
                 "Card. Scenario outlook: the outlook with the rate change and industrial load change below."),
                ("[Forecast revenue]", None,
                 "Card. Forecast revenue for the open months."),
                ("[Outlook against plan]", None,
                 "Card. Outlook against plan for the full year."),
            ),
            {"type": "slicer", "field": "RateChange[Rate change %]", "title": "Rate change on open months",
             "pos": (20, ROW1_Y, 300, 76),
             "alt": "Slicer. Sets the Rate change % parameter the scenario outlook reads."},
            {"type": "slicer", "field": "LoadChange[Industrial load change (MW)]", "title": "Industrial load change",
             "pos": (330, ROW1_Y, 296, 76),
             "alt": "Slicer. Sets the Industrial load change (MW) parameter the scenario outlook reads: a "
                    "data center arriving, or a large customer leaving."},
            {"type": "stacked_column", "x": "dim_month[month_name]",
             "y": ["[Actual this year]", "[Forecast revenue]"],
             "sort": ("dim_month[month_name]", "Ascending"),
             "title": "This year by month: actual, then forecast",
             "pos": (20, ROW1_Y + 86, 606, 186),
             "alt": "Stacked column chart titled This year by month. Plots Actual this year and "
                    "Forecast revenue by month name."},
            {"type": "bar", "x": "tornado[lever]", "y": ["[Tornado impact]"],
             "sort": ("tornado[lever]", "Ascending"),
             "title": "Full-year revenue from one step on each lever",
             "pos": (640, ROW1_Y, 620, 272),
             "alt": "Bar chart titled Full-year revenue from one step on each lever. Plots Tornado "
                    "impact by lever."},
            {"type": "matrix", "rows": "sensitivity[rate_change]", "columns_by": "sensitivity[extra_mw]",
             "values": ["[Full-year revenue ($M)]"], "totals": False,
             "title": "Full-year revenue: rate change down the side, new data-center MW across the top",
             "pos": (20, ROW2_Y, 1240, 258),
             "alt": "Matrix titled Full-year revenue by rate change and new data-center MW. Shows "
                    "Sensitivity outlook for each rate change by extra MW."},
        ],
    },
    # ----------------------------------------------------------------- 6
    {
        "name": "section_ferc",
        "display": "FERC P&L and peers",
        "visuals": [
            *cards(
                ("[Operating margin, latest year]", None,
                 "Card. Operating margin, latest year: operating income over operating revenue in the "
                 "latest FERC filing."),
                ("[Retail price, latest year (cents/kWh)]", None,
                 "Card. Retail price, latest year (cents/kWh): FERC retail revenue over retail MWh, in cents per kWh."),
                ("[Power cost per MWh, latest year]", None,
                 "Card. Power cost per MWh, latest year: fuel and purchased power over every MWh sold."),
                ("[PGE industrial load growth]", None,
                 "Card. PGE industrial load growth a year since the base year."),
            ),
            {"type": "matrix", "rows": "ferc_pnl[line]", "columns_by": "ferc_pnl[year]",
             "values": ["[P&L ($M)]"], "totals": False,
             "title": "Profit and loss from FERC Form 1, $ millions",
             "pos": (20, ROW1_Y, 1240, 318),
             "alt": "Matrix titled Profit and loss from FERC Form 1. Shows P&L in $ millions for "
                    "each line by year."},
            {"type": "line", "x": "ferc_years[year_label]", "y": ["[Operating margin %]"],
             "sort": ("ferc_years[year_label]", "Ascending"),
             "title": "Operating margin",
             "pos": (20, ROW2_Y + 46, 340, 212),
             "alt": "Line chart titled Operating margin. Plots Operating margin % by year label."},
            {"type": "bar", "x": "peers[utility]", "y": ["[Industrial load growth a year]"],
             "sort": ("[Industrial load growth a year]", "Descending"),
             "title": "Industrial load growth a year, 2019–2025",
             "pos": (374, ROW2_Y + 46, 340, 212),
             "alt": "Bar chart titled Industrial load growth a year. Plots Industrial load growth a "
                    "year by utility."},
            {"type": "table",
             "columns": ["peers[utility]", "[Peer ¢/kWh]", "[Peer margin]", "[Peer power $/MWh]"],
             "title": "Five Pacific Northwest utilities, 2025",
             "totals": False,
             "pos": (728, ROW2_Y + 46, 532, 212),
             "alt": "Table titled Five Pacific Northwest utilities. Lists utility, peer retail price "
                    "in cents per kWh, peer margin and peer power cost per MWh."},
            {"type": "slicer", "field": "peers[utility]", "title": "Utility",
             "pos": (1068, ROW2_Y, 192, 76),
             "alt": "Slicer. Filters the peer chart and table by utility."},
        ],
    },
    # ----------------------------------------------------------------- 7
    {
        "name": "section_tieout",
        "display": "EIA against FERC",
        "visuals": [
            *cards(
                ("[MWh gap %]", None,
                 "Card. MWh gap: EIA's monthly survey against the FERC filing, all full years."),
                ("[Revenue gap %]", None,
                 "Card. Revenue gap: EIA billed revenue against FERC booked revenue."),
                ("[EIA revenue]", None, "Card. EIA revenue: retail revenue summed from the monthly survey."),
                ("[FERC revenue]", None, "Card. FERC revenue: retail revenue as filed in Form 1."),
            ),
            {"type": "column", "x": "tieout[year_label]", "y": ["[Revenue gap %]"],
             "sort": ("tieout[year_label]", "Ascending"),
             "title": "Revenue: EIA's billed figure runs below FERC's booked one every year",
             "pos": (20, ROW1_Y, 760, 272),
             "alt": "Column chart titled Revenue, EIA's billed figure against FERC's booked one. "
                    "Plots Revenue gap % by year label."},
            {"type": "line", "x": "tieout[year_label]", "y": ["[EIA MWh]", "[FERC MWh]"],
             "sort": ("tieout[year_label]", "Ascending"),
             "title": "Retail MWh: two federal sources, one line",
             "pos": (794, ROW1_Y, 466, 272),
             "alt": "Line chart titled Retail MWh, two federal sources. Plots EIA MWh and FERC MWh "
                    "by year label; the two lines lie on top of each other."},
            {"type": "table",
             "columns": ["tieout[year_label]", "[EIA MWh]", "[FERC MWh]", "[MWh gap %]",
                         "[EIA revenue ($M)]", "[FERC revenue ($M)]", "[Revenue gap %]"],
             "title": "The reconciliation: MWh agree within 0.1%; revenue is billed against booked",
             "pos": (20, ROW2_Y, 1044, 258),
             "alt": "Table titled The reconciliation. Lists year label, EIA MWh, FERC MWh, MWh gap %, "
                    "EIA revenue, FERC revenue and revenue gap %."},
            {"type": "slicer", "field": "tieout[year_label]", "title": "Year",
             "pos": (1068, ROW2_Y, 192, 76),
             "alt": "Slicer. Filters the reconciliation by year label."},
        ],
    },
]
VISUAL_TYPES: dict[str, str] = {
    "card": "card",
    "bar": "clusteredBarChart",
    "column": "clusteredColumnChart",
    "stacked_column": "columnChart",
    "line": "lineChart",
    "area": "areaChart",
    "scatter": "scatterChart",
    "donut": "donutChart",
    "treemap": "treemap",
    "waterfall": "waterfallChart",
    "table": "tableEx",
    "matrix": "pivotTable",
    "slicer": "slicer",
    "gauge": "gauge",
    # The chrome report_chrome adds to every page.
    "page_header": "image",
    "nav": "actionButton",
    "filters_button": "actionButton",
    "panel_close": "actionButton",
    "panel_clear": "actionButton",
    "panel_background": "shape",
    "panel_title": "textbox",
}

from powerbi.report_chrome import add_chrome  # noqa: E402 -- applied to the pages above

PAGES = add_chrome(PAGES)
