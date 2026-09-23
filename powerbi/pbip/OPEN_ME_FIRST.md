# Opening this in Power BI Desktop

`PGEUtilityFPA.pbip` is a Power BI project in the **PBIR** format. The report is
one JSON file per visual and the semantic model is TMDL, so both can be reviewed
in a diff instead of being a binary file. The generator in `powerbi/` writes all
of it, and CI fails if the committed project and the spec drift apart.

## Before you open it

The model reads the CSVs in `tables/` through a parameter called **DataPath**.
When the project is generated, DataPath is set to an absolute path: the
repository root on the machine that generated it. On your machine it will point
somewhere else. You can fix that in either of two ways.

The first is to regenerate the project, which sets DataPath to wherever the
repository is:

```bash
python model/export_tables.py     # only if data/ changed
python -m powerbi.build_pbip
```

The second is to open `PGEUtilityFPA.SemanticModel/definition/expressions.tmdl`
and edit the one line at the top.

## Then

1. Open `PGEUtilityFPA.pbip` in Power BI Desktop.
2. **Refresh.** A `.pbip` stores the model *definition*, not its data. Until you
   refresh, every table is empty and the visuals are blank.
3. Wait about half a minute after the refresh bar clears, because the visuals
   are still querying.

## What is in it

The report has seven pages:

1. Overview
2. Plan and bridge
3. Weather
4. Price, volume and mix
5. Forecast and scenarios
6. FERC P&L and peers
7. EIA against FERC

It holds 117 visuals and a model of 17 tables plus a measures table, with two
what-if parameters: a rate change and an industrial load change in MW.

The tables are written by `model/export_tables.py` from the same reference model
the Excel workbook is held to. `tests/test_powerbi_tables.py` reads the
workbook's own cells against them, so the report and the workbook cannot
publish different numbers without a test failing.

The one calculation the report does itself is the scenario outlook behind the
sliders. That test replays the measure's DAX and holds it to the sensitivity
grid in every one of its 49 cells.
