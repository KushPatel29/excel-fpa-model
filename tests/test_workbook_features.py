"""The Excel techniques the README claims, proved from the saved file.

A claim like "built with Power Query and a Power Pivot data model" is easy to
make and easy to fake with pasted values. Each test here opens the part of the
.xlsx package that only exists if the feature is really there.
"""
from __future__ import annotations

import re
from datetime import datetime

import pytest
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula

import layout as L
from workbook_io import ROOT, block, parts, power_query_m, sheet_xml

ERRORS = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!",
          "#GETTING_DATA", "#FIELD!", "#BUSY!")


@pytest.fixture(scope="module")
def pkg():
    return parts()


def test_sheets_in_reading_order(values):
    assert values.sheetnames == L.SHEETS


@pytest.mark.parametrize("sheet", L.SHEETS)
def test_no_cell_holds_an_error(values, sheet):
    bad = [(c.coordinate, c.value) for row in values[sheet].iter_rows() for c in row
           if isinstance(c.value, str) and c.value.startswith(ERRORS)]
    assert not bad, f"{sheet}: {bad[:5]}"


# ------------------------------------------------------------- Power Query
def test_six_power_query_queries(pkg):
    m = power_query_m(pkg)
    for name in ("Sales", "Product", "Channel", "Region", "Calendar", "Budget"):
        assert re.search(rf"shared #?\"?{name}\"? =", m), name


def test_budget_query_unpivots_and_merges(pkg):
    m = power_query_m(pkg)
    budget = m[m.index("shared Budget"):]
    for step in ("Table.UnpivotOtherColumns", "Table.NestedJoin", "Table.ExpandTableColumn",
                 "Table.TransformColumnTypes"):
        assert step in budget, step


def test_calendar_query_reads_the_as_of_input(pkg):
    m = power_query_m(pkg)
    assert 'Excel.CurrentWorkbook(){[Name="AsOfMonth"]}' in m


def test_query_connections(pkg):
    xml = pkg["xl/connections.xml"].decode()
    for name in ("Sales", "Product", "Channel", "Region", "Calendar", "Budget"):
        assert f'name="Query - {name}"' in xml, name
    assert "ThisWorkbookDataModel" in xml


# ------------------------------------------------------- data model and DAX
def test_power_pivot_data_model_is_embedded(pkg):
    assert len(pkg["xl/model/item.data"]) > 100_000


def test_pivot_table_reads_dax_measures(pkg):
    cache = next(v.decode() for k, v in pkg.items() if k.startswith("xl/pivotCache/pivotCacheDefinition"))
    assert 'type="external"' in cache
    for measure in ("Revenue", "Revenue vs Budget %", "GM %", "Contribution %", "Contribution per Case",
                    "Revenue YoY %"):
        assert f'uniqueName="[Measures].[{measure}]"' in cache, measure


def test_three_slicers_on_the_data_model(pkg):
    caches = [v.decode() for k, v in pkg.items() if k.startswith("xl/slicerCaches/")]
    assert len(caches) == 3
    fields = " ".join(caches)
    for field in ("[Calendar].[year]", "[Product].[category]", "[Region].[region]"):
        assert field in fields, field


def test_cube_formulas(formulas):
    ws = formulas["Channels"]
    cube = [c.value for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and "CUBEVALUE" in c.value]
    assert len(cube) >= 20
    assert any("Revenue YoY %" in f for f in cube)


# --------------------------------------------------------- what-if analysis
def test_two_what_if_data_tables(pkg, formulas):
    xml = sheet_xml(pkg, formulas, "Scenarios")
    tables = re.findall(r'<f t="dataTable"[^>]*>', xml)
    assert len(tables) == 2
    assert any('dt2D="1"' in t for t in tables), "the price × volume grid is two-variable"
    assert any('r1="C4"' in t for t in tables), "the scenario table varies the scenario cell"


def test_scenario_picker_is_a_validated_list(formulas):
    dv = [d for d in formulas["Scenarios"].data_validations.dataValidation if "C4" in str(d.sqref)]
    assert dv and dv[0].type == "list" and dv[0].formula1 == "ScenarioList"


# ------------------------------------------------------- modern formulas
def test_named_lambda_functions(formulas):
    for name in ("SAFEDIV", "FAVVAR", "ISCLOSED", "DAYSOF", "MONEY", "SIGNMONEY"):
        assert formulas.defined_names[name].attr_text.startswith("_xlfn.LAMBDA("), name


@pytest.mark.parametrize("function,sheet", [
    ("XLOOKUP", "Forecast"), ("FORECAST.ETS", "Forecast"), ("SORTBY", "Dashboard"), ("LET", "Dashboard"),
    ("HSTACK", "Scenarios"), ("XMATCH", "PVM"), ("TRANSPOSE", "Dashboard"), ("SUMIFS", "PnL"),
])
def test_modern_functions_in_use(formulas, function, sheet):
    text = " ".join(str(c.value.text if hasattr(c.value, "text") else c.value)
                    for row in formulas[sheet].iter_rows() for c in row if c.value is not None)
    assert function in text, f"{function} not found on {sheet}"


def test_formulas_use_structured_references(formulas):
    text = " ".join(str(c.value) for row in formulas["PnL"].iter_rows() for c in row
                    if isinstance(c.value, str))
    assert "tbl_Sales[net_revenue]" in text and "tbl_Budget[revenue]" in text


# ------------------------------------------------------------- charts
def test_waterfall_chart(pkg):
    chartex = [v.decode() for k, v in pkg.items() if k.startswith("xl/charts/chartEx")]
    assert any('layoutId="waterfall"' in c for c in chartex)


def test_regular_charts(pkg):
    charts = [k for k in pkg if re.fullmatch(r"xl/charts/chart\d+\.xml", k)]
    assert len(charts) >= 5


@pytest.mark.parametrize("sheet", ["Dashboard", "PnL"])
def test_sparklines(pkg, formulas, sheet):
    assert "sparklineGroup" in sheet_xml(pkg, formulas, sheet)


def test_conditional_formatting(formulas):
    for sheet in ("PnL", "PVM", "Scenarios", "WorkingCapital", "Checks", "Dashboard"):
        assert len(formulas[sheet].conditional_formatting) > 0, sheet


# ------------------------------------------------- no numbers typed by hand
CALC_SHEETS = ["Dashboard", "PnL", "PVM", "Forecast", "WorkingCapital", "Channels"]


def excel_owned(ws) -> set[str]:
    """Cells whose numbers Excel wrote itself: dynamic-array spills (stored as plain
    values beside their anchor formula) and PivotTable output."""
    def cells(ref: str) -> set[str]:
        c1, r1, c2, r2 = range_boundaries(ref)
        return {f"{get_column_letter(c)}{r}" for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)}

    owned = set()
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, ArrayFormula) and c.value.ref:
                owned |= cells(c.value.ref)
    for pt in getattr(ws, "_pivots", []):
        owned |= cells(pt.location.ref)
    return owned


def typed_numbers(ws, allowed=frozenset()) -> list[str]:
    owned = excel_owned(ws) | set(allowed)
    return [c.coordinate for row in ws.iter_rows() for c in row
            if isinstance(c.value, (int, float, datetime)) and not isinstance(c.value, bool)
            and c.coordinate not in owned]


@pytest.mark.parametrize("sheet", CALC_SHEETS)
def test_calculation_sheets_hold_no_typed_numbers(formulas, sheet):
    """Every number on a calculation sheet is a formula."""
    typed = typed_numbers(formulas[sheet])
    assert not typed, f"{sheet} has typed numbers at {typed[:10]}"


def test_scenarios_types_only_its_inputs_and_axes(formulas):
    ws = formulas["Scenarios"]
    allowed = {"C5", "C6", "C35"} | {f"{c}25" for c in "CDEFGHI"} | {f"B{r}" for r in range(26, 33)}
    for ref in ("C18:G20", "C26:I32"):          # data-table interiors hold Excel's own results
        allowed |= {c.coordinate for row in ws[ref] for c in row}
    assert not typed_numbers(ws, allowed)


def test_inputs_are_styled_as_inputs(formulas):
    ws = formulas["Assumptions"]
    for addr in ("C6", "C7", "C10", "C12", "C13", "C14", "C19", "E23"):
        assert ws[addr].fill.fgColor.rgb.endswith("FFF5D6"), addr


# --------------------------------------------------------- deliverables
def test_board_pack_pdf():
    pdf = ROOT / "docs" / "board_pack.pdf"
    assert pdf.exists() and pdf.stat().st_size > 200_000
    assert len(re.findall(rb"/Type\s*/Page(?!s)", pdf.read_bytes())) >= 10


def test_workbook_opens_on_the_cover(values):
    assert values.active.title == "Cover"


def test_named_outputs_resolve(values):
    for name in ("pnl_outlook", "pvm_ytd", "sens_grid", "wc_lines", "ch_formulas", "checks_status"):
        assert block(values, name), name
