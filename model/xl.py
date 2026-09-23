"""Excel constants, the workbook's visual tokens and small COM helpers.

Kept apart from the sheet builders so each builder reads as a list of cells
and formulas rather than as COM plumbing.
"""
from __future__ import annotations

from datetime import date

# ---- Excel enumerations (numeric, so nothing depends on generated constants)
XL_WORKBOOK = 51
XL_CALC_AUTOMATIC, XL_CALC_MANUAL = -4105, -4135
XL_SRC_RANGE, XL_SRC_MODEL = 1, 4
XL_YES = 1
XL_ROW_FIELD, XL_DATA_FIELD = 1, 4
XL_EXTERNAL = 2
XL_LEFT, XL_CENTER, XL_RIGHT = -4131, -4108, -4152
XL_VTOP, XL_VCENTER = -4160, -4108
XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT = 7, 8, 9, 10
XL_CONTINUOUS, XL_DOUBLE = 1, -4119
XL_HAIRLINE, XL_THIN, XL_MEDIUM = 1, 2, -4138
XL_CELL_VALUE, XL_EXPRESSION = 1, 2
XL_EQUAL, XL_GREATER, XL_LESS = 3, 5, 6
XL_VALIDATE_DECIMAL, XL_VALIDATE_LIST, XL_VALIDATE_DATE = 2, 3, 4
XL_VALID_ALERT_STOP, XL_BETWEEN = 1, 1
XL_COLUMN_CLUSTERED, XL_COLUMN_STACKED = 51, 52
XL_BAR_CLUSTERED, XL_BAR_STACKED = 57, 58
XL_LINE, XL_LINE_MARKERS, XL_WATERFALL = 4, 65, 119
XL_LEGEND_BOTTOM, XL_LEGEND_TOP = -4107, -4160
XL_CATEGORY, XL_VALUE = 1, 2
XL_COLUMNS = 2
XL_PORTRAIT, XL_LANDSCAPE = 1, 2
XL_TYPE_PDF = 0
XL_SPARKLINE_LINE = 1
MSO_LINE_DASH = 4
MSO_THEME_ACCENT = {1: 5, 2: 6, 3: 7, 4: 8, 5: 9, 6: 10}

# ---- visual tokens -------------------------------------------------------------
INK = "#1D2433"        # body text
NAVY = "#1F3A5F"       # headings, bands, totals in charts
TEAL = "#0F766E"       # accent: budget lines, highlights
MUTED = "#5B6577"      # notes and captions
RULE = "#C9D1DC"       # hairlines
BAND = "#EEF2F7"       # column-header fill
TILE = "#F6F8FB"       # dashboard tile fill
FORECAST_FILL = "#EAF3FB"
INPUT_FILL = "#FFF5D6"
INPUT_FONT = "#1546C8"
FAV = "#1E7B34"
UNFAV = "#B42318"
AMBER = "#B7791F"
GREY = "#98A2B3"
PASS_FILL = "#DDF3E4"
FAIL_FILL = "#FBE0DC"

BODY_FONT = "Aptos Narrow"
TITLE_FONT = "Aptos Display"

# ---- number formats ------------------------------------------------------------
F_MONEY = '#,##0;(#,##0);"–"'
F_MONEY_M = '$#,##0.0,,"M";($#,##0.0,,"M");"–"'
F_PRICE = '$#,##0.00;($#,##0.00);"–"'
F_CASES = '#,##0;(#,##0);"–"'
F_PCT = '0.0%;(0.0%);"–"'
F_VAR = '+#,##0;(#,##0);"–"'
F_VARPCT = '+0.0%;(0.0%);"–"'
F_MULT = '0.000"x"'
F_DAYS = '0.0'
F_MONTH = "mmm yy"
F_MONTH_LONG = "mmmm yyyy"
F_DATE = "yyyy-mm-dd"


def rgb(hex_: str) -> int:
    """'#RRGGBB' to the BGR integer Excel's object model expects."""
    h = hex_.lstrip("#")
    return int(h[0:2], 16) | int(h[2:4], 16) << 8 | int(h[4:6], 16) << 16


def serial(iso: str) -> int:
    """ISO date to an Excel date serial."""
    y, m, d = (int(p) for p in iso[:10].split("-"))
    return (date(y, m, d) - date(1899, 12, 30)).days


def col(n: int) -> str:
    """1 -> 'A', 27 -> 'AA'."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def cnum(letter: str) -> int:
    n = 0
    for ch in letter:
        n = n * 26 + ord(ch) - 64
    return n


def shift(letter: str, by: int) -> str:
    return col(cnum(letter) + by)


# ---- cell helpers ----------------------------------------------------------------
# Early-bound COM turns parameterised properties into Get* methods: Range.Offset(1, 2)
# silently returns the wrong cell, so everything here uses GetOffset/GetResize.
def put(ws, addr: str, value) -> None:
    """Write a value, or a formula if it starts with '='."""
    if isinstance(value, str) and value.startswith("="):
        ws.Range(addr).Formula2 = value
    else:
        ws.Range(addr).Value = value


def fx(ws, addr: str, formula: str) -> None:
    """Formula2 into one cell or a block; relative references fill like Ctrl+Enter."""
    ws.Range(addr).Formula2 = formula


def column(ws, first: str, values) -> None:
    """Write a list down a column starting at `first` (e.g. 'B6')."""
    rng = ws.Range(first).GetResize(len(values), 1)
    rng.Value = tuple((v,) for v in values)


def row(ws, first: str, values) -> None:
    rng = ws.Range(first).GetResize(1, len(values))
    rng.Value = (tuple(values),)


def font(rng, size=None, bold=None, color=None, italic=None, name=None) -> None:
    f = rng.Font
    if name is not None:
        f.Name = name
    if size is not None:
        f.Size = size
    if bold is not None:
        f.Bold = bold
    if italic is not None:
        f.Italic = italic
    if color is not None:
        f.Color = rgb(color)


def fill(rng, color: str) -> None:
    rng.Interior.Color = rgb(color)


def edge(rng, which: int, weight=XL_THIN, color=RULE, style=XL_CONTINUOUS) -> None:
    b = rng.Borders(which)
    b.LineStyle = style
    if style != XL_DOUBLE:        # a double rule carries its own weight
        b.Weight = weight
    b.Color = rgb(color)


def box(rng, color=RULE) -> None:
    for e in (XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT):
        edge(rng, e, XL_THIN, color)


def numfmt(rng, fmt: str) -> None:
    rng.NumberFormat = fmt


def title(ws, text: str, desc: str | None = None, span: str | None = None) -> None:
    """Sheet title in B1 and a one-line description in B2. `span` wraps a long
    description across B2:{span}2 instead of letting it run off the page."""
    put(ws, "B1", text)
    font(ws.Range("B1"), size=18, bold=True, color=NAVY, name=TITLE_FONT)
    ws.Rows(1).RowHeight = 30
    if desc:
        put(ws, "B2", desc)
        font(ws.Range("B2"), size=10, color=MUTED)
        if span:
            ws.Range(f"B2:{span}2").Merge()
            ws.Range("B2").WrapText = True
            ws.Range("B2").VerticalAlignment = XL_VTOP
            ws.Rows(2).RowHeight = 28


def band(ws, r: int, text: str, c1: str = "B", c2: str = "P") -> None:
    """A navy section band across c1:c2 with white text."""
    rng = ws.Range(f"{c1}{r}:{c2}{r}")
    fill(rng, NAVY)
    put(ws, f"{c1}{r}", text)
    font(rng, bold=True, color="#FFFFFF", size=10)
    ws.Rows(r).RowHeight = 18
    rng.VerticalAlignment = XL_VCENTER


def header(ws, r: int, labels, c1: str = "B", wrap: bool = True) -> None:
    rng = ws.Range(f"{c1}{r}").GetResize(1, len(labels))
    rng.Value = (tuple(labels),)
    fill(rng, BAND)
    font(rng, bold=True, color=INK)
    rng.WrapText = wrap
    rng.VerticalAlignment = XL_VCENTER
    if len(labels) > 1:   # figures are right-aligned, so their headings are too
        ws.Range(f"{shift(c1, 1)}{r}").GetResize(1, len(labels) - 1).HorizontalAlignment = XL_RIGHT
    edge(rng, XL_EDGE_BOTTOM, XL_THIN, NAVY)


def input_cell(rng) -> None:
    """Inputs are blue on pale yellow, the FP&A convention for 'change me'."""
    fill(rng, INPUT_FILL)
    font(rng, color=INPUT_FONT)
    box(rng, "#E8D48A")


def total_row(rng, double: bool = False) -> None:
    font(rng, bold=True)
    edge(rng, XL_EDGE_TOP, XL_THIN, INK)
    if double:
        edge(rng, XL_EDGE_BOTTOM, color=INK, style=XL_DOUBLE)


def sign_colours(rng, reverse: bool = False) -> None:
    """Green when favourable, red when not (reverse for 'lower is better')."""
    good, bad = (UNFAV, FAV) if reverse else (FAV, UNFAV)
    c = rng.FormatConditions.Add(XL_CELL_VALUE, XL_GREATER, "=0")
    c.Font.Color = rgb(good)
    c = rng.FormatConditions.Add(XL_CELL_VALUE, XL_LESS, "=0")
    c.Font.Color = rgb(bad)


def widths(ws, spec: dict) -> None:
    """{'A': 2, 'B': 28, 'D:O': 11}"""
    for cols, w in spec.items():
        ws.Range(f"{cols.split(':')[0]}1:{cols.split(':')[-1]}1").EntireColumn.ColumnWidth = w


def place(ws, top_left: str, bottom_right: str):
    """Points (left, top, width, height) covering a cell block, for charts and slicers."""
    a, b = ws.Range(top_left), ws.Range(bottom_right)
    return a.Left, a.Top, b.Left + b.Width - a.Left, b.Top + b.Height - a.Top
