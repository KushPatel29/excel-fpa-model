"""Read the workbook through its defined names, so a test never hard-codes a cell."""
from __future__ import annotations

import base64
import io
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "workbook" / "Kestrel_Bay_FPA_Model.xlsx"


def block(wb, name: str) -> list[list]:
    """The values (or formulas) of a named range as a list of rows."""
    sheet, coord = next(wb.defined_names[name].destinations)
    cells = wb[sheet][coord]
    if not isinstance(cells, tuple):
        return [[cells.value]]
    if not isinstance(cells[0], tuple):
        cells = (cells,)
    return [[c.value for c in row] for row in cells]


def cell(wb, name: str):
    return block(wb, name)[0][0]


def column(wb, name: str) -> list:
    return [row[0] for row in block(wb, name)]


def parts() -> dict[str, bytes]:
    with zipfile.ZipFile(WORKBOOK) as z:
        return {n: z.read(n) for n in z.namelist()}


def sheet_xml(wb_parts: dict, wb_formulas, sheet: str) -> str:
    """The raw XML of a worksheet, found through the workbook's relationships."""
    index = wb_formulas.sheetnames.index(sheet) + 1
    rels = wb_parts["xl/_rels/workbook.xml.rels"].decode()
    book = wb_parts["xl/workbook.xml"].decode()
    rid = re.findall(r'<sheet [^>]*name="' + re.escape(sheet) + r'"[^>]*r:id="(rId\d+)"', book)
    if rid:
        target = re.search(r'Id="' + rid[0] + r'"[^>]*Target="([^"]+)"', rels) or \
            re.search(r'Target="([^"]+)"[^>]*Id="' + rid[0] + r'"', rels)
        path = "xl/" + target.group(1).lstrip("/").removeprefix("xl/")
        return wb_parts[path].decode("utf-8")
    return wb_parts[f"xl/worksheets/sheet{index}.xml"].decode("utf-8")


def power_query_m(wb_parts: dict) -> str:
    """The M code of every query, unpacked from the DataMashup part.

    DataMashup is base64 in a customXml item: a 4-byte version, a 4-byte
    length, then a zip whose Formulas/Section1.m holds the queries.
    """
    for name, raw in wb_parts.items():
        if not re.fullmatch(r"customXml/item\d+\.xml", name):
            continue
        text = raw.decode("utf-16") if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else raw.decode("utf-8")
        found = re.search(r"<DataMashup[^>]*>(.*)</DataMashup>", text, re.S)
        if not found:
            continue
        blob = base64.b64decode(found.group(1))
        size = int.from_bytes(blob[4:8], "little")
        with zipfile.ZipFile(io.BytesIO(blob[8:8 + size])) as z:
            return z.read("Formulas/Section1.m").decode("utf-8-sig")
    raise AssertionError("no DataMashup part: the workbook has no Power Query queries")


def money(x: float) -> str:
    """Python twin of the workbook's MONEY LAMBDA, for checking written commentary."""
    if abs(x) >= 999_500:
        return f"{'-' if x < 0 else ''}${abs(x) / 1e6:,.1f}M"
    return f"{'-' if x < 0 else ''}${abs(x) / 1e3:,.0f}K"


def signmoney(x: float) -> str:
    return ("−" if x < 0 else "+") + money(abs(x))
