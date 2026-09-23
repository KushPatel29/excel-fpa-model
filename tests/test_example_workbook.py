"""The second Excel model, kept as an example, is the version its own tests passed.

examples/kestrel-bay/ holds the synthetic distributor model that came before the
utility rebuild. Its build code and 182 tests live at the tag kestrel-bay-v1.
Pinning the file's hash to that tagged workbook means the copy shown here cannot
be swapped for an untested one without this failing.
"""
from __future__ import annotations

import hashlib
import warnings

import openpyxl

from workbook_io import ROOT

EXAMPLE = ROOT / "examples" / "kestrel-bay"
WORKBOOK = EXAMPLE / "Kestrel_Bay_FPA_Model.xlsx"
TAGGED_SHA256 = "733f5c2de03d36fdab58b90ad550ec6956fc3fd134f079142f495e6829c0865d"


def test_the_kept_workbook_is_the_tagged_one():
    assert hashlib.sha256(WORKBOOK.read_bytes()).hexdigest() == TAGGED_SHA256


def test_it_still_says_its_checks_pass():
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(WORKBOOK, data_only=True)
    sheet, cell = next(wb.defined_names["checks_summary"].destinations)
    assert wb[sheet][cell.replace("$", "")].value == "All 31 checks pass"
    assert {"WorkingCapital", "Channels", "PVM", "Scenarios"} <= set(wb.sheetnames)


def test_the_readme_points_at_the_tag_and_the_pictures_exist():
    text = (EXAMPLE / "README.md").read_text(encoding="utf-8")
    assert "kestrel-bay-v1" in text and "synthetic" in text.lower()
    for name in ("dashboard", "pvm", "working-capital", "channels"):
        assert (EXAMPLE / "img" / f"{name}.png").exists(), name
