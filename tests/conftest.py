"""Fixtures shared by the suite.

The workbook is read twice with openpyxl: once for the values Excel cached at
its last calculation, once for the formulas. No Excel is needed, so the suite
runs on Linux CI against the committed workbook.
"""
from __future__ import annotations

import sys
import warnings
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "model"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

WORKBOOK = ROOT / "workbook" / "PGE_Utility_FPA_Model.xlsx"


def _load(data_only: bool):
    import openpyxl

    with warnings.catch_warnings():
        # openpyxl warns about parts it does not model (slicers, the data model,
        # the waterfall chart); it still reads every cell and cached value.
        warnings.simplefilter("ignore")
        return openpyxl.load_workbook(WORKBOOK, data_only=data_only)


@pytest.fixture(scope="session")
def values():
    return _load(data_only=True)


@pytest.fixture(scope="session")
def formulas():
    return _load(data_only=False)


@pytest.fixture(scope="session")
def data():
    import reference

    return reference.load()
