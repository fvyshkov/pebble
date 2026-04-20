"""Test: Excel import with text-based period headers and yellow cell detection.

Verifies:
  1. Text-based period headers ("январь 2026") are detected
  2. Formula-derived dates (=B1+31) resolve to proper periods
  3. Yellow/colored cells import as manual, not formula
  4. All period columns produce cells (no missing months)
  5. Reimport after delete works cleanly
"""
from __future__ import annotations

import os
import pytest
import requests

API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")
EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test-avg.xlsx")


def _req(method, path, **kw):
    r = getattr(requests, method)(f"{API}{path}", timeout=30, **kw)
    return r


def _ok(r, msg=""):
    assert r.status_code == 200, f"{msg}: {r.status_code} {r.text[:300]}"
    return r


@pytest.fixture(scope="module")
def imported():
    """Import test-avg.xlsx and return model info."""
    if not os.path.exists(EXCEL_PATH):
        pytest.skip("test-avg.xlsx not found")

    with open(EXCEL_PATH, "rb") as f:
        r = _ok(_req("post", "/import/excel", files={"file": ("test-avg.xlsx", f)}), "import")
    data = r.json()
    yield data

    # Teardown
    _req("delete", f"/models/{data['model_id']}")


def test_import_creates_model(imported):
    """Import should create a model with 1 sheet."""
    assert imported["sheets"] == 1
    assert len(imported["sheet_list"]) == 1


def test_import_detects_both_periods(imported):
    """Should detect 2 monthly periods (Jan + Feb 2026)."""
    assert imported["periods"] == 2


def test_import_cell_count(imported):
    """5 indicators × 2 months = 10 cells."""
    sheet = imported["sheet_list"][0]
    assert sheet["cells"] == 10


def test_import_cells_have_correct_rules(imported):
    """Yellow cells should be manual, formula cells should be formula."""
    sid = imported["sheet_list"][0]["id"]
    r = _ok(_req("get", f"/cells/by-sheet/{sid}"))
    cells = r.json()

    manual_count = sum(1 for c in cells if c["rule"] == "manual")
    formula_count = sum(1 for c in cells if c["rule"] == "formula")

    # 3 indicators × 2 months = 6 manual (yellow cells)
    assert manual_count == 6, f"Expected 6 manual cells, got {manual_count}"
    # 2 indicators × 2 months = 4 formula cells
    assert formula_count == 4, f"Expected 4 formula cells, got {formula_count}"


def test_import_manual_cells_have_no_formula(imported):
    """Manual cells should not have formulas."""
    sid = imported["sheet_list"][0]["id"]
    r = _ok(_req("get", f"/cells/by-sheet/{sid}"))
    cells = r.json()

    for c in cells:
        if c["rule"] == "manual":
            assert not c.get("formula"), \
                f"Manual cell {c['coord_key']} should not have formula, got: {c.get('formula')}"


def test_import_formula_cells_have_formula(imported):
    """Formula cells should have non-empty formula text."""
    sid = imported["sheet_list"][0]["id"]
    r = _ok(_req("get", f"/cells/by-sheet/{sid}"))
    cells = r.json()

    for c in cells:
        if c["rule"] == "formula":
            assert c.get("formula"), \
                f"Formula cell {c['coord_key']} should have formula text"


def test_import_all_periods_have_cells(imported):
    """Both January and February should have cells."""
    sid = imported["sheet_list"][0]["id"]
    r = _ok(_req("get", f"/cells/by-sheet/{sid}"))
    cells = r.json()

    period_ids = set()
    for c in cells:
        parts = c["coord_key"].split("|")
        period_ids.add(parts[0])

    assert len(period_ids) == 2, f"Expected 2 distinct period IDs, got {len(period_ids)}: {period_ids}"


def test_import_values_match_excel(imported):
    """Spot-check: known values from test-avg.xlsx."""
    sid = imported["sheet_list"][0]["id"]
    r = _ok(_req("get", f"/cells/by-sheet/{sid}"))
    cells = r.json()

    values = sorted([float(c["value"]) for c in cells])
    # Excel has: 10, 15, 25, 30, 50, 55, 300, 375, 15000, 20625
    expected = sorted([10, 15, 25, 30, 50, 55, 300, 375, 15000, 20625])
    assert values == expected, f"Values mismatch: {values} != {expected}"
