"""Test: Excel import via streaming endpoint (same as UI).

Uses /import/excel-stream — the SAME endpoint the browser UI uses.
Verifies:
  1. Formula-derived dates (=B1+31) resolve to proper periods
  2. Yellow/colored cells import as manual, not formula
  3. All period columns produce cells (no missing months)
  4. Values match the Excel source
"""
from __future__ import annotations

import json
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


def _import_via_stream(path: str) -> dict:
    """Import using the streaming SSE endpoint (same as UI uses)."""
    with open(path, "rb") as f:
        r = requests.post(
            f"{API}/import/excel-stream",
            files={"file": (os.path.basename(path), f)},
            timeout=120,
            stream=True,
        )
    assert r.status_code == 200, f"stream import failed: {r.status_code}"
    last_data = None
    for line in r.iter_lines(decode_unicode=True):
        if line and line.startswith("data: "):
            last_data = json.loads(line[6:])
    assert last_data and "model_id" in last_data, \
        f"No model_id in stream response: {last_data}"
    return last_data


@pytest.fixture(scope="module")
def imported():
    """Import test-avg.xlsx via streaming endpoint and gather model info."""
    if not os.path.exists(EXCEL_PATH):
        pytest.skip("test-avg.xlsx not found")

    data = _import_via_stream(EXCEL_PATH)
    model_id = data["model_id"]

    # Get model tree to find sheets and analytics
    tree = _ok(_req("get", f"/models/{model_id}/tree"), "get tree").json()
    sheets = tree.get("sheets", [])
    assert len(sheets) > 0, "Model should have at least one sheet"

    sid = sheets[0]["id"]

    # Get period records
    analytics = tree.get("analytics", [])
    period_analytic = next((a for a in analytics if a.get("is_periods")), None)
    period_records = []
    if period_analytic:
        period_records = _ok(_req("get", f"/analytics/{period_analytic['id']}/records")).json()

    # Get cells
    cells = _ok(_req("get", f"/cells/by-sheet/{sid}")).json()

    result = {
        "model_id": model_id,
        "sheet_id": sid,
        "sheets": sheets,
        "period_records": period_records,
        "cells": cells,
    }
    yield result

    # Teardown
    _req("delete", f"/models/{model_id}")


def test_import_creates_model(imported):
    """Import should create a model with 1 sheet."""
    assert len(imported["sheets"]) == 1


def test_import_detects_both_periods(imported):
    """Should detect 2 monthly periods (Jan + Feb 2026)."""
    # Filter to leaf periods (no children = month-level)
    recs = imported["period_records"]
    parent_ids = {r["parent_id"] for r in recs if r.get("parent_id")}
    leaves = [r for r in recs if r["id"] not in parent_ids]
    assert len(leaves) == 2, \
        f"Expected 2 leaf periods (Jan+Feb), got {len(leaves)}: {[r.get('data_json') for r in leaves]}"


def test_import_cell_count(imported):
    """5 indicators × 2 months = 10 cells."""
    assert len(imported["cells"]) == 10, \
        f"Expected 10 cells, got {len(imported['cells'])}"


def test_import_cells_have_correct_rules(imported):
    """Yellow cells should be manual, formula cells should be formula."""
    cells = imported["cells"]
    manual_count = sum(1 for c in cells if c["rule"] == "manual")
    formula_count = sum(1 for c in cells if c["rule"] == "formula")

    # 3 indicators × 2 months = 6 manual (yellow cells in Excel)
    assert manual_count == 6, f"Expected 6 manual cells, got {manual_count}"
    # 2 indicators × 2 months = 4 formula cells
    assert formula_count == 4, f"Expected 4 formula cells, got {formula_count}"


def test_import_manual_cells_have_no_formula(imported):
    """Manual cells should not have formulas."""
    for c in imported["cells"]:
        if c["rule"] == "manual":
            assert not c.get("formula"), \
                f"Manual cell {c['coord_key']} should not have formula, got: {c.get('formula')}"


def test_import_formula_cells_have_formula(imported):
    """Formula cells should have non-empty formula text."""
    for c in imported["cells"]:
        if c["rule"] == "formula":
            assert c.get("formula"), \
                f"Formula cell {c['coord_key']} should have formula text"


def test_import_all_periods_have_cells(imported):
    """Both January and February should have cells."""
    period_ids = set()
    for c in imported["cells"]:
        parts = c["coord_key"].split("|")
        period_ids.add(parts[0])
    assert len(period_ids) == 2, \
        f"Expected 2 distinct period IDs, got {len(period_ids)}"


def test_import_values_match_excel(imported):
    """Spot-check: known values from test-avg.xlsx."""
    values = sorted([float(c["value"]) for c in imported["cells"]])
    expected = sorted([10, 15, 25, 30, 50, 55, 300, 375, 15000, 20625])
    assert values == expected, f"Values mismatch: {values} != {expected}"
