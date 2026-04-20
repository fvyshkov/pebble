"""Test: Excel import via streaming endpoint (same as UI).

Uses /import/excel-stream — the SAME endpoint the browser UI uses.
Verifies:
  1. Formula-derived dates (=B1+31) resolve to proper periods
  2. Yellow/colored cells import as manual, not formula
  3. All period columns produce cells (no missing months)
  4. Values match the Excel source
  5. Consolidation formulas extracted from year totals (non-SUM)
  6. Recalc produces correct consolidation values
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


def test_import_detects_all_periods(imported):
    """Should detect 3 monthly periods (Jan, Feb, Mar 2026)."""
    recs = imported["period_records"]
    parent_ids = {r["parent_id"] for r in recs if r.get("parent_id")}
    leaves = [r for r in recs if r["id"] not in parent_ids]
    assert len(leaves) == 3, \
        f"Expected 3 leaf periods, got {len(leaves)}"


def test_import_cell_count(imported):
    """5 indicators × 3 months = 15 cells."""
    assert len(imported["cells"]) == 15, \
        f"Expected 15 cells, got {len(imported['cells'])}"


def test_import_cells_have_correct_rules(imported):
    """Yellow cells should be manual, formula cells should be formula."""
    cells = imported["cells"]
    manual_count = sum(1 for c in cells if c["rule"] == "manual")
    formula_count = sum(1 for c in cells if c["rule"] == "formula")

    # 3 indicators × 3 months = 9 manual (yellow cells in Excel)
    assert manual_count == 9, f"Expected 9 manual cells, got {manual_count}"
    # 2 indicators × 3 months = 6 formula cells
    assert formula_count == 6, f"Expected 6 formula cells, got {formula_count}"


def test_import_manual_cells_have_no_formula(imported):
    """Manual cells should not have formulas."""
    for c in imported["cells"]:
        if c["rule"] == "manual":
            assert not c.get("formula"), \
                f"Manual cell {c['coord_key']} should not have formula, got: {c.get('formula')}"


def test_import_all_periods_have_cells(imported):
    """All 3 months should have cells."""
    period_ids = set()
    for c in imported["cells"]:
        parts = c["coord_key"].split("|")
        period_ids.add(parts[0])
    assert len(period_ids) == 3, \
        f"Expected 3 distinct period IDs, got {len(period_ids)}"


def test_consolidation_after_recalc(imported):
    """After recalc, year/quarter totals should use Excel consolidation formulas.

    Excel year column:
      количество партнеров: =B2+C2+D2 → SUM = 40
      среднее кол-во выдач: =E5/E2 → 705/40 = 17.625 (NOT SUM)
      ср. сумма выдачи: =E6/E5 → 35715/705 = 50.66 (NOT SUM)
      количество выдач: =B5+C5+D5 → SUM = 705
      выдача (сумма): =B6+C6+D6 → SUM = 35715
    """
    sid = imported["sheet_id"]
    _ok(_req("post", f"/cells/calculate/{sid}"), "recalc")
    cells = _ok(_req("get", f"/cells/by-sheet/{sid}")).json()

    # Build lookup: period_name → indicator_name → value
    by_name: dict[str, dict[str, float]] = {}
    for c in cells:
        parts = c["coord_key"].split("|")
        # Get names via API
        pid, iid = parts[0], parts[1]
        by_name.setdefault(pid, {})[iid] = float(c["value"])

    # Find year-level period (has children, is top-level)
    recs = imported["period_records"]
    parent_ids = {r["parent_id"] for r in recs if r.get("parent_id")}
    year_recs = [r for r in recs if not r.get("parent_id")]
    assert year_recs, "Should have a year-level period"
    year_id = year_recs[0]["id"]

    year_vals = by_name.get(year_id, {})
    assert len(year_vals) == 5, f"Year should have 5 indicators, got {len(year_vals)}"

    # Check expected values
    vals_list = sorted(year_vals.values())
    # Expected: 17.625, 40, 50.66, 705, 35715
    expected_approx = [17.625, 40.0, 50.66, 705.0, 35715.0]
    for exp, got in zip(sorted(expected_approx), vals_list):
        assert abs(exp - got) < 0.1, \
            f"Year consolidation mismatch: expected ~{exp}, got {got}. All: {vals_list}"
