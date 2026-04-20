"""Test: Excel import via streaming endpoint (same as UI).

Uses /import/excel-stream — the SAME endpoint the browser UI uses.
Verifies:
  1. Formula-derived dates (=B1+31) resolve to proper periods
  2. Yellow/colored cells import as manual, not formula
  3. All period columns produce cells (no missing months)
  4. Values match the Excel source
  5. Consolidation formulas extracted from year totals (non-SUM)
  6. Recalc produces correct consolidation values
  7. Formula rules sync: getAllIndicatorRules and getIndicatorRules return same formulas
  8. Adding a new analytic (подразделения) with hierarchy D1→D11,D12
  9. Large model (models.xlsx): all avg/rate indicators get consolidation formulas (not SUM)
"""
from __future__ import annotations

import json
import os
import pytest
import requests

API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")
EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test-avg.xlsx")
MODELS_XLSX = os.path.join(os.path.dirname(os.path.dirname(__file__)), "models.xlsx")

# Indicator name patterns that should NEVER be SUM-consolidated.
# These are rates, averages, shares — must use ratio formulas.
# Note: "% доход/расход" is interest income/expense (absolute) — not a rate.
_AVG_RATE_PATTERNS = ("средн", "ср. ", "ставка", "доля ", "на 1 ")


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


def test_formula_rules_sync(imported):
    """Formula rules from batch (grid) and individual (panel) endpoints must match.

    The grid uses GET /sheets/{sid}/indicator-rules-all (batch).
    The right panel uses GET /sheets/{sid}/indicators/{iid}/rules (individual).
    Both must return the same leaf AND consolidation formulas for each indicator.
    """
    sid = imported["sheet_id"]
    # Batch endpoint (used by grid)
    all_rules = _ok(_req("get", f"/sheets/{sid}/indicator-rules-all"), "batch rules").json()
    assert len(all_rules) > 0, "Should have formula rules for some indicators"

    # For each indicator, check individual endpoint matches batch
    for ind_id, batch_entry in all_rules.items():
        batch_leaf = batch_entry.get("leaf", "")
        batch_consol = batch_entry.get("consolidation", "")
        if not batch_leaf and not batch_consol:
            continue
        # Individual endpoint (used by right panel)
        ind_rules = _ok(
            _req("get", f"/sheets/{sid}/indicators/{ind_id}/rules"),
            f"individual rules for {ind_id}",
        ).json()
        panel_leaf = ind_rules.get("leaf", "")
        panel_consol = ind_rules.get("consolidation", "")
        assert panel_leaf == batch_leaf, \
            f"Leaf formula mismatch for {ind_id}: grid={batch_leaf!r}, panel={panel_leaf!r}"
        assert panel_consol == batch_consol, \
            f"Consolidation formula mismatch for {ind_id}: grid={batch_consol!r}, panel={panel_consol!r}"


def test_no_average_consolidation(imported):
    """AVERAGE consolidation should never be used — it's mathematically wrong
    for weighted averages. Excel AVERAGE formulas should be skipped so the LLM
    provides correct ratio formulas instead."""
    sid = imported["sheet_id"]
    all_rules = _ok(_req("get", f"/sheets/{sid}/indicator-rules-all"), "batch rules").json()

    for ind_id, entry in all_rules.items():
        consol = entry.get("consolidation", "")
        assert consol != "AVERAGE", \
            f"Indicator {ind_id} has AVERAGE consolidation — should be a ratio formula"


def test_formula_columns_separate(imported):
    """Grid should have separate leaf and consolidation formulas.

    After import, indicators should have:
    - 2 indicators with consolidation formulas (non-SUM: weighted avg, ratio)
    - 2 indicators with leaf formulas (from cell_data)
    - 1 indicator with neither (количество партнеров — pure manual)
    """
    sid = imported["sheet_id"]
    all_rules = _ok(_req("get", f"/sheets/{sid}/indicator-rules-all"), "batch rules").json()

    leaf_count = sum(1 for e in all_rules.values() if e.get("leaf"))
    consol_count = sum(1 for e in all_rules.values() if e.get("consolidation"))

    assert leaf_count == 2, f"Expected 2 indicators with leaf formulas, got {leaf_count}"
    assert consol_count == 2, f"Expected 2 indicators with consolidation formulas, got {consol_count}"

    # No indicator should have both leaf AND consolidation
    both_count = sum(1 for e in all_rules.values() if e.get("leaf") and e.get("consolidation"))
    assert both_count == 0, f"No indicator should have both leaf and consolidation, got {both_count}"


def test_manual_indicators_have_no_formula_rules(imported):
    """Indicators with only manual cells should NOT have formula rules."""
    sid = imported["sheet_id"]
    all_rules = _ok(_req("get", f"/sheets/{sid}/indicator-rules-all"), "batch rules").json()
    cells = imported["cells"]

    # Find indicators that are fully manual (all cells are manual)
    indicator_rules: dict[str, set] = {}
    for c in cells:
        parts = c["coord_key"].split("|")
        iid = parts[1] if len(parts) > 1 else parts[0]
        indicator_rules.setdefault(iid, set()).add(c["rule"])

    for iid, rules_set in indicator_rules.items():
        if rules_set == {"manual"}:
            # Fully manual indicator — should not have a leaf formula
            entry = all_rules.get(iid, {})
            assert not entry.get("leaf"), \
                f"Manual indicator {iid} should not have leaf formula, got: {entry.get('leaf')}"


def test_add_analytic_with_hierarchy(imported):
    """Add a new analytic 'подразделения' with D1 → D11, D12 and verify structure."""
    model_id = imported["model_id"]
    sid = imported["sheet_id"]

    # Create analytic
    analytic = _ok(_req("post", "/analytics", json={
        "model_id": model_id,
        "name": "Подразделения",
        "code": "divisions",
    }), "create analytic").json()
    aid = analytic["id"]

    # Add fields: name
    _ok(_req("post", f"/analytics/{aid}/fields", json={
        "name": "Наименование", "code": "name", "field_type": "string",
    }), "add name field")

    # Create root record D1
    d1 = _ok(_req("post", f"/analytics/{aid}/records", json={
        "data_json": {"name": "D1"},
    }), "create D1").json()

    # Create child records D11, D12 under D1
    d11 = _ok(_req("post", f"/analytics/{aid}/records", json={
        "parent_id": d1["id"],
        "data_json": {"name": "D11"},
    }), "create D11").json()
    d12 = _ok(_req("post", f"/analytics/{aid}/records", json={
        "parent_id": d1["id"],
        "data_json": {"name": "D12"},
    }), "create D12").json()

    # Verify hierarchy
    records = _ok(_req("get", f"/analytics/{aid}/records"), "list records").json()
    assert len(records) == 3, f"Expected 3 records, got {len(records)}"

    children = [r for r in records if r.get("parent_id") == d1["id"]]
    assert len(children) == 2, f"D1 should have 2 children, got {len(children)}"

    child_names = sorted(
        json.loads(r["data_json"]).get("name", "") if isinstance(r["data_json"], str)
        else r["data_json"].get("name", "")
        for r in children
    )
    assert child_names == ["D11", "D12"], f"Expected D11, D12; got {child_names}"

    # Bind analytic to the sheet
    _ok(_req("post", f"/sheets/{sid}/analytics", json={
        "analytic_id": aid,
    }), "bind analytic to sheet")

    # Verify tree now shows the new analytic
    tree = _ok(_req("get", f"/models/{model_id}/tree"), "get tree after analytic").json()
    analytic_ids = [a["id"] for a in tree.get("analytics", [])]
    assert aid in analytic_ids, "New analytic should appear in model tree"

    # Recalc still works with the new analytic
    _ok(_req("post", f"/cells/calculate/{sid}"), "recalc after adding analytic")
    cells = _ok(_req("get", f"/cells/by-sheet/{sid}")).json()
    assert len(cells) >= 15, f"Should still have at least 15 cells, got {len(cells)}"


def test_panel_sync_with_3_analytics(imported):
    """After adding a 3rd analytic, individual getIndicatorRules must still
    return the same leaf formula as batch getAllIndicatorRules.

    Regression: LIKE pattern '%|{id}' didn't match 3-part coord_keys
    like 'period|indicator|division', so the panel showed 'ручной ввод'
    while the grid showed the correct formula.
    """
    sid = imported["sheet_id"]
    # Batch endpoint (grid)
    all_rules = _ok(_req("get", f"/sheets/{sid}/indicator-rules-all"), "batch rules").json()

    # For each indicator with a leaf formula, check individual endpoint matches
    mismatches = []
    for ind_id, batch_entry in all_rules.items():
        batch_leaf = batch_entry.get("leaf", "")
        if not batch_leaf:
            continue
        # Individual endpoint (panel)
        ind_rules = _ok(
            _req("get", f"/sheets/{sid}/indicators/{ind_id}/rules"),
            f"individual rules for {ind_id}",
        ).json()
        panel_leaf = ind_rules.get("leaf", "")
        if panel_leaf != batch_leaf:
            mismatches.append(
                f"{ind_id}: grid={batch_leaf!r}, panel={panel_leaf!r}"
            )
    assert not mismatches, \
        f"Panel/grid leaf formula mismatch after adding 3rd analytic:\n" + "\n".join(mismatches)


# ---------------------------------------------------------------------------
# Large model tests (models.xlsx — multi-sheet financial model)
# ---------------------------------------------------------------------------

def _get_indicator_names(tree: dict) -> dict[str, str]:
    """Build {record_id: name} from analytics records."""
    names = {}
    for a in tree.get("analytics", []):
        if a.get("is_periods"):
            continue
        recs = _ok(_req("get", f"/analytics/{a['id']}/records"), "records").json()
        for r in recs:
            dj = r.get("data_json", {})
            if isinstance(dj, str):
                dj = json.loads(dj)
            nm = (dj or {}).get("name", "")
            if nm:
                names[r["id"]] = nm
    return names


@pytest.fixture(scope="module")
def large_model():
    """Import models.xlsx and gather per-sheet info."""
    if not os.path.exists(MODELS_XLSX):
        pytest.skip("models.xlsx not found")

    data = _import_via_stream(MODELS_XLSX)
    model_id = data["model_id"]
    tree = _ok(_req("get", f"/models/{model_id}/tree"), "get tree").json()
    sheets = tree.get("sheets", [])
    ind_names = _get_indicator_names(tree)

    info: dict[str, dict] = {}
    for sh in sheets:
        sid = sh["id"]
        rules = _ok(_req("get", f"/sheets/{sid}/indicator-rules-all"), "rules").json()
        # Attach names to rules
        named_rules = {}
        for ind_id, entry in rules.items():
            entry["name"] = ind_names.get(ind_id, "")
            named_rules[ind_id] = entry
        info[sh["name"]] = {
            "id": sid,
            "rules": named_rules,
            "indicator_count": len(named_rules),
        }

    yield {"model_id": model_id, "sheets": info}

    _req("delete", f"/models/{model_id}")


def test_large_model_imports_all_sheets(large_model):
    """models.xlsx should produce 7 sheets."""
    assert len(large_model["sheets"]) == 7, \
        f"Expected 7 sheets, got {len(large_model['sheets'])}"


def test_large_model_avg_indicators_have_formulas(large_model):
    """Every indicator with 'средн/ср./ставка/доля' in name must have
    a consolidation formula — never plain SUM."""
    missing = []
    for sheet_name, info in large_model["sheets"].items():
        for ind_id, entry in info["rules"].items():
            ind_name = entry.get("name", "").lower()
            consol = entry.get("consolidation", "")
            if any(p in ind_name for p in _AVG_RATE_PATTERNS):
                if not consol or consol == "SUM":
                    missing.append(f"{sheet_name}: {entry.get('name', ind_id)}")
    assert not missing, \
        f"Avg/rate indicators without consolidation formula:\n" + "\n".join(missing[:20])


def test_large_model_no_average_consolidation(large_model):
    """No indicator should have AVERAGE as consolidation — it's always wrong."""
    bad = []
    for sheet_name, info in large_model["sheets"].items():
        for ind_id, entry in info["rules"].items():
            if entry.get("consolidation") == "AVERAGE":
                bad.append(f"{sheet_name}: {entry.get('name', ind_id)}")
    assert not bad, \
        f"Indicators with AVERAGE consolidation:\n" + "\n".join(bad[:20])


def test_large_model_cross_sheet_consistency(large_model):
    """Same-named indicator on different sheets must have same consolidation type.

    If indicator X has a consolidation formula on sheet A, the same-named
    indicator on sheet B must also have one (not be left as default SUM).
    """
    # Build name → list of (sheet, formula) pairs
    name_rules: dict[str, list[tuple[str, str]]] = {}
    for sheet_name, info in large_model["sheets"].items():
        for ind_id, entry in info["rules"].items():
            nm = entry.get("name", "")
            if not nm:
                continue
            consol = entry.get("consolidation", "")
            name_rules.setdefault(nm, []).append((sheet_name, consol))

    # Check: if any sheet has a formula, all sheets should
    inconsistent = []
    for nm, entries in name_rules.items():
        if len(entries) < 2:
            continue
        has_formula = any(f and f != "SUM" for _, f in entries)
        has_empty = any(not f for _, f in entries)
        if has_formula and has_empty:
            sheets_with = [s for s, f in entries if f and f != "SUM"]
            sheets_without = [s for s, f in entries if not f]
            inconsistent.append(
                f"{nm}: has formula on [{', '.join(sheets_with)}] "
                f"but missing on [{', '.join(sheets_without)}]"
            )
    assert not inconsistent, \
        f"Cross-sheet inconsistency:\n" + "\n".join(inconsistent[:10])
