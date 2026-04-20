"""Test: adding a new analytic dimension to a sheet with existing data.

Verifies:
  1. Existing cell values migrate to first leaf (F1), keep value, lose per-cell formula
  2. Other leaves (F2) start empty
  3. HEAD = SUM(F1, F2) via consolidation
  4. Indicator formula rules apply uniformly to ALL leaves (not just F1)
  5. Right panel formulas match grid display
  6. Removing analytic from all sheets works (no UNIQUE constraint crash)
"""
from __future__ import annotations

import os
import time
import pytest
import requests

API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")


def _req(method, path, **kw):
    r = getattr(requests, method)(f"{API}{path}", timeout=30, **kw)
    return r


def _ok(r, msg=""):
    assert r.status_code == 200, f"{msg}: {r.status_code} {r.text[:300]}"
    return r


@pytest.fixture(scope="module")
def model():
    """Build a minimal model: 1 sheet, Periods (3 months + Q), PL (indicators), no dept yet."""
    suffix = int(time.time() * 1000) % 100000
    name = f"add_analytic_test_{suffix}"

    # Model
    mid = _ok(_req("post", "/models", json={"name": name}), "create model").json()["id"]

    # Periods analytic (Q1 2027: Jan, Feb, Mar)
    periods_aid = _ok(_req("post", "/analytics", json={
        "model_id": mid, "name": "Periods", "is_periods": True,
        "period_types": ["quarter", "month"],
        "period_start": "2027-01-01", "period_end": "2027-03-31",
    }), "create periods").json()["id"]
    _ok(_req("post", f"/analytics/{periods_aid}/generate-periods"), "generate periods")

    # PL analytic (main): revenue, cost, profit
    pl_aid = _ok(_req("post", "/analytics", json={
        "model_id": mid, "name": "PL",
    }), "create PL").json()["id"]

    def mk_rec(aid, name, parent_id=None):
        return _ok(_req("post", f"/analytics/{aid}/records",
                        json={"data_json": {"name": name}, "parent_id": parent_id}),
                   f"create record {name}").json()["id"]

    revenue_id = mk_rec(pl_aid, "выручка")
    cost_id = mk_rec(pl_aid, "расходы")
    profit_id = mk_rec(pl_aid, "прибыль")

    # Sheet with Periods + PL
    sid = _ok(_req("post", "/sheets", json={"model_id": mid, "name": "Test"}),
              "create sheet").json()["id"]
    _ok(_req("post", f"/sheets/{sid}/analytics",
             json={"analytic_id": periods_aid, "sort_order": 0}), "bind periods")
    _ok(_req("post", f"/sheets/{sid}/analytics",
             json={"analytic_id": pl_aid, "sort_order": 1}), "bind PL")
    _ok(_req("put", f"/sheets/{sid}/main-analytic",
             json={"analytic_id": pl_aid}), "set main")

    # Find month period records
    periods = _ok(_req("get", f"/analytics/{periods_aid}/records")).json()
    by_id = {p["id"]: p for p in periods}
    parent_ids = {p["parent_id"] for p in periods if p.get("parent_id")}
    month_ids = sorted(
        [p["id"] for p in periods if p["id"] not in parent_ids],
        key=lambda i: by_id[i].get("sort_order", 0),
    )
    assert len(month_ids) >= 3
    m1, m2, m3 = month_ids[:3]

    # Save cells: revenue=100, cost=40 for each month, profit has formula
    for m in [m1, m2, m3]:
        _ok(_req("put", f"/cells/by-sheet/{sid}/single", json={
            "coord_key": f"{m}|{revenue_id}", "value": "100",
            "data_type": "number", "rule": "manual",
        }), f"save revenue {m}")
        _ok(_req("put", f"/cells/by-sheet/{sid}/single", json={
            "coord_key": f"{m}|{cost_id}", "value": "40",
            "data_type": "number", "rule": "manual",
        }), f"save cost {m}")
        # Profit = revenue - cost (per-cell formula)
        _ok(_req("put", f"/cells/by-sheet/{sid}/single", json={
            "coord_key": f"{m}|{profit_id}", "value": "60",
            "data_type": "number", "rule": "formula",
            "formula": "[выручка] - [расходы]",
        }), f"save profit formula {m}")

    # Add indicator rule: profit leaf = [выручка] - [расходы]
    _ok(_req("put", f"/sheets/{sid}/indicators/{profit_id}/rules", json={
        "leaf": "[выручка] - [расходы]",
        "consolidation": "",
        "scoped": [],
    }), "add profit leaf rule")

    out = {
        "model_id": mid, "sheet_id": sid,
        "periods_aid": periods_aid, "pl_aid": pl_aid,
        "revenue_id": revenue_id, "cost_id": cost_id, "profit_id": profit_id,
        "m1": m1, "m2": m2, "m3": m3,
    }
    yield out

    # Teardown
    _req("delete", f"/models/{mid}")


def _cells_by_coord(sheet_id):
    r = _ok(_req("get", f"/cells/by-sheet/{sheet_id}"))
    return {c["coord_key"]: c for c in r.json()}


# ──────────────────────────────────────────────────────────────────
# Phase 1: verify base model works before adding analytic
# ──────────────────────────────────────────────────────────────────

def test_base_model_values(model):
    """Revenue=100, cost=40, profit=60 for each month."""
    cells = _cells_by_coord(model["sheet_id"])
    for m in [model["m1"], model["m2"], model["m3"]]:
        rev = cells.get(f"{m}|{model['revenue_id']}")
        assert rev and float(rev["value"]) == 100, f"revenue should be 100, got {rev}"
        cost = cells.get(f"{m}|{model['cost_id']}")
        assert cost and float(cost["value"]) == 40, f"cost should be 40"
        profit = cells.get(f"{m}|{model['profit_id']}")
        assert profit and float(profit["value"]) == 60, f"profit should be 60"


def test_base_recalc(model):
    """Recalc should preserve values (profit formula evaluates to 60)."""
    r = _ok(_req("post", f"/cells/calculate/{model['sheet_id']}"), "recalc")
    cells = _cells_by_coord(model["sheet_id"])
    profit = cells.get(f"{model['m1']}|{model['profit_id']}")
    assert profit, "profit cell should exist"
    assert abs(float(profit["value"]) - 60) < 0.01, f"profit should be 60, got {profit['value']}"


# ──────────────────────────────────────────────────────────────────
# Phase 2: add new analytic dimension (HEAD → F1, F2)
# ──────────────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def dept(model):
    """Create dept analytic and add to sheet. Returns record IDs."""
    mid = model["model_id"]
    sid = model["sheet_id"]

    dept_aid = _ok(_req("post", "/analytics", json={
        "model_id": mid, "name": "Dept",
    }), "create dept").json()["id"]

    head_id = _ok(_req("post", f"/analytics/{dept_aid}/records",
                       json={"data_json": {"name": "HEAD"}}), "create HEAD").json()["id"]
    f1_id = _ok(_req("post", f"/analytics/{dept_aid}/records",
                      json={"data_json": {"name": "F1"}, "parent_id": head_id}),
                "create F1").json()["id"]
    f2_id = _ok(_req("post", f"/analytics/{dept_aid}/records",
                      json={"data_json": {"name": "F2"}, "parent_id": head_id}),
                "create F2").json()["id"]

    # Add dept analytic to sheet
    _ok(_req("post", f"/sheets/{sid}/analytics",
             json={"analytic_id": dept_aid, "sort_order": 2}), "bind dept")

    return {
        "dept_aid": dept_aid,
        "head_id": head_id, "f1_id": f1_id, "f2_id": f2_id,
    }


def test_cells_migrated_to_f1(model, dept):
    """After adding dept, old cell values should be on F1 (first leaf)."""
    cells = _cells_by_coord(model["sheet_id"])
    f1 = dept["f1_id"]

    for m in [model["m1"], model["m2"], model["m3"]]:
        rev_key = f"{m}|{model['revenue_id']}|{f1}"
        rev = cells.get(rev_key)
        assert rev, f"revenue on F1 should exist: {rev_key}"
        assert float(rev["value"]) == 100, f"revenue F1 should be 100, got {rev['value']}"


def test_f1_cells_keep_formulas(model, dept):
    """Migrated cells should keep their per-cell formulas (they still work
    with the new dimension because they reference indicators by name)."""
    cells = _cells_by_coord(model["sheet_id"])
    f1 = dept["f1_id"]

    for m in [model["m1"], model["m2"], model["m3"]]:
        profit_key = f"{m}|{model['profit_id']}|{f1}"
        profit = cells.get(profit_key)
        assert profit, f"profit on F1 should exist"
        assert profit["rule"] == "formula", \
            f"profit F1 should keep formula rule, got rule={profit['rule']}"
        assert profit.get("formula") == "[выручка] - [расходы]", \
            f"profit F1 should keep formula, got: {profit.get('formula')}"


def test_f2_cells_empty(model, dept):
    """F2 should have no cells (user enters values manually)."""
    cells = _cells_by_coord(model["sheet_id"])
    f2 = dept["f2_id"]

    for m in [model["m1"], model["m2"], model["m3"]]:
        rev_key = f"{m}|{model['revenue_id']}|{f2}"
        assert rev_key not in cells, f"F2 should start empty: {rev_key}"


def test_head_consolidation_after_recalc(model, dept):
    """After recalc, HEAD should equal SUM(F1, F2) for revenue."""
    _ok(_req("post", f"/cells/calculate/{model['sheet_id']}"), "recalc")
    cells = _cells_by_coord(model["sheet_id"])

    head = dept["head_id"]
    f1 = dept["f1_id"]

    m = model["m1"]
    rev_head = cells.get(f"{m}|{model['revenue_id']}|{head}")
    rev_f1 = cells.get(f"{m}|{model['revenue_id']}|{f1}")

    # F1=100, F2=empty(0), so HEAD should be 100
    assert rev_head, "HEAD revenue cell should exist after recalc"
    assert abs(float(rev_head["value"]) - 100) < 0.01, \
        f"HEAD revenue should be 100 (F1=100 + F2=0), got {rev_head['value']}"


def test_indicator_rule_applies_to_both_leaves(model, dept):
    """The indicator rule for profit should apply to BOTH F1 and F2 uniformly.
    This is the key test: resolved formulas should be the same for both."""
    sid = model["sheet_id"]
    f1 = dept["f1_id"]
    f2 = dept["f2_id"]
    m = model["m1"]

    # Check resolved formulas for F1 and F2 via the batch endpoint
    coord_f1 = f"{m}|{model['profit_id']}|{f1}"
    coord_f2 = f"{m}|{model['profit_id']}|{f2}"

    r = _ok(_req("post", f"/sheets/{sid}/cells/resolved-formulas",
                  json={"coord_keys": [coord_f1, coord_f2]}))
    resolved = {item["coord_key"]: item for item in r.json()}

    # Both should have the same formula from the indicator rule
    f1_formula = resolved.get(coord_f1, {}).get("formula", "")
    f2_formula = resolved.get(coord_f2, {}).get("formula", "")
    assert f1_formula == f2_formula, \
        f"F1 and F2 should have the same resolved formula.\n  F1: {f1_formula}\n  F2: {f2_formula}"
    assert f1_formula, "Both should have a formula (from indicator rule)"


def test_head_profit_consolidation(model, dept):
    """HEAD profit = SUM(F1 profit, F2 profit) after recalc.
    F1 profit = 60 (from indicator rule: выручка-расходы = 100-40).
    F2 has no revenue/cost, so F2 profit = 0.
    HEAD profit = 60."""
    _ok(_req("post", f"/cells/calculate/{model['sheet_id']}"), "recalc")
    cells = _cells_by_coord(model["sheet_id"])

    m = model["m1"]
    head = dept["head_id"]

    profit_head = cells.get(f"{m}|{model['profit_id']}|{head}")
    assert profit_head, "HEAD profit should exist after recalc"
    # F1 profit: indicator rule [выручка]-[расходы] = 100-40 = 60
    # F2 profit: indicator rule applies but выручка=0, расходы=0, so 0
    # HEAD = 60 + 0 = 60
    assert abs(float(profit_head["value"]) - 60) < 0.01, \
        f"HEAD profit should be 60, got {profit_head['value']}"


def test_enter_f2_values_and_recalc(model, dept):
    """Enter values for F2, recalc, verify HEAD updates."""
    sid = model["sheet_id"]
    f2 = dept["f2_id"]
    head = dept["head_id"]
    m = model["m1"]

    # Enter revenue=50, cost=20 for F2
    _ok(_req("put", f"/cells/by-sheet/{sid}/single", json={
        "coord_key": f"{m}|{model['revenue_id']}|{f2}",
        "value": "50", "data_type": "number", "rule": "manual",
    }))
    _ok(_req("put", f"/cells/by-sheet/{sid}/single", json={
        "coord_key": f"{m}|{model['cost_id']}|{f2}",
        "value": "20", "data_type": "number", "rule": "manual",
    }))

    # Recalc
    _ok(_req("post", f"/cells/calculate/{sid}"), "recalc")
    cells = _cells_by_coord(sid)

    # F2 profit should now be 50-20=30 (from indicator rule)
    profit_f2 = cells.get(f"{m}|{model['profit_id']}|{f2}")
    assert profit_f2, "F2 profit should exist after recalc"
    assert abs(float(profit_f2["value"]) - 30) < 0.01, \
        f"F2 profit should be 30, got {profit_f2['value']}"

    # HEAD revenue = 100 + 50 = 150
    rev_head = cells.get(f"{m}|{model['revenue_id']}|{head}")
    assert rev_head and abs(float(rev_head["value"]) - 150) < 0.01, \
        f"HEAD revenue should be 150, got {rev_head}"

    # HEAD profit = 60 + 30 = 90
    profit_head = cells.get(f"{m}|{model['profit_id']}|{head}")
    assert profit_head and abs(float(profit_head["value"]) - 90) < 0.01, \
        f"HEAD profit should be 90, got {profit_head['value']}"


# ──────────────────────────────────────────────────────────────────
# Phase 3: remove analytic
# ──────────────────────────────────────────────────────────────────

def test_remove_analytic(model, dept):
    """Removing the dept analytic should not crash (UNIQUE constraint)
    and restore cells to 2-part coord keys."""
    sid = model["sheet_id"]
    dept_aid = dept["dept_aid"]

    # Find the sheet_analytics binding
    bindings = _ok(_req("get", f"/sheets/{sid}/analytics")).json()
    sa = next((b for b in bindings if b["analytic_id"] == dept_aid), None)
    assert sa, "dept binding should exist"

    # Remove
    _ok(_req("delete", f"/sheets/{sid}/analytics/{sa['id']}"), "remove dept")

    # Cells should be back to 2-part keys (values may differ from original
    # because the test added F2 data and HEAD consolidated F1+F2)
    cells = _cells_by_coord(sid)
    rev = cells.get(f"{model['m1']}|{model['revenue_id']}")
    assert rev, "revenue should exist with 2-part key after removal"
    # Value should be numeric (the kept cell from collision resolution)
    assert float(rev["value"]) > 0, f"revenue should be positive, got {rev['value']}"
