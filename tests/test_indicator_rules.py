"""P1: indicator formula rules — schema + resolver + API.

Scenario (from the plan at /Users/mac/.claude/plans/zippy-zooming-pelican.md):
    - Build a tiny model: 1 sheet, Periods(Q4 2026: M10/M11/M12) + PL(main)
      with 3 indicators (партнёры, выдачи, сред/партнёра) + Dep(D1 → D11, D12).
    - Indicator «сред/партнёра» gets consolidation rule = `[выдачи] / [партнёры]`
      (same-context ratio). On HEAD, recursion gives sum(выдачи)/sum(партнёры),
      NOT avg of leaf ratios.
    - A per-cell override on a HEAD cell beats the indicator rule (precedence).
    - A scoped rule «для Q4 2026» beats the base consolidation rule.
    - Promote-cell endpoint converts a per-cell formula into a scoped rule.
"""
from __future__ import annotations

import os
import time
import uuid

import pytest
import requests

API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")


def _req(method, path, **kw):
    return getattr(requests, method)(f"{API}{path}", timeout=30, **kw)


# ──────────────────────────────────────────────────────────────────
# Fixture: build a throw-away model with fully-controlled structure
# ──────────────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def model():
    suffix = int(time.time() * 1000) % 100000
    name = f"ifr_test_{suffix}"

    # 1. Model
    r = _req("post", "/models", json={"name": name})
    assert r.status_code == 200, r.text
    mid = r.json()["id"]

    # 2. Analytics: Periods, PL (main), Dep
    def mk_analytic(payload):
        r = _req("post", "/analytics", json={**payload, "model_id": mid})
        assert r.status_code == 200, r.text
        return r.json()["id"]

    periods_aid = mk_analytic({
        "name": "Periods",
        "is_periods": True,
        "period_types": ["quarter", "month"],
        "period_start": "2026-10-01",
        "period_end": "2026-12-31",
    })
    # Period records are only auto-created by /generate-periods.
    r = _req("post", f"/analytics/{periods_aid}/generate-periods")
    assert r.status_code == 200, r.text

    pl_aid = mk_analytic({"name": "PL"})
    dep_aid = mk_analytic({"name": "Dep"})

    # 3. Indicator records on PL (3 leafs, no hierarchy needed).
    def mk_rec(aid, name, parent_id=None):
        r = _req("post", f"/analytics/{aid}/records",
                 json={"data_json": {"name": name}, "parent_id": parent_id})
        assert r.status_code == 200, r.text
        return r.json()["id"]

    partners_id = mk_rec(pl_aid, "партнёры")
    loans_id = mk_rec(pl_aid, "выдачи")
    avg_id = mk_rec(pl_aid, "сред/партнёра")

    # 4. Dep records: D1 parent, D11 and D12 leaves.
    d1_id = mk_rec(dep_aid, "D1")
    d11_id = mk_rec(dep_aid, "D11", parent_id=d1_id)
    d12_id = mk_rec(dep_aid, "D12", parent_id=d1_id)

    # 5. Sheet + analytic bindings (Periods, PL=main, Dep).
    r = _req("post", "/sheets", json={"model_id": mid, "name": "Test"})
    assert r.status_code == 200, r.text
    sid = r.json()["id"]

    def bind(aid, order):
        r = _req("post", f"/sheets/{sid}/analytics",
                 json={"analytic_id": aid, "sort_order": order})
        assert r.status_code == 200, r.text

    bind(periods_aid, 0)
    bind(pl_aid, 1)
    bind(dep_aid, 2)

    # Mark PL as main.
    r = _req("put", f"/sheets/{sid}/main-analytic",
             json={"analytic_id": pl_aid})
    assert r.status_code == 200, r.text
    assert r.json().get("ok"), r.text

    # 6. Find period leaf records (M10, M11, M12).
    r = _req("get", f"/analytics/{periods_aid}/records")
    periods = r.json()
    # Period records are a tree: Quarter → months. We want leaf months.
    by_id = {p["id"]: p for p in periods}
    parent_ids = {p["parent_id"] for p in periods if p.get("parent_id")}
    month_ids = [p["id"] for p in periods if p["id"] not in parent_ids]
    # Sort by sort_order for deterministic M10/M11/M12 identification.
    month_ids.sort(key=lambda i: by_id[i].get("sort_order", 0))
    assert len(month_ids) >= 3, f"expected 3 months, got {len(month_ids)}"
    m10, m11, m12 = month_ids[:3]

    out = {
        "model_id": mid,
        "sheet_id": sid,
        "periods_aid": periods_aid,
        "pl_aid": pl_aid,
        "dep_aid": dep_aid,
        "partners_id": partners_id,
        "loans_id": loans_id,
        "avg_id": avg_id,
        "d1_id": d1_id,
        "d11_id": d11_id,
        "d12_id": d12_id,
        "m10": m10, "m11": m11, "m12": m12,
    }
    yield out

    # teardown
    _req("delete", f"/models/{mid}")


def _coord(p, pl, dep):
    return f"{p}|{pl}|{dep}"


def _save_cell(sheet_id, coord, value):
    r = _req("put", f"/cells/by-sheet/{sheet_id}/single", json={
        "coord_key": coord,
        "value": str(value),
        "data_type": "number",
        "rule": "manual",
    })
    assert r.status_code == 200, r.text
    return r.json()


def _cells_by_coord(sheet_id):
    r = _req("get", f"/cells/by-sheet/{sheet_id}")
    assert r.status_code == 200, r.text
    return {c["coord_key"]: c for c in r.json()}


# ──────────────────────────────────────────────────────────────────
# Tests
# ──────────────────────────────────────────────────────────────────

def test_main_analytic_flag_set(model):
    r = _req("get", f"/sheets/{model['sheet_id']}/main-analytic")
    assert r.status_code == 200
    assert r.json()["analytic_id"] == model["pl_aid"]


def test_consolidation_ratio_recurses_correctly(model):
    """сред/партнёра consolidation = [выдачи] / [партнёры]. On HEAD (D1) must
    give sum(выдачи)/sum(партнёры), NOT avg of leaf ratios."""
    sid = model["sheet_id"]
    # Leaf cells for M10 × {partners, loans} × {D11, D12}.
    # D11: partners=10, loans=300 → ratio 30
    # D12: partners=15, loans=375 → ratio 25
    # HEAD D1 should be (300+375)/(10+15) = 675/25 = 27, not (30+25)/2 = 27.5
    _save_cell(sid, _coord(model["m10"], model["partners_id"], model["d11_id"]), 10)
    _save_cell(sid, _coord(model["m10"], model["partners_id"], model["d12_id"]), 15)
    _save_cell(sid, _coord(model["m10"], model["loans_id"], model["d11_id"]), 300)
    _save_cell(sid, _coord(model["m10"], model["loans_id"], model["d12_id"]), 375)

    # Pre-create HEAD-level cell for сред/партнёра so recalc has a row to update.
    head_avg_coord = _coord(model["m10"], model["avg_id"], model["d1_id"])
    _save_cell(sid, head_avg_coord, 0)

    # Install consolidation rule on сред/партнёра.
    r = _req("put", f"/sheets/{sid}/indicators/{model['avg_id']}/rules", json={
        "leaf": "",
        "consolidation": "[выдачи] / [партнёры]",
        "scoped": [],
    })
    assert r.status_code == 200, r.text

    # Trigger recalc by saving a no-op cell.
    _save_cell(sid, _coord(model["m10"], model["partners_id"], model["d11_id"]), 10)

    cells = _cells_by_coord(sid)
    head_val = float(cells[head_avg_coord]["value"])
    assert abs(head_val - 27.0) < 1e-6, (
        f"HEAD avg should be 675/25 = 27 (weighted), got {head_val}"
    )


def test_scoped_rule_beats_base_consolidation(model):
    """Scoped rule with priority 100 must beat base consolidation."""
    sid = model["sheet_id"]
    head_avg_m10 = _coord(model["m10"], model["avg_id"], model["d1_id"])

    r = _req("put", f"/sheets/{sid}/indicators/{model['avg_id']}/rules", json={
        "leaf": "",
        "consolidation": "[выдачи] / [партнёры]",
        "scoped": [{
            "scope": {model["periods_aid"]: model["m10"]},  # only M10 rows
            "priority": 100,
            "formula": "999",  # fixed value — obvious marker
        }],
    })
    assert r.status_code == 200, r.text

    # Ensure the HEAD cell exists so the update goes somewhere.
    _save_cell(sid, head_avg_m10, 0)
    cells = _cells_by_coord(sid)
    assert float(cells[head_avg_m10]["value"]) == 999.0

    # Reset to just consolidation for later tests.
    _req("put", f"/sheets/{sid}/indicators/{model['avg_id']}/rules", json={
        "leaf": "",
        "consolidation": "[выдачи] / [партнёры]",
        "scoped": [],
    })


def test_per_cell_formula_beats_indicator_rule(model):
    """Explicit cell_data.rule='formula' takes precedence over indicator rules."""
    sid = model["sheet_id"]
    head_avg_m10 = _coord(model["m10"], model["avg_id"], model["d1_id"])

    # Per-cell formula: always 42.
    r = _req("put", f"/cells/by-sheet/{sid}/single", json={
        "coord_key": head_avg_m10,
        "value": "0",
        "data_type": "number",
        "rule": "formula",
        "formula": "42",
    })
    assert r.status_code == 200, r.text

    cells = _cells_by_coord(sid)
    assert float(cells[head_avg_m10]["value"]) == 42.0

    # Clean up — set back to manual so other tests see the rule-driven value.
    _req("put", f"/cells/by-sheet/{sid}/single", json={
        "coord_key": head_avg_m10,
        "value": "0",
        "data_type": "number",
        "rule": "manual",
        "formula": "",
    })


def test_resolved_formulas_endpoint_reports_source(model):
    sid = model["sheet_id"]
    head_avg_m10 = _coord(model["m10"], model["avg_id"], model["d1_id"])
    leaf_avg_m10 = _coord(model["m10"], model["avg_id"], model["d11_id"])

    r = _req("post", f"/sheets/{sid}/cells/resolved-formulas", json={
        "coord_keys": [head_avg_m10, leaf_avg_m10],
    })
    assert r.status_code == 200, r.text
    by_ck = {x["coord_key"]: x for x in r.json()}
    head = by_ck[head_avg_m10]
    leaf = by_ck[leaf_avg_m10]
    assert head["formula"] == "[выдачи] / [партнёры]"
    assert head["source"].startswith("rule:")
    assert head["kind"] == "consolidation"
    # Leaf: no leaf rule installed → manual.
    assert leaf["source"] == "manual"


def test_promote_cell_creates_scoped_rule(model):
    sid = model["sheet_id"]
    # Target: M11, сред/партнёра, D1 (some HEAD cell).
    coord = _coord(model["m11"], model["avg_id"], model["d1_id"])

    # Put a per-cell formula first, then promote it.
    _req("put", f"/cells/by-sheet/{sid}/single", json={
        "coord_key": coord,
        "value": "0",
        "data_type": "number",
        "rule": "formula",
        "formula": "777",
    })

    r = _req("post",
             f"/sheets/{sid}/indicators/{model['avg_id']}/rules/promote-cell",
             json={"coord_key": coord, "formula": "777"})
    assert r.status_code == 200, r.text
    payload = r.json()
    assert payload.get("ok")
    assert model["periods_aid"] in payload["scope"]
    assert payload["scope"][model["periods_aid"]] == model["m11"]
    assert payload["scope"][model["dep_aid"]] == model["d1_id"]
    # PL (main) must NOT appear in the scope.
    assert model["pl_aid"] not in payload["scope"]

    # Verify the rule is now listed and per-cell override is cleared.
    r = _req("get", f"/sheets/{sid}/indicators/{model['avg_id']}/rules")
    rules = r.json()
    scoped = rules["scoped"]
    assert any(s["formula"] == "777" for s in scoped), f"rule not promoted: {scoped}"

    cells = _cells_by_coord(sid)
    cell = cells[coord]
    assert cell["rule"] == "manual"
    assert (cell.get("formula") or "") == ""
    assert float(cell["value"]) == 777.0  # rule-driven value

    # Cleanup: clear scoped rules.
    _req("put", f"/sheets/{sid}/indicators/{model['avg_id']}/rules", json={
        "leaf": "",
        "consolidation": "[выдачи] / [партнёры]",
        "scoped": [],
    })
