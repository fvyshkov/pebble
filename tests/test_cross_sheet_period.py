"""Regression test: cross-sheet references with period modifiers.

Verifies that `[Sheet::indicator](периоды="предыдущий")` correctly returns
the previous period's value, not the current one.

Bug: _resolve_cross_sheet ignored ref["params"], so period modifiers
like (периоды="предыдущий") had no effect — always returned current period.
Fix: applied in formula_engine.py, _resolve_cross_sheet now reads params
and navigates prev_period map.

Run: pytest tests/test_cross_sheet_period.py -x -q
"""
from __future__ import annotations

import os
import time

import pytest
import requests

API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")


def _req(method, path, **kw):
    return getattr(requests, method)(f"{API}{path}", timeout=30, **kw)


def _mk_analytic(model_id, name, is_periods=False):
    r = _req("post", "/analytics", json={"model_id": model_id, "name": name, "is_periods": is_periods})
    assert r.status_code == 200, r.text
    return r.json()["id"]


def _mk_rec(aid, name, parent_id=None):
    r = _req("post", f"/analytics/{aid}/records",
             json={"data_json": {"name": name}, "parent_id": parent_id})
    assert r.status_code == 200, r.text
    return r.json()["id"]


@pytest.fixture(scope="module")
def model():
    """Build a 2-sheet model:

    Sheet A ("Параметры"): periods P1,P2,P3 × indicator "ставка"
      - P1: 100, P2: 200, P3: 300

    Sheet B ("Расчёт"): same periods × indicators:
      - "дельта": formula = [Параметры::ставка] - [Параметры::ставка](периоды="предыдущий")
        Expected: P1=100 (prev=0), P2=100 (200-100), P3=100 (300-200)
      - "сумма_назад2": formula = [Параметры::ставка](периоды="назад(2)")
        Expected: P1=0, P2=0, P3=100
    """
    suffix = int(time.time() * 1000) % 100000
    name = f"xsheet_period_test_{suffix}"

    # 1. Model
    r = _req("post", "/models", json={"name": name})
    assert r.status_code == 200, r.text
    mid = r.json()["id"]

    # 2. Shared periods analytic with manual records
    periods_aid = _mk_analytic(mid, "Периоды", is_periods=True)
    p1 = _mk_rec(periods_aid, "P1")
    p2 = _mk_rec(periods_aid, "P2")
    p3 = _mk_rec(periods_aid, "P3")

    # 3. Indicator analytics
    ind_a_aid = _mk_analytic(mid, "Показатели A")
    stavka_rid = _mk_rec(ind_a_aid, "ставка")

    ind_b_aid = _mk_analytic(mid, "Показатели B")
    delta_rid = _mk_rec(ind_b_aid, "дельта")
    nazad2_rid = _mk_rec(ind_b_aid, "сумма_назад2")

    # 4. Sheet A: "Параметры"
    r = _req("post", "/sheets", json={"model_id": mid, "name": "Параметры"})
    assert r.status_code == 200, r.text
    sheet_a = r.json()["id"]

    def bind(sid, aid, order):
        r = _req("post", f"/sheets/{sid}/analytics",
                 json={"analytic_id": aid, "sort_order": order})
        assert r.status_code == 200, r.text

    bind(sheet_a, periods_aid, 0)
    bind(sheet_a, ind_a_aid, 1)

    # 5. Sheet B: "Расчёт"
    r = _req("post", "/sheets", json={"model_id": mid, "name": "Расчёт"})
    assert r.status_code == 200, r.text
    sheet_b = r.json()["id"]

    bind(sheet_b, periods_aid, 0)
    bind(sheet_b, ind_b_aid, 1)

    # 6. Fill Sheet A: ставка P1=100, P2=200, P3=300
    cells_a = []
    for period_rid, val in [(p1, "100"), (p2, "200"), (p3, "300")]:
        ck = f"{period_rid}|{stavka_rid}"
        cells_a.append({"coord_key": ck, "value": val})
    r = _req("put", f"/cells/by-sheet/{sheet_a}",
             json={"cells": cells_a}, params={"no_recalc": "true"})
    assert r.status_code == 200, r.text

    # 7. Fill Sheet B with formulas (no_recalc so we control when calc runs)
    delta_formula = '[Параметры::ставка] - [Параметры::ставка](периоды="предыдущий")'
    nazad2_formula = '[Параметры::ставка](периоды="назад(2)")'

    cells_b = []
    for period_rid in [p1, p2, p3]:
        cells_b.append({
            "coord_key": f"{period_rid}|{delta_rid}",
            "value": "", "formula": delta_formula, "rule": "formula",
        })
        cells_b.append({
            "coord_key": f"{period_rid}|{nazad2_rid}",
            "value": "", "formula": nazad2_formula, "rule": "formula",
        })
    r = _req("put", f"/cells/by-sheet/{sheet_b}",
             json={"cells": cells_b}, params={"no_recalc": "true"})
    assert r.status_code == 200, r.text

    yield {
        "model_id": mid,
        "sheet_a": sheet_a,
        "sheet_b": sheet_b,
        "periods": {"P1": p1, "P2": p2, "P3": p3},
        "delta_rid": delta_rid,
        "nazad2_rid": nazad2_rid,
    }

    # Cleanup
    _req("delete", f"/models/{mid}")


def _get_cells(sheet_id):
    """Return {coord_key: value} for all cells in sheet (uuid-form keys).

    The API stores coord_keys in seq_id form for compactness; we re-expand to
    uuid form here so tests can look up by the same key they used on PUT.
    """
    bindings = _req("get", f"/sheets/{sheet_id}/analytics").json()
    seq_to_uuid: dict[str, str] = {}
    for b in bindings:
        aid = b.get("analytic_id") or b.get("id")
        if not aid:
            continue
        recs = _req("get", f"/analytics/{aid}/records").json()
        for r in recs:
            if r.get("seq_id") is not None:
                seq_to_uuid[str(r["seq_id"])] = r["id"]
    r = _req("get", f"/cells/by-sheet/{sheet_id}")
    assert r.status_code == 200, r.text
    result = {}
    for c in r.json():
        ck = c["coord_key"]
        uuid_ck = "|".join(seq_to_uuid.get(p, p) for p in ck.split("|"))
        try:
            result[uuid_ck] = float(c.get("value", 0) or 0)
        except (ValueError, TypeError):
            result[uuid_ck] = 0.0
    return result


def test_cross_sheet_previous_period(model):
    """[Sheet::indicator](периоды="предыдущий") returns previous period value."""
    # Trigger recalculation
    r = _req("post", f"/cells/calculate/{model['sheet_b']}")
    assert r.status_code == 200, r.text

    cells = _get_cells(model["sheet_b"])
    delta_rid = model["delta_rid"]
    periods = model["periods"]

    v1 = cells.get(f"{periods['P1']}|{delta_rid}", 0)
    v2 = cells.get(f"{periods['P2']}|{delta_rid}", 0)
    v3 = cells.get(f"{periods['P3']}|{delta_rid}", 0)

    # дельта = current - previous
    # P1: 100 - 0 = 100  (no previous period → 0)
    # P2: 200 - 100 = 100
    # P3: 300 - 200 = 100
    assert abs(v1 - 100.0) < 0.01, f"P1 delta should be 100, got {v1}"
    assert abs(v2 - 100.0) < 0.01, f"P2 delta should be 100, got {v2}"
    assert abs(v3 - 100.0) < 0.01, f"P3 delta should be 100, got {v3}"


def test_cross_sheet_nazad_n(model):
    """[Sheet::indicator](периоды="назад(2)") returns value 2 periods back."""
    cells = _get_cells(model["sheet_b"])
    nazad2_rid = model["nazad2_rid"]
    periods = model["periods"]

    v1 = cells.get(f"{periods['P1']}|{nazad2_rid}", 0)
    v2 = cells.get(f"{periods['P2']}|{nazad2_rid}", 0)
    v3 = cells.get(f"{periods['P3']}|{nazad2_rid}", 0)

    # назад(2): value from 2 periods back
    # P1: no period 2 back → 0
    # P2: no period 2 back → 0
    # P3: P1 value = 100
    assert abs(v1) < 0.01, f"P1 назад(2) should be 0, got {v1}"
    assert abs(v2) < 0.01, f"P2 назад(2) should be 0, got {v2}"
    assert abs(v3 - 100.0) < 0.01, f"P3 назад(2) should be 100, got {v3}"
