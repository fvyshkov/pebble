"""Slice consolidation: when leaf-indicator values are placed on ONE
specific (department × version) slice, the engine must compute parent-indicator
and parent-period consolidations *on that same slice* — and leave other slices
untouched (zero/empty).

This is the core invariant for arbitrary added analytics: dropping data into
a single concrete combo of non-period dimensions should reproduce within that
combo the same totals you would see in a 2D model. Other combos stay clean.

Run:
  PEBBLE_API=http://127.0.0.1:8011/api pytest tests/test_slice_consolidation.py -x -s
"""
from __future__ import annotations

import os
import time
import pytest
import requests

API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")


def _req(method, path, **kw):
    return getattr(requests, method)(f"{API}{path}", timeout=30, **kw)


def _ok(r, msg=""):
    assert r.status_code == 200, f"{msg}: {r.status_code} {r.text[:300]}"
    return r


def _cells(sheet_id):
    return {c["coord_key"]: c for c in _ok(_req("get", f"/cells/by-sheet/{sheet_id}")).json()}


@pytest.fixture(scope="module")
def slice_model():
    """Sheet with 4 axes:
        period (Q1 + months), indicator (Itog parent → A, B, C),
        dept (D1, D2), ver (V1, V2).
    """
    suffix = int(time.time() * 1000) % 100000
    name = f"slice_consol_test_{suffix}"

    mid = _ok(_req("post", "/models", json={"name": name})).json()["id"]

    # Periods Q1 + months
    p_aid = _ok(_req("post", "/analytics", json={
        "model_id": mid, "name": "Periods", "is_periods": True,
        "period_types": ["quarter", "month"],
        "period_start": "2027-01-01", "period_end": "2027-03-31",
    })).json()["id"]
    _ok(_req("post", f"/analytics/{p_aid}/generate-periods"))

    # Indicator analytic with parent Itog and 3 children
    ind_aid = _ok(_req("post", "/analytics", json={
        "model_id": mid, "name": "Ind",
    })).json()["id"]

    def mk(aid, name, parent=None):
        return _ok(_req("post", f"/analytics/{aid}/records",
                        json={"data_json": {"name": name}, "parent_id": parent})
                   ).json()["id"]

    itog = mk(ind_aid, "Itog")
    a = mk(ind_aid, "A", parent=itog)
    b = mk(ind_aid, "B", parent=itog)
    c = mk(ind_aid, "C", parent=itog)

    # Sheet, bind period+indicator first
    sid = _ok(_req("post", "/sheets", json={"model_id": mid, "name": "S"})).json()["id"]
    _ok(_req("post", f"/sheets/{sid}/analytics",
             json={"analytic_id": p_aid, "sort_order": 0}))
    _ok(_req("post", f"/sheets/{sid}/analytics",
             json={"analytic_id": ind_aid, "sort_order": 1}))
    _ok(_req("put", f"/sheets/{sid}/main-analytic",
             json={"analytic_id": ind_aid}))

    # Add dept + version analytics (no children — leaf-only, no consolidation)
    dept_aid = _ok(_req("post", "/analytics", json={"model_id": mid, "name": "Dept"})).json()["id"]
    d1 = mk(dept_aid, "D1")
    d2 = mk(dept_aid, "D2")

    ver_aid = _ok(_req("post", "/analytics", json={"model_id": mid, "name": "Ver"})).json()["id"]
    v1 = mk(ver_aid, "V1")
    v2 = mk(ver_aid, "V2")

    _ok(_req("post", f"/sheets/{sid}/analytics",
             json={"analytic_id": dept_aid, "sort_order": 2}))
    _ok(_req("post", f"/sheets/{sid}/analytics",
             json={"analytic_id": ver_aid, "sort_order": 3}))

    # Find leaf months
    periods = _ok(_req("get", f"/analytics/{p_aid}/records")).json()
    by_id = {p["id"]: p for p in periods}
    parent_ids = {p["parent_id"] for p in periods if p.get("parent_id")}
    months = sorted(
        [p["id"] for p in periods if p["id"] not in parent_ids],
        key=lambda i: by_id[i].get("sort_order", 0),
    )
    quarter = next(p["id"] for p in periods if p["id"] in parent_ids)
    assert len(months) >= 3
    m1, m2, m3 = months[:3]

    out = {
        "model_id": mid, "sheet_id": sid,
        "p_aid": p_aid, "ind_aid": ind_aid,
        "itog": itog, "a": a, "b": b, "c": c,
        "d1": d1, "d2": d2, "v1": v1, "v2": v2,
        "m1": m1, "m2": m2, "m3": m3, "quarter": quarter,
    }
    yield out
    _req("delete", f"/models/{mid}")


def _put(sid, coord, value):
    _ok(_req("put", f"/cells/by-sheet/{sid}/single", json={
        "coord_key": coord, "value": str(value),
        "data_type": "number", "rule": "manual",
    }), f"put {coord}={value}")


def test_slice_consolidation_after_recalc(slice_model):
    """Place A=10, B=20, C=30 on (D1, V1) for each month.
    After recalc:
      - leaf cells on (D1,V1) keep their values
      - Itog × month × (D1,V1) = 60
      - Itog × Q1 × (D1,V1) = 180
      - any-indicator × any-period × (D2,*) or (*,V2) = 0/missing
    """
    m = slice_model
    sid = m["sheet_id"]

    for mid_p in [m["m1"], m["m2"], m["m3"]]:
        _put(sid, f"{mid_p}|{m['a']}|{m['d1']}|{m['v1']}", 10)
        _put(sid, f"{mid_p}|{m['b']}|{m['d1']}|{m['v1']}", 20)
        _put(sid, f"{mid_p}|{m['c']}|{m['d1']}|{m['v1']}", 30)

    _ok(_req("post", f"/cells/calculate/{sid}"), "recalc")
    cells = _cells(sid)

    # Leaf cells preserved on (D1,V1)
    for mid_p in [m["m1"], m["m2"], m["m3"]]:
        for ind, expected in [(m["a"], 10), (m["b"], 20), (m["c"], 30)]:
            ck = f"{mid_p}|{ind}|{m['d1']}|{m['v1']}"
            cell = cells.get(ck)
            assert cell, f"missing leaf {ck}"
            assert abs(float(cell["value"]) - expected) < 0.01, \
                f"{ck}: expected {expected}, got {cell['value']}"

    # Parent-indicator on (D1,V1) for each month = 60
    for mid_p in [m["m1"], m["m2"], m["m3"]]:
        ck = f"{mid_p}|{m['itog']}|{m['d1']}|{m['v1']}"
        cell = cells.get(ck)
        assert cell, f"Itog × month × (D1,V1) missing: {ck}"
        assert abs(float(cell["value"]) - 60) < 0.01, \
            f"Itog={cell['value']}, expected 60"

    # Parent-period (quarter) × Itog × (D1,V1) = 180
    ck_q_itog = f"{m['quarter']}|{m['itog']}|{m['d1']}|{m['v1']}"
    cell = cells.get(ck_q_itog)
    assert cell, f"Q1 × Itog × (D1,V1) missing: {ck_q_itog}"
    assert abs(float(cell["value"]) - 180) < 0.01, \
        f"Q1 × Itog: {cell['value']}, expected 180"

    # Parent-period × leaf-indicator A on (D1,V1) = 30 (10+10+10)
    ck_q_a = f"{m['quarter']}|{m['a']}|{m['d1']}|{m['v1']}"
    cell = cells.get(ck_q_a)
    assert cell, f"Q1 × A × (D1,V1) missing"
    assert abs(float(cell["value"]) - 30) < 0.01, \
        f"Q1 × A: {cell['value']}, expected 30"


def test_other_slices_remain_zero(slice_model):
    """Cells on (D2,*) and (*,V2) slices should have no positive values."""
    m = slice_model
    sid = m["sheet_id"]
    cells = _cells(sid)

    # Any cell where dept=D2 OR ver=V2 should be 0 or empty
    for ck, cell in cells.items():
        parts = ck.split("|")
        if len(parts) != 4: continue
        period, ind, dept, ver = parts
        if dept == m["d2"] or ver == m["v2"]:
            try:
                v = float(cell["value"])
            except (ValueError, TypeError):
                continue
            assert abs(v) < 0.01, \
                f"unexpected non-zero on inactive slice {ck} = {cell['value']}"
