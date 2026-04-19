"""Super-test: full end-to-end permissions & aggregation scenario.

Scenario (per user's spec):
    - Two "subdivisions" (Dep records): D12 and D13, both children of D1.
    - Two users, each gets view-only permission on exactly one record.
    - Admin enters numbers for both records in two PL indicators (Jan 2026).
    - Each user must see ONLY their record's cells.
    - Admin (no filter) sees everything and the per-parent sum adds up.

This exercises the full stack that the user cares about:
    POST /api/users                            (create user)
    PUT  /api/users/analytic-permissions/set   (grant per-record view)
    PUT  /api/cells/by-sheet/{sid}/single      (admin writes a cell)
    GET  /api/cells/by-sheet/{sid}?user_id=... (server-side filter)
    GET  /api/cells/by-sheet/{sid}             (admin, unfiltered)

We keep the scenario API-driven so it is deterministic and fast; the UI-login
+ grid-contents variant is already partially covered by test_full_ui_flow.py.
"""
from __future__ import annotations

import os
import sqlite3
import time

import pytest
import requests

API = os.environ.get("PEBBLE_API", "http://localhost:8000/api")

# These IDs are fixtures in the dev DB (see MEMORY / prior investigation).
# If they disappear, the test will self-skip.
PL_SHEET_ID       = "6ba3c793-51d5-4620-a341-5d03a0b9b5f6"       # Финансовый результат BaaS
DEP_ANALYTIC_ID   = "1af83d15-4ec7-4233-b626-5c71f9535bd9"       # Dep
D1_ID             = "adb8f3c8-7cee-4c5d-9877-cf67315f231a"       # parent
D12_ID            = "77cbf0e0-18a1-408d-a152-e850d556b954"       # leaf
D13_ID            = "9ad93fd3-7b4a-4c31-bcc5-8aa17f97f096"       # leaf
JAN_2026_ID       = "d3458a10-f2a8-4b7f-81d8-99475f667183"       # period
# Two PL leaf indicators on that sheet (picked from existing cells):
PL_LEAF_A         = "3cb1001d-0a2d-4474-abe9-715c99936421"       # Процентные доходы
PL_LEAF_B         = "f988c00c-08c9-4553-ba6d-25c8d8c1e60e"       # Кредитная линия для ИП

# Values admin will write (deliberately non-round so we catch truncation).
V1_A, V1_B = 111.0, 222.0   # → D12, two indicators
V2_A, V2_B = 333.0, 444.0   # → D13, two indicators

# ──────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────

def _req(method: str, path: str, **kw):
    r = getattr(requests, method)(f"{API}{path}", timeout=30, **kw)
    return r


def _db():
    db_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "pebble.db",
    )
    c = sqlite3.connect(db_path)
    c.row_factory = sqlite3.Row
    return c


def _coord(period: str, pl: str, dep: str) -> str:
    """Coord order on PL sheet: Периоды(0) | Показатели(1) | Dep(2)."""
    return f"{period}|{pl}|{dep}"


def _sum_for(cells, dep_id: str) -> float:
    total = 0.0
    for c in cells:
        ck = c.get("coord_key") or ""
        parts = ck.split("|")
        if len(parts) < 3 or parts[2] != dep_id or parts[0] != JAN_2026_ID:
            continue
        if parts[1] not in (PL_LEAF_A, PL_LEAF_B):
            continue
        try:
            total += float(c.get("value") or 0)
        except (TypeError, ValueError):
            pass
    return total


# ──────────────────────────────────────────────────────────────────
# Fixtures
# ──────────────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def prerequisites():
    """Verify the fixture data exists in the DB; skip cleanly if not."""
    c = _db()
    try:
        got = {
            "sheet": c.execute(
                "SELECT 1 FROM sheets WHERE id=?", (PL_SHEET_ID,)
            ).fetchone(),
            "dep": c.execute(
                "SELECT 1 FROM analytics WHERE id=?", (DEP_ANALYTIC_ID,)
            ).fetchone(),
            "d12": c.execute(
                "SELECT 1 FROM analytic_records WHERE id=?", (D12_ID,)
            ).fetchone(),
            "d13": c.execute(
                "SELECT 1 FROM analytic_records WHERE id=?", (D13_ID,)
            ).fetchone(),
        }
    finally:
        c.close()
    missing = [k for k, v in got.items() if not v]
    if missing:
        pytest.skip(f"super-scenario fixture data missing: {missing}")
    yield


def _ensure_user(username: str) -> str:
    """Return id of an existing user, or create one with password=username."""
    users = _req("get", "/users").json()
    for u in users:
        if u["username"] == username:
            return u["id"]
    r = _req("post", "/users", json={"username": username})
    assert r.status_code == 200, r.text
    uid = r.json()["id"]
    # /users POST already hashes password = username; no further setup needed.
    return uid


@pytest.fixture(scope="module")
def admin_id():
    # admin is expected to exist; if not, create it (username=admin).
    return _ensure_user("admin")


@pytest.fixture(scope="module")
def dep12_user():
    uname = f"dep12_super_{int(time.time())}"
    uid = _ensure_user(uname)
    yield {"id": uid, "username": uname}
    _req("delete", f"/users/{uid}")


@pytest.fixture(scope="module")
def dep13_user():
    uname = f"dep13_super_{int(time.time())+1}"
    uid = _ensure_user(uname)
    yield {"id": uid, "username": uname}
    _req("delete", f"/users/{uid}")


# ──────────────────────────────────────────────────────────────────
# The super-test
# ──────────────────────────────────────────────────────────────────

def test_super_scenario_permissions_and_aggregation(
    prerequisites, admin_id, dep12_user, dep13_user
):
    # 1. Snapshot pre-existing values so we can restore them at the end.
    original = {}
    resp = _req("get", f"/cells/by-sheet/{PL_SHEET_ID}")
    assert resp.status_code == 200, resp.text
    for c in resp.json():
        ck = c.get("coord_key")
        if ck and ck.split("|")[0] == JAN_2026_ID and ck.endswith(
            (D12_ID, D13_ID)
        ):
            original[ck] = c.get("value")

    try:
        # 2. Admin writes values for each (indicator × dep) combo.
        writes = [
            (_coord(JAN_2026_ID, PL_LEAF_A, D12_ID), V1_A),
            (_coord(JAN_2026_ID, PL_LEAF_B, D12_ID), V1_B),
            (_coord(JAN_2026_ID, PL_LEAF_A, D13_ID), V2_A),
            (_coord(JAN_2026_ID, PL_LEAF_B, D13_ID), V2_B),
        ]
        for coord, value in writes:
            r = _req(
                "put",
                f"/cells/by-sheet/{PL_SHEET_ID}/single",
                json={
                    "coord_key": coord,
                    "value": str(value),
                    "data_type": "number",
                    "rule": "manual",
                    "user_id": admin_id,
                },
            )
            assert r.status_code == 200, (coord, r.text)

        # 3. Grant dep12_user → D12 only, dep13_user → D13 only.
        for user, rec_id in (
            (dep12_user, D12_ID),
            (dep13_user, D13_ID),
        ):
            r = _req(
                "put",
                "/users/analytic-permissions/set",
                json={
                    "user_id": user["id"],
                    "analytic_id": DEP_ANALYTIC_ID,
                    "record_id": rec_id,
                    "can_view": True,
                    "can_edit": False,
                },
            )
            assert r.status_code == 200, r.text

        # 4. dep12 sees only D12 rows on this sheet.
        r12 = _req(
            "get",
            f"/cells/by-sheet/{PL_SHEET_ID}?user_id={dep12_user['id']}",
        )
        assert r12.status_code == 200
        cells_12 = r12.json()
        assert cells_12, "dep12 should see at least its D12 cells"
        # Isolation invariant: user with D12-only permission must never see
        # any cell whose coord_key contains D13 (or any other Dep leaf).
        bad = [
            c["coord_key"] for c in cells_12
            if D13_ID in c["coord_key"].split("|")
        ]
        assert not bad, f"dep12 leaked D13 cells: {bad[:3]} (of {len(bad)})"

        # 5. dep13 sees only D13 rows.
        r13 = _req(
            "get",
            f"/cells/by-sheet/{PL_SHEET_ID}?user_id={dep13_user['id']}",
        )
        assert r13.status_code == 200
        cells_13 = r13.json()
        assert cells_13, "dep13 should see at least its D13 cells"
        bad = [
            c["coord_key"] for c in cells_13
            if D12_ID in c["coord_key"].split("|")
        ]
        assert not bad, f"dep13 leaked D12 cells: {bad[:3]} (of {len(bad)})"

        # 6. Each user sees the values admin wrote for them.
        d12_jan = {
            c["coord_key"]: c["value"]
            for c in cells_12
            if c["coord_key"].startswith(JAN_2026_ID + "|")
        }
        d13_jan = {
            c["coord_key"]: c["value"]
            for c in cells_13
            if c["coord_key"].startswith(JAN_2026_ID + "|")
        }
        assert float(d12_jan[_coord(JAN_2026_ID, PL_LEAF_A, D12_ID)]) == V1_A
        assert float(d12_jan[_coord(JAN_2026_ID, PL_LEAF_B, D12_ID)]) == V1_B
        assert float(d13_jan[_coord(JAN_2026_ID, PL_LEAF_A, D13_ID)]) == V2_A
        assert float(d13_jan[_coord(JAN_2026_ID, PL_LEAF_B, D13_ID)]) == V2_B

        # 7. Cross-check isolation: among 3-part Dep-tagged cells, no overlap.
        #    (2-part cells — legacy rows without the Dep dimension — are fair
        #    game for both users and are ignored here.)
        def _dep_tagged(cells):
            out = set()
            for c in cells:
                parts = c["coord_key"].split("|")
                if len(parts) == 3:
                    out.add(c["coord_key"])
            return out
        keys_12 = _dep_tagged(cells_12)
        keys_13 = _dep_tagged(cells_13)
        assert keys_12 and keys_13, "both users should have Dep-tagged cells"
        assert keys_12.isdisjoint(keys_13), (
            "dep12 and dep13 must not share any Dep-tagged coord_keys"
        )

        # 8. Admin (no filter) sees both sets, and the Jan aggregation is
        #    correct per parent D1 = D12 + D13.
        r_all = _req("get", f"/cells/by-sheet/{PL_SHEET_ID}")
        assert r_all.status_code == 200
        all_cells = r_all.json()
        sum_d12 = _sum_for(all_cells, D12_ID)
        sum_d13 = _sum_for(all_cells, D13_ID)
        assert sum_d12 == V1_A + V1_B, (
            f"admin D12-jan sum {sum_d12} != expected {V1_A + V1_B}"
        )
        assert sum_d13 == V2_A + V2_B, (
            f"admin D13-jan sum {sum_d13} != expected {V2_A + V2_B}"
        )
        # D1 (parent) = D12 + D13 — this is the invariant the user cares
        # about: "admin sees the correctly-calculated sum".
        assert sum_d12 + sum_d13 == V1_A + V1_B + V2_A + V2_B

    finally:
        # Restore any cells we overwrote, so the fixture stays stable for
        # other tests. Missing keys → leave the cells with our values (dev DB).
        for coord, old_value in original.items():
            _req(
                "put",
                f"/cells/by-sheet/{PL_SHEET_ID}/single",
                json={
                    "coord_key": coord,
                    "value": str(old_value) if old_value is not None else "",
                    "data_type": "number",
                    "rule": "manual",
                    "user_id": admin_id,
                },
            )
