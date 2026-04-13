"""Test: add department dimension to the VERIFIED model (all 7 sheets).

Steps:
  1. Use existing VERIFIED model with 20K+ cells and 7 sheets
  2. Create analytic "Подразделение": Головной → Филиал 1, Филиал 2
  3. Bind to all 7 sheets (sort_order=2)
  4. For each existing cell: copy to Филиал 1 (original), Филиал 2 (randomised ×0.5–1.5)
  5. Recalculate formulas for each branch independently
  6. Compute Головной = dep1 + dep2 for every coordinate
  7. Create users dep1, dep2 with analytic permissions
  8. Show results: admin (all 3 depts), dep1 (Филиал 1 only), dep2 (Филиал 2 only)

Model stays in the DB after test so it's visible in the UI.
"""

import pytest
import requests
import sqlite3
import json
import random

API = "http://localhost:8000/api"
MODEL_NAME = "VERIFIED"
DB_PATH = "pebble.db"


def _api(method, path, **kw):
    r = getattr(requests, method)(f"{API}{path}", **kw, timeout=120)
    assert r.status_code == 200, f"{method.upper()} {path} → {r.status_code}: {r.text[:300]}"
    return r


def _to_float(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


# ── Module-level setup: runs once ─────────────────────────────────────

def _has_dept_analytic():
    """Check if dept analytic already exists on VERIFIED model."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    row = conn.execute("""
        SELECT a.id FROM analytics a
        JOIN models m ON m.id = a.model_id
        WHERE m.name = ? AND a.code = 'dept'
    """, (MODEL_NAME,)).fetchone()
    conn.close()
    return row["id"] if row else None


def _get_model_id():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT id FROM models WHERE name=?", (MODEL_NAME,)).fetchone()
    conn.close()
    assert row, f"Model '{MODEL_NAME}' not found"
    return row["id"]


def _setup_departments():
    """Create dept analytic, bind to sheets, populate data. Idempotent."""
    model_id = _get_model_id()

    existing = _has_dept_analytic()
    if existing:
        # Already done — just return IDs
        return _extract_existing_ids(model_id, existing)

    # Create Подразделение analytic
    dept_a = _api("post", "/analytics", json={
        "model_id": model_id, "name": "Подразделение", "code": "dept",
        "data_type": "sum",
    }).json()
    aid = dept_a["id"]

    head = _api("post", f"/analytics/{aid}/records", json={
        "data_json": {"name": "Головной"},
    }).json()
    dep1 = _api("post", f"/analytics/{aid}/records", json={
        "parent_id": head["id"], "data_json": {"name": "Филиал 1"},
    }).json()
    dep2 = _api("post", f"/analytics/{aid}/records", json={
        "parent_id": head["id"], "data_json": {"name": "Филиал 2"},
    }).json()

    # Bind to all sheets
    sheets = _api("get", f"/sheets/by-model/{model_id}").json()
    for s in sheets:
        _api("post", f"/sheets/{s['id']}/analytics", json={
            "analytic_id": aid, "sort_order": 2,
        })

    # Populate cells for each sheet
    _populate_all_sheets(sheets, dep1["id"], dep2["id"], head["id"])

    # Create users
    _ensure_user("dep1", aid, dep1["id"])
    _ensure_user("dep2", aid, dep2["id"])

    return {
        "model_id": model_id,
        "dept_aid": aid,
        "head_id": head["id"],
        "dep1_id": dep1["id"],
        "dep2_id": dep2["id"],
        "sheets": sheets,
    }


def _populate_all_sheets(sheets, dep1_id, dep2_id, head_id):
    """For each sheet: read existing cells, create dep1/dep2/head copies."""
    random.seed(42)
    BATCH = 400

    for s in sheets:
        sid = s["id"]
        print(f"  Заполняю {s['name']}...", end="", flush=True)

        cells = _api("get", f"/cells/by-sheet/{sid}").json()
        if not cells:
            print(" (пусто)")
            continue

        dep1_cells = []
        dep2_cells = []

        for c in cells:
            ck = c["coord_key"]
            # Safety: skip if already has dept dimension
            if ck.count("|") >= 2:
                continue

            val = c.get("value")
            rule = c.get("rule") or "manual"
            formula = c.get("formula") or ""
            dtype = c.get("data_type") or "number"

            # dep1 = original
            dep1_cells.append({
                "coord_key": f"{ck}|{dep1_id}",
                "value": val, "data_type": dtype,
                "rule": rule, "formula": formula,
            })

            # dep2 = randomised manual, same formulas
            if rule == "manual" and val:
                try:
                    fval = float(val)
                    mult = random.uniform(0.5, 1.5)
                    if fval == int(fval):
                        new_val = str(round(fval * mult))
                    else:
                        new_val = str(round(fval * mult, 6))
                    dep2_cells.append({
                        "coord_key": f"{ck}|{dep2_id}",
                        "value": new_val, "data_type": dtype,
                        "rule": "manual", "formula": "",
                    })
                except ValueError:
                    dep2_cells.append({
                        "coord_key": f"{ck}|{dep2_id}",
                        "value": val, "data_type": dtype,
                        "rule": "manual", "formula": "",
                    })
            else:
                dep2_cells.append({
                    "coord_key": f"{ck}|{dep2_id}",
                    "value": val, "data_type": dtype,
                    "rule": rule, "formula": formula,
                })

        # Save dep1 + dep2
        for batch in [dep1_cells, dep2_cells]:
            for i in range(0, len(batch), BATCH):
                _api("put", f"/cells/by-sheet/{sid}", json={"cells": batch[i:i+BATCH]})

        print(f" dep1={len(dep1_cells)}, dep2={len(dep2_cells)}", flush=True)

    # Recalculate all sheets (formulas compute independently per dept branch)
    for s in sheets:
        print(f"  Расчёт {s['name']}...", end="", flush=True)
        r = _api("post", f"/cells/calculate/{s['id']}")
        print(f" computed={r.json().get('computed', 0)}")

    # Compute head = dep1 + dep2
    for s in sheets:
        sid = s["id"]
        all_cells = _api("get", f"/cells/by-sheet/{sid}").json()

        dep1_vals = {}
        dep2_vals = {}
        for c in all_cells:
            ck = c["coord_key"]
            if ck.endswith(f"|{dep1_id}"):
                dep1_vals[ck[:-(len(dep1_id)+1)]] = c["value"]
            elif ck.endswith(f"|{dep2_id}"):
                dep2_vals[ck[:-(len(dep2_id)+1)]] = c["value"]

        head_cells = []
        for base in set(dep1_vals) | set(dep2_vals):
            v1 = _to_float(dep1_vals.get(base, "0"))
            v2 = _to_float(dep2_vals.get(base, "0"))
            total = v1 + v2
            head_cells.append({
                "coord_key": f"{base}|{head_id}",
                "value": str(round(total, 6)) if total != 0 else "0",
                "data_type": "number", "rule": "manual",
            })

        for i in range(0, len(head_cells), BATCH):
            _api("put", f"/cells/by-sheet/{sid}", json={"cells": head_cells[i:i+BATCH]})

        print(f"  Головной {s['name']}: {len(head_cells)} ячеек")


def _ensure_user(username, dept_aid, dept_rid):
    r = requests.post(f"{API}/users", json={"username": username}, timeout=10)
    if r.status_code == 200:
        u = r.json()
    else:
        users = _api("get", "/users").json()
        u = next((x for x in users if x["username"] == username), None)
        assert u, f"Cannot create/find user {username}"
    _api("post", f"/users/{u['id']}/reset-password", json={"password": f"{username}pass"})
    _api("put", "/users/analytic-permissions/set", json={
        "user_id": u["id"], "analytic_id": dept_aid,
        "record_id": dept_rid,
        "can_view": True, "can_edit": True,
    })
    return u


def _extract_existing_ids(model_id, dept_aid):
    sheets = _api("get", f"/sheets/by-model/{model_id}").json()
    drecs = _api("get", f"/analytics/{dept_aid}/records").json()
    head = next(r for r in drecs if r["parent_id"] is None)
    children = sorted(
        [r for r in drecs if r["parent_id"] == head["id"]],
        key=lambda r: json.loads(r["data_json"])["name"] if isinstance(r["data_json"], str) else r["data_json"]["name"]
    )
    return {
        "model_id": model_id,
        "dept_aid": dept_aid,
        "head_id": head["id"],
        "dep1_id": children[0]["id"],
        "dep2_id": children[1]["id"],
        "sheets": sheets,
    }


# ── Fixtures ──────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def ctx():
    """Setup departments on VERIFIED model (idempotent)."""
    return _setup_departments()


@pytest.fixture(scope="module")
def users():
    all_users = _api("get", "/users").json()
    dep1 = next((u for u in all_users if u["username"] == "dep1"), None)
    dep2 = next((u for u in all_users if u["username"] == "dep2"), None)
    assert dep1 and dep2, "Users dep1/dep2 not found"
    return {"dep1": dep1, "dep2": dep2}


# ── Helpers ───────────────────────────────────────────────────────────

def _get_indicator_names(sheet_id):
    """Return {record_id: name} for non-period, non-dept analytics."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("""
        SELECT ar.id, json_extract(ar.data_json, '$.name') as name
        FROM sheet_analytics sa
        JOIN analytics a ON a.id = sa.analytic_id
        JOIN analytic_records ar ON ar.analytic_id = a.id
        WHERE sa.sheet_id = ? AND a.is_periods = 0 AND a.code != 'dept'
    """, (sheet_id,)).fetchall()
    conn.close()
    return {r["id"]: r["name"] for r in rows}


def _get_period_names(sheet_id):
    """Return {leaf_record_id: name} for period analytics."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("""
        SELECT ar.id, ar.parent_id, json_extract(ar.data_json, '$.name') as name
        FROM sheet_analytics sa
        JOIN analytics a ON a.id = sa.analytic_id
        JOIN analytic_records ar ON ar.analytic_id = a.id
        WHERE sa.sheet_id = ? AND a.is_periods = 1
    """, (sheet_id,)).fetchall()
    conn.close()
    parent_ids = {r["id"] for r in rows if any(x["parent_id"] == r["id"] for x in rows)}
    return {r["id"]: r["name"] for r in rows if r["id"] not in parent_ids}


def _print_sheet(title, cells, sheet_id, dept_ids, dept_names, max_periods=6, max_indicators=15):
    """Print readable grid for one sheet."""
    ind_names = _get_indicator_names(sheet_id)
    per_names = _get_period_names(sheet_id)

    # Build grid: {(period_id, ind_id, dept_id): value}
    grid = {}
    for c in cells:
        parts = c["coord_key"].split("|")
        if len(parts) < 3:
            continue
        pid, iid, did = parts[0], parts[1], parts[2]
        if did in dept_ids and pid in per_names and iid in ind_names:
            grid[(pid, iid, did)] = c["value"]

    if not grid:
        print(f"  (нет данных для {title})")
        return

    # Sort periods by name, take first N
    used_pids = sorted({p for p, _, _ in grid}, key=lambda p: per_names.get(p, ""))[:max_periods]
    used_iids = sorted({i for _, i, _ in grid}, key=lambda i: ind_names.get(i, ""))[:max_indicators]

    col_w = 13
    ind_w = 32

    print(f"\n{'=' * 90}")
    print(f"  {title}")
    print(f"{'=' * 90}")

    for did, dname in zip(dept_ids, dept_names):
        sub = {(p, i): grid.get((p, i, did), "") for p in used_pids for i in used_iids}
        if not any(sub.values()):
            continue

        short_per = {}
        for pid in used_pids:
            n = per_names[pid]
            parts = n.split()
            short_per[pid] = (parts[0][:3] + " " + parts[-1][-2:]) if len(parts) >= 2 else n[:8]

        print(f"\n  ── {dname} ──")
        header = f"  {'Показатель':<{ind_w}}"
        for pid in used_pids:
            header += f"{short_per[pid]:>{col_w}}"
        print(header)
        print(f"  {'─' * (ind_w + col_w * len(used_pids))}")

        for iid in used_iids:
            iname = ind_names[iid][:ind_w-1]
            row = f"  {iname:<{ind_w}}"
            has = False
            for pid in used_pids:
                val = grid.get((pid, iid, did), "")
                if val:
                    has = True
                    try:
                        fv = float(val)
                        if abs(fv) >= 1000:
                            row += f"{fv:>{col_w},.0f}"
                        elif fv == int(fv):
                            row += f"{int(fv):>{col_w}}"
                        else:
                            row += f"{fv:>{col_w},.4f}"
                    except ValueError:
                        row += f"{val:>{col_w}}"
                else:
                    row += f"{'':>{col_w}}"
            if has:
                print(row)


# ── Tests ─────────────────────────────────────────────────────────────

def test_dept_created(ctx):
    """Verify department analytic created and bound."""
    assert ctx["dept_aid"]
    assert ctx["head_id"]
    assert ctx["dep1_id"]
    assert ctx["dep2_id"]
    print(f"\n✓ Подразделение: Головной → Филиал 1, Филиал 2")
    print(f"  Привязано к {len(ctx['sheets'])} листам")


def test_cell_counts(ctx):
    """Show cell counts per sheet per department."""
    dep1_id, dep2_id, head_id = ctx["dep1_id"], ctx["dep2_id"], ctx["head_id"]

    print(f"\n{'=' * 70}")
    print(f"  {'Лист':<40} {'Ф1':>7} {'Ф2':>7} {'Гол':>7} {'Всего':>8}")
    print(f"  {'─' * 69}")

    total = [0, 0, 0, 0]
    for s in ctx["sheets"]:
        cells = _api("get", f"/cells/by-sheet/{s['id']}").json()
        cnt = {dep1_id: 0, dep2_id: 0, head_id: 0, "old": 0}
        for c in cells:
            parts = c["coord_key"].split("|")
            if len(parts) >= 3 and parts[2] in cnt:
                cnt[parts[2]] += 1
            elif len(parts) < 3:
                cnt["old"] += 1
        t = cnt[dep1_id] + cnt[dep2_id] + cnt[head_id]
        print(f"  {s['name'][:38]:<40} {cnt[dep1_id]:>7} {cnt[dep2_id]:>7} {cnt[head_id]:>7} {t:>8}")
        total[0] += cnt[dep1_id]; total[1] += cnt[dep2_id]; total[2] += cnt[head_id]; total[3] += t

    print(f"  {'─' * 69}")
    print(f"  {'ИТОГО':<40} {total[0]:>7} {total[1]:>7} {total[2]:>7} {total[3]:>8}")


def test_admin_view(ctx):
    """Admin: show one sheet with all 3 departments (first 6 periods, first 15 indicators)."""
    # Show BaaS.1 (кредитование) as example
    target = next(s for s in ctx["sheets"] if s.get("excel_code") == "BaaS.1")
    cells = _api("get", f"/cells/by-sheet/{target['id']}").json()

    dept_ids = [ctx["head_id"], ctx["dep1_id"], ctx["dep2_id"]]
    dept_names = ["Головной (свод)", "Филиал 1", "Филиал 2"]

    _print_sheet(f"АДМИН — {target['name']}", cells, target["id"], dept_ids, dept_names)

    dept_set = {c["coord_key"].split("|")[2] for c in cells if c["coord_key"].count("|") >= 2}
    assert ctx["head_id"] in dept_set
    assert ctx["dep1_id"] in dept_set
    assert ctx["dep2_id"] in dept_set


def test_dep1_view(ctx, users):
    """dep1: sees only Филиал 1 on BaaS.1."""
    target = next(s for s in ctx["sheets"] if s.get("excel_code") == "BaaS.1")
    uid = users["dep1"]["id"]
    cells = _api("get", f"/cells/by-sheet/{target['id']}?user_id={uid}").json()

    _print_sheet(f"dep1 — {target['name']}", cells, target["id"],
                 [ctx["dep1_id"]], ["Филиал 1"])

    for c in cells:
        parts = c["coord_key"].split("|")
        if len(parts) >= 3:
            assert parts[2] == ctx["dep1_id"], f"dep1 видит чужие: {parts[2]}"


def test_dep2_view(ctx, users):
    """dep2: sees only Филиал 2 on BaaS.1."""
    target = next(s for s in ctx["sheets"] if s.get("excel_code") == "BaaS.1")
    uid = users["dep2"]["id"]
    cells = _api("get", f"/cells/by-sheet/{target['id']}?user_id={uid}").json()

    _print_sheet(f"dep2 — {target['name']}", cells, target["id"],
                 [ctx["dep2_id"]], ["Филиал 2"])

    for c in cells:
        parts = c["coord_key"].split("|")
        if len(parts) >= 3:
            assert parts[2] == ctx["dep2_id"], f"dep2 видит чужие: {parts[2]}"


def test_head_is_sum(ctx):
    """Verify Головной = Ф1 + Ф2 across all sheets."""
    dep1_id, dep2_id, head_id = ctx["dep1_id"], ctx["dep2_id"], ctx["head_id"]
    total_checked = 0

    for s in ctx["sheets"]:
        cells = _api("get", f"/cells/by-sheet/{s['id']}").json()
        by_base = {}
        for c in cells:
            parts = c["coord_key"].split("|")
            if len(parts) < 3:
                continue
            base = "|".join(parts[:2])
            by_base.setdefault(base, {})[parts[2]] = c["value"]

        checked = 0
        for base, depts in by_base.items():
            if head_id in depts and dep1_id in depts and dep2_id in depts:
                h = _to_float(depts[head_id])
                d1 = _to_float(depts[dep1_id])
                d2 = _to_float(depts[dep2_id])
                exp = d1 + d2
                if abs(exp) > 0.001:
                    tol = abs(exp) * 0.01 + 0.01
                    assert abs(h - exp) <= tol, f"{s['name']}: Гол({h}) != Ф1({d1})+Ф2({d2}) at {base}"
                    checked += 1
        total_checked += checked

    assert total_checked > 0
    print(f"\n✓ Головной = Ф1 + Ф2: проверено {total_checked} ячеек по {len(ctx['sheets'])} листам")


def test_users_exist(users):
    """Verify dep1, dep2 users."""
    assert users["dep1"]["username"] == "dep1"
    assert users["dep2"]["username"] == "dep2"
    print(f"\n✓ dep1 (id={users['dep1']['id'][:8]}…), dep2 (id={users['dep2']['id'][:8]}…)")
