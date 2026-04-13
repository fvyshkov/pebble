"""Test: admin hides a sheet from user → user doesn't see it."""
import pytest
import requests
import sqlite3
import time

API = "http://localhost:8000/api"


def _api(method, path, **kwargs):
    resp = getattr(requests, method)(f"{API}{path}", **kwargs, timeout=30)
    return resp


@pytest.fixture(scope="module")
def db():
    conn = sqlite3.connect("pebble.db")
    conn.row_factory = sqlite3.Row
    yield conn
    conn.close()


@pytest.fixture(scope="module")
def test_user():
    resp = _api("post", "/users", json={"username": f"hide_test_{int(time.time())}"})
    user = resp.json()
    yield user
    _api("delete", f"/users/{user['id']}")


@pytest.fixture(scope="module")
def verified_model(db):
    m = db.execute("SELECT id FROM models WHERE name='VERIFIED'").fetchone()
    if not m:
        pytest.skip("VERIFIED model not found")
    return m["id"]


def test_user_sees_all_sheets_by_default(test_user, verified_model):
    """New user sees all sheets by default."""
    resp = _api("get", f"/users/{test_user['id']}/accessible-sheets")
    assert resp.status_code == 200
    models = resp.json()
    verified = next((m for m in models if m["id"] == verified_model), None)
    assert verified, "VERIFIED model not visible"
    assert len(verified["sheets"]) == 7


def test_hide_sheet_from_user(test_user, verified_model, db):
    """Admin hides a sheet → user can't see it."""
    sheets = db.execute("SELECT id, name FROM sheets WHERE model_id=? ORDER BY sort_order", (verified_model,)).fetchall()
    hidden_sheet = sheets[0]  # Hide first sheet

    # Deny view
    resp = _api("put", f"/users/permissions/by-sheet/{hidden_sheet['id']}", json={
        "user_id": test_user["id"], "can_view": False, "can_edit": False,
    })
    assert resp.status_code == 200

    # User should see 6 sheets now
    resp = _api("get", f"/users/{test_user['id']}/accessible-sheets")
    models = resp.json()
    verified = next((m for m in models if m["id"] == verified_model), None)
    assert verified
    visible_ids = [s["id"] for s in verified["sheets"]]
    assert hidden_sheet["id"] not in visible_ids
    assert len(verified["sheets"]) == 6

    # Restore
    _api("put", f"/users/permissions/by-sheet/{hidden_sheet['id']}", json={
        "user_id": test_user["id"], "can_view": True, "can_edit": True,
    })


def test_hide_all_sheets_hides_model(test_user, verified_model, db):
    """If ALL sheets hidden → model disappears from list."""
    sheets = db.execute("SELECT id FROM sheets WHERE model_id=?", (verified_model,)).fetchall()

    # Hide all
    for s in sheets:
        _api("put", f"/users/permissions/by-sheet/{s['id']}", json={
            "user_id": test_user["id"], "can_view": False, "can_edit": False,
        })

    resp = _api("get", f"/users/{test_user['id']}/accessible-sheets")
    models = resp.json()
    verified = next((m for m in models if m["id"] == verified_model), None)
    assert verified is None, "Model should be hidden when all sheets are hidden"

    # Restore all
    for s in sheets:
        _api("put", f"/users/permissions/by-sheet/{s['id']}", json={
            "user_id": test_user["id"], "can_view": True, "can_edit": True,
        })


def test_read_only_sheet(test_user, verified_model, db):
    """Sheet with can_view=True, can_edit=False appears but is read-only."""
    sheets = db.execute("SELECT id FROM sheets WHERE model_id=? ORDER BY sort_order", (verified_model,)).fetchall()
    sheet = sheets[0]

    _api("put", f"/users/permissions/by-sheet/{sheet['id']}", json={
        "user_id": test_user["id"], "can_view": True, "can_edit": False,
    })

    resp = _api("get", f"/users/{test_user['id']}/accessible-sheets")
    models = resp.json()
    verified = next((m for m in models if m["id"] == verified_model), None)
    s = next((s for s in verified["sheets"] if s["id"] == sheet["id"]), None)
    assert s is not None, "Sheet should be visible"
    assert s["can_edit"] is False, "Sheet should be read-only"

    # Restore
    _api("put", f"/users/permissions/by-sheet/{sheet['id']}", json={
        "user_id": test_user["id"], "can_view": True, "can_edit": True,
    })
