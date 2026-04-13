"""Permission tests: sheet access, analytic record access, cell filtering."""
import pytest
import requests
import sqlite3

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
    import time
    resp = _api("post", "/users", json={"username": f"pytest_{int(time.time())}"})
    assert resp.status_code == 200, f"Create user failed: {resp.status_code} {resp.text}"
    user = resp.json()
    yield user
    _api("delete", f"/users/{user['id']}")


def test_list_users():
    resp = _api("get", "/users")
    assert resp.status_code == 200
    assert len(resp.json()) >= 1


def test_sheet_permissions(db, test_user):
    sheet = db.execute("SELECT id FROM sheets LIMIT 1").fetchone()
    if not sheet:
        pytest.skip("No sheets")
    # Deny view
    resp = _api("put", f"/users/permissions/by-sheet/{sheet['id']}", json={
        "user_id": test_user["id"], "can_view": False, "can_edit": False
    })
    assert resp.status_code == 200

    # Check accessible sheets
    resp = _api("get", f"/users/{test_user['id']}/accessible-sheets")
    accessible_ids = []
    for m in resp.json():
        for s in m["sheets"]:
            accessible_ids.append(s["id"])
    assert sheet["id"] not in accessible_ids

    # Restore
    _api("put", f"/users/permissions/by-sheet/{sheet['id']}", json={
        "user_id": test_user["id"], "can_view": True, "can_edit": True
    })


def test_analytic_permissions(db, test_user):
    rec = db.execute(
        "SELECT ar.id, ar.analytic_id FROM analytic_records ar "
        "JOIN analytics a ON a.id = ar.analytic_id WHERE a.is_periods = 0 LIMIT 1"
    ).fetchone()
    if not rec:
        pytest.skip("No analytic records")

    resp = _api("put", "/users/analytic-permissions/set", json={
        "user_id": test_user["id"],
        "analytic_id": rec["analytic_id"],
        "record_id": rec["id"],
        "can_view": True, "can_edit": False,
    })
    assert resp.status_code == 200

    # Check permissions returned
    resp = _api("get", f"/users/{test_user['id']}/analytic-permissions")
    assert resp.status_code == 200


def test_cell_filtering_by_user(db, test_user):
    sheet = db.execute("SELECT id FROM sheets LIMIT 1").fetchone()
    if not sheet:
        pytest.skip("No sheets")
    # Without user_id — all cells
    resp1 = _api("get", f"/cells/by-sheet/{sheet['id']}")
    # With user_id — filtered (or same if no restrictions)
    resp2 = _api("get", f"/cells/by-sheet/{sheet['id']}?user_id={test_user['id']}")
    assert resp1.status_code == 200
    assert resp2.status_code == 200
    assert len(resp2.json()) <= len(resp1.json())


def test_accessible_sheets_have_excel_code(db, test_user):
    resp = _api("get", f"/users/{test_user['id']}/accessible-sheets")
    assert resp.status_code == 200
    for model in resp.json():
        for sheet in model["sheets"]:
            # excel_code should be present (may be empty for old sheets)
            assert "excel_code" in sheet, f"Sheet {sheet['name']} missing excel_code"
