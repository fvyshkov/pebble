"""API tests: CRUD, recalculation, export."""
import pytest
import requests
import json
import os

API = "http://localhost:8000/api"


def _api(method, path, **kwargs):
    resp = getattr(requests, method)(f"{API}{path}", **kwargs, timeout=30)
    assert resp.status_code == 200, f"{method.upper()} {path}: {resp.status_code} {resp.text[:200]}"
    return resp.json()


@pytest.fixture(scope="module")
def test_model():
    """Create a test model, yield it, delete after."""
    m = _api("post", "/models", json={"name": "pytest_model"})
    yield m
    try:
        _api("delete", f"/models/{m['id']}")
    except Exception:
        pass


def test_list_models():
    data = _api("get", "/models")
    assert isinstance(data, list)
    assert len(data) >= 1


def test_create_update_delete_model():
    m = _api("post", "/models", json={"name": "tmp_test"})
    assert m["id"]
    _api("put", f"/models/{m['id']}", json={"name": "tmp_renamed", "description": "test"})
    tree = _api("get", f"/models/{m['id']}/tree")
    assert tree["name"] == "tmp_renamed"
    _api("delete", f"/models/{m['id']}")


def test_create_sheet(test_model):
    s = _api("post", "/sheets", json={"model_id": test_model["id"], "name": "test_sheet"})
    assert s["id"]
    sheets = _api("get", f"/sheets/by-model/{test_model['id']}")
    assert any(sh["name"] == "test_sheet" for sh in sheets)


def test_create_analytic(test_model):
    a = _api("post", "/analytics", json={"model_id": test_model["id"], "name": "test_analytic"})
    assert a["id"]


def test_recalculate():
    """Test recalculation on VERIFIED model."""
    import sqlite3
    db = sqlite3.connect("pebble.db")
    db.row_factory = sqlite3.Row
    model = db.execute("SELECT id FROM models WHERE name='VERIFIED'").fetchone()
    if not model:
        pytest.skip("VERIFIED model not found")
    sheet = db.execute("SELECT id FROM sheets WHERE model_id=? LIMIT 1", (model["id"],)).fetchone()
    result = _api("post", f"/cells/calculate/{sheet['id']}")
    assert result["computed"] > 0


def test_export_model():
    """Test Excel export."""
    import sqlite3
    db = sqlite3.connect("pebble.db")
    db.row_factory = sqlite3.Row
    model = db.execute("SELECT id FROM models WHERE name='VERIFIED'").fetchone()
    if not model:
        pytest.skip("VERIFIED model not found")
    resp = requests.get(f"{API}/excel/models/{model['id']}/export", timeout=30)
    assert resp.status_code == 200
    assert len(resp.content) > 1000  # should be a real xlsx
    assert "spreadsheet" in resp.headers.get("content-type", "")


def test_cell_save_triggers_recalc():
    """Saving a cell should trigger recalculation."""
    import sqlite3
    db = sqlite3.connect("pebble.db")
    db.row_factory = sqlite3.Row
    model = db.execute("SELECT id FROM models WHERE name='VERIFIED'").fetchone()
    if not model:
        pytest.skip("VERIFIED model not found")
    sheet = db.execute("SELECT id FROM sheets WHERE model_id=? LIMIT 1", (model["id"],)).fetchone()
    # Get a manual cell
    cell = db.execute(
        "SELECT coord_key FROM cell_data WHERE sheet_id=? AND rule='manual' LIMIT 1",
        (sheet["id"],)
    ).fetchone()
    if not cell:
        pytest.skip("No manual cells")
    old_val = db.execute("SELECT value FROM cell_data WHERE sheet_id=? AND coord_key=?",
                         (sheet["id"], cell["coord_key"])).fetchone()["value"]
    result = _api("put", f"/cells/by-sheet/{sheet['id']}", json={
        "cells": [{"coord_key": cell["coord_key"], "value": "999"}]
    })
    assert result.get("computed", 0) > 0
    # Restore original value
    _api("put", f"/cells/by-sheet/{sheet['id']}", json={
        "cells": [{"coord_key": cell["coord_key"], "value": old_val}]
    })
