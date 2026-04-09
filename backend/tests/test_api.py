"""Fast API tests covering all major CRUD flows."""
import pytest
from httpx import AsyncClient

pytestmark = pytest.mark.asyncio


# ──── Models ────────────────────────────────────────────────────

async def test_create_and_list_models(client: AsyncClient):
    r = await client.post("/api/models", json={"name": "Бюджет", "description": "Тест"})
    assert r.status_code == 200
    model = r.json()
    assert model["name"] == "Бюджет"
    assert model["id"]

    r = await client.get("/api/models")
    assert r.status_code == 200
    assert any(m["id"] == model["id"] for m in r.json())


async def test_update_model(client: AsyncClient):
    r = await client.post("/api/models", json={"name": "M1"})
    mid = r.json()["id"]
    r = await client.put(f"/api/models/{mid}", json={"name": "M1-Updated"})
    assert r.status_code == 200
    assert r.json()["name"] == "M1-Updated"


async def test_delete_model(client: AsyncClient):
    r = await client.post("/api/models", json={"name": "ToDelete"})
    mid = r.json()["id"]
    r = await client.delete(f"/api/models/{mid}")
    assert r.status_code == 200
    r = await client.get("/api/models")
    assert not any(m["id"] == mid for m in r.json())


async def test_model_tree(client: AsyncClient):
    r = await client.post("/api/models", json={"name": "TreeTest"})
    mid = r.json()["id"]
    r = await client.get(f"/api/models/{mid}/tree")
    assert r.status_code == 200
    tree = r.json()
    assert tree["id"] == mid
    assert "sheets" in tree
    assert "analytics" in tree


# ──── Analytics ─────────────────────────────────────────────────

async def test_analytics_crud(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "AM"})).json()
    mid = m["id"]

    # Create
    r = await client.post("/api/analytics", json={"model_id": mid, "name": "Продукты", "code": "products"})
    assert r.status_code == 200
    a = r.json()
    aid = a["id"]
    assert a["name"] == "Продукты"

    # List
    r = await client.get(f"/api/analytics/by-model/{mid}")
    assert len(r.json()) >= 1

    # Get
    r = await client.get(f"/api/analytics/{aid}")
    assert r.json()["code"] == "products"

    # Update
    r = await client.put(f"/api/analytics/{aid}", json={"name": "Продукты 2"})
    assert r.json()["name"] == "Продукты 2"

    # Delete
    r = await client.delete(f"/api/analytics/{aid}")
    assert r.status_code == 200


# ──── Analytic Fields ───────────────────────────────────────────

async def test_fields_crud(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "FM"})).json()
    a = (await client.post("/api/analytics", json={"model_id": m["id"], "name": "F", "code": "f"})).json()
    aid = a["id"]

    r = await client.post(f"/api/analytics/{aid}/fields", json={"name": "Цена", "code": "price", "data_type": "number"})
    assert r.status_code == 200
    fid = r.json()["id"]

    r = await client.get(f"/api/analytics/{aid}/fields")
    assert len(r.json()) >= 1

    r = await client.put(f"/api/analytics/{aid}/fields/{fid}", json={"name": "Цена2"})
    assert r.json()["name"] == "Цена2"

    r = await client.delete(f"/api/analytics/{aid}/fields/{fid}")
    assert r.status_code == 200


# ──── Analytic Records ──────────────────────────────────────────

async def test_records_crud(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "RM"})).json()
    a = (await client.post("/api/analytics", json={"model_id": m["id"], "name": "R", "code": "r"})).json()
    aid = a["id"]

    r = await client.post(f"/api/analytics/{aid}/records", json={"data_json": {"name": "Запись 1"}})
    assert r.status_code == 200
    rid = r.json()["id"]

    # Child record
    r = await client.post(f"/api/analytics/{aid}/records", json={"parent_id": rid, "data_json": {"name": "Потомок"}})
    child_id = r.json()["id"]
    assert r.json()["parent_id"] == rid

    r = await client.get(f"/api/analytics/{aid}/records")
    assert len(r.json()) == 2

    r = await client.put(f"/api/analytics/{aid}/records/{rid}", json={"data_json": {"name": "Обновлено"}})
    assert r.status_code == 200

    r = await client.delete(f"/api/analytics/{aid}/records/{child_id}")
    assert r.status_code == 200


async def test_records_bulk(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "BM"})).json()
    a = (await client.post("/api/analytics", json={"model_id": m["id"], "name": "B", "code": "b"})).json()
    aid = a["id"]

    # bulk endpoint expects a raw list
    r = await client.post(f"/api/analytics/{aid}/records/bulk", json=[
        {"data_json": {"name": "A"}},
        {"data_json": {"name": "B"}},
        {"data_json": {"name": "C"}},
    ])
    assert r.status_code == 200
    assert len(r.json()["created"]) == 3


async def test_generate_periods(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "PM"})).json()
    a = (await client.post("/api/analytics", json={
        "model_id": m["id"], "name": "Periods", "code": "periods",
        "is_periods": True, "period_types": ["year", "quarter", "month"],
        "period_start": "2026-01-01", "period_end": "2026-06-30",
    })).json()
    aid = a["id"]

    r = await client.post(f"/api/analytics/{aid}/generate-periods")
    assert r.status_code == 200
    recs = r.json()
    # Should have: 1 year + 2 quarters + 6 months = 9
    assert len(recs) >= 9
    # Check hierarchy: first record should have no parent (year level)
    years = [rec for rec in recs if rec["parent_id"] is None]
    assert len(years) == 1


# ──── Sheets ────────────────────────────────────────────────────

async def test_sheets_crud(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "SM"})).json()
    mid = m["id"]

    r = await client.post("/api/sheets", json={"model_id": mid, "name": "Лист 1"})
    assert r.status_code == 200
    sid = r.json()["id"]

    r = await client.get(f"/api/sheets/by-model/{mid}")
    assert len(r.json()) >= 1

    r = await client.put(f"/api/sheets/{sid}", json={"name": "Лист обновлен"})
    assert r.json()["name"] == "Лист обновлен"

    r = await client.delete(f"/api/sheets/{sid}")
    assert r.status_code == 200


async def test_sheet_analytics_binding(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "SAM"})).json()
    a1 = (await client.post("/api/analytics", json={"model_id": m["id"], "name": "A1", "code": "a1"})).json()
    a2 = (await client.post("/api/analytics", json={"model_id": m["id"], "name": "A2", "code": "a2"})).json()
    s = (await client.post("/api/sheets", json={"model_id": m["id"], "name": "S"})).json()
    sid = s["id"]

    # Add analytics
    r = await client.post(f"/api/sheets/{sid}/analytics", json={"analytic_id": a1["id"]})
    sa1_id = r.json()["id"]
    r = await client.post(f"/api/sheets/{sid}/analytics", json={"analytic_id": a2["id"]})
    sa2_id = r.json()["id"]

    r = await client.get(f"/api/sheets/{sid}/analytics")
    assert len(r.json()) == 2

    # Reorder
    r = await client.put(f"/api/sheets/{sid}/analytics-reorder", json={"ordered_ids": [sa2_id, sa1_id]})
    assert r.status_code == 200

    # Remove
    r = await client.delete(f"/api/sheets/{sid}/analytics/{sa1_id}")
    assert r.status_code == 200
    r = await client.get(f"/api/sheets/{sid}/analytics")
    assert len(r.json()) == 1


async def test_view_settings(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "VSM"})).json()
    s = (await client.post("/api/sheets", json={"model_id": m["id"], "name": "VS"})).json()
    sid = s["id"]

    settings = {"order": ["id1", "id2"], "pinned": {"id1": "val"}}
    r = await client.put(f"/api/sheets/{sid}/view-settings", json={"settings": settings})
    assert r.status_code == 200

    r = await client.get(f"/api/sheets/{sid}/view-settings")
    data = r.json()
    assert data["order"] == ["id1", "id2"]


# ──── Cells ─────────────────────────────────────────────────────

async def test_cells_save_and_read(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "CM"})).json()
    s = (await client.post("/api/sheets", json={"model_id": m["id"], "name": "CS"})).json()
    sid = s["id"]

    r = await client.put(f"/api/cells/by-sheet/{sid}", json={"cells": [
        {"coord_key": "a|b", "value": "100", "data_type": "sum"},
        {"coord_key": "a|c", "value": "200", "data_type": "sum"},
    ]})
    assert r.status_code == 200

    r = await client.get(f"/api/cells/by-sheet/{sid}")
    cells = r.json()
    assert len(cells) == 2
    vals = {c["coord_key"]: c["value"] for c in cells}
    assert vals["a|b"] == "100"
    assert vals["a|c"] == "200"


async def test_cell_update_overwrites(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "CU"})).json()
    s = (await client.post("/api/sheets", json={"model_id": m["id"], "name": "CU"})).json()
    sid = s["id"]

    await client.put(f"/api/cells/by-sheet/{sid}", json={"cells": [{"coord_key": "x|y", "value": "10", "data_type": "sum"}]})
    await client.put(f"/api/cells/by-sheet/{sid}", json={"cells": [{"coord_key": "x|y", "value": "20", "data_type": "sum"}]})

    r = await client.get(f"/api/cells/by-sheet/{sid}")
    vals = {c["coord_key"]: c["value"] for c in r.json()}
    assert vals["x|y"] == "20"


async def test_cell_history(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "CH"})).json()
    s = (await client.post("/api/sheets", json={"model_id": m["id"], "name": "CH"})).json()
    sid = s["id"]

    await client.put(f"/api/cells/by-sheet/{sid}", json={"cells": [{"coord_key": "h|k", "value": "1", "data_type": "sum"}]})
    await client.put(f"/api/cells/by-sheet/{sid}", json={"cells": [{"coord_key": "h|k", "value": "2", "data_type": "sum"}]})

    r = await client.get(f"/api/cells/history/{sid}/h%7Ck")
    assert r.status_code == 200
    assert len(r.json()) >= 1


# ──── Users ─────────────────────────────────────────────────────

async def test_users_crud(client: AsyncClient):
    r = await client.get("/api/users")
    initial = len(r.json())

    r = await client.post("/api/users", json={"username": "TestUser"})
    assert r.status_code == 200
    uid = r.json()["id"]

    r = await client.get("/api/users")
    assert len(r.json()) == initial + 1

    r = await client.delete(f"/api/users/{uid}")
    assert r.status_code == 200


async def test_sheet_permissions(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "PP"})).json()
    s = (await client.post("/api/sheets", json={"model_id": m["id"], "name": "PP"})).json()
    sid = s["id"]
    u = (await client.post("/api/users", json={"username": "PermUser"})).json()
    uid = u["id"]

    r = await client.put(f"/api/users/permissions/by-sheet/{sid}", json={
        "user_id": uid, "can_view": True, "can_edit": False,
    })
    assert r.status_code == 200

    r = await client.get(f"/api/users/permissions/by-sheet/{sid}")
    perms = r.json()
    user_perm = next((p for p in perms if p["user_id"] == uid), None)
    assert user_perm is not None
    assert user_perm["can_edit"] == 0


# ──── Excel I/O ─────────────────────────────────────────────────

async def test_excel_export_import(client: AsyncClient):
    m = (await client.post("/api/models", json={"name": "EX"})).json()
    # Use ASCII name to avoid latin-1 encoding issue in Content-Disposition header
    a = (await client.post("/api/analytics", json={"model_id": m["id"], "name": "Goods", "code": "goods"})).json()
    aid = a["id"]

    # Add a field
    await client.post(f"/api/analytics/{aid}/fields", json={"name": "Name", "code": "name", "data_type": "string"})

    # Add records
    await client.post(f"/api/analytics/{aid}/records", json={"data_json": {"name": "Apples"}})
    parent = (await client.post(f"/api/analytics/{aid}/records", json={"data_json": {"name": "Vegetables"}})).json()
    await client.post(f"/api/analytics/{aid}/records", json={"parent_id": parent["id"], "data_json": {"name": "Carrots"}})

    # Export
    r = await client.get(f"/api/excel/analytics/{aid}/export")
    assert r.status_code == 200
    assert "spreadsheetml" in r.headers["content-type"]
    xlsx_bytes = r.content
    assert len(xlsx_bytes) > 100

    # Import back (clears and re-creates)
    import io
    r = await client.post(
        f"/api/excel/analytics/{aid}/import",
        files={"file": ("test.xlsx", io.BytesIO(xlsx_bytes), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert r.status_code == 200
    recs = r.json()
    assert len(recs) == 3
    # Verify hierarchy preserved
    children = [rec for rec in recs if rec["parent_id"] is not None]
    assert len(children) == 1  # Carrots under Vegetables


# ──── Health ────────────────────────────────────────────────────

async def test_health(client: AsyncClient):
    r = await client.get("/api/health")
    assert r.status_code == 200
    assert r.json()["status"] == "ok"
