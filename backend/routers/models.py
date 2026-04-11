import uuid
from fastapi import APIRouter
from pydantic import BaseModel
from backend.db import get_db

router = APIRouter(prefix="/api/models", tags=["models"])


class ModelIn(BaseModel):
    name: str = ""
    description: str = ""


@router.get("")
async def list_models():
    db = get_db()
    rows = await db.execute_fetchall("SELECT * FROM models ORDER BY created_at")
    return [dict(r) for r in rows]


@router.post("")
async def create_model(body: ModelIn):
    db = get_db()
    mid = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO models (id, name, description) VALUES (?, ?, ?)",
        (mid, body.name, body.description),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM models WHERE id = ?", (mid,))
    return dict(row[0])


@router.put("/{model_id}")
async def update_model(model_id: str, body: ModelIn):
    db = get_db()
    await db.execute(
        "UPDATE models SET name=?, description=?, updated_at=datetime('now') WHERE id=?",
        (body.name, body.description, model_id),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM models WHERE id = ?", (model_id,))
    return dict(row[0])


@router.delete("/{model_id}")
async def delete_model(model_id: str):
    db = get_db()
    await db.execute("DELETE FROM models WHERE id = ?", (model_id,))
    await db.commit()
    return {"ok": True}


@router.get("/{model_id}/tree")
async def get_model_tree(model_id: str):
    db = get_db()
    model_rows = await db.execute_fetchall("SELECT * FROM models WHERE id = ?", (model_id,))
    if not model_rows:
        return {"error": "not found"}
    model = dict(model_rows[0])
    sheets = await db.execute_fetchall(
        "SELECT * FROM sheets WHERE model_id = ? ORDER BY sort_order, created_at", (model_id,)
    )
    analytics = await db.execute_fetchall(
        "SELECT * FROM analytics WHERE model_id = ? ORDER BY sort_order", (model_id,)
    )
    return {
        **model,
        "sheets": [dict(s) for s in sheets],
        "analytics": [dict(a) for a in analytics],
    }
