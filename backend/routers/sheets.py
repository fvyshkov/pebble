import uuid
from fastapi import APIRouter
from pydantic import BaseModel
from backend.db import get_db

router = APIRouter(prefix="/api/sheets", tags=["sheets"])


class SheetIn(BaseModel):
    model_id: str | None = None
    name: str = ""
    excel_code: str = ""


class SheetAnalyticIn(BaseModel):
    analytic_id: str
    sort_order: int = 0
    is_fixed: bool = False
    fixed_record_id: str | None = None


class ReorderIn(BaseModel):
    ordered_ids: list[str]


# --- Sheets CRUD ---

@router.get("/by-model/{model_id}")
async def list_sheets(model_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT * FROM sheets WHERE model_id = ? ORDER BY sort_order, created_at", (model_id,)
    )
    return [dict(r) for r in rows]


@router.put("/reorder/{model_id}")
async def reorder_sheets(model_id: str, body: ReorderIn):
    db = get_db()
    for i, sheet_id in enumerate(body.ordered_ids):
        await db.execute(
            "UPDATE sheets SET sort_order=? WHERE id=? AND model_id=?",
            (i, sheet_id, model_id),
        )
    await db.commit()
    return {"ok": True}


@router.post("")
async def create_sheet(body: SheetIn):
    db = get_db()
    sid = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO sheets (id, model_id, name, excel_code) VALUES (?, ?, ?, ?)",
        (sid, body.model_id, body.name, body.excel_code),
    )
    # Auto-grant permissions to all existing users
    users = await db.execute_fetchall("SELECT id FROM users")
    for u in users:
        spid = str(uuid.uuid4())
        await db.execute(
            "INSERT OR IGNORE INTO sheet_permissions (id, sheet_id, user_id, can_view, can_edit) VALUES (?,?,?,1,1)",
            (spid, sid, u["id"]),
        )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM sheets WHERE id = ?", (sid,))
    return dict(row[0])


@router.put("/{sheet_id}")
async def update_sheet(sheet_id: str, body: SheetIn):
    db = get_db()
    fields = ["name=?", "updated_at=datetime('now')"]
    params: list = [body.name]
    if body.excel_code:
        fields.append("excel_code=?")
        params.append(body.excel_code)
    params.append(sheet_id)
    await db.execute(f"UPDATE sheets SET {', '.join(fields)} WHERE id=?", params)
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM sheets WHERE id = ?", (sheet_id,))
    return dict(row[0])


@router.delete("/{sheet_id}")
async def delete_sheet(sheet_id: str):
    db = get_db()
    await db.execute("DELETE FROM sheets WHERE id = ?", (sheet_id,))
    await db.commit()
    return {"ok": True}


# --- Sheet-Analytic bindings ---

@router.get("/{sheet_id}/analytics")
async def list_sheet_analytics(sheet_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        """SELECT sa.*, a.name as analytic_name, a.icon as analytic_icon
           FROM sheet_analytics sa
           JOIN analytics a ON a.id = sa.analytic_id
           WHERE sa.sheet_id = ? ORDER BY sa.sort_order""",
        (sheet_id,),
    )
    return [dict(r) for r in rows]


@router.post("/{sheet_id}/analytics")
async def add_sheet_analytic(sheet_id: str, body: SheetAnalyticIn):
    db = get_db()
    said = str(uuid.uuid4())
    await db.execute(
        "INSERT INTO sheet_analytics (id, sheet_id, analytic_id, sort_order, is_fixed, fixed_record_id) VALUES (?,?,?,?,?,?)",
        (said, sheet_id, body.analytic_id, body.sort_order, int(body.is_fixed), body.fixed_record_id),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM sheet_analytics WHERE id = ?", (said,))
    return dict(row[0])


@router.put("/{sheet_id}/analytics/{sa_id}")
async def update_sheet_analytic(sheet_id: str, sa_id: str, body: SheetAnalyticIn):
    db = get_db()
    await db.execute(
        "UPDATE sheet_analytics SET sort_order=?, is_fixed=?, fixed_record_id=? WHERE id=? AND sheet_id=?",
        (body.sort_order, int(body.is_fixed), body.fixed_record_id, sa_id, sheet_id),
    )
    await db.commit()
    row = await db.execute_fetchall("SELECT * FROM sheet_analytics WHERE id = ?", (sa_id,))
    return dict(row[0])


@router.delete("/{sheet_id}/analytics/{sa_id}")
async def remove_sheet_analytic(sheet_id: str, sa_id: str):
    db = get_db()
    await db.execute("DELETE FROM sheet_analytics WHERE id = ? AND sheet_id = ?", (sa_id, sheet_id))
    await db.commit()
    return {"ok": True}


@router.put("/{sheet_id}/analytics-reorder")
async def reorder_sheet_analytics(sheet_id: str, body: ReorderIn):
    db = get_db()
    for i, sa_id in enumerate(body.ordered_ids):
        await db.execute(
            "UPDATE sheet_analytics SET sort_order=? WHERE id=? AND sheet_id=?",
            (i, sa_id, sheet_id),
        )
    await db.commit()
    return {"ok": True}


# ── View settings ──

class ViewSettingsIn(BaseModel):
    settings: dict


@router.get("/{sheet_id}/view-settings")
async def get_view_settings(sheet_id: str):
    db = get_db()
    rows = await db.execute_fetchall(
        "SELECT settings FROM sheet_view_settings WHERE sheet_id = ?", (sheet_id,)
    )
    if rows:
        import json
        return json.loads(rows[0]["settings"])
    return {}


@router.put("/{sheet_id}/view-settings")
async def save_view_settings(sheet_id: str, body: ViewSettingsIn):
    db = get_db()
    import json
    settings_json = json.dumps(body.settings, ensure_ascii=False)
    existing = await db.execute_fetchall(
        "SELECT sheet_id FROM sheet_view_settings WHERE sheet_id = ?", (sheet_id,)
    )
    if existing:
        await db.execute(
            "UPDATE sheet_view_settings SET settings = ? WHERE sheet_id = ?",
            (settings_json, sheet_id),
        )
    else:
        await db.execute(
            "INSERT INTO sheet_view_settings (sheet_id, settings) VALUES (?, ?)",
            (sheet_id, settings_json),
        )
    await db.commit()
    return {"ok": True}
