import uuid
from fastapi import APIRouter
from pydantic import BaseModel
from backend.db import get_db

router = APIRouter(prefix="/api/users", tags=["users"])


class UserIn(BaseModel):
    username: str


@router.get("")
async def list_users():
    db = get_db()
    rows = await db.execute_fetchall("SELECT id, username, created_at FROM users ORDER BY username")
    return [dict(r) for r in rows]


@router.post("")
async def create_user(body: UserIn):
    db = get_db()
    uid = str(uuid.uuid4())
    # Default password = username
    await db.execute(
        "INSERT INTO users (id, username, password) VALUES (?, ?, ?)",
        (uid, body.username, body.username),
    )
    # Grant access to all existing sheets
    sheets = await db.execute_fetchall("SELECT id FROM sheets")
    for s in sheets:
        spid = str(uuid.uuid4())
        await db.execute(
            "INSERT OR IGNORE INTO sheet_permissions (id, sheet_id, user_id, can_view, can_edit) VALUES (?,?,?,1,1)",
            (spid, s["id"], uid),
        )
    await db.commit()
    return {"id": uid, "username": body.username}


@router.delete("/{user_id}")
async def delete_user(user_id: str):
    db = get_db()
    await db.execute("DELETE FROM users WHERE id = ?", (user_id,))
    await db.commit()
    return {"ok": True}


class PasswordIn(BaseModel):
    password: str


@router.post("/{user_id}/reset-password")
async def reset_password(user_id: str, body: PasswordIn):
    db = get_db()
    rows = await db.execute_fetchall("SELECT id FROM users WHERE id = ?", (user_id,))
    if not rows:
        return {"error": "not found"}
    await db.execute("UPDATE users SET password = ? WHERE id = ?", (body.password, user_id))
    await db.commit()
    return {"ok": True}


# ── Sheet permissions ──

@router.get("/permissions/by-sheet/{sheet_id}")
async def get_sheet_permissions(sheet_id: str):
    db = get_db()
    # Return all users with their permissions for this sheet
    users = await db.execute_fetchall("SELECT id, username FROM users ORDER BY username")
    perms = await db.execute_fetchall(
        "SELECT user_id, can_view, can_edit FROM sheet_permissions WHERE sheet_id = ?", (sheet_id,)
    )
    perm_map = {p["user_id"]: dict(p) for p in perms}
    result = []
    for u in users:
        p = perm_map.get(u["id"], {"can_view": 1, "can_edit": 1})
        result.append({"user_id": u["id"], "username": u["username"], "can_view": p["can_view"], "can_edit": p["can_edit"]})
    return result


@router.get("/{user_id}/accessible-sheets")
async def get_accessible_sheets(user_id: str):
    """Return models and sheets the user can view, with edit flag."""
    db = get_db()
    rows = await db.execute_fetchall("""
        SELECT m.id as model_id, m.name as model_name,
               s.id as sheet_id, s.name as sheet_name,
               COALESCE(sp.can_view, 1) as can_view,
               COALESCE(sp.can_edit, 1) as can_edit
        FROM sheets s
        JOIN models m ON m.id = s.model_id
        LEFT JOIN sheet_permissions sp ON sp.sheet_id = s.id AND sp.user_id = ?
        WHERE COALESCE(sp.can_view, 1) = 1
        ORDER BY m.name, s.name
    """, (user_id,))
    # Group by model
    models: dict = {}
    for r in rows:
        mid = r["model_id"]
        if mid not in models:
            models[mid] = {"id": mid, "name": r["model_name"], "sheets": []}
        models[mid]["sheets"].append({
            "id": r["sheet_id"], "name": r["sheet_name"],
            "can_edit": bool(r["can_edit"]),
        })
    return list(models.values())


class PermissionIn(BaseModel):
    user_id: str
    can_view: bool = True
    can_edit: bool = True


@router.put("/permissions/by-sheet/{sheet_id}")
async def set_sheet_permission(sheet_id: str, body: PermissionIn):
    db = get_db()
    existing = await db.execute_fetchall(
        "SELECT id FROM sheet_permissions WHERE sheet_id = ? AND user_id = ?",
        (sheet_id, body.user_id),
    )
    if existing:
        await db.execute(
            "UPDATE sheet_permissions SET can_view=?, can_edit=? WHERE sheet_id=? AND user_id=?",
            (int(body.can_view), int(body.can_edit), sheet_id, body.user_id),
        )
    else:
        spid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO sheet_permissions (id, sheet_id, user_id, can_view, can_edit) VALUES (?,?,?,?,?)",
            (spid, sheet_id, body.user_id, int(body.can_view), int(body.can_edit)),
        )
    await db.commit()
    return {"ok": True}
