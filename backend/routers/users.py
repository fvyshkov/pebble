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
    rows = await db.execute_fetchall("SELECT id, username, created_at, can_admin FROM users ORDER BY username")
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


@router.put("/{user_id}")
async def update_user(user_id: str, body: UserIn):
    db = get_db()
    await db.execute("UPDATE users SET username = ? WHERE id = ?", (body.username, user_id))
    await db.commit()
    return {"ok": True}


@router.delete("/{user_id}")
async def delete_user(user_id: str):
    db = get_db()
    await db.execute("DELETE FROM users WHERE id = ?", (user_id,))
    await db.commit()
    return {"ok": True}


class AdminIn(BaseModel):
    can_admin: bool


@router.put("/{user_id}/admin")
async def set_admin(user_id: str, body: AdminIn):
    db = get_db()
    await db.execute("UPDATE users SET can_admin = ? WHERE id = ?", (int(body.can_admin), user_id))
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


@router.get("/{user_id}/all-permissions")
async def get_all_permissions(user_id: str):
    """Return ALL models > sheets with can_view/can_edit for this user."""
    db = get_db()
    rows = await db.execute_fetchall("""
        SELECT m.id as model_id, m.name as model_name,
               s.id as sheet_id, s.name as sheet_name,
               COALESCE(sp.can_view, 1) as can_view,
               COALESCE(sp.can_edit, 1) as can_edit
        FROM sheets s
        JOIN models m ON m.id = s.model_id
        LEFT JOIN sheet_permissions sp ON sp.sheet_id = s.id AND sp.user_id = ?
        ORDER BY m.name, s.name
    """, (user_id,))
    models: dict = {}
    for r in rows:
        mid = r["model_id"]
        if mid not in models:
            models[mid] = {"id": mid, "name": r["model_name"], "sheets": []}
        models[mid]["sheets"].append({
            "id": r["sheet_id"], "name": r["sheet_name"],
            "can_view": bool(r["can_view"]), "can_edit": bool(r["can_edit"]),
        })
    return list(models.values())


@router.get("/{user_id}/allowed-records/{sheet_id}")
async def get_allowed_records(user_id: str, sheet_id: str):
    """Return analytic_id → list of allowed record_ids for this user on this sheet.
    Only includes analytics that have explicit permissions set.
    If an analytic has no permissions → user sees all (no restriction).
    """
    db = get_db()
    # Get analytics bound to this sheet
    bindings = await db.execute_fetchall(
        "SELECT sa.analytic_id FROM sheet_analytics sa WHERE sa.sheet_id = ? ORDER BY sa.sort_order",
        (sheet_id,),
    )
    result = {}
    for b in bindings:
        aid = b["analytic_id"]
        perms = await db.execute_fetchall(
            "SELECT record_id, can_view FROM analytic_record_permissions WHERE user_id = ? AND analytic_id = ? AND can_view = 1",
            (user_id, aid),
        )
        if perms:
            # User has explicit permissions → restrict to these records
            result[aid] = [p["record_id"] for p in perms]
        # If no permissions set → no restriction (user sees all)
    return result


class PermissionIn(BaseModel):
    user_id: str
    can_view: bool = True
    can_edit: bool = True


# ── Analytic record permissions ──

@router.get("/{user_id}/analytic-permissions")
async def get_analytic_permissions(user_id: str):
    """Return all analytic record permissions for a user, grouped by model > analytic."""
    db = get_db()
    rows = await db.execute_fetchall("""
        SELECT m.id as model_id, m.name as model_name,
               a.id as analytic_id, a.name as analytic_name,
               ar.id as record_id, ar.parent_id,
               json_extract(ar.data_json, '$.name') as record_name,
               COALESCE(arp.can_view, 0) as can_view,
               COALESCE(arp.can_edit, 0) as can_edit
        FROM analytic_records ar
        JOIN analytics a ON a.id = ar.analytic_id
        JOIN models m ON m.id = a.model_id
        LEFT JOIN analytic_record_permissions arp
            ON arp.record_id = ar.id AND arp.user_id = ?
        WHERE ar.parent_id IS NULL
        ORDER BY m.name, a.sort_order, ar.sort_order
    """, (user_id,))
    models: dict = {}
    for r in rows:
        mid = r["model_id"]
        if mid not in models:
            models[mid] = {"id": mid, "name": r["model_name"], "analytics": {}}
        aid = r["analytic_id"]
        if aid not in models[mid]["analytics"]:
            models[mid]["analytics"][aid] = {"id": aid, "name": r["analytic_name"], "records": []}
        models[mid]["analytics"][aid]["records"].append({
            "id": r["record_id"], "name": r["record_name"],
            "can_view": bool(r["can_view"]), "can_edit": bool(r["can_edit"]),
        })
    result = []
    for m in models.values():
        m["analytics"] = list(m["analytics"].values())
        result.append(m)
    return result


class AnalyticPermissionIn(BaseModel):
    user_id: str
    analytic_id: str
    record_id: str
    can_view: bool = True
    can_edit: bool = False


@router.put("/analytic-permissions")
async def set_analytic_permission(body: AnalyticPermissionIn):
    db = get_db()
    existing = await db.execute_fetchall(
        "SELECT id FROM analytic_record_permissions WHERE user_id = ? AND record_id = ?",
        (body.user_id, body.record_id),
    )
    if existing:
        await db.execute(
            "UPDATE analytic_record_permissions SET can_view=?, can_edit=? WHERE user_id=? AND record_id=?",
            (int(body.can_view), int(body.can_edit), body.user_id, body.record_id),
        )
    else:
        pid = str(uuid.uuid4())
        await db.execute(
            "INSERT INTO analytic_record_permissions (id, user_id, analytic_id, record_id, can_view, can_edit) VALUES (?,?,?,?,?,?)",
            (pid, body.user_id, body.analytic_id, body.record_id, int(body.can_view), int(body.can_edit)),
        )
    await db.commit()
    return {"ok": True}


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
